#!/usr/bin/env python3
"""
Chakra HQ — Selenium Audit Bot
Full end-to-end flow:
  Looker Studio (Chrome Profile 6) → Lead ID click → Google redirect →
  ChakraHQ scrape → GPT-4o audit → Audit Form - LQ (Google Form) fill
"""

import json
import logging
import os
import re
import subprocess
import sys
import time
from pathlib import Path
from typing import Optional
from urllib.parse import unquote

from dotenv import load_dotenv
from openai import OpenAI
from selenium import webdriver
from selenium.common.exceptions import (
    NoSuchElementException,
    StaleElementReferenceException,
    TimeoutException,
    WebDriverException,
)
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait

# ── Paths ─────────────────────────────────────────────────────────────────────
SCRIPT_DIR = Path(__file__).parent

# ── Config ────────────────────────────────────────────────────────────────────
LOOKER_URL = (
    "https://datastudio.google.com/u/0/reporting/"
    "0b0ea9bb-0051-4901-bafa-42923b72b30f/page/p_goiiob3ltd"
)
DEFAULT_CHROME_USER_DATA = SCRIPT_DIR / ".chrome-user-data"
DEFAULT_CHROME_PROFILE = "Default"
LEGACY_CHROME_USER_DATA = Path(r"C:\Users\Ravi\AppData\Local\Google\Chrome\User Data")
LEGACY_CHROME_PROFILE = "Profile 6"

LOOKER_LOAD_WAIT  = 300   # seconds — 5 minutes for Looker Studio
PAGE_LOAD_WAIT    = 60    # seconds — ChakraHQ page
IMPLICIT_WAIT     = 10    # seconds — implicit element wait

# ── Logging ───────────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.StreamHandler(sys.stdout),
        logging.FileHandler(SCRIPT_DIR / "audit_bot.log", encoding="utf-8"),
    ],
)
log = logging.getLogger("audit_bot")


# ── Custom exceptions ─────────────────────────────────────────────────────────
class AuditBotError(Exception):
    """Raised when the bot cannot continue."""


def _env_flag(name: str) -> bool:
    return os.getenv(name, "").strip().lower() in {"1", "true", "yes", "on"}


def _chrome_profile_config() -> tuple[Path, str, bool]:
    """Return user-data dir, profile name, and whether it is a shared Chrome profile."""
    use_existing_profile = _env_flag("AUDIT_BOT_USE_EXISTING_CHROME_PROFILE")
    default_user_data = LEGACY_CHROME_USER_DATA if use_existing_profile else DEFAULT_CHROME_USER_DATA
    default_profile = LEGACY_CHROME_PROFILE if use_existing_profile else DEFAULT_CHROME_PROFILE

    user_data = Path(os.getenv("AUDIT_BOT_CHROME_USER_DATA", str(default_user_data))).expanduser()
    profile = os.getenv("AUDIT_BOT_CHROME_PROFILE", default_profile).strip() or default_profile
    return user_data, profile, use_existing_profile


# ── Bot ───────────────────────────────────────────────────────────────────────
class AuditBot:
    def __init__(self, openai_key: str, system_prompt: str) -> None:
        self.openai         = OpenAI(api_key=openai_key, timeout=90.0)
        self.system_prompt  = system_prompt
        self.driver: Optional[webdriver.Chrome] = None
        self.looker_handle: Optional[str]        = None
        self.lead_row_idx: int                   = 0  # 0-based index of processed row

    # ── Driver ────────────────────────────────────────────────────────────────
    @staticmethod
    def _kill_chrome(kill_all_chrome: bool, user_data_dir: Path) -> None:
        """Terminate automation leftovers, and optionally all Chrome for shared profiles."""
        for exe in ["chromedriver.exe"]:
            subprocess.run(
                ["taskkill", "/F", "/IM", exe],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                timeout=10,
            )
        if kill_all_chrome:
            subprocess.run(
                ["taskkill", "/F", "/IM", "chrome.exe"],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                timeout=10,
            )
        else:
            escaped_profile = str(user_data_dir).replace("'", "''")
            subprocess.run(
                [
                    "powershell",
                    "-NoProfile",
                    "-Command",
                    (
                        f"$needle = '{escaped_profile}'; "
                        "Get-CimInstance Win32_Process -Filter \"name = 'chrome.exe'\" | "
                        "Where-Object { $_.CommandLine -and $_.CommandLine -like \"*--user-data-dir=$needle*\" } | "
                        "ForEach-Object { Stop-Process -Id $_.ProcessId -Force }"
                    ),
                ],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                timeout=10,
            )
        time.sleep(2)
        if kill_all_chrome:
            log.info("Killed existing Chrome/ChromeDriver processes.")
        else:
            log.info("Killed existing ChromeDriver processes.")

    @staticmethod
    def _clear_profile_locks(user_data_dir: Path, profile: str) -> None:
        """Remove all lock files that block Chrome startup."""
        lock_names = ("SingletonLock", "SingletonCookie", "SingletonSocket")
        removed: list[str] = []
        for name in lock_names:
            for candidate in (
                user_data_dir / name,
                user_data_dir / profile / name,
            ):
                try:
                    if candidate.exists():
                        candidate.unlink()
                        removed.append(str(candidate))
                except OSError:
                    pass
        # Remove LevelDB LOCK file from the profile root — left behind by force-kill
        leveldb_lock = user_data_dir / profile / "LOCK"
        try:
            if leveldb_lock.exists():
                leveldb_lock.unlink()
                removed.append(str(leveldb_lock))
        except OSError:
            pass
        if removed:
            log.info(f"Removed lock files: {removed}")

    @staticmethod
    def _fix_profile_exit_type(user_data_dir: Path, profile: str) -> None:
        """Set exit_type to Normal so Chrome skips crash-recovery on launch.

        When Chrome is force-killed it writes exit_type='Crashed'. On the next
        launch Chrome enters session-restore mode and hangs waiting for UI
        interaction, which makes ChromeDriver's initial connection time out.
        """
        prefs_path = user_data_dir / profile / "Preferences"
        if not prefs_path.exists():
            return
        try:
            prefs = json.loads(prefs_path.read_text(encoding="utf-8"))
            profile_section = prefs.setdefault("profile", {})
            if profile_section.get("exit_type") != "Normal":
                profile_section["exit_type"] = "Normal"
                profile_section["exited_cleanly"] = True
                prefs_path.write_text(
                    json.dumps(prefs, ensure_ascii=False),
                    encoding="utf-8",
                )
                log.info("Patched Preferences: exit_type → Normal")
        except Exception as exc:
            log.warning(f"Could not patch Preferences: {exc}")

    def _build_driver(self) -> webdriver.Chrome:
        user_data_dir, profile, is_shared_profile = _chrome_profile_config()
        user_data_dir.mkdir(parents=True, exist_ok=True)

        self._kill_chrome(kill_all_chrome=is_shared_profile, user_data_dir=user_data_dir)
        self._clear_profile_locks(user_data_dir, profile)
        self._fix_profile_exit_type(user_data_dir, profile)

        opts = Options()
        opts.add_argument(f"--user-data-dir={user_data_dir}")
        opts.add_argument(f"--profile-directory={profile}")
        opts.add_argument("--start-maximized")
        opts.add_argument("--disable-blink-features=AutomationControlled")
        opts.add_argument("--no-sandbox")
        opts.add_argument("--disable-dev-shm-usage")
        opts.add_argument("--disable-gpu")
        opts.add_argument("--disable-extensions")
        opts.add_argument("--no-first-run")
        opts.add_argument("--no-default-browser-check")
        opts.add_argument("--disable-background-networking")
        opts.add_argument("--safebrowsing-disable-auto-update")
        opts.add_experimental_option("excludeSwitches", ["enable-automation"])
        opts.add_experimental_option("useAutomationExtension", False)

        service_env = os.environ.copy()
        service_env.pop("CHROME_CRASHPAD_PIPE_NAME", None)
        service = Service(ChromeDriverManager().install(), env=service_env)
        try:
            driver = webdriver.Chrome(service=service, options=opts)
        except WebDriverException as exc:
            raise AuditBotError(
                "Chrome could not start. The bot uses a dedicated automation "
                f"profile at {user_data_dir}. If you force the real Chrome profile "
                "with AUDIT_BOT_USE_EXISTING_CHROME_PROFILE=1, close Chrome first "
                "or remove that override."
            ) from exc
        driver.implicitly_wait(IMPLICIT_WAIT)
        log.info(f"Chrome launched with profile '{profile}' at {user_data_dir}.")
        return driver

    def _wait(self, timeout: int = PAGE_LOAD_WAIT) -> WebDriverWait:
        return WebDriverWait(self.driver, timeout)

    def _screenshot(self, name: str) -> None:
        """Save a debug screenshot."""
        path = SCRIPT_DIR / f"debug_{name}.png"
        try:
            self.driver.save_screenshot(str(path))
            log.info(f"Screenshot saved: {path.name}")
        except Exception:
            pass

    # ── Frame traversal ───────────────────────────────────────────────────────
    def _reset_frame(self) -> None:
        self.driver.switch_to.default_content()

    def _find_lead_links_in_current_frame(self) -> list:
        """Return <a> elements whose text starts with CHQ- in current frame."""
        for xpath in [
            "//a[contains(text(),'CHQ-')]",
            "//a[starts-with(normalize-space(text()),'CHQ')]",
            "//*[contains(@class,'cell')]//a[contains(text(),'CHQ')]",
        ]:
            els = self.driver.find_elements(By.XPATH, xpath)
            if els:
                return els
        return []

    def _dfs_enter_frame_with_leads(self, depth: int = 0) -> bool:
        """
        Depth-first search all iframes looking for CHQ- lead links.
        Leaves the driver focused inside the matching iframe if found.
        Returns True on success.
        """
        if depth > 7:
            return False

        iframes = self.driver.find_elements(By.TAG_NAME, "iframe")
        for idx in range(len(iframes)):
            try:
                # Re-query because switching frames refreshes stale refs
                iframes = self.driver.find_elements(By.TAG_NAME, "iframe")
                if idx >= len(iframes):
                    break
                self.driver.switch_to.frame(iframes[idx])

                if self._find_lead_links_in_current_frame():
                    log.info(f"Lead links found in iframe (depth={depth}, idx={idx}).")
                    return True

                if self._dfs_enter_frame_with_leads(depth + 1):
                    return True

                self.driver.switch_to.parent_frame()
            except (StaleElementReferenceException, Exception):
                try:
                    self.driver.switch_to.parent_frame()
                except Exception:
                    self._reset_frame()

        return False

    def _locate_lead_links(self) -> list:
        """Find CHQ- links anywhere on page (top context + all iframes)."""
        # Check top-level first
        links = self._find_lead_links_in_current_frame()
        if links:
            return links

        # DFS into iframes
        self._reset_frame()
        if self._dfs_enter_frame_with_leads():
            return self._find_lead_links_in_current_frame()

        return []

    # ── Step 1: Open Looker Studio ────────────────────────────────────────────
    def open_looker_studio(self) -> None:
        log.info("Navigating to Looker Studio …")
        self.driver.get(LOOKER_URL)
        self.looker_handle = self.driver.current_window_handle
        log.info(
            f"Waiting {LOOKER_LOAD_WAIT // 60} min "
            f"({LOOKER_LOAD_WAIT}s) for the report to fully load …"
        )
        time.sleep(LOOKER_LOAD_WAIT)
        log.info("Load wait complete.")
        self._screenshot("01_looker_loaded")

    # ── Step 2: Scroll to the table section ──────────────────────────────────
    def scroll_to_table(self) -> None:
        log.info("Scrolling down to find the ticket details table …")
        self._reset_frame()

        for scroll_num in range(20):
            links = self._locate_lead_links()
            if links:
                log.info(f"Table visible after {scroll_num} scroll steps.")
                try:
                    self.driver.execute_script(
                        "arguments[0].scrollIntoView({block:'center', behavior:'smooth'});",
                        links[0],
                    )
                except Exception:
                    pass
                time.sleep(1)
                self._screenshot("02_table_visible")
                return

            # Scroll the outer document
            self._reset_frame()
            self.driver.execute_script("window.scrollBy(0, 350);")
            time.sleep(1)

            # Also try scrolling inside iframes
            try:
                self._dfs_enter_frame_with_leads()
                self.driver.execute_script("window.scrollBy(0, 350);")
                self._reset_frame()
            except Exception:
                self._reset_frame()

        log.warning(
            "Table not confirmed visible after max scroll — proceeding anyway."
        )

    # ── Step 3: Click first Lead ID ───────────────────────────────────────────
    def click_first_lead_id(self) -> str:
        log.info("Locating and clicking the first Lead ID …")
        links = self._locate_lead_links()
        if not links:
            self._screenshot("err_no_lead_links")
            raise AuditBotError(
                "No Lead ID links (CHQ-XXXXXX) found in Looker Studio table."
            )

        lead_el   = links[0]
        lead_text = lead_el.text.strip()
        log.info(f"Lead ID selected: {lead_text}")

        # Record which row index this is so we can find the LQ form link later.
        # Looker Studio uses div-based layout, so we look for sibling wrappers.
        try:
            all_links = self._locate_lead_links()
            for i, el in enumerate(all_links):
                if el == lead_el:
                    self.lead_row_idx = i
                    break
        except Exception:
            self.lead_row_idx = 0

        old_handles = set(self.driver.window_handles)
        lead_el.click()
        log.info("Lead ID clicked — waiting for navigation …")
        time.sleep(4)

        new_handles = set(self.driver.window_handles) - old_handles
        if new_handles:
            self.driver.switch_to.window(new_handles.pop())
            log.info(f"New tab: {self.driver.current_url[:80]}")
        else:
            log.info(f"Same tab: {self.driver.current_url[:80]}")

        return lead_text

    # ── Step 4: Resolve Google redirect ──────────────────────────────────────
    def handle_google_redirect(self) -> None:
        time.sleep(2)
        url   = self.driver.current_url
        title = self.driver.title

        is_redirect = (
            "google.com/url" in url
            or "Redirect Notice" in title
            or "redirect" in title.lower()
        )
        if not is_redirect:
            log.info("No redirect page — already on destination.")
            return

        log.info(f"Google redirect detected (title='{title}') — resolving …")
        self._screenshot("03_redirect_page")

        # Strategy A: click the chakrahq link in the redirect body
        for xpath in [
            "//a[contains(@href,'chakrahq.com')]",
            "//a[contains(@href,'app.chakrahq')]",
            "//div[@id='main']//a",
        ]:
            try:
                el = self._wait(15).until(
                    EC.presence_of_element_located((By.XPATH, xpath))
                )
                href = el.get_attribute("href") or ""
                log.info(f"Clicking redirect link: {href[:80]}")
                el.click()
                time.sleep(3)
                return
            except TimeoutException:
                continue

        # Strategy B: extract the ?q= parameter and navigate directly
        m = re.search(r"[?&]q=([^&]+)", url)
        if m:
            target = unquote(m.group(1))
            log.info(f"Direct navigation to: {target[:80]}")
            self.driver.get(target)
            time.sleep(3)
            return

        raise AuditBotError(f"Cannot resolve Google redirect from URL: {url[:120]}")

    # ── Step 5: Wait for ChakraHQ page ───────────────────────────────────────
    def wait_for_chakrahq(self) -> None:
        log.info("Waiting for ChakraHQ lead page to load …")
        try:
            self._wait(PAGE_LOAD_WAIT).until(
                EC.any_of(
                    EC.url_contains("chakrahq.com"),
                    EC.presence_of_element_located(
                        (By.XPATH, "//*[contains(text(),'Lead ID') or contains(text(),'Lead Id')]")
                    ),
                    EC.presence_of_element_located(
                        (By.XPATH, "//*[contains(text(),'Lead Source')]")
                    ),
                    EC.presence_of_element_located(
                        (By.XPATH, "//*[contains(text(),'CHQ-')]")
                    ),
                )
            )
        except TimeoutException:
            log.warning("Timeout waiting for ChakraHQ — scraping whatever is loaded.")

        # Extra dwell time for dynamic content (React/Vue hydration)
        time.sleep(5)
        self._screenshot("04_chakrahq_page")

    # ── Step 6: Scrape ChakraHQ lead ─────────────────────────────────────────
    def _text(self, xpath: str) -> str:
        """Return element text or empty string."""
        try:
            return self.driver.find_element(By.XPATH, xpath).text.strip()
        except NoSuchElementException:
            return ""

    def _label_value(self, label: str) -> str:
        """
        Extract value adjacent to a label.
        Tries multiple DOM patterns (key:value divs, table rows, dl/dt/dd).
        """
        xpaths = [
            f"//*[normalize-space(text())='{label}']/following-sibling::*[1]",
            f"//*[normalize-space(text())='{label}']/../following-sibling::*[1]",
            f"//*[normalize-space(text())='{label}']/following::div[1]",
            f"//*[normalize-space(text())='{label}']/following::span[1]",
            f"//td[normalize-space(text())='{label}']/following-sibling::td[1]",
            f"//dt[normalize-space(text())='{label}']/following-sibling::dd[1]",
        ]
        for xp in xpaths:
            val = self._text(xp)
            if val and val.lower() != label.lower():
                return val
        return ""

    def scrape_chakrahq(self) -> dict:
        log.info("Scraping lead data …")
        data: dict = {"page_url": self.driver.current_url}

        # Status (look for the active button in the lead header)
        for status in ("Completed", "Closed", "Reopen", "Open"):
            try:
                self.driver.find_element(
                    By.XPATH,
                    f"//button[contains(normalize-space(text()),'{status}')] | "
                    f"//span[contains(@class,'status') and contains(text(),'{status}')]",
                )
                data["lead_status"] = status
                break
            except NoSuchElementException:
                pass

        # Assigned / Qualified By
        data["assigned_to"] = self._text(
            "//*[contains(@class,'agent') or contains(@class,'assigned')]"
            "[not(contains(text(),'Assigned'))]"
        ) or self._label_value("Qualified By")

        # Summary section
        for label in (
            "Phone", "Skills", "Created At", "Updated At",
            "Lead Source", "Lead ID", "Contact", "Qualified By",
        ):
            data[label] = self._label_value(label)

        # Lead Details
        for label in (
            "Name", "Gemstone", "Gemstone Weight (Carat)", "Purpose",
            "Budget", "Customer Type", "Category", "Super Category",
            "Order ID", "Final Order Value", "Lead Type",
            "First Task Completed Timestamp", "Discount %",
        ):
            data[label] = self._label_value(label)

        # Qualification
        for label in ("Qualification Status", "Primary Disposition", "Secondary Disposition"):
            data[label] = self._label_value(label)

        # Skill / tag chips
        try:
            chips = self.driver.find_elements(
                By.XPATH,
                "//*[contains(@class,'chip') or contains(@class,'tag') "
                "or contains(@class,'skill') or contains(@class,'badge')]",
            )
            data["skill_tags"] = list({c.text.strip() for c in chips if c.text.strip()})
        except Exception:
            data["skill_tags"] = []

        # History — scroll into view, then collect entries
        try:
            self.driver.execute_script(
                "var el = document.querySelector('[class*=history],[id*=history]');"
                "if(el) el.scrollIntoView({behavior:'smooth', block:'center'});"
            )
            time.sleep(2)
        except Exception:
            pass

        history: list = []
        try:
            history_items = self.driver.find_elements(
                By.XPATH,
                "//*[contains(text(),'History')]/..//div[@class] | "
                "//*[contains(text(),'History')]/following::div[position()<80]",
            )
            seen: set = set()
            for el in history_items[:60]:
                t = el.text.strip()
                if t and len(t) > 5 and t not in seen:
                    seen.add(t)
                    history.append(t)
                    if len(history) >= 30:
                        break
        except Exception as e:
            log.warning(f"History partial: {e}")

        data["history_entries"] = history

        # Full page body text as fallback context (capped at 12 000 chars)
        try:
            data["full_page_text"] = (
                self.driver.find_element(By.TAG_NAME, "body").text[:12_000]
            )
        except Exception:
            data["full_page_text"] = ""

        log.info(f"Scraped {len(data)} top-level keys.")
        return data

    # ── Step 7: GPT audit review ──────────────────────────────────────────────
    def get_gpt_review(self, lead_data: dict) -> dict:
        log.info("Sending lead data to GPT-4o …")
        user_msg = (
            "Below is the complete data scraped from a ChakraHQ lead page.\n"
            "Audit the LQ agent's performance and return your assessment "
            "as a single JSON object with the keys specified in the system prompt.\n\n"
            f"```json\n{json.dumps(lead_data, indent=2, ensure_ascii=False)}\n```"
        )

        resp = self.openai.chat.completions.create(
            model="gpt-4o",
            messages=[
                {"role": "system", "content": self.system_prompt},
                {"role": "user",   "content": user_msg},
            ],
            temperature=0.1,
            max_tokens=2000,
            response_format={"type": "json_object"},
        )

        raw = resp.choices[0].message.content or "{}"
        log.info(f"GPT response: {raw[:300]} …")

        try:
            return json.loads(raw)
        except json.JSONDecodeError:
            m = re.search(r"\{.*\}", raw, re.DOTALL)
            return json.loads(m.group(0)) if m else {"raw_response": raw}

    # ── Step 8: Switch back to Looker Studio ─────────────────────────────────
    def switch_to_looker(self) -> None:
        log.info("Switching back to Looker Studio tab …")
        if self.looker_handle and self.looker_handle in self.driver.window_handles:
            self.driver.switch_to.window(self.looker_handle)
        else:
            # Fallback: find the Looker Studio tab by URL
            for h in self.driver.window_handles:
                self.driver.switch_to.window(h)
                url = self.driver.current_url
                if "datastudio" in url or "lookerstudio" in url:
                    self.looker_handle = h
                    log.info("Looker Studio tab found by URL scan.")
                    break
        time.sleep(2)
        log.info(f"Active tab: {self.driver.current_url[:80]}")

    # ── Step 9: Click Audit Form - LQ link ───────────────────────────────────
    def click_audit_form_lq(self, lead_id: str) -> None:
        log.info(f"Locating Audit Form - LQ link for {lead_id} …")
        self._reset_frame()
        self._dfs_enter_frame_with_leads()  # re-enter the iframe that has the table
        time.sleep(1)

        old_handles = set(self.driver.window_handles)

        # ── Strategy A: find a "Click" anchor in the same visual row as lead_id
        #    Looker Studio uses div-based layout; sibling divs represent cells.
        #    The LQ column is the 2nd "Click" link in the same row container.
        try:
            # Find the element containing the lead_id text
            lead_el = self.driver.find_element(
                By.XPATH, f"//a[normalize-space(text())='{lead_id}']"
            )
            # Walk up to the row container (try several ancestor levels)
            for levels_up in range(2, 9):
                ancestor_xpath = "/".join([".."] * levels_up)
                try:
                    row_el = lead_el.find_element(By.XPATH, ancestor_xpath)
                    click_links = row_el.find_elements(
                        By.XPATH, ".//a[normalize-space(text())='Click']"
                    )
                    if len(click_links) >= 2:
                        log.info(
                            f"Strategy A: {len(click_links)} 'Click' links found "
                            f"at ancestor depth {levels_up}. Clicking index 1 (LQ)."
                        )
                        click_links[1].click()
                        self._await_new_tab(old_handles)
                        return
                    elif len(click_links) == 1:
                        # Only one found at this level; keep going up
                        continue
                except Exception:
                    continue
        except NoSuchElementException:
            pass

        # ── Strategy B: gather all "Click" links; group by rows of 3 (Sales/LQ/Quick)
        #    Each data row has exactly 3 audit form links.
        #    Row N → indices 3N, 3N+1 (LQ), 3N+2
        try:
            all_clicks = self.driver.find_elements(
                By.XPATH, "//a[normalize-space(text())='Click']"
            )
            if all_clicks:
                lq_idx = self.lead_row_idx * 3 + 1  # index 1 = LQ column
                if lq_idx < len(all_clicks):
                    log.info(
                        f"Strategy B: clicking all_clicks[{lq_idx}] "
                        f"(row {self.lead_row_idx}, LQ=+1)."
                    )
                    all_clicks[lq_idx].click()
                    self._await_new_tab(old_handles)
                    return
                elif len(all_clicks) >= 2:
                    # Fallback to index 1 if grouping math is off
                    log.info("Strategy B fallback: clicking all_clicks[1].")
                    all_clicks[1].click()
                    self._await_new_tab(old_handles)
                    return
        except Exception as e:
            log.warning(f"Strategy B failed: {e}")

        # ── Strategy C: look for elements with aria-label or data-label containing LQ
        try:
            lq_el = self.driver.find_element(
                By.XPATH,
                "//a[contains(@aria-label,'LQ') or contains(@title,'LQ') "
                "or contains(@data-label,'LQ')]",
            )
            log.info("Strategy C: aria/title label match for LQ.")
            lq_el.click()
            self._await_new_tab(old_handles)
            return
        except NoSuchElementException:
            pass

        self._screenshot("err_lq_link_not_found")
        raise AuditBotError(
            "Could not locate the Audit Form - LQ 'Click' link in the Looker Studio table. "
            "Check debug_err_lq_link_not_found.png for the current page state."
        )

    def _await_new_tab(self, old_handles: set) -> None:
        """Wait up to 10 s for a new tab and switch to it."""
        deadline = time.time() + 10
        while time.time() < deadline:
            new = set(self.driver.window_handles) - old_handles
            if new:
                self.driver.switch_to.window(new.pop())
                log.info(f"Switched to new tab: {self.driver.current_url[:80]}")
                return
            time.sleep(0.5)
        log.info(f"No new tab opened — staying on: {self.driver.current_url[:80]}")

    # ── Step 10: Fill Google Form ─────────────────────────────────────────────
    def fill_google_form(self, review: dict) -> None:
        log.info("Filling Google Form …")

        try:
            self._wait(30).until(
                EC.any_of(
                    EC.presence_of_element_located((By.XPATH, "//form")),
                    EC.presence_of_element_located(
                        (By.XPATH, "//*[contains(@class,'freebirdFormviewer')]")
                    ),
                    EC.presence_of_element_located(
                        (By.XPATH, "//div[@role='listitem']")
                    ),
                )
            )
        except TimeoutException:
            log.warning("Google Form did not load cleanly — attempting anyway.")

        time.sleep(3)
        self.driver.execute_script("window.scrollTo(0,0);")
        self._screenshot("05_google_form")

        flat = _flatten(review)

        # Collect all question blocks (Google Forms uses several class patterns)
        question_blocks = self.driver.find_elements(
            By.XPATH,
            "//div[contains(@class,'freebirdFormviewerViewItemsItemItem')]"
            " | //div[@class='Qr7Oae']"
            " | //div[@role='listitem']",
        )
        log.info(f"Found {len(question_blocks)} form question blocks.")

        for block in question_blocks:
            try:
                q_text = self._block_label(block).lower()
                answer = _best_answer(q_text, flat, review)
                if not answer:
                    continue
                self._fill_block(block, q_text, answer)
                time.sleep(0.4)
            except StaleElementReferenceException:
                continue
            except Exception as e:
                log.warning(f"Could not fill block ('{q_text[:30]}'): {e}")

        # Submit
        try:
            submit = self.driver.find_element(
                By.XPATH,
                "//div[@role='button'][contains(.,'Submit')] | "
                "//span[text()='Submit']/ancestor::div[@role='button'] | "
                "//input[@type='submit']",
            )
            log.info("Submitting form …")
            submit.click()
            time.sleep(4)
            self._screenshot("06_form_submitted")
            log.info("Form submitted.")
        except NoSuchElementException:
            log.warning("Submit button not found — please submit manually.")

    @staticmethod
    def _block_label(block) -> str:
        """Extract the question label text from a Google Form block element."""
        for cls in (
            "freebirdFormviewerViewItemsItemItemTitle",
            "freebirdFormviewerViewItemsTextTextItemTitle",
            "M7eMe",
            "APjFqb",
        ):
            try:
                return block.find_element(By.XPATH, f".//*[contains(@class,'{cls}')]").text.strip()
            except NoSuchElementException:
                pass
        return block.text.split("\n")[0].strip()

    @staticmethod
    def _fill_block(block, q_text: str, answer: str) -> None:
        """Fill a Google Form field: text input, textarea, radio, checkbox, or dropdown."""
        # Text / number inputs
        for css in ("input[type='text']", "textarea", "input[type='number']"):
            try:
                inp = block.find_element(By.CSS_SELECTOR, css)
                inp.clear()
                inp.send_keys(answer)
                log.info(f"  [text] '{q_text[:35]}' → '{answer[:40]}'")
                return
            except NoSuchElementException:
                pass

        # Radio buttons (linear scale or MCQ)
        try:
            radios = block.find_elements(
                By.XPATH, ".//div[@role='radio'] | .//label[.//input[@type='radio']]"
            )
            if radios:
                for r in radios:
                    if answer.lower() in r.text.lower() or r.text.strip() == answer.strip():
                        r.click()
                        log.info(f"  [radio-text] '{q_text[:35]}' → '{r.text.strip()}'")
                        return
                # Numeric fallback — linear scale (answer = "7" → click 7th option)
                if answer.isdigit():
                    idx = max(0, min(int(answer) - 1, len(radios) - 1))
                    radios[idx].click()
                    log.info(f"  [radio-idx] '{q_text[:35]}' → index {idx}")
                return
        except Exception:
            pass

        # Checkboxes
        try:
            cbs = block.find_elements(By.XPATH, ".//div[@role='checkbox']")
            for cb in cbs:
                if answer.lower() in cb.text.lower():
                    cb.click()
                    log.info(f"  [checkbox] '{q_text[:35]}' → '{cb.text.strip()}'")
                    return
        except Exception:
            pass

        # Dropdown
        try:
            dd = block.find_element(
                By.XPATH, ".//div[@role='listbox'] | .//select"
            )
            dd.click()
            time.sleep(0.4)
            opt = block.find_element(
                By.XPATH,
                f".//li[contains(normalize-space(text()),'{answer}')] | "
                f".//option[contains(normalize-space(text()),'{answer}')]",
            )
            opt.click()
            log.info(f"  [dropdown] '{q_text[:35]}' → '{answer[:40]}'")
        except Exception:
            pass

    # ── Orchestration ─────────────────────────────────────────────────────────
    def run(self) -> None:
        self.driver = self._build_driver()
        try:
            self.open_looker_studio()       # Step 1 — open + wait 5 min
            self.scroll_to_table()          # Step 2 — scroll into table view
            lead_id = self.click_first_lead_id()    # Step 3 — click Lead ID
            self.handle_google_redirect()   # Step 4 — bypass redirect notice
            self.wait_for_chakrahq()        # Step 5 — wait for lead page
            lead_data = self.scrape_chakrahq()      # Step 6 — scrape
            lead_data["_meta_lead_id"] = lead_id
            _save_json(SCRIPT_DIR / "lead_data.json", lead_data)

            review = self.get_gpt_review(lead_data) # Step 7 — GPT audit
            _save_json(SCRIPT_DIR / "gpt_review.json", review)

            self.switch_to_looker()         # Step 8 — back to Looker Studio
            self.click_audit_form_lq(lead_id)       # Step 9 — click LQ form link
            self.fill_google_form(review)   # Step 10 — fill form

            log.info("=== Audit Bot completed successfully ===")

        except AuditBotError as e:
            log.error(f"Bot halted: {e}")
            self._screenshot("err_fatal")
            raise
        except Exception as e:
            log.exception(f"Unexpected error: {e}")
            self._screenshot("err_unexpected")
            raise
        finally:
            time.sleep(6)
            if self.driver:
                self.driver.quit()


# ── Pure helpers ──────────────────────────────────────────────────────────────
def _flatten(d: dict, prefix: str = "") -> dict:
    """Flatten a nested dict into dot-notation string keys."""
    out: dict = {}
    for k, v in d.items():
        key = f"{prefix}.{k}" if prefix else str(k)
        if isinstance(v, dict):
            out.update(_flatten(v, key))
        elif isinstance(v, list):
            out[key] = "; ".join(str(i) for i in v)
        else:
            out[key] = str(v) if v is not None else ""
    return out


# Map of question substring keywords → flat/review dict keys to probe
_KEYWORD_MAP: dict = {
    "lead id":       ["_meta_lead_id", "Lead ID"],
    "agent":         ["assigned_to", "Qualified By", "agent_name"],
    "phone":         ["Phone"],
    "score":         ["audit_score", "score", "overall_score"],
    "remark":        ["overall_remark", "remarks", "feedback"],
    "greeting":      ["greeting_score", "greeting"],
    "need":          ["need_identification_score"],
    "product":       ["product_knowledge_score"],
    "objection":     ["objection_handling_score"],
    "closing":       ["closing_score"],
    "compliance":    ["compliance_score"],
    "followup":      ["followup_score"],
    "follow":        ["followup_score"],
    "grade":         ["grade"],
    "budget":        ["Budget", "budget"],
    "category":      ["Category"],
    "gemstone":      ["Gemstone"],
    "purpose":       ["Purpose"],
    "disposition":   ["Primary Disposition", "primary_disposition"],
    "status":        ["lead_status", "Qualification Status"],
    "qualified":     ["is_lead_correctly_qualified", "Qualification Status"],
    "action":        ["recommended_action"],
    "recommend":     ["recommended_action"],
    "strength":      ["strengths"],
    "positive":      ["strengths"],
    "improve":       ["improvements"],
    "issue":         ["improvements"],
    "gap":           ["improvements"],
    "missed":        ["improvements"],
    "name":          ["Name", "agent_name"],
    "source":        ["Lead Source"],
    "order":         ["Order ID", "Final Order Value"],
    "review":        ["overall_remark"],
    "audit":         ["audit_score"],
    "overall":       ["overall_remark", "audit_score"],
}


def _best_answer(question: str, flat: dict, review: dict) -> str:
    """Return the most relevant value for a form question."""
    for kw, keys in _KEYWORD_MAP.items():
        if kw in question:
            for k in keys:
                if k in review:
                    v = review[k]
                    if isinstance(v, list):
                        return "; ".join(str(x) for x in v)
                    return str(v) if v is not None else ""
                for fk, fv in flat.items():
                    if k.lower() in fk.lower() and fv:
                        return fv
    # Fallback: first non-empty string in review
    for v in review.values():
        if isinstance(v, str) and v:
            return v
    return ""


def _save_json(path: Path, data: dict) -> None:
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)
    log.info(f"Saved: {path.name}")


# ── Entry point ───────────────────────────────────────────────────────────────
def main() -> None:
    load_dotenv(SCRIPT_DIR / ".env")

    api_key = os.getenv("OPENAI_API_KEY", "").strip()
    if not api_key:
        raise EnvironmentError("OPENAI_API_KEY is not set in .env")

    prompt_path = SCRIPT_DIR / "prompt.txt"
    if not prompt_path.exists():
        raise FileNotFoundError(f"prompt.txt not found at {prompt_path}")
    system_prompt = prompt_path.read_text(encoding="utf-8").strip()
    if not system_prompt:
        raise ValueError("prompt.txt is empty — add your audit instructions.")

    log.info("=== Chakra HQ Audit Bot starting ===")
    AuditBot(openai_key=api_key, system_prompt=system_prompt).run()


if __name__ == "__main__":
    main()

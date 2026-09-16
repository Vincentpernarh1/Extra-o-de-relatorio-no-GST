import asyncio
import os
import json
import queue
import signal
import sys
import threading
import time
import tkinter as tk
from tkinter import messagebox, scrolledtext, ttk

# When built as a windowed .exe (console=False in the .spec), sys.stdout/stderr
# are None. Any print() call then raises AttributeError and kills the process
# before any window is ever shown - it just silently vanishes. Redirect to a
# log file next to the executable so prints don't crash and logs are visible.
if getattr(sys, "frozen", False) and (sys.stdout is None or sys.stderr is None):
    _log_path = os.path.join(os.path.dirname(sys.executable), "rpa_log.txt")
    _log_file = open(_log_path, "a", encoding="utf-8", buffering=1)
    sys.stdout = _log_file
    sys.stderr = _log_file

from playwright.async_api import async_playwright, TimeoutError as PlaywrightTimeoutError

BASE_DIR = os.getcwd()
DOWNLOAD_DIR = os.path.join(BASE_DIR, "Downloads_Auxiliar")
CONSOLIDATED_DIR = os.path.join(BASE_DIR, "Arquivos_Consolidados")
URL = "https://grouppurchasing.fiat.com/irj/portal/gssm?standAlone=true&sapDocumentRenderingMode=Edge&HistoryMode=1&TarTitle=Source%20Package%20Management&windowId=WID1703160349761&NavMode=0"
FRAME_NAME = "ivuFrm_page0ivu1"


def _load_credentials():
    cred_path = os.path.join(BASE_DIR, "Credencial", "usuario.json")
    with open(cred_path, "r") as f:
        credenciais = json.load(f)
    username = credenciais.get("Usuario", "")
    password = credenciais.get("Senha", "")
    if not username or not password:
        raise ValueError("Usuario ou senha ausente em Credencial/usuario.json")
    return username, password


LOGIN_XPATH = '//*[@id="logonuidfield"]'
PASSWORD_XPATH = '//*[@id="logonpassfield"]'
LOGIN_BUTTON_XPATH = '/html/body/span/table/tbody/tr/td/table/tbody/tr/td/form/table/tbody/tr[2]/td[3]/table/tbody/tr[5]/td[2]/input'
APPLICATION_XPATH = '/html/body/table/tbody/tr[1]/td/table/tbody/tr[1]/td/div[2]/ul[2]/div/li[2]'
GLOBAL_SOURCING_TOOL_XPATH = '/html/body/table/tbody/tr[1]/td/table/tbody/tr[2]/td/table/tbody/tr[1]/td/table/tbody/tr/td/table/tbody/tr/td/div/div/table/tbody/tr[2]/td/table/tbody/tr[1]/td[1]'
SOURCE_PACKAGE_MANAGEMENT_XPATH = '/html/body/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr/td/table/tbody/tr/td/div/div/table/tbody/tr[4]/td/table/tbody/tr[1]/td[1]/img'
REPORTING_PACKAGE_XPATH = '/html/body/table/tbody/tr/td/div/table/tbody/tr/td/div/table/tbody/tr/td/div/table/tbody/tr[2]/td/div/table/tbody/tr[1]/td/table/tbody/tr/td[2]/div/div[3]/div[1]'
REPORTING_DISPLAY_XPATH = '/html/body/table/tbody/tr/td/div/table/tbody/tr/td/div/table/tbody/tr/td/div/table/tbody/tr[2]/td/div/table/tbody/tr[3]/td/div[3]/div/table/tbody/tr/td/div/div/div/div/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/span/span/table/tbody/tr/td/span/span[1]/div/div/table/tbody/tr/td[1]/div/span/span/div'
REPORTING_MODIFICATION_XPATH = '/html/body/table/tbody/tr/td/div/table/tbody/tr/td/div/table/tbody/tr/td/div/table/tbody/tr[2]/td/div/table/tbody/tr[3]/td/div[3]/div/table/tbody/tr/td/div/div/div/div/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/span/span/table/tbody/tr/td/span/span[2]/table/tbody/tr/td/div/table/tbody/tr/td/div/table/tbody/tr[1]/td/div/div/div[2]/span/span/table/tbody/tr/td/div/table/tbody/tr[16]/td[3]/span/input'
REPORTING_STATUS_XPATH = '/html/body/table/tbody/tr/td/div/table/tbody/tr/td/div/table/tbody/tr/td/div/table/tbody/tr[2]/td/div/table/tbody/tr[3]/td/div[3]/div/table/tbody/tr/td/div/div/div/div/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/span/span/table/tbody/tr/td/span/span[2]/table/tbody/tr/td/div/table/tbody/tr/td/div/table/tbody/tr[1]/td/div/div/div[2]/span/span/table/tbody/tr/td/div/table/tbody/tr[13]/td[3]/span/input'
REPORTING_VIEW_XPATH = '/html/body/table/tbody/tr/td/div/table/tbody/tr/td/div/table/tbody/tr/td/div/table/tbody/tr[2]/td/div/table/tbody/tr[3]/td/div[3]/div/table/tbody/tr/td/div/div/div/div/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/span/span/table/tbody/tr/td/span/span[4]/table/tbody/tr/td/span/span[1]/div/div/div/span/span/table/tbody/tr[1]/td/div/div/div/div[1]/span[1]/span[2]/input'
REPORTING_SEARCH_XPATH = '/html/body/table/tbody/tr/td/div/table/tbody/tr/td/div/table/tbody/tr/td/div/table/tbody/tr[2]/td/div/table/tbody/tr[3]/td/div[3]/div/table/tbody/tr/td/div/div/div/div/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/span/span/table/tbody/tr/td/span/span[2]/table/tbody/tr/td/div/table/tbody/tr/td/div/table/tbody/tr[2]/td/div/span[1]/div'
REPORTING_DOWNLOAD_XPATH = '/html/body/table/tbody/tr/td/div/table/tbody/tr/td/div/table/tbody/tr/td/div/table/tbody/tr[2]/td/div/table/tbody/tr[3]/td/div[3]/div/table/tbody/tr/td/div/div/div/div/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/span/span/table/tbody/tr/td/span/span[4]/table/tbody/tr/td/span/span[1]/div/div/div/span/span/table/tbody/tr[1]/td/div/div/div/div[1]/span[3]/div'
BREADCRUMB_GST_XPATH = '//*[@id="gssm_breadcrumb"]/div/a[2]'
SOURCING_MANAGEMENT_XPATH = '/html/body/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr/td/table/tbody/tr/td/div/div/table/tbody/tr[6]/td[1]/table/tbody/tr[1]/td[1]/img'
SOURCE_PROCESS_DASHBOARD_XPATH = '/html/body/table/tbody/tr/td/div/table/tbody/tr/td/div/table/tbody/tr/td/div/table/tbody/tr[2]/td/div/table/tbody/tr[1]/td/table/tbody/tr/td[2]/div/div[2]/div[1]'
SOURCING_MODIFICATION_XPATH = '/html/body/table/tbody/tr/td/div/table/tbody/tr/td/div/table/tbody/tr/td/div/table/tbody/tr[2]/td/div/table/tbody/tr[3]/td/div[2]/div/table/tbody/tr/td/div/div/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/span/span/table/tbody/tr/td/span/span[2]/table/tbody/tr/td/div/table/tbody/tr/td/div/table/tbody/tr[1]/td/div/div/div[2]/span/span/table/tbody/tr/td/div/table/tbody/tr[11]/td[3]/span/input'
SOURCING_REGION_XPATH = '/html/body/table/tbody/tr/td/div/table/tbody/tr/td/div/table/tbody/tr/td/div/table/tbody/tr[2]/td/div/table/tbody/tr[3]/td/div[2]/div/table/tbody/tr/td/div/div/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/span/span/table/tbody/tr/td/span/span[2]/table/tbody/tr/td/div/table/tbody/tr/td/div/table/tbody/tr[1]/td/div/div/div[2]/span/span/table/tbody/tr/td/div/table/tbody/tr[8]/td[3]/span/input'
SOURCING_VIEW_XPATH = '/html/body/table/tbody/tr/td/div/table/tbody/tr/td/div/table/tbody/tr/td/div/table/tbody/tr[2]/td/div/table/tbody/tr[3]/td/div[2]/div/table/tbody/tr/td/div/div/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/span/span/table/tbody/tr/td/span/span[4]/table/tbody/tr/td/span/span[1]/div/div/div/span/span/table/tbody/tr[1]/td/div/div/div/div[1]/span[1]/span[2]/input'
SOURCING_SEARCH_XPATH = '/html/body/table/tbody/tr/td/div/table/tbody/tr/td/div/table/tbody/tr/td/div/table/tbody/tr[2]/td/div/table/tbody/tr[3]/td/div[2]/div/table/tbody/tr/td/div/div/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/span/span/table/tbody/tr/td/span/span[2]/table/tbody/tr/td/div/table/tbody/tr/td/div/table/tbody/tr[2]/td/div/span[1]/div'
SOURCING_DOWNLOAD_XPATH = '/html/body/table/tbody/tr/td/div/table/tbody/tr/td/div/table/tbody/tr/td/div/table/tbody/tr[2]/td/div/table/tbody/tr[3]/td/div[2]/div/table/tbody/tr/td/div/div/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/span/span/table/tbody/tr/td/span/span[4]/table/tbody/tr/td/span/span[1]/div/div/div/span/span/table/tbody/tr[1]/td/div/div/div/div[1]/span[9]/div'
LOADING_XPATH = '//*[@id="ur-loading-itm2"]'


def _dialog_root():
    """Creates an invisible-but-mapped Tk root forced to the foreground so
    message boxes don't end up hidden behind the browser or other windows.
    A withdrawn (unmapped) root can make its transient dialog fail to paint
    at all, so we keep the root "shown" via 0 alpha instead of withdraw()."""
    root = tk.Tk()
    root.geometry("1x1+0+0")
    root.attributes("-alpha", 0.0)
    root.attributes("-topmost", True)
    root.deiconify()
    root.lift()
    root.update()
    return root


def show_error(title, message):
    """Last-resort dialog for failures that happen outside the main app
    window (e.g. the window itself failing to construct)."""
    root = _dialog_root()
    try:
        messagebox.showerror(title, message, parent=root)
    finally:
        root.destroy()


def _locator(scope, xpath):
    return scope.locator(f"xpath={xpath}")


async def _wait_visible(scope, xpath, timeout=60000):
    locator = _locator(scope, xpath)
    await locator.wait_for(state="visible", timeout=timeout)
    return locator


async def _click(scope, xpath, timeout=60000):
    locator = await _wait_visible(scope, xpath, timeout)
    await locator.click()
    return locator


async def _type_text(scope, xpath, value, timeout=60000):
    locator = await _wait_visible(scope, xpath, timeout)
    # Use press_sequentially like Selenium's send_keys for readonly fields
    await locator.press_sequentially(value, delay=50)
    return locator


async def _choose_autocomplete(page, scope, xpath, value):
    # Match Selenium behavior: just send keys directly to the element
    locator = _locator(scope, xpath)
    await locator.wait_for(state="visible", timeout=60000)
    # Send keys directly (like Selenium's send_keys) - triggers events even on readonly
    await locator.press_sequentially(value, delay=50)
    await page.wait_for_timeout(1000)
    # Use global keyboard actions like Selenium's ActionChains
    await page.keyboard.press("Enter")
    await page.wait_for_timeout(1000)
    await page.keyboard.press("ArrowUp")
    await page.wait_for_timeout(1000)
    await page.keyboard.press("ArrowUp")
    await page.wait_for_timeout(1000)
    await page.keyboard.press("Enter")
    await page.wait_for_timeout(1000)


async def _wait_for_loading(page):
    # Wait for spinner to disappear. Returns immediately if already hidden or not in DOM.
    loading = page.locator(f"xpath={LOADING_XPATH}")
    try:
        await loading.wait_for(state="hidden", timeout=900000)
    except PlaywrightTimeoutError:
        pass


async def _frame(page):
    # Wait for network to settle after navigation
    try:
        await page.wait_for_load_state("networkidle", timeout=900000)
    except PlaywrightTimeoutError:
        pass

    # Wait for any iframe with ivu in the name to appear before scanning
    try:
        await page.locator('iframe[name*="ivu"], iframe[name*="Source"], iframe[name*="Sourcing"]').first.wait_for(state="attached", timeout=30000)
    except PlaywrightTimeoutError:
        pass

    # Give extra time for iframe to be injected
    await page.wait_for_timeout(3000)

    # Find the iframe by various possible names
    all_frames = page.frames
    target_frame_name = None

    for frame in all_frames:
        frame_name = frame.name if frame.name else ""

        if frame_name == FRAME_NAME:
            target_frame_name = FRAME_NAME
            break
        elif "ivu" in frame_name.lower() or "page0ivu" in frame_name.lower():
            target_frame_name = frame_name
            break
        elif "Source Package Management" in frame_name or "Sourcing Management" in frame_name:
            target_frame_name = frame_name
            break

    if not target_frame_name:
        target_frame_name = FRAME_NAME

    iframe_locator = page.locator(f'iframe[name="{target_frame_name}"]')
    await iframe_locator.wait_for(state="attached", timeout=60000)
    await page.wait_for_timeout(3000)
    return page.frame_locator(f'iframe[name="{target_frame_name}"]')


def _clean_download_dir():
    os.makedirs(DOWNLOAD_DIR, exist_ok=True)
    for filename in os.listdir(DOWNLOAD_DIR):
        if filename.lower().endswith(".xlsx"):
            try:
                os.remove(os.path.join(DOWNLOAD_DIR, filename))
            except OSError:
                pass


async def _clear_and_type(scope, xpath, value, timeout=20000):
    """Selects all existing content with Ctrl+A then types the new value."""
    locator = await _wait_visible(scope, xpath, timeout)
    await locator.click()
    await locator.press("Control+a")
    await locator.press("Delete")
    await locator.press_sequentially(value, delay=50)
    return locator


async def _save_download(page, scope, click_xpath, output_path):
    async with page.expect_download(timeout=600000) as download_info:
        await _click(scope, click_xpath, timeout=60000)
    download = await download_info.value
    await download.save_as(output_path)


def get_playwright_browser_path():
    """Get Playwright browser path for bundled .exe or development."""
    if getattr(sys, 'frozen', False):
        # Running in bundled .exe
        base_path = sys._MEIPASS
        chromium_path = os.path.join(base_path, "ms-playwright", "chromium-1228", "chrome-win64", "chrome.exe")
    else:
        # Running in development
        base_path = os.path.join(os.path.expanduser("~"), "AppData", "Local")
        chromium_path = os.path.join(
            base_path,
            "ms-playwright",
            "chromium-1228",
            "chrome-win64",
            "chrome.exe"
        )
   
    if chromium_path and not os.path.exists(chromium_path):
        raise FileNotFoundError(f"Chromium executable not found at {chromium_path}")

    return chromium_path


async def _launch_browser(playwright):
    launch_options = {
        "headless": False,
        "args": ["--start-maximized"],
    }

    # Try to get bundled Playwright browser path
    try:
        executable_path = get_playwright_browser_path()
        launch_options["executable_path"] = executable_path
        return await playwright.chromium.launch(**launch_options)
    except (FileNotFoundError, Exception):
        # Fallback to system Chrome or installed Chromium
        try:
            return await playwright.chromium.launch(channel="chrome", **launch_options)
        except Exception:
            try:
                return await playwright.chromium.launch(**launch_options)
            except Exception as exc:
                raise RuntimeError(
                    "Nao foi possivel iniciar o navegador com Playwright. "
                    "Verifique se o Google Chrome esta instalado ou execute 'playwright install chromium'."
                ) from exc


async def _maximize_window(context, page):
    """--start-maximized only affects Chromium's initial default window;
    context.new_page() opens a brand-new window/target that never inherits
    it. Force the actual window via CDP instead, which works regardless of
    how the page was created."""
    try:
        cdp = await context.new_cdp_session(page)
        window_info = await cdp.send("Browser.getWindowForTarget")
        await cdp.send("Browser.setWindowBounds", {
            "windowId": window_info["windowId"],
            "bounds": {"windowState": "maximized"},
        })
    except Exception:
        pass


async def _new_page(context):
    page = await context.new_page()
    page.set_default_timeout(20000)
    page.set_default_navigation_timeout(600000)
    await _maximize_window(context, page)
    return page


async def run_source_package_report(context):
    page = await _new_page(context)
    await page.goto(URL, wait_until="domcontentloaded")
    await page.wait_for_timeout(2000)

    await _click(page, APPLICATION_XPATH)
    await page.wait_for_timeout(2000)
    await _click(page, GLOBAL_SOURCING_TOOL_XPATH)
    await page.wait_for_timeout(2000)
    await _click(page, SOURCE_PACKAGE_MANAGEMENT_XPATH)
    await page.wait_for_timeout(3000)

    frame = await _frame(page)
    await page.pause()
    await _click(frame, REPORTING_PACKAGE_XPATH)
    await page.wait_for_timeout(3000)
    await _wait_for_loading(page)
    await page.wait_for_timeout(2000)
    await _click(frame, REPORTING_DISPLAY_XPATH)
    await page.wait_for_timeout(5000)
    await _choose_autocomplete(page, frame, REPORTING_MODIFICATION_XPATH, "Semana Anterior")
    await page.wait_for_timeout(2000)
    await _type_text(frame, REPORTING_STATUS_XPATH, "Technical Data Completed")
    await page.wait_for_timeout(3000)
    await _type_text(frame, REPORTING_VIEW_XPATH, "RPA")
    await page.wait_for_timeout(3000)
    await _click(frame, REPORTING_SEARCH_XPATH)
    await page.wait_for_timeout(5000)
    await _save_download(
        page,
        frame,
        REPORTING_DOWNLOAD_XPATH,
        os.path.join(DOWNLOAD_DIR, "01_source_package_management.xlsx"),
    )
    await page.close()


async def run_sourcing_reports(context):
    """Runs all 3 sourcing region reports sequentially on one page.
    The SAP portal maintains server-side session state per cookie session,
    so concurrent search requests would overwrite each other's filters.
    """
    page = await _new_page(context)
    await page.goto(URL, wait_until="domcontentloaded")
    await page.wait_for_timeout(2000)

    await _click(page, APPLICATION_XPATH)
    await page.wait_for_timeout(2000)
    await _click(page, GLOBAL_SOURCING_TOOL_XPATH)
    await page.wait_for_timeout(3000)
    await _click(page, SOURCING_MANAGEMENT_XPATH)
    await page.wait_for_timeout(30000)

    frame = await _frame(page)
    await _click(frame, SOURCE_PROCESS_DASHBOARD_XPATH)
    await page.wait_for_timeout(3000)
    await _wait_for_loading(page)
    await page.wait_for_timeout(1000)
    await _choose_autocomplete(page, frame, SOURCING_MODIFICATION_XPATH, "Semana Anterior")
    await page.wait_for_timeout(1000)
    await _type_text(frame, SOURCING_VIEW_XPATH, "RPA")
    await page.wait_for_timeout(1000)

    regions = [
        ("Global",  "02_sourcing_management_global.xlsx"),
        ("LATAM",   "03_sourcing_management_latam.xlsx"),
        ("Neutral", "04_sourcing_management_neutral.xlsx"),
    ]
    for region, filename in regions:
        # triple_click selects any existing text before typing the new region
        await _clear_and_type(frame, SOURCING_REGION_XPATH, region)
        await page.wait_for_timeout(1000)
        await _click(frame, SOURCING_SEARCH_XPATH)
        await page.wait_for_timeout(3000)
        await _wait_for_loading(page)
        await page.wait_for_timeout(1000)
        await _save_download(
            page,
            frame,
            SOURCING_DOWNLOAD_XPATH,
            os.path.join(DOWNLOAD_DIR, filename),
        )
        await page.wait_for_timeout(3000)

    await page.close()


async def async_main(log=print, progress=lambda value: None):
    os.makedirs(CONSOLIDATED_DIR, exist_ok=True)
    _clean_download_dir()

    log("Carregando credenciais...")
    progress(5)
    username, password = _load_credentials()

    async with async_playwright() as playwright:
        log("Iniciando navegador...")
        progress(10)
        browser = await _launch_browser(playwright)
        context = await browser.new_context(accept_downloads=True, viewport=None, locale="pt-BR")
        page = await _new_page(context)

        log("Acessando portal...")
        progress(15)
        await page.goto(URL, wait_until="domcontentloaded")
        log("Realizando login...")
        await _click(page, LOGIN_XPATH)
        await _type_text(page, LOGIN_XPATH, username)
        await page.wait_for_timeout(1000)
        await _click(page, PASSWORD_XPATH)
        await _type_text(page, PASSWORD_XPATH, password)
        await page.wait_for_timeout(2000)
        await _click(page, LOGIN_BUTTON_XPATH)
        await page.wait_for_timeout(3000)
        log("Login realizado com sucesso.")
        progress(25)

        # --- Report 01: Source Package Management ---
        log("Baixando relatório: Source Package Management...")
        progress(30)
        await _click(page, APPLICATION_XPATH)
        await page.wait_for_timeout(2000)
        await _click(page, GLOBAL_SOURCING_TOOL_XPATH)
        await page.wait_for_timeout(2000)
        await _click(page, SOURCE_PACKAGE_MANAGEMENT_XPATH)
        await page.wait_for_timeout(3000)
        frame = await _frame(page)
        await _click(frame, REPORTING_PACKAGE_XPATH)
        await page.wait_for_timeout(3000)
        await _wait_for_loading(frame)
        await page.wait_for_timeout(2000)
        await _click(frame, REPORTING_DISPLAY_XPATH)
        await page.wait_for_timeout(5000)
        await _choose_autocomplete(page, frame, REPORTING_MODIFICATION_XPATH, "Semana Anterior")
        await page.wait_for_timeout(2000)
        await _type_text(frame, REPORTING_STATUS_XPATH, "Technical Data Completed")
        await page.wait_for_timeout(3000)
        await _type_text(frame, REPORTING_VIEW_XPATH, "RPA")
        await page.wait_for_timeout(3000)
        await _click(frame, REPORTING_SEARCH_XPATH)
        await page.wait_for_timeout(5000)
        await _wait_for_loading(frame)
        log("Aguardando download...")
        await _save_download(
            page,
            frame,
            REPORTING_DOWNLOAD_XPATH,
            os.path.join(DOWNLOAD_DIR, "01_source_package_management.xlsx"),
        )
        log("Relatório salvo como: 01_source_package_management.xlsx")
        progress(45)
        await page.wait_for_timeout(3000)

        # --- Reports 02-04: Sourcing Management (Global, LATAM, Neutral) ---
        await _click(page, BREADCRUMB_GST_XPATH)
        await _wait_for_loading(page)
        await page.wait_for_timeout(1000)
        await _click(page, SOURCING_MANAGEMENT_XPATH)
        await _wait_for_loading(page)
        await page.wait_for_timeout(1000)

        frame = await _frame(page)
        await _click(frame, SOURCE_PROCESS_DASHBOARD_XPATH)
        await page.wait_for_timeout(3000)
        await _wait_for_loading(frame)
        await page.wait_for_timeout(1000)
        await _choose_autocomplete(page, frame, SOURCING_MODIFICATION_XPATH, "Semana Anterior")
        await page.wait_for_timeout(1000)
        await _type_text(frame, SOURCING_VIEW_XPATH, "RPA")
        await page.wait_for_timeout(1000)

        regions = [
            ("Global",  "02_sourcing_management_global.xlsx"),
            ("LATAM",   "03_sourcing_management_latam.xlsx"),
            ("Neutral", "04_sourcing_management_neutral.xlsx"),
        ]
        for i, (region, filename) in enumerate(regions):
            log(f"Baixando relatório: Sourcing Management ({region})...")
            await _clear_and_type(frame, SOURCING_REGION_XPATH, region)
            await page.wait_for_timeout(1000)
            await _click(frame, SOURCING_SEARCH_XPATH)
            await page.wait_for_timeout(3000)
            await _wait_for_loading(frame)
            await page.wait_for_timeout(1000)
            log("Aguardando download...")
            await _save_download(
                page,
                frame,
                SOURCING_DOWNLOAD_XPATH,
                os.path.join(DOWNLOAD_DIR, filename),
            )
            log(f"Relatório salvo como: {filename}")
            progress(45 + (i + 1) * 15)
            await page.wait_for_timeout(3000)
            await _wait_for_loading(frame)
            await page.wait_for_timeout(1000)

        log("Fechando navegador...")
        progress(95)
        await page.close()
        await context.close()
        await browser.close()


def signal_handler(sig, frame_arg):
    print("\n\nInterrompido pelo usuário (Ctrl+C). Encerrando...")
    sys.exit(0)


# DHL / STELLANTIS brand colors, shared with the other RPA tools (e.g. RPA FIAPE).
DHL_YELLOW = "#FFCC00"
STELLANTIS_BLUE = "#003DA5"
STELLANTIS_ORANGE = "#FF6600"

EXPECTED_FILES = [
    "01_source_package_management.xlsx",
    "02_sourcing_management_global.xlsx",
    "03_sourcing_management_latam.xlsx",
    "04_sourcing_management_neutral.xlsx",
]


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Ferramenta de Automação e Processamento RPA")
        self.geometry("700x550")
        self.resizable(True, True)

        icon_path = os.path.join(BASE_DIR, "Credencial", "icon.ico")
        if os.path.exists(icon_path):
            try:
                self.iconbitmap(icon_path)
            except tk.TclError:
                pass

        style = ttk.Style()
        style.theme_use("clam")
        style.configure("TButton", background=STELLANTIS_BLUE, foreground="white", relief="flat", padding=6)
        style.map("TButton", background=[("active", STELLANTIS_ORANGE)])
        style.configure(
            "TProgressbar", background=STELLANTIS_BLUE, troughcolor="#E8E8E8",
            bordercolor="#CCCCCC", lightcolor=STELLANTIS_ORANGE, darkcolor=STELLANTIS_BLUE,
        )

        self._queue = queue.Queue()
        self._extraction_running = False
        self._processing_running = False

        self._build_widgets()
        self._bring_to_front()

    def _build_widgets(self):
        container = tk.Frame(self, bg="white")
        container.pack(fill=tk.BOTH, expand=True)

        header_frame = tk.Frame(container, bg=STELLANTIS_BLUE, height=80)
        header_frame.pack(fill=tk.X)
        header_frame.pack_propagate(False)
        tk.Label(
            header_frame, text="🤖 Ferramenta de Extração GST",
            font=("Segoe UI", 16, "bold"), fg="white", bg=STELLANTIS_BLUE,
        ).pack(anchor="w", padx=15, pady=(10, 2))
        tk.Label(
            header_frame, text="Extração e Processamento de Relatórios GST",
            font=("Segoe UI", 9), fg=DHL_YELLOW, bg=STELLANTIS_BLUE,
        ).pack(anchor="w", padx=15, pady=(0, 10))

        main_frame = ttk.Frame(container, padding="13")
        main_frame.pack(fill=tk.BOTH, expand=True)

        self.status_var = tk.StringVar(value="Pronto para iniciar. Clique em 'Iniciar Extração'.")
        self.status_label = ttk.Label(
            main_frame, textvariable=self.status_var, font=("Segoe UI", 11), foreground=STELLANTIS_BLUE,
        )
        self.status_label.pack(pady=(2, 5), padx=1, fill=tk.X)

        self.progress_bar = ttk.Progressbar(main_frame, orient="horizontal", length=400, mode="determinate")
        self.progress_bar.pack(pady=10, padx=5, fill=tk.X)

        button_frame = ttk.Frame(main_frame)
        button_frame.pack(pady=4, fill=tk.X)

        self.start_button = ttk.Button(
            button_frame, text="▶ Iniciar Extração", command=self._start_extraction, style="TButton",
        )
        self.start_button.pack(side=tk.LEFT, padx=5)

        self.process_button = ttk.Button(
            button_frame, text="▶ Processar Arquivos", command=self._start_processing,
            style="TButton", state="disabled",
        )
        self.process_button.pack(side=tk.RIGHT, padx=5)

        log_frame = ttk.LabelFrame(main_frame, text="📋 Log de Atividades", padding="13")
        log_frame.pack(pady=0, padx=2, fill=tk.BOTH, expand=True)

        self.log_text = scrolledtext.ScrolledText(
            log_frame, wrap=tk.WORD, width=80, height=10,
            font=("Consolas", 11), bg="#F5F5F5", fg="#333333",
        )
        self.log_text.pack(fill=tk.BOTH, expand=True)

        footer_frame = tk.Frame(container, bg=STELLANTIS_BLUE, height=34)
        footer_frame.pack(fill=tk.X, side=tk.BOTTOM)
        footer_frame.pack_propagate(False)

        left_footer = tk.Frame(footer_frame, bg=STELLANTIS_BLUE)
        left_footer.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=15, pady=10)
        tk.Label(left_footer, text="🚚 DHL", font=("Segoe UI", 11, "bold"), fg=DHL_YELLOW, bg=STELLANTIS_BLUE).pack(side=tk.LEFT, padx=1)
        tk.Label(left_footer, text="→", font=("Segoe UI", 12, "bold"), fg=DHL_YELLOW, bg=STELLANTIS_BLUE).pack(side=tk.LEFT, padx=3)
        tk.Label(left_footer, text="STELLANTIS 🏢", font=("Segoe UI", 11, "bold"), fg=STELLANTIS_ORANGE, bg=STELLANTIS_BLUE).pack(side=tk.LEFT, padx=3)

        right_footer = tk.Frame(footer_frame, bg=STELLANTIS_BLUE)
        right_footer.pack(side=tk.RIGHT, padx=15, pady=10)
        tk.Label(right_footer, text="Desenvolvido por: Vincent Pernarh", font=("Segoe UI", 9), fg="white", bg=STELLANTIS_BLUE).pack(anchor="e")

    def _bring_to_front(self):
        self.lift()
        self.attributes("-topmost", True)
        self.after(300, lambda: self.attributes("-topmost", False))
        self.focus_force()

    # -- thread-safe messaging: worker threads only ever push onto the queue,
    # the Tk widgets are only ever touched from the main thread via _poll_queue.
    def _log(self, message):
        self._queue.put(("status", message))

    def _progress(self, value):
        self._queue.put(("progress", value))

    def _append_log(self, message):
        self.log_text.insert(tk.END, f"{time.strftime('%H:%M:%S')} - {message}\n")
        self.log_text.see(tk.END)

    def _poll_queue(self):
        """Drains the queue and updates widgets. Mirrors the other RPA tools'
        update_gui(): keeps rescheduling itself every 100ms until a *_done
        message arrives, then stops - the next _start_* call restarts it."""
        stop = False
        try:
            while True:
                kind, payload = self._queue.get_nowait()
                if kind == "status":
                    self.status_var.set(payload)
                    self._append_log(payload)
                elif kind == "progress":
                    self.progress_bar["value"] = payload
                elif kind == "extraction_done":
                    self._on_extraction_done(payload)
                    stop = True
                elif kind == "processing_done":
                    self._on_processing_done(payload)
                    stop = True
        except queue.Empty:
            pass
        if not stop:
            self.after(100, self._poll_queue)

    def _start_extraction(self):
        if self._extraction_running:
            return
        self._extraction_running = True
        self.start_button.config(state="disabled")
        self.process_button.config(state="disabled")
        self.progress_bar["value"] = 0
        self.status_var.set("🚀 Iniciando extração...")
        threading.Thread(target=self._run_extraction, daemon=True).start()
        self._poll_queue()

    def _run_extraction(self):
        try:
            asyncio.run(async_main(log=self._log, progress=self._progress))
            time.sleep(2)
            downloaded = [f for f in EXPECTED_FILES if os.path.exists(os.path.join(DOWNLOAD_DIR, f))]
            self._log(f"✅ Arquivos baixados: {len(downloaded)}/{len(EXPECTED_FILES)}")
            self._queue.put(("extraction_done", True))
        except Exception as exc:
            self._log(f"❌ Erro na extração: {exc}")
            self._queue.put(("extraction_done", False))

    def _on_extraction_done(self, success):
        self._extraction_running = False
        self.start_button.config(state="normal")
        if success:
            self.progress_bar["value"] = 100
            self.status_var.set("Processo Concluído!")
            self.process_button.config(state="normal")
        else:
            self.status_var.set("❌ Erro na extração. Veja o log.")
        self._bring_to_front()

    def _start_processing(self):
        if self._processing_running:
            return
        self._processing_running = True
        self.start_button.config(state="disabled")
        self.process_button.config(state="disabled")
        self.progress_bar["value"] = 0
        self.status_var.set("🚀 Processando arquivos...")
        threading.Thread(target=self._run_processing, daemon=True).start()
        self._poll_queue()

    def _run_processing(self):
        import main_organizar
        success = main_organizar.run(BASE_DIR, log=self._log, progress=self._progress)
        self._queue.put(("processing_done", success))

    def _on_processing_done(self, success):
        self._processing_running = False
        self.start_button.config(state="normal")
        self.process_button.config(state="normal")
        if success:
            self.progress_bar["value"] = 100
            self.status_var.set("Processo Concluído!")
        else:
            self.status_var.set("❌ Erro no processamento. Veja o log.")
        self._bring_to_front()


if __name__ == "__main__":
    # Register signal handler for Ctrl+C
    signal.signal(signal.SIGINT, signal_handler)

    try:
        app = App()
        app.mainloop()
    except Exception as exc:
        print(exc)
        show_error("Erro", f"Erro inesperado ao iniciar a aplicação: {exc}")

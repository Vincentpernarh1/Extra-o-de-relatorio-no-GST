import os
import json
import signal
import sys
import subprocess
import time
from tkinter import messagebox

from playwright.sync_api import TimeoutError as PlaywrightTimeoutError
from playwright.sync_api import sync_playwright

BASE_DIR = os.getcwd()  # Get the current working directory
DOWNLOAD_DIR = os.path.join(BASE_DIR, "Downloads_Auxiliar")
CONSOLIDATED_DIR = os.path.join(BASE_DIR, "Arquivos_Consolidados")
URL = "https://grouppurchasing.fiat.com/irj/portal/gssm?standAlone=true&sapDocumentRenderingMode=Edge&HistoryMode=1&TarTitle=Source%20Package%20Management&windowId=WID1703160349761&NavMode=0"
FRAME_NAME = "ivuFrm_page0ivu1"

USERNAME = ""
PASSWORD = ""

with open(os.path.join(BASE_DIR, "Credencial", "usuario.json"), "r") as f:
    credenciais = json.load(f)
    USERNAME = credenciais.get("Usuario", USERNAME)
    PASSWORD = credenciais.get("Senha", PASSWORD)

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


def _locator(scope, xpath):
    return scope.locator(f"xpath={xpath}")


def _wait_visible(scope, xpath, timeout=10000):
    locator = _locator(scope, xpath)
    locator.wait_for(state="visible", timeout=timeout)
    return locator


def _click(scope, xpath, timeout=10000):
    locator = _wait_visible(scope, xpath, timeout)
    locator.click()
    return locator


def _type_text(scope, xpath, value, timeout=10000):
    locator = _wait_visible(scope, xpath, timeout)
    # Use press_sequentially like Selenium's send_keys for readonly fields
    locator.press_sequentially(value, delay=50)
    return locator


def _choose_autocomplete(page, scope, xpath, value):
    # Match Selenium behavior: just send keys directly to the element
    locator = _locator(scope, xpath)
    locator.wait_for(state="visible", timeout=10000)
    # Send keys directly (like Selenium's send_keys) - triggers events even on readonly
    locator.press_sequentially(value, delay=50)
    page.wait_for_timeout(1000)
    # Use global keyboard actions like Selenium's ActionChains
    page.keyboard.press("Enter")
    page.wait_for_timeout(1000)
    page.keyboard.press("ArrowUp")
    page.wait_for_timeout(1000)
    page.keyboard.press("ArrowUp")
    page.wait_for_timeout(1000)
    page.keyboard.press("Enter")
    page.wait_for_timeout(1000)


def _wait_for_loading(page):
    loading = page.locator(f"xpath={LOADING_XPATH}")
    try:
        loading.wait_for(state="visible", timeout=10000)
        loading.wait_for(state="hidden", timeout=300000)
    except PlaywrightTimeoutError:
        pass


def _frame(page):
    # Wait for network to settle after navigation
    try:
        page.wait_for_load_state("networkidle", timeout=10000)
    except PlaywrightTimeoutError:
        pass
    
    # Give extra time for iframe to be injected
    page.wait_for_timeout(3000)
    
    # Find the iframe by various possible names
    all_frames = page.frames
    target_frame_name = None
    
    for frame in all_frames:
        frame_name = frame.name if frame.name else ""
        
        # Check for various possible frame names
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
        # If still not found, use the fallback
        target_frame_name = FRAME_NAME
    
    iframe_locator = page.locator(f'iframe[name="{target_frame_name}"]')
    iframe_locator.wait_for(state="attached", timeout=30000)
    page.wait_for_timeout(1000)
    return page.frame_locator(f'iframe[name="{target_frame_name}"]')


def _clean_download_dir():
    os.makedirs(DOWNLOAD_DIR, exist_ok=True)
    for filename in os.listdir(DOWNLOAD_DIR):
        if filename.lower().endswith(".xlsx"):
            try:
                os.remove(os.path.join(DOWNLOAD_DIR, filename))
            except OSError:
                pass


def _save_download(page, scope, click_xpath, output_path):
    download_trigger = _wait_visible(scope, click_xpath)
    with page.expect_download(timeout=300000) as download_info:
        download_trigger.click()
    download = download_info.value
    download.save_as(output_path)


def get_playwright_browser_path():
    """Get Playwright browser path for bundled .exe or development."""
    if getattr(sys, 'frozen', False):
        # Running in bundled .exe
        base_path = sys._MEIPASS
        chromium_path = os.path.join(base_path, "ms-playwright", "chromium-1187", "chrome-win", "chrome.exe")
    else:
        # Running in development
        base_path = os.path.join(os.path.expanduser("~"), "AppData", "Local")
        chromium_path = os.path.join(
            base_path,
            "ms-playwright",
            "chromium-1187",
            "chrome-win",
            "chrome.exe"
        )
   
    if chromium_path and not os.path.exists(chromium_path):
        raise FileNotFoundError(f"Chromium executable not found at {chromium_path}")

    return chromium_path


def _launch_browser(playwright):
    launch_options = {
        "headless": False,
        "args": ["--start-maximized"],
    }

    # Try to get bundled Playwright browser path
    try:
        executable_path = get_playwright_browser_path()
        launch_options["executable_path"] = executable_path
        return playwright.chromium.launch(**launch_options)
    except (FileNotFoundError, Exception):
        # Fallback to system Chrome or installed Chromium
        try:
            return playwright.chromium.launch(channel="chrome", **launch_options)
        except Exception:
            try:
                return playwright.chromium.launch(**launch_options)
            except Exception as exc:
                raise RuntimeError(
                    "Nao foi possivel iniciar o navegador com Playwright. "
                    "Verifique se o Google Chrome esta instalado ou execute 'playwright install chromium'."
                ) from exc


def signal_handler(sig, frame_arg):
    print("\n\nInterrompido pelo usuário (Ctrl+C). Encerrando...")
    sys.exit(0)


if __name__ == "__main__":
    # Register signal handler for Ctrl+C
    signal.signal(signal.SIGINT, signal_handler)

    resposta = messagebox.askquestion("Confirmação", "Deseja prosseguir com a extração?")
    if resposta == "yes":
        os.makedirs(CONSOLIDATED_DIR, exist_ok=True)
        _clean_download_dir()

        browser = None
        try:
            with sync_playwright() as playwright:
                browser = _launch_browser(playwright)
                context = browser.new_context(accept_downloads=True, viewport=None)
                page = context.new_page()
                page.set_default_timeout(10000)
                page.set_default_navigation_timeout(300000)

                page.goto(URL, wait_until="domcontentloaded")

                _click(page, LOGIN_XPATH)
                _type_text(page, LOGIN_XPATH, USERNAME)
                page.wait_for_timeout(1000)
                _click(page, PASSWORD_XPATH)
                _type_text(page, PASSWORD_XPATH, PASSWORD)
                page.wait_for_timeout(1000)
                _click(page, LOGIN_BUTTON_XPATH)
                page.wait_for_timeout(3000)

                _click(page, APPLICATION_XPATH)
                page.wait_for_timeout(1000)
                _click(page, GLOBAL_SOURCING_TOOL_XPATH)
                page.wait_for_timeout(1000)
                
                _click(page, SOURCE_PACKAGE_MANAGEMENT_XPATH)
                page.wait_for_timeout(1000)

                frame = _frame(page)
                _click(frame, REPORTING_PACKAGE_XPATH)
                page.wait_for_timeout(1000)
                _wait_for_loading(page)
                page.wait_for_timeout(1000)
                _click(frame, REPORTING_DISPLAY_XPATH)
                page.wait_for_timeout(5000)
                _choose_autocomplete(page, frame, REPORTING_MODIFICATION_XPATH, "Semana Anterior")
                page.wait_for_timeout(1000)
                _type_text(frame, REPORTING_STATUS_XPATH, "Technical Data Completed")
                page.wait_for_timeout(1000)
                _type_text(frame, REPORTING_VIEW_XPATH, "RPA")
                page.wait_for_timeout(2000)
                _click(frame, REPORTING_SEARCH_XPATH)
                page.wait_for_timeout(5000)
                _save_download(
                    page,
                    frame,
                    REPORTING_DOWNLOAD_XPATH,
                    os.path.join(DOWNLOAD_DIR, "01_source_package_management.xlsx"),
                )
                page.wait_for_timeout(5000)

                _click(page, BREADCRUMB_GST_XPATH)
                page.wait_for_timeout(1000)
                _click(page, SOURCING_MANAGEMENT_XPATH)
                page.wait_for_timeout(1000)

                frame = _frame(page)
                _click(frame, SOURCE_PROCESS_DASHBOARD_XPATH)
                page.wait_for_timeout(3000)
                _wait_for_loading(page)
                page.wait_for_timeout(1000)
                _choose_autocomplete(page, frame, SOURCING_MODIFICATION_XPATH, "Semana Anterior")
                page.wait_for_timeout(1000)
                _type_text(frame, SOURCING_REGION_XPATH, "Global")
                page.wait_for_timeout(1000)
                _type_text(frame, SOURCING_VIEW_XPATH, "RPA")
                page.wait_for_timeout(3000)
                _click(frame, SOURCING_SEARCH_XPATH)
                page.wait_for_timeout(3000)
                _wait_for_loading(page)
                page.wait_for_timeout(1000)
                _save_download(
                    page,
                    frame,
                    SOURCING_DOWNLOAD_XPATH,
                    os.path.join(DOWNLOAD_DIR, "02_sourcing_management_global.xlsx"),
                )
                page.wait_for_timeout(3000)

                _wait_for_loading(page)
                page.wait_for_timeout(1000)
                _type_text(frame, SOURCING_REGION_XPATH, "LATAM")
                page.wait_for_timeout(1000)
                _click(frame, SOURCING_SEARCH_XPATH)
                page.wait_for_timeout(3000)
                _wait_for_loading(page)
                page.wait_for_timeout(1000)
                _save_download(
                    page,
                    frame,
                    SOURCING_DOWNLOAD_XPATH,
                    os.path.join(DOWNLOAD_DIR, "03_sourcing_management_latam.xlsx"),
                )
                page.wait_for_timeout(3000)

                _wait_for_loading(page)
                page.wait_for_timeout(1000)
                _type_text(frame, SOURCING_REGION_XPATH, "Neutral")
                page.wait_for_timeout(1000)
                _click(frame, SOURCING_SEARCH_XPATH)
                page.wait_for_timeout(3000)
                _wait_for_loading(page)
                page.wait_for_timeout(1000)
                _save_download(
                    page,
                    frame,
                    SOURCING_DOWNLOAD_XPATH,
                    os.path.join(DOWNLOAD_DIR, "04_sourcing_management_neutral.xlsx"),
                )
                page.wait_for_timeout(3000)

                # Cleanup in proper order
                page.close()
                context.close()
            
            # Give files time to fully write to disk
            time.sleep(2)
            
            # Verify downloads completed
            expected_files = [
                "01_source_package_management.xlsx",
                "02_sourcing_management_global.xlsx",
                "03_sourcing_management_latam.xlsx",
                "04_sourcing_management_neutral.xlsx"
            ]
            downloaded = [f for f in expected_files if os.path.exists(os.path.join(DOWNLOAD_DIR, f))]
            print(f"Arquivos baixados: {len(downloaded)}/4")
            
            messagebox.showinfo("Sucesso", "Extração realizada com sucesso !")
            
            # Ask user if they want to process the files
            resposta_processar = messagebox.askquestion(
                "Processar Arquivos", 
                "Deseja processar os arquivos baixados agora?"
            )
            
            if resposta_processar == "yes":
                try:
                    # Run main_organizar.py to process the files (skip confirmation dialog)
                    organizar_script = os.path.join(BASE_DIR, "main_organizar.py")
                    subprocess.run(
                        [sys.executable, organizar_script, '--skip-confirmation'],
                        cwd=BASE_DIR,  # Ensure it runs in the correct directory
                        check=True
                    )
                except subprocess.CalledProcessError:
                    messagebox.showerror("Erro", "Erro ao processar os arquivos!")
                    messagebox.showinfo("Aviso", "Verifique os arquivos em Downloads_Auxiliar e tente novamente.")
                except Exception as exc:
                    messagebox.showerror("Erro", f"Erro ao processar: {exc}")
            else:
                messagebox.showinfo("Info", "Você pode processar os arquivos depois executando main_organizar.py")
                
        except KeyboardInterrupt:
            print("\n\nExtração cancelada pelo usuário.")
            sys.exit(0)
        except Exception as exc:
            print(exc)
            messagebox.showerror("Erro", "Erro ao executar a extração !")
            messagebox.showinfo("Aviso", "Tente novamente! Caso o erro persista, entre em contato com o suporte.")
    else:
        messagebox.showwarning("Cancelado", "Operação cancelada !")

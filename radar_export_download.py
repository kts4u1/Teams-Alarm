"""
RADAR - Export 자동 다운로드 (콘솔 창 없음, 완전 자동)

동작 방식
---------
1. exe를 더블클릭하면 콘솔 창 없이 Edge 브라우저 창만 뜹니다.
2. 뜬 창에서 평소처럼 로그인만 하시면 됩니다 (Enter 누를 필요 없음).
3. 로그인이 끝나 Export 버튼이 있는 화면이 나타나면, 프로그램이 자동으로
   감지해서 버튼을 클릭하고 다운로드를 진행합니다.
4. 완료되면 알림창이 하나 뜹니다 - 그것만 닫으면 끝입니다.

사전 준비
---------
pip install selenium webdriver-manager

빌드 명령어 (콘솔 창 없이 만들기)
----------------------------------
py -m PyInstaller --onefile --windowed --name RadarExport --collect-all selenium radar_export_download.py

다운로드된 파일은 바탕화면의 RadarDownloads 폴더에 저장됩니다.
로그인이 오래 걸려서 시간 초과가 뜨면, 아래 LOGIN_WAIT_TIMEOUT 값을 늘려주세요.
"""

import os
import time
import glob
import ctypes
from selenium import webdriver
from selenium.webdriver.edge.service import Service
from webdriver_manager.microsoft import EdgeChromiumDriverManager

LOGIN_URL = "https://portal.radar-digitalservices.hsp.philips.com/rmw/systemlist"
EXPORT_BUTTON_TEXT = "Export"
LOGIN_WAIT_TIMEOUT = 180  # 로그인 + 화면 로딩에 최대 기다려줄 시간(초)
DOWNLOAD_DIR = os.path.join(os.path.expanduser("~"), "Desktop", "RadarDownloads")


def show_message(text, title="RADAR 다운로드"):
    try:
        ctypes.windll.user32.MessageBoxW(0, text, title, 0)
    except Exception:
        pass  # Windows가 아닌 환경이면 조용히 넘어감


def setup_driver():
    os.makedirs(DOWNLOAD_DIR, exist_ok=True)
    options = webdriver.EdgeOptions()
    prefs = {
        "download.default_directory": DOWNLOAD_DIR,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True,
    }
    options.add_experimental_option("prefs", prefs)
    return webdriver.Edge(service=Service(EdgeChromiumDriverManager().install()), options=options)


FIND_AND_CLICK_JS = """
function findByText(root, text) {
    const all = root.querySelectorAll('*');
    for (const el of all) {
        if (el.shadowRoot) {
            const found = findByText(el.shadowRoot, text);
            if (found) return found;
        }
        if (el.children.length === 0 && el.textContent && el.textContent.trim() === text) {
            return el;
        }
    }
    return null;
}
const target = findByText(document, arguments[0]);
if (target) {
    target.click();
    return true;
}
return false;
"""


def wait_and_click_export(driver, timeout):
    """로그인이 끝나 Export 요소가 나타날 때까지 기다렸다가, 나타나면 바로 클릭합니다."""
    end_time = time.time() + timeout
    while time.time() < end_time:
        try:
            found = driver.execute_script(FIND_AND_CLICK_JS, EXPORT_BUTTON_TEXT)
        except Exception:
            found = False
        if found:
            return True
        time.sleep(1)
    return False


def wait_for_download(before_files, timeout=60):
    end_time = time.time() + timeout
    while time.time() < end_time:
        current_files = set(glob.glob(os.path.join(DOWNLOAD_DIR, "*")))
        new_files = current_files - before_files
        finished = [f for f in new_files if not f.endswith((".crdownload", ".tmp", ".partial"))]
        if finished:
            return finished[0]
        time.sleep(1)
    return None


def main():
    try:
        driver = setup_driver()
        driver.get(LOGIN_URL)

        before_files = set(glob.glob(os.path.join(DOWNLOAD_DIR, "*")))

        clicked = wait_and_click_export(driver, LOGIN_WAIT_TIMEOUT)
        if not clicked:
            driver.quit()
            show_message(
                f"{LOGIN_WAIT_TIMEOUT}초 안에 Export 버튼을 찾지 못했습니다.\n"
                "로그인이 늦어졌거나 화면이 바뀌었을 수 있습니다."
            )
            return

        downloaded_file = wait_for_download(before_files)
        driver.quit()

        if downloaded_file:
            show_message(f"다운로드 완료!\n\n{downloaded_file}")
        else:
            show_message(f"다운로드를 확인하지 못했습니다. 이 폴더를 확인해주세요:\n{DOWNLOAD_DIR}")

    except Exception as e:
        show_message(f"오류가 발생했습니다:\n{e}")


if __name__ == "__main__":
    main()

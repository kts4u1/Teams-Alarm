"""
RADAR - GA/GC(Gradient Amplifier / Gradient Coil Circuit Filling Pressure) 최신 값 조회

사람이 하는 순서 그대로 화면을 조작합니다.
  시스템 목록에서 SRN 검색 -> 번호 클릭 -> Dashboards -> System Cooling Dashboard
  -> 화면이 불러오는 그래프 데이터(fetchGeneratedPlots 응답)를 읽음

Edge 창이 뜨면 평소처럼 로그인만 하세요. 이후는 자동으로 진행됩니다.

사용법
------
pip install selenium webdriver-manager
py radar_gradient_pressure.py 24020 24062
"""

import base64
import json
import sys
import time
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.edge.service import Service
from webdriver_manager.microsoft import EdgeChromiumDriverManager

PORTAL_URL = "https://portal.radar-digitalservices.hsp.philips.com/rmw/systemlist"
PLOTS_API = "fetchGeneratedPlots"
LOGIN_WAIT_TIMEOUT = 180
STEP_TIMEOUT = 40

GA_PLOT = "LCC_GradAmpFillingPress"
GC_PLOT = "LCC_GradCoilFillingPress"
LOWER_BOUNDARY = 1.2  # 대시보드 그래프의 빨간 선 (bar)
UPPER_BOUNDARY = 2.5

# 보이는 요소 중 '자기 자신의 글자'가 text와 같은 첫 요소를 돌려줍니다.
FIND_BY_TEXT_JS = """
const text = arguments[0];
function own(el) {
    return Array.from(el.childNodes).filter(n => n.nodeType === 3)
        .map(n => n.textContent).join('').trim();
}
function search(root) {
    for (const el of root.querySelectorAll('*')) {
        if (el.shadowRoot) {
            const f = search(el.shadowRoot);
            if (f) return f;
        }
        if (own(el) === text && el.offsetParent !== null) return el;
    }
    return null;
}
return search(document);
"""

# 목록 맨 위 검색칸 중 첫 번째(System #)
SEARCH_BOX_JS = """
return Array.from(document.querySelectorAll('input'))
    .filter(i => i.placeholder === 'Search' && i.offsetParent !== null)[0] || null;
"""


def latest_value(plots, plot_id):
    """응답에서 plot_id의 가장 최근 (날짜, 값, 상태)를 돌려줍니다. 데이터가 없으면 None."""
    key = f"{plot_id}_0"
    for item in plots:
        if key not in item:
            continue
        data = item[key][0].get("data")
        if not isinstance(data, dict):  # "Not Applicable"
            return None
        for series in data.get("plotData", []):
            if series.get("name") != "Calculated Data" or "x" not in series:
                continue
            points = [(x, y) for x, y in zip(series["x"], series["y"]) if y is not None]
            if not points:
                return None
            date, value = points[-1]
            if value < LOWER_BOUNDARY:
                status = "LOW"
            elif value > UPPER_BOUNDARY:
                status = "HIGH"
            else:
                status = "OK"
            return {"date": date[:10], "value": value, "status": status}
    return None


def parse_gradient_pressure(plots):
    return {"GA": latest_value(plots, GA_PLOT), "GC": latest_value(plots, GC_PLOT)}


def setup_driver(profile_dir=None):
    options = webdriver.EdgeOptions()
    if profile_dir:  # 로그인 상태를 유지하는 전용 프로필
        options.add_argument(f"--user-data-dir={profile_dir}")
    for cap in ("ms:loggingPrefs", "goog:loggingPrefs"):
        options.set_capability(cap, {"performance": "ALL"})
    driver = webdriver.Edge(service=Service(EdgeChromiumDriverManager().install()), options=options)
    driver.execute_cdp_cmd("Network.enable", {})
    return driver


def wait_for(fn, timeout, what):
    end_time = time.time() + timeout
    while time.time() < end_time:
        try:
            result = fn()
            if result:
                return result
        except Exception:
            pass
        time.sleep(1)
    raise RuntimeError(f"시간 초과: {what}")


def click(driver, element):
    try:
        element.click()
    except Exception:
        driver.execute_script("arguments[0].click()", element)


def click_text(driver, text, timeout=STEP_TIMEOUT):
    element = wait_for(lambda: driver.execute_script(FIND_BY_TEXT_JS, text), timeout,
                       f"화면에서 '{text}'를 찾지 못했습니다")
    click(driver, element)


def open_cooling_dashboard(driver, srn, first):
    driver.get(PORTAL_URL)
    box_timeout = LOGIN_WAIT_TIMEOUT if first else STEP_TIMEOUT
    box = wait_for(lambda: driver.execute_script(SEARCH_BOX_JS), box_timeout,
                   "시스템 목록 화면(로그인 필요)")
    box.clear()
    box.send_keys(str(srn) + Keys.ENTER)
    click_text(driver, str(srn))
    click_text(driver, "Dashboards")
    driver.get_log("performance")  # 이전 기록 비우기
    click_text(driver, "System Cooling Dashboard")


def capture_plots(driver, srn, timeout=60):
    """화면이 보낸 fetchGeneratedPlots 요청 중 deviceId가 srn인 것의 응답을 읽습니다."""
    requests, finished = {}, set()
    end_time = time.time() + timeout
    while time.time() < end_time:
        for entry in driver.get_log("performance"):
            message = json.loads(entry["message"])["message"]
            method, params = message.get("method"), message.get("params", {})
            if method == "Network.requestWillBeSent":
                request = params.get("request", {})
                if request.get("url", "").endswith(PLOTS_API):
                    try:
                        device = json.loads(request.get("postData", "{}")).get("deviceId")
                    except ValueError:
                        device = None
                    requests[params["requestId"]] = str(device)
            elif method == "Network.loadingFinished":
                finished.add(params["requestId"])
        for request_id, device in requests.items():
            if device == str(srn) and request_id in finished:
                body = driver.execute_cdp_cmd("Network.getResponseBody", {"requestId": request_id})
                text = body["body"]
                if body.get("base64Encoded"):
                    text = base64.b64decode(text).decode("utf-8")
                return json.loads(text)
        time.sleep(1)
    raise RuntimeError("그래프 데이터 응답을 받지 못했습니다.")


def get_gradient_pressure(device_ids, driver=None):
    own_driver = driver is None
    if own_driver:
        driver = setup_driver()
    results = {}
    try:
        for i, srn in enumerate(device_ids):
            try:
                open_cooling_dashboard(driver, srn, first=(i == 0))
                results[srn] = parse_gradient_pressure(capture_plots(driver, srn))
            except Exception as e:
                print(f"[{srn}] 조회 실패: {e}", flush=True)
                results[srn] = {"error": str(e)}
        return results
    finally:
        if own_driver:
            driver.quit()


if __name__ == "__main__":
    ids = sys.argv[1:] or ["24020"]
    print(json.dumps(get_gradient_pressure(ids), ensure_ascii=False, indent=2))

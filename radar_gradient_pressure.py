"""
RADAR - GA/GC(Gradient Amplifier / Gradient Coil Circuit Filling Pressure) 최신 값 조회

Edge 창이 뜨면 평소처럼 로그인만 하세요. 로그인이 끝나면 자동으로 값을 가져옵니다.

사용법
------
pip install selenium webdriver-manager
py radar_gradient_pressure.py 24020 24062
"""

import json
import sys
import time
from selenium import webdriver
from selenium.webdriver.edge.service import Service
from webdriver_manager.microsoft import EdgeChromiumDriverManager

PORTAL_URL = "https://portal.radar-digitalservices.hsp.philips.com/rmw/systemlist"
PLOTS_API = "/api/v1/systemDetails/fetchGeneratedPlots"
LOGIN_WAIT_TIMEOUT = 180

GA_PLOT = "LCC_GradAmpFillingPress"
GC_PLOT = "LCC_GradCoilFillingPress"
LOWER_BOUNDARY = 1.2  # 대시보드 그래프의 빨간 선 (bar)
UPPER_BOUNDARY = 2.5

FETCH_JS = """
const done = arguments[arguments.length - 1];
fetch(arguments[0], {
    method: 'POST',
    credentials: 'include',
    headers: {'Content-Type': 'application/json', 'Accept': 'application/json'},
    body: arguments[1],
}).then(async r => done({status: r.status, body: await r.text()}))
  .catch(e => done({status: 0, body: String(e)}));
"""


def build_payload(device_id):
    return json.dumps({
        "date": time.strftime("%Y-%m-%dT%H:%M:%S.000Z", time.gmtime()),
        "deviceId": str(device_id),
        "plotIds": [GA_PLOT, GC_PLOT],
    })


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
    return webdriver.Edge(service=Service(EdgeChromiumDriverManager().install()), options=options)


def fetch_plots(driver, device_id):
    result = driver.execute_async_script(FETCH_JS, PLOTS_API, build_payload(device_id))
    if result["status"] != 200:
        raise RuntimeError(f"HTTP {result['status']}: {result['body'][:200]}")
    return json.loads(result["body"])


def wait_for_login(driver, device_id, timeout):
    end_time = time.time() + timeout
    last_error = None
    while time.time() < end_time:
        try:
            fetch_plots(driver, device_id)
            return
        except Exception as e:
            if str(e) != last_error:  # 같은 오류는 한 번만 보여줍니다
                print(f"[재시도 중] {e}", flush=True)
                last_error = str(e)
            time.sleep(3)
    raise RuntimeError(f"로그인 후 데이터를 가져오지 못했습니다. 마지막 오류: {last_error}")


def get_gradient_pressure(device_ids, driver=None):
    own_driver = driver is None
    if own_driver:
        driver = setup_driver()
    try:
        driver.get(PORTAL_URL)
        wait_for_login(driver, device_ids[0], LOGIN_WAIT_TIMEOUT)
        return {d: parse_gradient_pressure(fetch_plots(driver, d)) for d in device_ids}
    finally:
        if own_driver:
            driver.quit()


if __name__ == "__main__":
    ids = sys.argv[1:] or ["24020"]
    print(json.dumps(get_gradient_pressure(ids), ensure_ascii=False, indent=2))

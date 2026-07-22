import requests
import urllib3

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

urls = [
    "https://10.14.173.54:8055/api/sensor-readings/ingest",
    "https://10.14.173.54:8053/api/sensor-readings/ingest",
    "http://10.14.173.54:8053/api/sensor-readings/ingest",
    "http://10.14.173.54:8055/api/sensor-readings/ingest"
]

for url in urls:
    try:
        print(f"Testing {url} ...")
        resp = requests.post(url, json={}, verify=False, timeout=3)
        print(f"Status: {resp.status_code}")
    except Exception as e:
        print(f"Error: {e}")

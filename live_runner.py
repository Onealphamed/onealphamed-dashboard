import time
import subprocess

print("🚀 Live dashboard started...")

while True:
    print("🔄 Fetching latest data from Google Sheet...")
    subprocess.run(["python", "generate_dashboard.py"])
    time.sleep(10)  # every 10 seconds
#!/usr/bin/env python3
import requests
import sys
from pathlib import Path

firmware_path = Path(r"C:\HubVoiceSat\.esphome\build\hubvoice-sat\.pio\build\hubvoice-sat\firmware.ota.bin")
device_ip = "192.168.4.135"
upload_url = f"http://{device_ip}/update"

if not firmware_path.exists():
    print(f"ERROR: Firmware file not found at {firmware_path}")
    sys.exit(1)

print(f"Uploading firmware to http://{device_ip}")
print(f"Firmware size: {firmware_path.stat().st_size} bytes")

try:
    with open(firmware_path, 'rb') as f:
        files = {'file': f}
        response = requests.post(upload_url, files=files, timeout=120)
        print(f"Upload response: {response.status_code}")
        print(response.text[:500] if response.text else "(no response body)")
        if response.status_code == 200:
            print("\n✓ Firmware uploaded successfully!")
        else:
            print(f"\n✗ Upload failed with status {response.status_code}")
            sys.exit(1)
except requests.exceptions.Timeout:
    print("✓ Upload timed out (likely rebooting device) - this is normal!")
except requests.exceptions.ConnectionError as e:
    print(f"✓ Connection closed by device (likely rebooting) - upload probably successful!")
except Exception as e:
    print(f"ERROR: {e}")
    sys.exit(1)

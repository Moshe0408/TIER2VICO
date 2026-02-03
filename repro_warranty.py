import requests
import json
import uuid

BASE_URL = "http://localhost:8000"
EMAIL = "moshei1@verifone.com"
PASS = "123456"

s = requests.Session()

# 1. Login
print("Logging in...")
resp = s.post(f"{BASE_URL}/login", data={"email": EMAIL, "password": PASS})
print(f"Login Status: {resp.status_code}")
if resp.status_code != 200:
    print("Login failed")
    exit(1)

# 2. Get Stats
print("Fetching stats...")
resp = s.get(f"{BASE_URL}/api/stats")
print(f"Stats Status: {resp.status_code}")

try:
    data = resp.json()
    integrations = data.get("Integrations", [])
    print(f"Total Integrations: {len(integrations)}")
    
    # Check for specific warranty fields
    warranty_count = sum(1 for x in integrations if x.get("WarrantyStatus"))
    print(f"Items with WarrantyStatus: {warranty_count}")
    
    # Check sample
    if integrations:
        print("Sample Item:", json.dumps(integrations[0], indent=2, ensure_ascii=False))
        
except Exception as e:
    print(f"Error parsing JSON: {e}")
    print("Response text partial:", resp.text[:500])

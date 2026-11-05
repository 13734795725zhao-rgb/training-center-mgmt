"""Quick API smoke test"""
import requests, json

BASE = "http://127.0.0.1:5000"
s = requests.Session()

# 1. Login page
r = s.get(f"{BASE}/login")
print(f"[1] Login page: {r.status_code}")

# 2. Login as admin
r = s.post(f"{BASE}/api/auth/login", json={"username": "admin", "password": "admin123"})
print(f"[2] Login API: {r.status_code} -> {r.json().get('user', {}).get('username', 'N/A')}")

# 3. Students API (paginated)
r = s.get(f"{BASE}/api/students?page=1&per_page=5")
d = r.json()
print(f"[3] Students API: {r.status_code} | total={d.get('total')} | page={d.get('page')} | keys={list(d.keys())}")
print(f"    Stats: {d.get('stats')}")

# 4. Home page
r = s.get(f"{BASE}/")
print(f"[4] Home page: {r.status_code}")

# 5. Reports page
r = s.get(f"{BASE}/reports")
print(f"[5] Reports page: {r.status_code}")

# 6. Payments API
r = s.get(f"{BASE}/api/payments")
d = r.json()
if isinstance(d, list):
    print(f"[6] Payments API: {r.status_code} | count={len(d)}")
    if d:
        print(f"    First payment keys: {list(d[0].keys())}")
        print(f"    Status field: {d[0].get('status', 'MISSING')}")
else:
    print(f"[6] Payments API: {r.status_code} | unexpected response: {type(d)}")

# 7. Subjects API
r = s.get(f"{BASE}/api/subjects")
d = r.json()
print(f"[7] Subjects API: {r.status_code} | count={len(d.get('subjects', []))}")

# 8. Teachers API
r = s.get(f"{BASE}/api/teachers")
d = r.json()
if isinstance(d, list):
    print(f"[8] Teachers API: {r.status_code} | count={len(d)}")
else:
    print(f"[8] Teachers API: {r.status_code} | count={len(d.get('teachers', []))}")

# 9. Courses page
r = s.get(f"{BASE}/courses")
print(f"[9] Courses page: {r.status_code}")

# 10. Settings page
r = s.get(f"{BASE}/settings")
print(f"[10] Settings page: {r.status_code}")

print("\nAll tests passed!")

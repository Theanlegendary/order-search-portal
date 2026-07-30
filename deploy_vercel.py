#!/usr/bin/env python3
"""
deploy_vercel.py — sets env vars then deploys to Vercel production
"""
import subprocess
import json
import os
import sys

# Load values from parent config
config_path = os.path.join(os.path.dirname(__file__), "..", "config.json")
with open(config_path, encoding="utf-8") as f:
    cfg = json.load(f)["api"]

env_vars = {
    "API_URL":        cfg["url"],
    "API_BEARER":     cfg["bearer_token"],
    "API_BRANCH":     cfg["branch_code"],
    "API_CLIENT_ID":  cfg.get("x_client_id", "TMS_ANDROID"),
    "API_REFERER":    cfg.get("referer", "https://opsexpress.metfone.com.kh/"),
    "CACHE_TTL_SEC":  "300",
    "SEARCH_PASSWORD": "",   # leave blank = no password; set to a value to add protection
}

portal_dir = os.path.dirname(os.path.abspath(__file__))

print("=== Setting Vercel environment variables ===")
for key, val in env_vars.items():
    if not val:
        continue
    short = val[:30] + "..." if len(val) > 30 else val
    print(f"  {key} = {short}")
    # Remove old value first (ignore errors)
    subprocess.run(
        ["npx", "-y", "vercel", "env", "rm", key, "production", "--yes"],
        cwd=portal_dir, capture_output=True, shell=True
    )
    # Add new value
    result = subprocess.run(
        ["npx", "-y", "vercel", "env", "add", key, "production"],
        input=val.encode(),
        cwd=portal_dir, capture_output=True, shell=True
    )
    if result.returncode != 0:
        print(f"  [WARN] {key}: {result.stderr.decode()[:80]}")
    else:
        print(f"  [OK] {key}")

print("\n=== Deploying to Vercel production ===")
result = subprocess.run(
    ["npx", "-y", "vercel", "--prod", "--yes"],
    cwd=portal_dir, capture_output=False, shell=True
)
if result.returncode == 0:
    print("\n[SUCCESS] Deployed to Vercel!")
else:
    print("\n[ERROR] Deploy failed. Check output above.")
    sys.exit(1)

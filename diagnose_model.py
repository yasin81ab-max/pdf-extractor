#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
diagnose_model.py — عیب‌یابیِ سریعِ مدل/کلید (با VPN روشن اجرا کن).

بدونِ باز کردنِ کلِ برنامه، یک ListModels و یک generateContentِ واقعی می‌زند و
پاسخ/خطای واقعیِ Gemini را چاپ می‌کند تا در ۲ ثانیه بفهمی:
  ۱) شبکه/VPN وصل است یا نه،
  ۲) مدلِ انتخاب‌شده (مثلاً gemini-3.1-flash-lite) واقعاً برای کلیدِ تو وجود دارد یا 404 می‌دهد،
  ۳) اگر خطا هست، متنِ دقیقِ خطا چیست.

اجرا:  python diagnose_model.py
"""
import json
import sys
from pathlib import Path

import requests

# کنسولِ ویندوز پیش‌فرض cp1252 است و کاراکترِ فارسی را نمی‌پذیرد → UTF-8 اجباری
for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

CFG = Path(__file__).with_name("config.json")
BASE = "https://generativelanguage.googleapis.com/v1beta"


def main() -> int:
    cfg = json.loads(CFG.read_text(encoding="utf-8")) if CFG.exists() else {}
    keys = cfg.get("api_keys") or cfg.get("gemini_api_keys") or []
    model = cfg.get("gemini_model") or "gemini-3.1-flash-lite"
    if not keys:
        print("✗ هیچ کلیدی در config.json پیدا نشد (api_keys).")
        return 1

    key = keys[0]
    headers = {"x-goog-api-key": key, "Content-Type": "application/json"}
    print(f"model = {model}")
    print(f"keys  = {len(keys)} کلید (تستِ کلیدِ #۱)")

    # ── ۱) ListModels: آیا مدلِ انتخابی واقعاً وجود دارد؟ ──
    try:
        r = requests.get(f"{BASE}/models?pageSize=1000", headers=headers, timeout=30)
        ctype = r.headers.get("Content-Type", "")
        if r.status_code == 403 and "html" in ctype.lower():
            print("\n✗ HTTP 403 (HTML) — دسترسی به Google مسدود است. VPN را روشن کن و دوباره اجرا کن.")
            return 2
        r.raise_for_status()
        names = sorted(
            m["name"].replace("models/", "")
            for m in r.json().get("models", [])
            if "generateContent" in m.get("supportedGenerationMethods", [])
        )
        print(f"\n✓ ListModels OK — {len(names)} مدل با generateContent در دسترس است")
        flash = [n for n in names if "flash" in n or "gemini-3" in n]
        print("  flash/3.x:", ", ".join(flash) or "(هیچ)")
        exists = model in names
        print(f"  مدلِ انتخابی «{model}» موجود است؟ →", "بله ✓" if exists else "خیر ✗ (این مدل 404 می‌دهد)")
        if not exists and flash:
            print(f"  پیشنهاد: یکی از این‌ها را در تنظیمات انتخاب کن، مثلاً «{flash[0]}».")
    except requests.HTTPError as e:
        print(f"\n✗ ListModels HTTP {e.response.status_code}: {e.response.text[:250]}")
    except Exception as e:
        print(f"\n✗ ListModels error: {type(e).__name__}: {e}")

    # ── ۲) یک generateContentِ واقعی با همان مدل ──
    body = {
        "contents": [{"role": "user", "parts": [{"text": "فقط بنویس: سلام"}]}],
        "generationConfig": {"temperature": 0.1, "maxOutputTokens": 65536},
    }
    try:
        r = requests.post(f"{BASE}/models/{model}:generateContent", headers=headers, json=body, timeout=60)
        print(f"\ngenerateContent → HTTP {r.status_code}")
        if r.status_code == 200:
            txt = r.json()["candidates"][0]["content"]["parts"][0]["text"]
            print(f"✓ مدل کار می‌کند. پاسخ: {txt.strip()[:80]!r}")
        else:
            print(f"✗ متنِ دقیقِ خطا: {r.text[:400]}")
            low = r.text.lower()
            if r.status_code == 400 and ("maxoutputtokens" in low or "output token" in low):
                print("  → همان مشکلِ سقفِ توکنِ خروجی است. V16 خودکار به 8192 کم می‌کند و درست کار خواهد کرد.")
    except Exception as e:
        print(f"✗ generateContent error: {type(e).__name__}: {e}")
    return 0


if __name__ == "__main__":
    sys.exit(main())

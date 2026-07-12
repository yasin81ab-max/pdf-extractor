#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
make_drive_token.py — یک‌بار اجرا کن تا «رفرش‌توکنِ» اکانتِ اختصاصیِ جمع‌آوری ساخته شود.

⚠ مهم: با همان اکانتِ گوگلی لاگین کن که فولدرِ جمع‌آوری در آن است (اکانتِ جدا از دیتای شخصی).
خروجی: مقدارِ drive_refresh_token که باید در config (build/config_dist.json) قرار بگیرد
تا اپ بدونِ لاگینِ کاربر، فایلِ همه‌ی کاربران را در آن فولدر آپلود کند.

پیش‌نیاز: oauth_client_id و oauth_client_secret باید در config.json باشند.
اجرا:  python make_drive_token.py
"""
import json
import sys
from pathlib import Path

for _s in (sys.stdout, sys.stderr):
    try:
        _s.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

try:
    from google_auth_oauthlib.flow import InstalledAppFlow
except ImportError:
    print("✗ کتابخانه‌ی google-auth-oauthlib نصب نیست:  pip install google-auth-oauthlib")
    sys.exit(1)

SCOPES = ["https://www.googleapis.com/auth/drive"]
HERE = Path(__file__).parent
CFG = HERE / "config.json"
DIST = HERE / "build" / "config_dist.json"


def main() -> int:
    cfg = json.loads(CFG.read_text(encoding="utf-8")) if CFG.exists() else {}
    cid = (cfg.get("oauth_client_id") or "").strip()
    csec = (cfg.get("oauth_client_secret") or "").strip()
    if not cid or not csec:
        print("✗ oauth_client_id / oauth_client_secret در config.json پیدا نشد.")
        return 1

    client_config = {
        "installed": {
            "client_id": cid,
            "client_secret": csec,
            "auth_uri": "https://accounts.google.com/o/oauth2/auth",
            "token_uri": "https://oauth2.googleapis.com/token",
            "redirect_uris": ["http://localhost"],
        }
    }

    print(">> مرورگر باز می‌شود.")
    print(">> با اکانتِ اختصاصیِ جمع‌آوری لاگین کن (نه اکانتِ شخصی‌ات).\n")
    flow = InstalledAppFlow.from_client_config(client_config, SCOPES)
    # select_account: صفحه‌ی انتخابِ اکانت را مجبور کن | consent+offline: حتماً refresh_token بده
    creds = flow.run_local_server(port=0, access_type="offline", prompt="select_account consent")

    rt = getattr(creds, "refresh_token", None)
    if not rt:
        print("✗ رفرش‌توکن دریافت نشد. دوباره امتحان کن (باید prompt=consent باشد).")
        return 2

    print("\n" + "=" * 60)
    print("رفرش‌توکن دریافت شد ✓")
    print("=" * 60)

    # خودکار در config_dist.json (که نصب‌کننده به‌عنوانِ config.json می‌فرستد) نوشته می‌شود
    if DIST.exists():
        d = json.loads(DIST.read_text(encoding="utf-8"))
        d["drive_refresh_token"] = rt
        DIST.write_text(json.dumps(d, ensure_ascii=False, indent=2), encoding="utf-8")
        print(f"✓ در {DIST} ذخیره شد.")
    else:
        print(f"⚠ {DIST} پیدا نشد. این توکن را دستی زیرِ drive_refresh_token بگذار:\n{rt}")
    return 0


if __name__ == "__main__":
    sys.exit(main())

from __future__ import annotations

import argparse
import importlib.util
import json
import platform
import subprocess
import sys
from pathlib import Path

from env_utils import get_env_var


def check_module(name: str) -> dict[str, object]:
    return {"name": name, "ok": importlib.util.find_spec(name) is not None}


def check_url(url: str) -> dict[str, object]:
    try:
        from fetch_buyer_assets import fetch_url

        final_url, body, content_type = fetch_url(url, timeout=15)
        return {
            "url": url,
            "ok": True,
            "final_url": final_url,
            "content_type": content_type,
            "bytes": len(body),
        }
    except Exception as exc:
        return {"url": url, "ok": False, "error": f"{exc.__class__.__name__}: {exc}"}


def check_powerpoint_com() -> dict[str, object]:
    if platform.system().lower() != "windows":
        return {"ok": False, "reason": "not_windows"}
    command = (
        "$ErrorActionPreference='Stop'; "
        "$pp = New-Object -ComObject PowerPoint.Application; "
        "$pp.Quit(); "
        "'ok'"
    )
    try:
        result = subprocess.run(
            ["powershell", "-NoProfile", "-Command", command],
            capture_output=True,
            text=True,
            timeout=30,
        )
        return {
            "ok": result.returncode == 0,
            "stdout": result.stdout.strip(),
            "stderr": result.stderr.strip(),
        }
    except Exception as exc:
        return {"ok": False, "error": f"{exc.__class__.__name__}: {exc}"}


def build_report() -> dict[str, object]:
    return {
        "python": sys.version,
        "platform": platform.platform(),
        "openai_api_key_visible": bool(get_env_var("OPENAI_API_KEY")),
        "buyer_research_model": get_env_var("BUYER_RESEARCH_MODEL") or "gpt-4.1",
        "curl_fallback_enabled": bool(get_env_var("BUYER_BOARD_ENABLE_CURL_FALLBACK")),
        "modules": [check_module(name) for name in ("openai", "pptx", "PIL", "cairosvg")],
        "network": [
            check_url("https://www.scatec.com"),
            check_url("https://html.duckduckgo.com/html/?q=scatec"),
        ],
        "powerpoint_com": check_powerpoint_com(),
    }


def main() -> int:
    parser = argparse.ArgumentParser(description="Check buyer-board-layout runtime readiness.")
    parser.add_argument("--output", default="buyer-board-doctor-report.json", help="Diagnostic report path")
    args = parser.parse_args()

    report = build_report()
    output = Path(args.output).resolve()
    output.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8-sig")
    print(output)
    print(json.dumps(report, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

"""Installation helper for the UNYC mobile RIO scraper."""

from __future__ import annotations

import subprocess
import sys


def run(command: list[str]) -> None:
    print(">", " ".join(command))
    subprocess.run(command, check=True)


def main() -> int:
    try:
        run([sys.executable, "-m", "pip", "install", "-r", "requirements.txt"])
        run([sys.executable, "-m", "playwright", "install", "chromium"])
    except subprocess.CalledProcessError as exc:
        print(f"Installation interrompue (code {exc.returncode}).")
        return exc.returncode or 1

    print("Installation terminée.")
    print(f"Lancement : {sys.executable} unyc_automation.py")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

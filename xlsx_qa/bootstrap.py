"""Virtual environment bootstrap utilities."""

from __future__ import annotations

import os
import subprocess
import sys
from typing import Iterable

PACKAGE_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_ROOT = os.path.dirname(PACKAGE_DIR)
VENV_DIR = os.path.join(PROJECT_ROOT, "venv")
VENV_PYTHON = os.path.join(VENV_DIR, "bin", "python")
REQUIRED_PACKAGES = ("openpyxl",)


def _in_project_venv() -> bool:
    """Return True if running inside the managed virtual environment."""
    return os.path.abspath(sys.executable) == os.path.abspath(VENV_PYTHON)


def _venv_exists() -> bool:
    """Return True when the managed virtual environment already exists."""
    pyvenv_cfg = os.path.join(VENV_DIR, "pyvenv.cfg")
    return os.path.isdir(VENV_DIR) and os.path.isfile(VENV_PYTHON) and os.path.isfile(pyvenv_cfg)


def _create_venv() -> None:
    """Create the project virtual environment in-place."""
    print(f"Creating virtual environment in {VENV_DIR}...")
    subprocess.check_call([sys.executable, "-m", "venv", VENV_DIR])
    print("Virtual environment created.")


def _reexec_in_venv(argv: list[str]) -> None:
    """Replace current process with venv Python, preserving user arguments."""
    print("Switching to virtual environment...")
    os.execv(VENV_PYTHON, [VENV_PYTHON] + argv)


def _ensure_packages(packages: Iterable[str]) -> None:
    """Install required packages that are not yet importable."""
    missing: list[str] = []
    for package in packages:
        try:
            __import__(package)
        except ImportError:
            missing.append(package)

    if not missing:
        return

    print(f"Installing missing packages: {', '.join(missing)}...")
    subprocess.check_call([sys.executable, "-m", "pip", "install", "--quiet", *missing])
    print("Packages installed.")


def bootstrap(argv: list[str] | None = None) -> None:
    """Ensure the script runs inside the managed virtual environment."""
    argv = argv or sys.argv

    if not _in_project_venv():
        if not _venv_exists():
            _create_venv()
        _reexec_in_venv(argv)

    _ensure_packages(REQUIRED_PACKAGES)

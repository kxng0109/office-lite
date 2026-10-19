"""Office-lite lightweight conversion package."""

from importlib.metadata import PackageNotFoundError, version

try:
    __version__ = version("office-lite")
except PackageNotFoundError:  # pragma: no cover - importable without install
    __version__ = "0.0.0+unknown"

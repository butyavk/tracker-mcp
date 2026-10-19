"""MCP server over a job-application tracker: Excel, CSV or Google Sheets."""

from importlib.metadata import PackageNotFoundError, version

try:
    __version__ = version("tracker-mcp")
except PackageNotFoundError:  # running from a source tree without installing
    __version__ = "0.0.0"

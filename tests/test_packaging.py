"""The release files agree with each other (the release workflow checks the tag too)."""

import json
import tomllib
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]


def test_versions_and_names_match():
    version = tomllib.loads((ROOT / "pyproject.toml").read_text(encoding="utf-8"))["project"]["version"]
    server = json.loads((ROOT / "server.json").read_text(encoding="utf-8"))
    manifest = json.loads((ROOT / "mcpb" / "manifest.json").read_text(encoding="utf-8"))
    assert server["version"] == server["packages"][0]["version"] == manifest["version"] == version
    headings = [line.split()[1] for line in (ROOT / "CHANGELOG.md").read_text(encoding="utf-8").splitlines()
                if line.startswith("## ")]
    assert version in headings
    # The registry accepts the PyPI package only if its README names the server.
    assert f"mcp-name: {server['name']}" in (ROOT / "README.md").read_text(encoding="utf-8")


def test_the_plugin_runs_a_locked_published_version():
    # The plugin runs tracker-mcp from PyPI with every dependency locked. It may lag
    # behind this version, because the release workflow moves it only once PyPI serves
    # the new one. Its manifest, pin and lock must agree.
    plugin = json.loads((ROOT / "plugin" / ".claude-plugin" / "plugin.json").read_text(encoding="utf-8"))
    project = tomllib.loads((ROOT / "plugin" / "pyproject.toml").read_text(encoding="utf-8"))["project"]
    lock = tomllib.loads((ROOT / "plugin" / "uv.lock").read_text(encoding="utf-8"))
    locked = {p["name"]: p for p in lock["package"]}
    v = plugin["version"]
    assert project["version"] == v and project["dependencies"] == [f"tracker-mcp=={v}"]
    assert locked["tracker-mcp"]["version"] == v
    assert locked["tracker-mcp"]["source"] == {"registry": "https://pypi.org/simple"}
    assert plugin["mcpServers"]["tracker"]["command"] == "uv"
    assert plugin["mcpServers"]["tracker"]["args"] == [
        "run", "--project", "${CLAUDE_PLUGIN_ROOT}", "--frozen", "tracker-mcp"]
    package = tomllib.loads((ROOT / "pyproject.toml").read_text(encoding="utf-8"))["project"]["version"]
    assert tuple(map(int, v.split("."))) <= tuple(map(int, package.split(".")))
    # the directory reads files up to 256 KiB
    assert (ROOT / "plugin" / "uv.lock").stat().st_size < 256 * 1024


def test_the_extension_lists_the_real_tools():
    import anyio
    from mcp import Client

    from tracker_mcp import server

    async def names() -> set[str]:
        async with Client(server.mcp) as c:
            return {t.name for t in (await c.list_tools()).tools}

    manifest = json.loads((ROOT / "mcpb" / "manifest.json").read_text(encoding="utf-8"))
    assert {t["name"] for t in manifest["tools"]} == anyio.run(names)

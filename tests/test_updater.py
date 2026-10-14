import io
import json

from binder_builder import updater


class _Resp(io.BytesIO):
    def __enter__(self):
        return self

    def __exit__(self, *a):
        return False


def _fake(payload):
    return lambda req, timeout=0: _Resp(json.dumps(payload).encode())


def test_version_compare():
    assert updater.is_newer("v0.3.0", "0.2.9")
    assert updater.is_newer("v1.0.0", "0.9.9")
    assert not updater.is_newer("v0.2.0", "0.2.0")
    assert not updater.is_newer("v0.1.9", "0.2.0")


def test_check_for_update_newer(monkeypatch):
    monkeypatch.setattr(updater.urllib.request, "urlopen", _fake({
        "tag_name": "v9.0.0", "body": "Notes", "html_url": "https://github.com/x/y/releases/v9.0.0",
        "assets": [{"name": "BinderBuilder-9.0.0-win64.zip",
                    "browser_download_url": "https://github.com/x/y/releases/download/v9.0.0/BinderBuilder-9.0.0-win64.zip"}],
    }))
    info = updater.check_for_update("0.2.0")
    assert info.version == "9.0.0" and info.notes == "Notes" and info.asset_url.endswith(".zip")


def test_check_for_update_same_or_offline(monkeypatch):
    monkeypatch.setattr(updater.urllib.request, "urlopen", _fake({"tag_name": "v0.2.0", "assets": []}))
    assert updater.check_for_update("0.2.0") is None

    def offline(*a, **k):
        raise OSError("no network")

    monkeypatch.setattr(updater.urllib.request, "urlopen", offline)
    assert updater.check_for_update("0.2.0") is None


def test_not_frozen_cannot_self_update():
    assert updater.can_self_update() is False

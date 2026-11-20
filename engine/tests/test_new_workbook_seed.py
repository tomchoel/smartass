"""New workbooks: a copy of the seed when one carries the sensitivity label; else the hidden Excel (old path)."""
import zipfile
from pathlib import Path

import openpyxl
import pytest

from smartass.core import config
from smartass.xl import target

CUSTOM = ('<?xml version="1.0" encoding="UTF-8"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/'
          '2006/custom-properties"><property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" '
          'name="MSIP_Label_x_Enabled"><vt:lpwstr xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/'
          'docPropsVTypes">true</vt:lpwstr></property></Properties>')


def _xlsx(path: Path, custom: str | None) -> Path:
    openpyxl.Workbook().save(path)
    if custom is not None:
        src = path.with_suffix(".tmp")
        path.replace(src)
        with zipfile.ZipFile(src) as zi, zipfile.ZipFile(path, "w") as zo:
            for n in zi.namelist():
                zo.writestr(n, zi.read(n))
            zo.writestr("docProps/custom.xml", custom)
        src.unlink()
    return path


class FakeHost:
    def __init__(self):
        self.calls = 0

    def on_host(self, fn):
        self.calls += 1
        return fn()


@pytest.fixture
def env(tmp_path, monkeypatch):
    monkeypatch.setattr(config, "NEW_WORKBOOK_SEED", tmp_path / "seed.xlsx")
    monkeypatch.setattr(config, "LABEL_ID", "x")
    monkeypatch.setattr(target, "_new_path", lambda title, taken=(): tmp_path / "out" / f"{title}.xlsx")
    (tmp_path / "out").mkdir()
    return tmp_path


def test_seed_with_label_is_copied_and_host_untouched(env, monkeypatch):
    _xlsx(config.NEW_WORKBOOK_SEED, CUSTOM)
    monkeypatch.setattr(target, "_new_analysis_file_host",
                        lambda h, p: pytest.fail("hidden Excel must not be used"))
    h = FakeHost()
    p = target._new_analysis_file(h, "A")
    assert h.calls == 0
    assert p.read_bytes() == config.NEW_WORKBOOK_SEED.read_bytes()


def test_no_seed_runs_old_path_once_and_writes_seed(env, monkeypatch):
    made = []

    def fake(h, path):
        _xlsx(path, CUSTOM)
        made.append(path)
        return path
    monkeypatch.setattr(target, "_new_analysis_file_host", fake)
    h = FakeHost()
    p = target._new_analysis_file(h, "A")
    assert h.calls == 1 and made == [p]
    assert config.NEW_WORKBOOK_SEED.read_bytes() == p.read_bytes()


@pytest.mark.parametrize("custom", [None, CUSTOM.replace("MSIP_Label_x_Enabled", "Other")])
def test_seed_without_label_is_ignored(env, monkeypatch, custom):
    _xlsx(config.NEW_WORKBOOK_SEED, custom)
    monkeypatch.setattr(target, "_new_analysis_file_host", lambda h, path: _xlsx(path, CUSTOM))
    h = FakeHost()
    target._new_analysis_file(h, "A")
    assert h.calls == 1
    assert target._seed_has_label(config.NEW_WORKBOOK_SEED)   # rewritten from the labelled file


def test_seed_with_another_label_id_is_ignored(env, monkeypatch):
    _xlsx(config.NEW_WORKBOOK_SEED, CUSTOM)
    monkeypatch.setattr(config, "LABEL_ID", "Other-Id")
    assert not target._seed_has_label(config.NEW_WORKBOOK_SEED)
    monkeypatch.setattr(config, "LABEL_ID", "X")                # case-insensitive
    assert target._seed_has_label(config.NEW_WORKBOOK_SEED)


def test_two_copies_same_title_get_different_files(env, monkeypatch):
    from smartass.core import session
    monkeypatch.undo()                                          # real _new_path
    monkeypatch.setattr(config, "NEW_WORKBOOK_SEED", env / "seed.xlsx")
    monkeypatch.setattr(config, "LABEL_ID", "x")
    monkeypatch.setattr(session, "project_dir", lambda: env / "out")
    _xlsx(config.NEW_WORKBOOK_SEED, CUSTOM)
    first = target._new_path("T")
    seq = iter([first, first])                                  # both calls picked the same name before either copied
    real = target._new_path
    monkeypatch.setattr(target, "_new_path", lambda title, taken=frozenset(): next(seq, None) or real(title, taken))
    a = target._new_analysis_file(FakeHost(), "T")
    b = target._new_analysis_file(FakeHost(), "T")
    assert a != b and a.exists() and b.exists()


def test_cleanup_permission_error_does_not_escape(env, monkeypatch):
    _xlsx(config.NEW_WORKBOOK_SEED, CUSTOM)
    monkeypatch.setattr(target.shutil, "copyfile", lambda a, b: (_ for _ in ()).throw(OSError("x")))
    monkeypatch.setattr(Path, "unlink", lambda self, missing_ok=False: (_ for _ in ()).throw(PermissionError("locked")))
    monkeypatch.setattr(target, "_new_analysis_file_host", lambda h, path: path)
    assert target._new_analysis_file(FakeHost(), "A").name == "A.xlsx"


def test_failed_copy_falls_back(env, monkeypatch):
    _xlsx(config.NEW_WORKBOOK_SEED, CUSTOM)
    monkeypatch.setattr(target.shutil, "copyfile", lambda a, b: (_ for _ in ()).throw(OSError("x")))
    monkeypatch.setattr(target, "_new_analysis_file_host", lambda h, path: _xlsx(path, CUSTOM))
    h = FakeHost()
    target._new_analysis_file(h, "A")
    assert h.calls == 1


LABELINFO = ('<?xml version="1.0" encoding="utf-8" standalone="yes"?><clbl:labelList xmlns:clbl="http://schemas.'
             'microsoft.com/office/2020/mipLabelMetadata"><clbl:label id="{X}" enabled="1" method="Privileged" '
             'siteId="{S}" contentBits="2" removed="0" /></clbl:labelList>')


def test_a_seed_labelled_in_its_labelinfo_part_is_found(env, monkeypatch):
    """Phase 3c: Excel on this tenant keeps the label in docMetadata/LabelInfo.xml (Office's 2020 form), not in
    custom properties - the seed it saved has none, so it was never copied (live, 20.11.2026)."""
    seed = config.NEW_WORKBOOK_SEED
    openpyxl.Workbook().save(seed)
    src = seed.with_suffix(".tmp")
    seed.replace(src)
    with zipfile.ZipFile(src) as zi, zipfile.ZipFile(seed, "w") as zo:
        for n in zi.namelist():
            zo.writestr(n, zi.read(n))
        zo.writestr("docMetadata/LabelInfo.xml", LABELINFO)
    src.unlink()
    assert target._seed_has_label(seed)
    monkeypatch.setattr(config, "LABEL_ID", "other")
    assert not target._seed_has_label(seed)


def test_the_label_properties_are_ours(monkeypatch):
    """The sensitivity label as custom document properties (the form Excel reads when there is no LabelInfo part,
    live 20.11.2026): our label id, site, name and content bits from config, assigned as today's SetLabel assigns it
    (privileged), set now."""
    import datetime
    monkeypatch.setattr(config, "LABEL_ID", "f604")
    monkeypatch.setattr(config, "LABEL_SITE", "1676")
    monkeypatch.setattr(config, "LABEL_NAME", "Intern")
    monkeypatch.setattr(config, "LABEL_CONTENT_BITS", "0")
    when = datetime.datetime(2026, 11, 20, 16, 30, tzinfo=datetime.timezone.utc)
    p = target.label_properties(when)
    k = "MSIP_Label_f604_"
    assert {n[len(k):]: v for n, v in p.items() if n.startswith(k)} == {
        "Enabled": "true", "SetDate": "2026-11-20T16:30:00Z", "Method": "Privileged", "Name": "Intern",
        "SiteId": "1676", "ActionId": p[k + "ActionId"], "ContentBits": "0"}
    assert len(p) == 7 and len(p[k + "ActionId"]) == 36


def test_the_label_name_and_content_bits_are_configured_beside_its_id():
    """Review minor 6: another tenant's label is one config change ([label] in defaults.toml / config.toml)."""
    import tomllib
    label = tomllib.loads((Path(config.__file__).resolve().parents[2] / "defaults.toml").read_text("utf-8"))["label"]
    assert (label["name"], str(label["content_bits"])) == ("Internal", "2")
    assert config.resolve({}, {"label": label}, "SV_LABEL_NAME", "label.name", None) == "Internal"

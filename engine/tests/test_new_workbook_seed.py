"""New workbooks with no Excel running: a copy of the seed when one carries the sensitivity label; else a file built
without Excel (target._build_first_file), kept as the seed."""
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


class Builds:
    """target._build_first_file, recorded: each build writes a labelled file (or what made gives)."""

    def __init__(self, monkeypatch, made=None):
        self.calls = 0
        self.made = made or (lambda path: _xlsx(path, CUSTOM))

        def build(path):
            self.calls += 1
            return self.made(path)
        monkeypatch.setattr(target, "_build_first_file", build)


@pytest.fixture
def env(tmp_path, monkeypatch):
    monkeypatch.setattr(config, "NEW_WORKBOOK_SEED", tmp_path / "seed.xlsx")
    monkeypatch.setattr(config, "LABEL_ID", "x")
    monkeypatch.setattr(target, "_new_path", lambda title, taken=(): tmp_path / "out" / f"{title}.xlsx")
    (tmp_path / "out").mkdir()
    return tmp_path


def test_seed_with_label_is_copied_and_nothing_built(env, monkeypatch):
    _xlsx(config.NEW_WORKBOOK_SEED, CUSTOM)
    b = Builds(monkeypatch, made=lambda p: pytest.fail("a seed with our label is copied, nothing built"))
    p = target._new_analysis_file("A")
    assert b.calls == 0
    assert p.read_bytes() == config.NEW_WORKBOOK_SEED.read_bytes()


def test_no_seed_builds_the_file_once_and_writes_seed(env, monkeypatch):
    b = Builds(monkeypatch)
    p = target._new_analysis_file("A")
    assert b.calls == 1 and p.exists()
    assert config.NEW_WORKBOOK_SEED.read_bytes() == p.read_bytes()


@pytest.mark.parametrize("custom", [None, CUSTOM.replace("MSIP_Label_x_Enabled", "Other")])
def test_seed_without_label_is_ignored(env, monkeypatch, custom):
    _xlsx(config.NEW_WORKBOOK_SEED, custom)
    b = Builds(monkeypatch)
    target._new_analysis_file("A")
    assert b.calls == 1
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
    a = target._new_analysis_file("T")
    b = target._new_analysis_file("T")
    assert a != b and a.exists() and b.exists()


def test_cleanup_permission_error_does_not_escape(env, monkeypatch):
    _xlsx(config.NEW_WORKBOOK_SEED, CUSTOM)
    monkeypatch.setattr(target.shutil, "copyfile", lambda a, b: (_ for _ in ()).throw(OSError("x")))
    monkeypatch.setattr(Path, "unlink", lambda self, missing_ok=False: (_ for _ in ()).throw(PermissionError("locked")))
    Builds(monkeypatch, made=lambda path: path)
    assert target._new_analysis_file("A").name.startswith("A")


def test_failed_copy_falls_back(env, monkeypatch):
    _xlsx(config.NEW_WORKBOOK_SEED, CUSTOM)
    monkeypatch.setattr(target.shutil, "copyfile", lambda a, b: (_ for _ in ()).throw(OSError("x")))
    b = Builds(monkeypatch)
    target._new_analysis_file("A")
    assert b.calls == 1


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


def _parts(path: Path) -> dict[str, str]:
    with zipfile.ZipFile(path) as z:
        return {n: z.read(n).decode("utf-8", "replace") for n in z.namelist() if n.endswith(".xml")}


def test_the_first_file_is_built_without_excel_labelled_with_its_sources_log(env, monkeypatch):
    """Task 6, cold path (b): no hidden Excel - the file is built by render/xlsx (the writer of the 0.5.0 new
    workbook): one furnished Sources sheet, our label as custom document properties, the stored base and the
    template's theme when they are known."""
    from smartass.render.xlsx import Base
    from smartass.xl import styling
    monkeypatch.setattr(config, "LABEL_SITE", "1676")
    monkeypatch.setattr(target, "new_base", lambda: Base("Aptos Narrow", 11.0, 20, 64))
    monkeypatch.setattr(styling, "template_theme", lambda: None)
    p = target._build_first_file(env / "out" / "First.xlsx")
    parts = _parts(p)
    assert 'name="Sources"' in parts["xl/workbook.xml"] and len(openpyxl.load_workbook(p).sheetnames) == 1
    assert "MSIP_Label_x_Enabled" in parts["docProps/custom.xml"] and "MSIP_Label_x_SiteId" in parts["docProps/custom.xml"]
    assert "Aptos Narrow" in parts["xl/styles.xml"]
    assert target._seed_has_label(p)


def test_without_a_label_configured_the_first_file_carries_no_label(env, monkeypatch):
    """No label configured (LABEL_ID empty): the built file has no label properties, and no seed is ever taken
    (_seed_has_label is False without a label id, as before). Before Task 6 the hidden Excel called Office's label API
    with the empty id instead: what that did was not observed (unverified)."""
    from smartass.xl import styling
    monkeypatch.setattr(config, "LABEL_ID", "")
    monkeypatch.setattr(target, "new_base", lambda: None)
    monkeypatch.setattr(styling, "template_theme", lambda: None)
    p = target._new_analysis_file("A")
    assert "MSIP_Label" not in _parts(p).get("docProps/custom.xml", "") and not target._seed_has_label(p)
    assert 'name="Sources"' in _parts(p)["xl/workbook.xml"]


def test_with_no_excel_no_seed_and_no_stored_base_the_user_excel_is_launched_on_the_template(env, monkeypatch):
    """Task 6, cold path (b), the route that needs no stored base (Task 5 report (ii')): an Excel launched on the TN
    template as for a parallel agent, its workbook saved as (Office's label API in the user's Excel, as any new
    workbook added in a running Excel); from_template, so its look stores the base for the next new workbook."""
    from smartass.xl import jobs, styling
    tpl = env / "tn_template.xltx"
    tpl.write_bytes(b"x")
    monkeypatch.setattr(styling, "template_path", lambda: tpl)
    monkeypatch.setattr(target, "new_base", lambda: None)
    monkeypatch.setattr(jobs, "hold_turn", lambda *a, **k: True)
    book = object()
    spawned = []
    monkeypatch.setattr(target, "_spawn_excel", lambda t, known: spawned.append((t, known)) or ((5, 50, "app"), book))
    saved = []
    monkeypatch.setattr(target, "_save_new", lambda app, wb, title: saved.append((app, wb, title)) or env / "x.xlsx")
    monkeypatch.setattr(target, "_new_analysis_file", lambda title: pytest.fail("the template route builds no file"))
    t = target._new_target(None, "Q", [], set(), build_new=True)
    assert spawned == [(tpl, set())] and saved == [("app", book, "Q")]
    assert t.wb is book and t.from_template and t.new and t.created and t.made is None

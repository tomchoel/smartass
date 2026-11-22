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
    """target._build_first_file, recorded: each build writes a labelled file (or what made gives); based: it was built
    in a stored base (what _build_first_file returns)."""

    def __init__(self, monkeypatch, made=None, based=True):
        self.calls = 0
        self.made = made or (lambda path: _xlsx(path, CUSTOM))

        def build(path):
            self.calls += 1
            self.made(path)
            return based
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
    p = env / "out" / "First.xlsx"
    assert target._build_first_file(p) is True                  # built in the stored base
    parts = _parts(p)
    assert 'name="Sources"' in parts["xl/workbook.xml"] and len(openpyxl.load_workbook(p).sheetnames) == 1
    custom = parts["docProps/custom.xml"]
    assert "MSIP_Label_x_Enabled" in custom and "MSIP_Label_x_SiteId" in custom
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
    monkeypatch.setattr(target, "_spawn_excel",
                        lambda t, known, limit=None: spawned.append((t, known, limit)) or ((5, 50, "app"), book))
    saved = []
    monkeypatch.setattr(target, "_save_new", lambda app, wb, title: saved.append((app, wb, title)) or env / "x.xlsx")
    monkeypatch.setattr(target, "_new_analysis_file", lambda title: pytest.fail("the template route builds no file"))
    t = target._new_target("Q", [], set(), build_new=True)
    assert spawned == [(tpl, set(), target.OPEN_TIMEOUT_S)] and saved == [("app", book, "Q")]   # a cold Excel's time
    assert t.wb is book and t.from_template and t.new and t.created and t.made is None


# ---- Task 6b: the seed is made only from a file built in a stored base, and written whole

def test_a_first_file_built_without_a_stored_base_is_not_kept_as_the_seed(env, monkeypatch):
    """Without a stored base the built file is a plain Calibri workbook: kept as the seed, every later cold start would
    copy it for good (the template route, whose look stores the base, would never run again)."""
    b = Builds(monkeypatch, based=False)
    p = target._new_analysis_file("A")
    assert b.calls == 1 and p.exists() and not config.NEW_WORKBOOK_SEED.exists()
    target._new_analysis_file("B")
    assert b.calls == 2 and not config.NEW_WORKBOOK_SEED.exists()


@pytest.mark.parametrize("col_px, based", [(None, False), (63, False), (64, True)], ids=["no base", "odd", "stored"])
def test_the_first_file_says_whether_it_was_built_in_the_stored_base(env, monkeypatch, col_px, based):
    """The base actually used: none stored, or one whose column width is not a standard one (dropped), is not."""
    from smartass.render.xlsx import Base
    from smartass.xl import styling
    monkeypatch.setattr(target, "new_base", lambda: None if col_px is None else Base("Aptos Narrow", 11.0, 20, col_px))
    monkeypatch.setattr(styling, "template_theme", lambda: None)
    assert target._build_first_file(env / "out" / "First.xlsx") is based
    p = target._new_analysis_file("A")
    assert p.exists() and config.NEW_WORKBOOK_SEED.exists() is based


def test_the_seed_is_written_through_a_temp_file_of_its_own_and_replaced(env, monkeypatch):
    """Two cold jobs can write the seed at once: each copies to a temp file of its own beside the seed, then
    os.replace - a reader never sees a torn seed."""
    copies = []
    real = target.shutil.copyfile
    monkeypatch.setattr(target.shutil, "copyfile", lambda a, b: copies.append(Path(b)) or real(a, b))
    Builds(monkeypatch)
    a = target._new_analysis_file("A")
    first = config.NEW_WORKBOOK_SEED.read_bytes()
    config.NEW_WORKBOOK_SEED.unlink()
    target._new_analysis_file("B")
    seed = config.NEW_WORKBOOK_SEED
    assert len(copies) == 2 and len(set(copies)) == 2
    assert all(c.parent == seed.parent and c != seed and not c.exists() for c in copies)
    assert first == a.read_bytes() and target._seed_has_label(seed)


def test_a_seed_replace_windows_refuses_for_a_moment_is_tried_again(env, monkeypatch):
    """Replacing a seed that was just read (_seed_has_label) can be refused for a moment on Windows (WinError 5: the
    full suite saw it once): tried again, as a delete is (DELETE_TRIES)."""
    _xlsx(config.NEW_WORKBOOK_SEED, None)                      # a seed without our label: rebuilt and replaced
    labelled = _xlsx(env / "labelled.xlsx", CUSTOM)             # made before os.replace is faked (_xlsx uses it)
    Builds(monkeypatch, made=lambda p: p.write_bytes(labelled.read_bytes()))
    calls, real = [], target.os.replace

    def replace(a, b):
        calls.append(b)
        if len(calls) <= 2:
            raise PermissionError(13, "Access is denied", str(b))
        real(a, b)
    monkeypatch.setattr(target.os, "replace", replace)
    monkeypatch.setattr(target.time, "sleep", lambda s: None)
    target._new_analysis_file("A")
    assert len(calls) == 3 and target._seed_has_label(config.NEW_WORKBOOK_SEED)
    assert sorted(x.name for x in env.iterdir()) == ["labelled.xlsx", "out", "seed.xlsx"]


def test_a_failed_seed_write_leaves_no_temp_file(env, monkeypatch):
    Builds(monkeypatch, made=lambda p: p.write_bytes(b"built"))
    monkeypatch.setattr(target.os, "replace", lambda a, b: (_ for _ in ()).throw(PermissionError("locked")))
    monkeypatch.setattr(target.time, "sleep", lambda s: None)
    p = target._new_analysis_file("A")
    assert p.exists() and not config.NEW_WORKBOOK_SEED.exists()
    assert sorted(x.name for x in env.iterdir()) == ["out"]


# ---- Task 9 (final review B, M4): the cold first workbook gets open_new's read-only check

@pytest.mark.parametrize("read_only", [True, False])
def test_a_cold_first_workbook_excel_opened_read_only_is_a_clear_error(env, monkeypatch, read_only):
    """Excel opens a file whose sensitivity label it does not know read-only, and its Save then writes nothing,
    without an error (live, 20.11.2026): as open_new refuses that, so does the cold route - never an answer 'saved'
    with nothing written."""
    _xlsx(config.NEW_WORKBOOK_SEED, CUSTOM)
    Builds(monkeypatch, made=lambda p: pytest.fail("a seed with our label is copied"))

    class Book:
        ReadOnly = read_only
    opened = []

    def open_path(path, excels, insts=None):
        opened.append(path)
        return target.Target("app", Book(), 7, 70)
    monkeypatch.setattr(target, "_open_path", open_path)
    if read_only:
        with pytest.raises(target.HostError, match="read-only") as e:
            target._new_target("A", [], set())
        assert opened[0].name in str(e.value) and str(opened[0]) in str(e.value)
        assert opened[0].exists()                         # nothing closed or deleted: the text says what to do
    else:
        t = target._new_target("A", [], set())
        assert t.created and t.new and t.wb.ReadOnly is False

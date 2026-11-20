"""Building the add-in (xl/inproc/build.py) with a fake compiler: the hash and names, the atomic build, pruning, the
compile-error marker."""
import os
import shutil
import time
from pathlib import Path
from types import SimpleNamespace

import pytest

from smartass.core import config
from smartass.xl.inproc import batch, build


@pytest.fixture
def src(tmp_path, monkeypatch):
    """Copies of the sources (Runner.cs, the vendored files) and a home of its own; csc 'exists'."""
    s = tmp_path / "src"
    (s / "vendor").mkdir(parents=True)
    (s / "Runner.cs").write_text('// [ExcelFunction(Name = "SmartassRun___HASH__")] const string Hash = "__HASH__"; '
                                 'const string JobsFolder = @"__JOBS__";', encoding="utf-8")
    (s / "vendor" / "ExcelDna64.xll").write_bytes(b"MZ xll \x00\x01")
    (s / "vendor" / "ExcelDna.Integration.dll").write_bytes(b"MZ dll \x00\x02")
    monkeypatch.setattr(build, "RUNNER_CS", s / "Runner.cs")
    monkeypatch.setattr(build, "XLL", s / "vendor" / "ExcelDna64.xll")
    monkeypatch.setattr(build, "INTEGRATION", s / "vendor" / "ExcelDna.Integration.dll")
    csc = tmp_path / "csc.exe"
    csc.write_bytes(b"")
    monkeypatch.setattr(build, "CSC", csc)
    monkeypatch.setattr(config, "INPROC_DIR", tmp_path / "home" / "inproc")
    return s


@pytest.fixture
def compiler(monkeypatch):
    """A fake csc: records each command line and the source it was given, writes the -out: dll."""
    seen = []

    def compile_(cmd, cwd):
        out = Path(next(a for a in cmd if a.startswith("-out:"))[5:])
        seen.append((cmd, Path(cmd[-1]).read_text(encoding="utf-8"), cwd))
        out.write_bytes(b"compiled")
        return 0, ""
    monkeypatch.setattr(build, "_compile", compile_)
    return seen


def test_the_hash_covers_the_sources_and_the_batch_version(src, monkeypatch):
    h = build.hash8()
    assert len(h) == 8 and int(h, 16) >= 0
    assert build.hash8() == h
    for f in ("Runner.cs", "vendor/ExcelDna64.xll", "vendor/ExcelDna.Integration.dll"):
        p = src / f
        before = p.read_bytes()
        p.write_bytes(before + b" ")
        assert build.hash8() != h, f
        p.write_bytes(before)
    monkeypatch.setattr(batch, "VERSION", batch.VERSION + 1)
    assert build.hash8() != h


@pytest.mark.parametrize("part", ["DNA", "CONFIG", "FLAGS", "INPROC_DIR"])
def test_the_hash_covers_the_dna_the_config_the_flags_and_the_jobs_folder(src, monkeypatch, tmp_path, part):
    h = build.hash8()
    if part == "FLAGS":
        monkeypatch.setattr(build, "FLAGS", (*build.FLAGS, "-debug-"))
    elif part == "INPROC_DIR":
        monkeypatch.setattr(config, "INPROC_DIR", tmp_path / "elsewhere")
    else:
        monkeypatch.setattr(build, part, getattr(build, part) + "<!-- -->")
    assert build.hash8() != h


def test_line_endings_do_not_change_the_hash(src):
    cs = src / "Runner.cs"
    cs.write_bytes(b"line one\nline two\n")
    h = build.hash8()
    cs.write_bytes(b"line one\r\nline two\r\n")
    assert build.hash8() == h


def test_a_build_is_a_folder_named_by_its_hash_with_the_add_in_its_dna_config_and_dll(src, compiler):
    xll, h = build.ensure()
    name = f"smartass_run_{h}"
    folder = config.INPROC_DIR / name
    assert xll == folder / f"{name}.xll"
    assert xll.read_bytes() == (src / "vendor" / "ExcelDna64.xll").read_bytes()
    assert (folder / f"{name}.dll").read_bytes() == b"compiled"
    dna = (folder / f"{name}.dna").read_text(encoding="utf-8")
    assert f'Path="{name}.dll"' in dna and 'LoadFromBytes="true"' in dna and 'RuntimeVersion="v4.0"' in dna
    cfg = (folder / f"{name}.xll.config").read_text(encoding="utf-8")
    assert 'name="ExcelDna.Integration" switchValue="Off"' in cfg
    assert '<add name="LogDisplay"' in cfg and 'initializeData="Off"' in cfg
    (cmd, source, cwd), = compiler
    jobs = build.jobs_folder()
    assert source == (f'// [ExcelFunction(Name = "SmartassRun_{h}")] const string Hash = "{h}"; '
                      f'const string JobsFolder = @"{jobs}";')
    assert jobs == config.INPROC_DIR / "jobs" and jobs.is_absolute()
    assert cmd[0] == str(build.CSC)
    assert {"-nologo", "-target:library", "-langversion:5", f"-r:{build.INTEGRATION}",
            f"-r:{build.CSC.parent / 'System.Web.Extensions.dll'}"} <= set(cmd)
    refs = [a[3:] for a in cmd if a.startswith("-r:")]
    assert len(refs) == 2 and all(Path(r).is_absolute() for r in refs)    # final review S1: nothing from elsewhere
    assert Path(cwd) == Path(cmd[-1]).parent != Path.cwd()               # csc runs in the temporary build folder
    assert sorted(p.name for p in config.INPROC_DIR.iterdir()) == [name]       # no temporary folder left
    assert build.ensure() == (xll, h) and len(compiler) == 1                     # built once


def test_a_build_finished_meanwhile_by_another_job_is_used(src, monkeypatch):
    h = build.hash8()
    name = f"smartass_run_{h}"

    def other_job_wins(cmd, cwd):
        out = Path(next(a for a in cmd if a.startswith("-out:"))[5:])
        out.write_bytes(b"mine")
        shutil.copytree(out.parent, config.INPROC_DIR / name)        # the other job renamed its build into place
        (config.INPROC_DIR / name / f"{name}.dll").write_bytes(b"theirs")
        return 0, ""
    monkeypatch.setattr(build, "_compile", other_job_wins)
    xll, _ = build.ensure()
    assert (xll.parent / f"{name}.dll").read_bytes() == b"theirs"
    assert sorted(p.name for p in config.INPROC_DIR.iterdir()) == [name]


def test_older_builds_are_pruned_unless_their_add_in_is_locked(src, compiler):
    root = config.INPROC_DIR
    for old in ("smartass_run_00000001", "smartass_run_00000002"):
        (root / old).mkdir(parents=True)
        for ext in (".xll", ".dna", ".dll", ".xll.config"):
            (root / old / f"{old}{ext}").write_bytes(b"x")
    (root / "smartass_run_00000003.failed").write_text("old error", encoding="utf-8")
    locked = open(root / "smartass_run_00000002" / "smartass_run_00000002.xll", "rb")   # loaded in some Excel
    try:
        xll, h = build.ensure()
    finally:
        locked.close()
    left = sorted(p.name for p in root.iterdir())
    assert left == sorted([f"smartass_run_{h}", "smartass_run_00000002"])
    assert len(list((root / "smartass_run_00000002").iterdir())) == 4              # kept whole


def test_a_compile_error_is_kept_in_a_marker_and_not_retried(src, monkeypatch):
    calls = []

    def fails(cmd, cwd):
        calls.append(cmd)
        return 1, "Runner.cs(3,1): error CS1002: ; expected"
    monkeypatch.setattr(build, "_compile", fails)
    for _ in range(2):
        with pytest.raises(build.BuildError, match="CS1002"):
            build.ensure()
    assert len(calls) == 1
    marker = config.INPROC_DIR / f"smartass_run_{build.hash8()}.failed"
    assert "CS1002" in marker.read_text(encoding="utf-8")
    assert sorted(p.name for p in config.INPROC_DIR.iterdir()) == [marker.name]


def test_a_compile_error_marker_expires_after_a_day(src, compiler):
    marker = config.INPROC_DIR / f"smartass_run_{build.hash8()}.failed"
    marker.parent.mkdir(parents=True)
    marker.write_text("csc was killed by the virus scanner", encoding="utf-8")
    with pytest.raises(build.BuildError, match="virus"):
        build.ensure()
    os.utime(marker, (time.time() - 2 * 86400, time.time() - 2 * 86400))
    xll, _ = build.ensure()
    assert xll.is_file() and not marker.exists() and len(compiler) == 1


def test_no_compiler_is_a_build_error_without_a_marker(src, compiler, tmp_path, monkeypatch):
    monkeypatch.setattr(build, "CSC", tmp_path / "nowhere" / "csc.exe")
    with pytest.raises(build.BuildError, match="compiler"):
        build.ensure()
    assert compiler == []
    assert not config.INPROC_DIR.exists() or not list(config.INPROC_DIR.glob("*.failed"))


def test_the_compiler_runs_in_the_given_folder_without_a_window(monkeypatch, tmp_path):
    """Final review S1: csc looks for a reference that is not fully qualified in its working directory first: that is
    the temporary build folder, never wherever the server runs."""
    seen = {}

    def run(cmd, **kw):
        seen.update(kw, cmd=cmd)
        return SimpleNamespace(returncode=0, stdout="ok", stderr="")
    monkeypatch.setattr(build.subprocess, "run", run)
    assert build._compile(["csc.exe", "x.cs"], tmp_path) == (0, "ok")
    assert seen["cwd"] == tmp_path and seen["creationflags"] == build.subprocess.CREATE_NO_WINDOW
    assert seen["cmd"] == ["csc.exe", "x.cs"] and seen["timeout"] == build.COMPILE_LIMIT_S


def test_the_shipped_sources_are_there():
    for p in (build.RUNNER_CS, build.XLL, build.INTEGRATION, build.XLL.parent / "LICENSE-ExcelDna.txt"):
        assert p.is_file(), p
    cs = build.RUNNER_CS.read_text(encoding="utf-8")
    assert 'Name = "SmartassRun___HASH__"' in cs and 'Name = "SmartassPing___HASH__"' in cs
    assert 'const string JobsFolder = @"__JOBS__";' in cs

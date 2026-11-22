"""The fresh-home proof (light-fast phase 4 Task 3): scripts/check_fresh_home.py makes a home and a user profile in a
temp folder of its own, runs `smartass setup`'s steps there and then `smartass doctor`, each in a fresh Python with
its fakes (empty process list, a kill and a registry write as tripwires, "not set" registry readers, an EPM stub) and
fails closed when any place a stage can write is outside that folder. Here it runs whole (a child process: the
suite's own guards do not reach it - its tripwires do), and its judges are checked on made-up doctor results."""
import importlib.util
import os
import subprocess
import sys
from pathlib import Path

SCRIPT = Path(__file__).resolve().parents[2] / "scripts" / "check_fresh_home.py"
NOT_SET_UP = "EPM access is not set up ... add them to config.toml, then run `smartass login`."
LOGIN_STEP = "Run `smartass login` in a terminal (it signs you in to EPM, then builds the atlas)."


def _script():
    spec = importlib.util.spec_from_file_location("check_fresh_home", SCRIPT)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def _checks(**over):
    base = {"config": ("OK", "", ""), "excel": ("OK", "", ""), "smartview_startup": ("OK", "", ""),
            "epm": ("OK", "", ""), "oauth": ("WARN", "plan1: not signed in", "Run `smartass login` (one approval)."),
            "atlas": ("FAIL", "no atlas on this PC yet (x: missing)", LOGIN_STEP),
            "knowledge": ("WARN", "9 metrics; members not checked: waits for the atlas",
                          LOGIN_STEP[:-1] + "; the knowledge is checked against it after."),
            "output": ("OK", "", ""), "leftovers": ("OK", "", "")}
    base.update(over)
    return {"checks": [{"name": n, "status": s, "detail": d, "fix": f} for n, (s, d, f) in base.items()],
            "not_set_up": NOT_SET_UP}


def test_the_judge_takes_only_the_sign_in_and_the_atlas_with_one_command():
    m = _script()
    assert m.judge(_checks()) == []
    assert m.judge(_checks(output=("WARN", "x", "y")))                                # anything else: a problem
    assert m.judge(_checks(knowledge=("WARN", "waits for the atlas", "Run `smartass atlas build`.")))   # not the same
    assert m.judge(_checks(knowledge=("FAIL", "2 build error(s)", "Fix the semantic files.")))
    two = "Run `smartass login` or `smartass atlas build`."
    assert m.judge(_checks(atlas=("FAIL", "no atlas on this PC yet", two)))
    assert m.judge(_checks(oauth=("FAIL", "not set up", NOT_SET_UP))) != []


def test_before_the_oauth_step_the_judge_wants_the_set_up_text():
    m = _script()
    assert m.judge_before_oauth(_checks(oauth=("FAIL", "not set up", NOT_SET_UP))) == []
    assert m.judge_before_oauth(_checks()) != []                                       # a plain WARN is not it
    assert m.readme_names_the_oauth_step_before_login()


def test_a_stage_refuses_when_a_place_it_could_write_is_outside_its_temp_folder(tmp_path):
    env = {**os.environ, "SMARTASS_HOME": str(tmp_path / "elsewhere")}
    r = subprocess.run([sys.executable, str(SCRIPT), "--stage", "setup", "--root", str(tmp_path / "root")],
                       env=env, capture_output=True, text=True, encoding="utf-8", timeout=120)
    assert r.returncode == 2 and "REFUSED - outside the temp folder" in r.stdout and "home:" in r.stdout
    assert not (tmp_path / "elsewhere" / "config.toml").exists()


def test_a_fresh_home_comes_up_with_only_the_sign_in_and_the_atlas_to_do(monkeypatch, capsys):
    """The driver runs here, its stages in fresh Pythons as by hand, with Excel pointed at a file that exists: the
    suite runs on PCs without desktop Excel (CI). Run by hand, the script keeps the real check (EXCEL.EXE looked for
    on this PC) and says so."""
    m = _script()
    env = m._env
    monkeypatch.setattr(m, "_env", lambda root: {**env(root), "SV_EXCEL_EXE": sys.executable})
    assert m.main([]) == 0
    out = capsys.readouterr().out
    assert "OK: a fresh home comes up with only the sign-in and the atlas to do" in out
    assert "PROBLEM" not in out and "untouched outside the temp folder: 5 of 5" in out
    assert "epm check (a stub)" in out and "process list (empty)" in out
    excel = [line for line in out.splitlines() if line.split()[1:2] == ["excel"]]
    assert len(excel) == 2 and all(line.split()[0] == "OK" and sys.executable in line for line in excel), excel

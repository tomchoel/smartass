# engine/tests/v2/test_v2_autosave.py
from smartass import pipeline


class WB:
    def __init__(self, on):
        self.AutoSaveOn = on
        self.Name = "x.xlsx"
        self.log = []

    def Save(self):
        self.log.append(("save", self.AutoSaveOn))


class NoAutoSaveWB:
    """An Excel without AutoSave: reading the property raises."""
    Name = "x.xlsx"

    @property
    def AutoSaveOn(self):
        raise RuntimeError("no autosave")


class T:
    def __init__(self, on, ours, wb=None):
        self.wb = wb if wb is not None else WB(on)
        self.created = ours
        self.path = "C:/analyser/v2/p/x.xlsx" if ours else "C:/Users/me/own.xlsx"


def test_our_file_autosave_off_while_writing_then_back_on_and_it_saves_the_file(monkeypatch):
    """AutoSave on before the job: switched back on afterwards, it saves our file itself - no Save call."""
    monkeypatch.setattr(pipeline, "_is_ours", lambda t: t.created)
    t = T(True, ours=True)
    with pipeline.autosave_paused(t) as notes:
        assert t.wb.AutoSaveOn is False and notes.autosave is True
        pipeline._save(t, notes.autosave)
    assert t.wb.log == []
    assert t.wb.AutoSaveOn is True and notes == []


def test_our_file_without_autosave_is_saved(monkeypatch):
    monkeypatch.setattr(pipeline, "_is_ours", lambda t: t.created)
    t = T(False, ours=True)
    with pipeline.autosave_paused(t) as notes:
        assert notes.autosave is False
        pipeline._save(t, notes.autosave)
    assert t.wb.log == [("save", False)] and t.wb.AutoSaveOn is False


def test_the_users_workbook_is_never_saved(monkeypatch):
    monkeypatch.setattr(pipeline, "_is_ours", lambda t: t.created)
    for on in (True, False):
        t = T(on, ours=False)
        with pipeline.autosave_paused(t) as notes:
            pipeline._save(t, notes.autosave)
        assert t.wb.log == []


def test_user_workbook_autosave_left_off_with_a_note(monkeypatch):
    monkeypatch.setattr(pipeline, "_is_ours", lambda t: t.created)
    t = T(True, ours=False)
    with pipeline.autosave_paused(t) as notes:
        assert t.wb.AutoSaveOn is False
    assert t.wb.AutoSaveOn is False
    assert notes and "AutoSave was switched off" in notes[0]


def test_autosave_already_off_or_unsupported_is_quiet(monkeypatch):
    monkeypatch.setattr(pipeline, "_is_ours", lambda t: t.created)
    t = T(False, ours=False)
    with pipeline.autosave_paused(t) as notes:
        pass
    assert notes == []

    b = T(True, ours=True, wb=NoAutoSaveWB())
    with pipeline.autosave_paused(b) as notes:
        pass
    assert notes == []

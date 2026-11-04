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


def test_our_file_autosave_off_while_writing_then_restored(monkeypatch):
    monkeypatch.setattr(pipeline, "_is_ours", lambda t: t.created)
    t = T(True, ours=True)
    with pipeline.autosave_paused(t) as notes:
        assert t.wb.AutoSaveOn is False
        pipeline._save(t)
    assert t.wb.log == [("save", False)]
    assert t.wb.AutoSaveOn is True and notes == []


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

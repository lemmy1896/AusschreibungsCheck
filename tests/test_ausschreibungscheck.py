import datetime as dt
import io
import json
import sys
import zipfile
from pathlib import Path

from openpyxl import load_workbook

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import ausschreibungscheck as ac  # noqa: E402

PROFIL = Path(__file__).resolve().parents[1] / "profil_ifs.yaml"


def release(ocid, titel, beschreibung="", cpv="71322000", region="DE929", plattform="evergabe-online.de"):
    return {
        "ocid": ocid,
        "date": "2026-10-12T08:00:00Z",
        "parties": [{
            "id": "b1", "name": "Stadt Hannover", "roles": ["buyer"],
            "address": {"postalCode": "30159", "locality": "Hannover", "region": region},
        }],
        "tender": {
            "title": titel,
            "description": beschreibung,
            "items": [{"classification": {"scheme": "CPV", "id": cpv}}],
            "tenderPeriod": {"endDate": "2099-11-05T12:00:00+01:00"},
            "submissionMethodDetails": f"https://www.{plattform}/tenderdetails.html?id=123",
        },
    }


def export_zip(releases):
    puffer = io.BytesIO()
    with zipfile.ZipFile(puffer, "w") as zf:
        zf.writestr("export.json", json.dumps({"releases": releases}))
    return puffer.getvalue()


def test_bewertung_trifft_fachthema():
    profil = ac.Profil.laden(PROFIL)
    b = ac.aus_release(release("1", "Generalentwässerungsplanung Stadt Laatzen", "hydraulische Berechnung des Kanalnetzes"))
    t = ac.bewerten(b, profil)
    assert t is not None
    assert "Generalentwässerungsplanung" in t.gruende
    assert "CPV 71322000" in t.gruende
    assert "Region DE9" in t.gruende
    assert b.link.startswith("https://www.evergabe-online.de/")
    assert b.ort == "30159 Hannover"
    assert b.frist == "2099-11-05 12:00"


def test_fremdes_thema_und_ausschluss():
    profil = ac.Profil.laden(PROFIL)
    assert ac.bewerten(ac.aus_release(release("2", "Lieferung von Büromöbeln", cpv="39130000")), profil) is None
    assert ac.bewerten(ac.aus_release(release("3", "Kanalreinigung Regenwasserkanalnetz")), profil) is None


def test_abkuerzung_nur_als_wort():
    profil = ac.Profil.laden(PROFIL)
    b = ac.aus_release(release("4", "GEPÄCKBÄNDER Flughafen", cpv="42000000", region="DE1"))
    assert ac.bewerten(b, profil) is None
    b = ac.aus_release(release("5", "Fortschreibung GEP und Fremdwasserkonzept", cpv="42000000", region="DE1"))
    assert ac.bewerten(b, profil) is not None


def test_plattformfilter():
    profil = ac.Profil.laden(PROFIL)
    assert ac.passt_plattform(ac.aus_release(release("6", "x")), profil)
    assert not ac.passt_plattform(ac.aus_release(release("7", "x", plattform="vergabe.niedersachsen.de")), profil)


def test_tage_nachholen():
    heute = dt.date(2026, 10, 13)
    assert len(ac.zu_pruefende_tage({"letzter_tag": None}, heute, None)) == 30
    assert ac.zu_pruefende_tage({"letzter_tag": "2026-10-10"}, heute, None) == [dt.date(2026, 10, 11), dt.date(2026, 10, 12)]
    assert ac.zu_pruefende_tage({"letzter_tag": "2026-10-12"}, heute, None) == []


def test_abgelaufene_frist_wird_ausgelassen(tmp_path, monkeypatch):
    (tmp_path / "profil_ifs.yaml").write_text(PROFIL.read_text(encoding="utf-8"), encoding="utf-8")
    alt = release("alt", "Regenklärbecken Süd")
    alt["tender"]["tenderPeriod"]["endDate"] = "2020-01-01T10:00:00+01:00"
    zipdatei = tmp_path / "x.zip"
    zipdatei.write_bytes(export_zip([alt, release("neu", "Regenklärbecken Nord")]))
    monkeypatch.setattr(ac, "basisordner", lambda: tmp_path)
    assert ac.main(["--zip", str(zipdatei), "--nicht-oeffnen"]) == 0
    ws = load_workbook(next((tmp_path / "ergebnisse").glob("*.xlsx")))["Treffer"]
    assert [z[2] for z in ws.iter_rows(min_row=2, values_only=True)] == ["Regenklärbecken Nord"]


def test_gesamtlauf(tmp_path, monkeypatch):
    (tmp_path / "profil_ifs.yaml").write_text(PROFIL.read_text(encoding="utf-8"), encoding="utf-8")
    zipdatei = tmp_path / "2026-10-12.zip"
    zipdatei.write_bytes(export_zip([
        release("a", "Retentionsbodenfilter Volksparkstraße, Objektplanung LPH 1-6"),
        release("b", "Abflussmessungen Dattelner Mühlenbach", cpv="71351000", region="DEA36"),
        release("c", "Regenrückhaltebecken Nord", plattform="vergabe.niedersachsen.de"),
        release("d", "Reinigung Verwaltungsgebäude", cpv="90910000"),
    ]))
    monkeypatch.setattr(ac, "basisordner", lambda: tmp_path)
    assert ac.main(["--zip", str(zipdatei), "--nicht-oeffnen"]) == 0
    listen = list((tmp_path / "ergebnisse").glob("*.xlsx"))
    assert len(listen) == 1
    ws = load_workbook(listen[0])["Treffer"]
    zeilen = list(ws.iter_rows(min_row=2, values_only=True))
    assert {z[2] for z in zeilen} == {"Retentionsbodenfilter Volksparkstraße, Objektplanung LPH 1-6",
                                      "Abflussmessungen Dattelner Mühlenbach"}
    assert all(z[0] == "neu" for z in zeilen)

    # zweiter Lauf: dieselben Treffer sind nicht mehr neu
    for f in listen:
        f.unlink()
    assert ac.main(["--zip", str(zipdatei), "--nicht-oeffnen"]) == 0
    ws = load_workbook(next((tmp_path / "ergebnisse").glob("*.xlsx")))["Treffer"]
    assert all(z[0] is None for z in ws.iter_rows(min_row=2, values_only=True))


EFORMS = b"""<?xml version="1.0" encoding="UTF-8"?>
<ContractNotice xmlns="urn:oasis:names:specification:ubl:schema:xsd:ContractNotice-2"
  xmlns:cac="urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2"
  xmlns:cbc="urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2">
  <cac:ProcurementProjectLot><cac:TenderingProcess>
    <cac:TenderSubmissionDeadlinePeriod><cbc:EndDate>2026-11-20+01:00</cbc:EndDate><cbc:EndTime>10:00:00+01:00</cbc:EndTime></cac:TenderSubmissionDeadlinePeriod>
  </cac:TenderingProcess></cac:ProcurementProjectLot>
  <cac:ProcurementProjectLot><cac:TenderingProcess>
    <cac:TenderSubmissionDeadlinePeriod><cbc:EndDate>2026-11-05+01:00</cbc:EndDate><cbc:EndTime>12:00:00+01:00</cbc:EndTime></cac:TenderSubmissionDeadlinePeriod>
  </cac:TenderingProcess></cac:ProcurementProjectLot>
</ContractNotice>"""


def test_frist_aus_eforms():
    puffer = io.BytesIO()
    with zipfile.ZipFile(puffer, "w") as zf:
        zf.writestr("abc-01.xml", EFORMS)
    assert ac.fristen_aus_zip(puffer.getvalue()) == {"abc-01": "2026-11-05T12:00"}

    r = release("x", "Regenklärbecken")
    del r["tender"]["tenderPeriod"]
    puffer = io.BytesIO()
    with zipfile.ZipFile(puffer, "w") as zf:
        zf.writestr("abc-01.json", json.dumps(r))
    b = ac.aus_release(ac.releases_aus_zip(puffer.getvalue(), {"abc-01": "2026-11-05T12:00"})[0])
    assert b.frist == "2026-11-05 12:00"


def test_plattform_nur_aus_unterlagenlink():
    profil = ac.Profil.laden(PROFIL)
    r = release("y", "x", plattform="vergabe.vmstart.de")
    r["parties"][0]["contactPoint"] = {"url": "https://www.evergabe-online.de"}
    assert not ac.passt_plattform(ac.aus_release(r), profil)


def test_plattformen_nach_host():
    profil = ac.Profil.laden(PROFIL)
    def plattform(url):
        r = release("p", "x")
        r["tender"]["submissionMethodDetails"] = url
        return ac.plattform_von(ac.aus_release(r), profil.plattformen)
    assert plattform("https://www.evergabe.de/unterlagen/1") == "evergabe.de"
    assert plattform("https://www.deutsche-evergabe.de/dashboards/x") is None
    assert plattform("https://www.dtvp.de/Satellite/notice/X/documents") == "DTVP"
    assert plattform("https://www.subreport.de/E33255198") == "subreport ELViS"
    assert plattform("https://www.meinauftrag.rib.de/public/x") == "RIB meinauftrag"
    assert plattform("https://www.evergabe-online.de/tenderdetails.html?id=1") == "e-Vergabe des Bundes"


def test_zusatzbegriffe_lesen():
    text = "Starkregen\nKlärschlamm: 25\n  # Kommentar\n\nÜberflutung = 3"
    assert ac.zusatzbegriffe_lesen(text, 10) == {"Starkregen": 10, "Klärschlamm": 25, "Überflutung": 3}


def test_auswahl_und_zusatzbegriffe(tmp_path):
    profil = tmp_path / "profil_ifs.yaml"
    profil.write_text(PROFIL.read_text(encoding="utf-8"), encoding="utf-8")
    zipdatei = tmp_path / "x.zip"
    zipdatei.write_bytes(export_zip([
        release("a", "Regenklärbecken Nord", plattform="dtvp.de"),
        release("b", "Regenklärbecken Süd"),
        release("c", "Klärschlammverwertung", cpv="90513000", region="DE1"),
        release("d", "Retentionsbodenfilter", plattform="vergabe.niedersachsen.de"),
    ]))

    def titel(excel):
        return sorted(z[2] for z in load_workbook(excel)["Treffer"].iter_rows(min_row=2, values_only=True))

    excel = ac.ausfuehren(profil, tmp_path, plattformen=["DTVP"], zips=[zipdatei], log=lambda _: None)
    assert titel(excel) == ["Regenklärbecken Nord"]
    excel = ac.ausfuehren(profil, tmp_path, zusatzbegriffe={"Klärschlamm": 10}, zips=[zipdatei], log=lambda _: None)
    assert titel(excel) == ["Klärschlammverwertung", "Regenklärbecken Nord", "Regenklärbecken Süd"]
    excel = ac.ausfuehren(profil, tmp_path, plattformen=[], alle_plattformen=True, zips=[zipdatei], log=lambda _: None)
    assert titel(excel) == ["Regenklärbecken Nord", "Regenklärbecken Süd", "Retentionsbodenfilter"]
    plattformen = {z[2]: z[3] for z in load_workbook(excel)["Treffer"].iter_rows(min_row=2, values_only=True)}
    assert plattformen == {"Regenklärbecken Nord": "DTVP", "Regenklärbecken Süd": "e-Vergabe des Bundes",
                           "Retentionsbodenfilter": None}


def test_fenster_startet(tmp_path):
    import pytest
    tk = pytest.importorskip("tkinter")
    import gui
    try:
        root = tk.Tk()
    except tk.TclError:
        pytest.skip("keine Anzeige verfügbar")
    profil = tmp_path / "profil_ifs.yaml"
    profil.write_text(PROFIL.read_text(encoding="utf-8"), encoding="utf-8")
    fenster = gui.Fenster(root, tmp_path, profil)
    fenster.begriffe_text.insert("1.0", "Starkregen: 7")
    plattformen, alle, begriffe = fenster.auswahl()
    assert len(plattformen) == 5 and not alle and begriffe == "Starkregen: 7"
    root.update()
    root.destroy()


def test_abruf_wiederholt_bei_abweisung(tmp_path, monkeypatch):
    import urllib.error
    aufrufe = []

    class Antwort:
        def __enter__(self):
            return self
        def __exit__(self, *a):
            return False
        def read(self):
            return b"zipdaten"

    def urlopen(anfrage, timeout):
        aufrufe.append(anfrage.full_url)
        if len(aufrufe) < 3:
            raise urllib.error.URLError(ConnectionRefusedError(111, "Connection refused"))
        return Antwort()

    monkeypatch.setattr(ac.urllib.request, "urlopen", urlopen)
    monkeypatch.setattr(ac.time, "sleep", lambda s: None)
    meldungen = []
    assert ac.tagesexport(dt.date(2026, 10, 1), tmp_path, log=meldungen.append) == b"zipdaten"
    assert len(aufrufe) == 3 and len(meldungen) == 2
    assert (tmp_path / "2026-10-01.ocds.zip").read_bytes() == b"zipdaten"

    aufrufe.clear()
    monkeypatch.setattr(ac.urllib.request, "urlopen",
                        lambda a, timeout: (_ for _ in ()).throw(urllib.error.URLError("weg")))
    import pytest
    with pytest.raises(ac.AbrufFehler):
        ac.tagesexport(dt.date(2026, 10, 2), tmp_path, log=meldungen.append)

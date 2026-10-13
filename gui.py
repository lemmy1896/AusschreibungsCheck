"""Programmfenster für AusschreibungsCheck."""

from __future__ import annotations

import datetime as dt
import json
import queue
import threading
import tkinter as tk
import traceback
from pathlib import Path
from tkinter import messagebox, scrolledtext, ttk

import ausschreibungscheck as ac

ERSTER_EXPORTTAG = dt.date(2023, 10, 25)  # Start des Datenservice Öffentlicher Einkauf


class Einstellungen:
    def __init__(self, pfad: Path):
        self.pfad = pfad
        self.daten: dict = {}
        try:
            self.daten = json.loads(pfad.read_text(encoding="utf-8"))
        except (OSError, ValueError):
            pass

    def speichern(self) -> None:
        self.pfad.parent.mkdir(exist_ok=True)
        self.pfad.write_text(json.dumps(self.daten, ensure_ascii=False, indent=1), encoding="utf-8")


def datum_lesen(text: str) -> dt.date | None:
    for format_ in ("%d.%m.%Y", "%d.%m.%y", "%Y-%m-%d"):
        try:
            return dt.datetime.strptime(text.strip(), format_).date()
        except ValueError:
            pass
    return None


class Fenster:
    def __init__(self, root: tk.Tk, basis: Path, profilpfad: Path):
        self.root = root
        self.basis = basis
        self.profilpfad = profilpfad
        self.profil = ac.Profil.laden(profilpfad)
        self.einstellungen = Einstellungen(basis / "daten" / "einstellungen.json")
        self.meldungen: queue.Queue = queue.Queue()
        self.letzte_liste: Path | None = None
        self.begriffe: dict[str, int] = dict(self.profil.begriffe)
        self._alte_zusatzbegriffe_uebernehmen()

        root.title(f"AusschreibungsCheck – {self.profil.name}")
        root.geometry("860x800")
        root.minsize(700, 640)
        haupt = ttk.Frame(root, padding=12)
        haupt.pack(fill="both", expand=True)
        haupt.columnconfigure(0, weight=1)
        haupt.columnconfigure(1, weight=1)

        self._plattformen_bauen(haupt)
        self._zeitraum_bauen(haupt)
        self._schlagwoerter_bauen(haupt)

        knoepfe = ttk.Frame(haupt)
        knoepfe.grid(row=2, column=0, columnspan=2, sticky="ew")
        self.start_knopf = ttk.Button(knoepfe, text="Suche starten", command=self.starten)
        self.start_knopf.pack(side="left")
        self.oeffnen_knopf = ttk.Button(knoepfe, text="Ergebnisliste öffnen", command=self.liste_oeffnen,
                                        state="disabled")
        self.oeffnen_knopf.pack(side="left", padx=8)
        ttk.Button(knoepfe, text="Ergebnisordner", command=self.ordner_oeffnen).pack(side="left")
        self.fortschritt = ttk.Progressbar(knoepfe, mode="indeterminate", length=160)
        self.fortschritt.pack(side="right")

        self.protokoll = scrolledtext.ScrolledText(haupt, height=8, state="disabled", font=("Consolas", 9))
        self.protokoll.grid(row=3, column=0, columnspan=2, sticky="nsew", pady=(10, 0))
        haupt.rowconfigure(1, weight=3)
        haupt.rowconfigure(3, weight=1)

        self.root.after(100, self.meldungen_abholen)

    # ------------------------------------------------------------------ Aufbau

    def _plattformen_bauen(self, haupt) -> None:
        rahmen = ttk.LabelFrame(haupt, text="Plattformen", padding=8)
        rahmen.grid(row=0, column=0, sticky="nsew", padx=(0, 6))
        gewaehlt = self.einstellungen.daten.get("plattformen")
        self.plattform_vars: dict[str, tk.BooleanVar] = {}
        for name in self.profil.plattformen:
            var = tk.BooleanVar(value=gewaehlt is None or name in gewaehlt)
            ttk.Checkbutton(rahmen, text=name, variable=var).pack(anchor="w")
            self.plattform_vars[name] = var
        ttk.Separator(rahmen).pack(fill="x", pady=6)
        self.alle_var = tk.BooleanVar(value=self.einstellungen.daten.get("alle_plattformen", False))
        ttk.Checkbutton(rahmen, text="Alle Plattformen (auch nicht aufgeführte)",
                        variable=self.alle_var).pack(anchor="w")

    def _zeitraum_bauen(self, haupt) -> None:
        rahmen = ttk.LabelFrame(haupt, text="Zeitraum", padding=8)
        rahmen.grid(row=0, column=1, sticky="nsew", padx=(6, 0))
        self.zeitraum_var = tk.StringVar(value=self.einstellungen.daten.get("zeitraum", "neu"))
        ttk.Radiobutton(rahmen, text="Seit dem letzten Lauf (beim ersten Mal 30 Tage)", value="neu",
                        variable=self.zeitraum_var, command=self._zeitraum_umschalten).pack(anchor="w")
        zeile = ttk.Frame(rahmen)
        zeile.pack(anchor="w", pady=(4, 0))
        ttk.Radiobutton(zeile, text="Ab Datum:", value="datum", variable=self.zeitraum_var,
                        command=self._zeitraum_umschalten).pack(side="left")
        vorgabe = self.einstellungen.daten.get(
            "ab_datum", (dt.date.today() - dt.timedelta(days=ac.ERSTLAUF_TAGE)).strftime("%d.%m.%Y"))
        self.datum_var = tk.StringVar(value=vorgabe)
        self.datum_feld = ttk.Entry(zeile, textvariable=self.datum_var, width=12)
        self.datum_feld.pack(side="left", padx=(4, 0))
        ttk.Label(zeile, text="TT.MM.JJJJ", foreground="#555").pack(side="left", padx=6)
        ttk.Label(rahmen, foreground="#555", text=(
            "Ein früheres Datum lohnt sich nach geänderten\nSchlagwörtern oder Plattformen. "
            "Bereits geladene\nTage kommen aus dem Zwischenspeicher.")).pack(anchor="w", pady=(6, 0))
        self._zeitraum_umschalten()

    def _schlagwoerter_bauen(self, haupt) -> None:
        rahmen = ttk.LabelFrame(haupt, text="Schlagwörter", padding=8)
        rahmen.grid(row=1, column=0, columnspan=2, sticky="nsew", pady=10)
        rahmen.columnconfigure(0, weight=1)
        rahmen.rowconfigure(1, weight=1)
        ttk.Label(rahmen, foreground="#555", text=(
            "Gesucht wird in Titel und Beschreibung. Die Punkte aller gefundenen Schlagwörter, CPV-Codes und "
            f"Regionen werden addiert; ab {self.profil.mindestpunkte} Punkten erscheint eine Ausschreibung in "
            "der Liste.\nAbkürzungen in Großbuchstaben (z. B. GEP) zählen nur als ganzes Wort, alles andere "
            "auch als Wortteil. Änderungen werden sofort im Profil gespeichert.")).grid(row=0, column=0, columnspan=2, sticky="w")

        liste = ttk.Frame(rahmen)
        liste.grid(row=1, column=0, sticky="nsew", pady=(6, 0))
        liste.columnconfigure(0, weight=1)
        liste.rowconfigure(0, weight=1)
        self.tabelle = ttk.Treeview(liste, columns=("wort", "punkte"), show="headings", height=10)
        self.tabelle.heading("wort", text="Schlagwort", command=lambda: self._sortieren("wort"))
        self.tabelle.heading("punkte", text="Punkte", command=lambda: self._sortieren("punkte"))
        self.tabelle.column("wort", width=320)
        self.tabelle.column("punkte", width=70, anchor="e")
        self.tabelle.grid(row=0, column=0, sticky="nsew")
        leiste = ttk.Scrollbar(liste, orient="vertical", command=self.tabelle.yview)
        leiste.grid(row=0, column=1, sticky="ns")
        self.tabelle.configure(yscrollcommand=leiste.set)
        self.tabelle.bind("<<TreeviewSelect>>", self._auswahl_uebernehmen)
        self.tabelle.bind("<Delete>", lambda _: self.begriff_entfernen())

        bearbeiten = ttk.Frame(rahmen, padding=(10, 6, 0, 0))
        bearbeiten.grid(row=1, column=1, sticky="n")
        ttk.Label(bearbeiten, text="Schlagwort").grid(row=0, column=0, sticky="w")
        self.wort_var = tk.StringVar()
        wort_feld = ttk.Entry(bearbeiten, textvariable=self.wort_var, width=28)
        wort_feld.grid(row=1, column=0, columnspan=2, sticky="ew")
        wort_feld.bind("<Return>", lambda _: self.begriff_speichern())
        ttk.Label(bearbeiten, text="Punkte").grid(row=2, column=0, sticky="w", pady=(6, 0))
        self.punkte_var = tk.StringVar(value=str(self.profil.mindestpunkte))
        ttk.Spinbox(bearbeiten, from_=1, to=100, textvariable=self.punkte_var, width=6).grid(
            row=3, column=0, sticky="w")
        ttk.Button(bearbeiten, text="Hinzufügen / Punkte ändern", command=self.begriff_speichern).grid(
            row=4, column=0, columnspan=2, sticky="ew", pady=(10, 0))
        ttk.Button(bearbeiten, text="Entfernen", command=self.begriff_entfernen).grid(
            row=5, column=0, columnspan=2, sticky="ew", pady=(4, 0))
        self.anzahl_label = ttk.Label(bearbeiten, foreground="#555")
        self.anzahl_label.grid(row=6, column=0, columnspan=2, sticky="w", pady=(10, 0))
        self._sortierung = ("punkte", True)
        self._tabelle_fuellen()

    # ------------------------------------------------------------------ Schlagwörter

    def _alte_zusatzbegriffe_uebernehmen(self) -> None:
        """Frühere Version hatte ein separates Feld für zusätzliche Schlagwörter; die wandern ins Profil."""
        text = self.einstellungen.daten.pop("zusatzbegriffe", "")
        zusatz = ac.zusatzbegriffe_lesen(text, self.profil.mindestpunkte)
        if zusatz:
            self.begriffe.update(zusatz)
            ac.begriffe_speichern(self.profilpfad, self.begriffe)
        self.einstellungen.speichern()

    def _tabelle_fuellen(self, markieren: str | None = None) -> None:
        spalte, absteigend = self._sortierung
        if spalte == "punkte":
            eintraege = sorted(self.begriffe.items(), key=lambda e: (-e[1], e[0].lower()))
            if not absteigend:
                eintraege.reverse()
        else:
            eintraege = sorted(self.begriffe.items(), key=lambda e: e[0].lower(), reverse=absteigend)
        self.tabelle.delete(*self.tabelle.get_children())
        for wort, punkte in eintraege:
            self.tabelle.insert("", "end", iid=wort, values=(wort, punkte))
        if markieren and self.tabelle.exists(markieren):
            self.tabelle.selection_set(markieren)
            self.tabelle.see(markieren)
        self.anzahl_label.configure(text=f"{len(self.begriffe)} Schlagwörter")

    def _sortieren(self, spalte: str) -> None:
        alt_spalte, absteigend = self._sortierung
        self._sortierung = (spalte, not absteigend if spalte == alt_spalte else spalte == "punkte")
        self._tabelle_fuellen()

    def _auswahl_uebernehmen(self, _=None) -> None:
        auswahl = self.tabelle.selection()
        if auswahl:
            self.wort_var.set(auswahl[0])
            self.punkte_var.set(str(self.begriffe[auswahl[0]]))

    def begriff_speichern(self) -> None:
        wort = " ".join(self.wort_var.get().split())
        if not wort:
            messagebox.showinfo("AusschreibungsCheck", "Bitte ein Schlagwort eingeben.")
            return
        try:
            punkte = int(self.punkte_var.get())
            if punkte < 1:
                raise ValueError
        except ValueError:
            messagebox.showinfo("AusschreibungsCheck", "Punkte bitte als ganze Zahl ab 1 eingeben.")
            return
        self.begriffe[wort] = punkte
        self._begriffe_sichern(f"„{wort}“ mit {punkte} Punkten gespeichert.")
        self._tabelle_fuellen(markieren=wort)

    def begriff_entfernen(self) -> None:
        auswahl = list(self.tabelle.selection())
        if not auswahl:
            wort = self.wort_var.get().strip()
            auswahl = [wort] if wort in self.begriffe else []
        if not auswahl:
            messagebox.showinfo("AusschreibungsCheck", "Bitte zuerst ein Schlagwort in der Liste auswählen.")
            return
        for wort in auswahl:
            self.begriffe.pop(wort, None)
        self.wort_var.set("")
        self._begriffe_sichern("Entfernt: " + ", ".join(f"„{w}“" for w in auswahl))
        self._tabelle_fuellen()

    def _begriffe_sichern(self, meldung: str) -> None:
        try:
            ac.begriffe_speichern(self.profilpfad, self.begriffe)
            self.schreiben(meldung)
        except OSError as fehler:
            messagebox.showerror("AusschreibungsCheck", f"Das Profil konnte nicht gespeichert werden:\n{fehler}")

    # ------------------------------------------------------------------ Ablauf

    def _zeitraum_umschalten(self) -> None:
        self.datum_feld.configure(state="normal" if self.zeitraum_var.get() == "datum" else "disabled")

    def auswahl(self) -> tuple[list[str], bool]:
        return [n for n, v in self.plattform_vars.items() if v.get()], self.alle_var.get()

    def startdatum(self) -> tuple[bool, dt.date | None]:
        """(gültig, Datum); Datum None heißt: seit dem letzten Lauf."""
        if self.zeitraum_var.get() != "datum":
            return True, None
        datum = datum_lesen(self.datum_var.get())
        gestern = dt.date.today() - dt.timedelta(days=1)
        if datum is None:
            messagebox.showwarning("AusschreibungsCheck", "Bitte das Datum als TT.MM.JJJJ eingeben, z. B. 01.09.2026.")
            return False, None
        if datum > gestern:
            messagebox.showwarning("AusschreibungsCheck", f"Das Datum darf höchstens {gestern:%d.%m.%Y} sein "
                                   "(der heutige Export ist noch nicht vollständig).")
            return False, None
        if datum < ERSTER_EXPORTTAG:
            messagebox.showwarning("AusschreibungsCheck", f"Daten gibt es erst ab {ERSTER_EXPORTTAG:%d.%m.%Y}.")
            return False, None
        tage = (gestern - datum).days + 1
        if tage > 90 and not messagebox.askyesno(
                "AusschreibungsCheck", f"Ab {datum:%d.%m.%Y} sind {tage} Tage zu laden. Noch nicht "
                "zwischengespeicherte Tage brauchen jeweils einige Sekunden. Trotzdem starten?"):
            return False, None
        return True, datum

    def starten(self) -> None:
        plattformen, alle = self.auswahl()
        if not plattformen and not alle:
            messagebox.showwarning("AusschreibungsCheck", "Bitte mindestens eine Plattform auswählen.")
            return
        gueltig, von = self.startdatum()
        if not gueltig:
            return
        self.einstellungen.daten.update(plattformen=plattformen, alle_plattformen=alle,
                                        zeitraum=self.zeitraum_var.get(), ab_datum=self.datum_var.get())
        self.einstellungen.speichern()

        self.start_knopf.configure(state="disabled")
        self.fortschritt.start(12)
        self.schreiben("—" * 40)
        if von:
            self.schreiben(f"Prüfe Ausschreibungen ab {von:%d.%m.%Y}.")
        threading.Thread(target=self.suche, args=(plattformen, alle, von), daemon=True).start()

    def suche(self, plattformen, alle, von) -> None:
        try:
            excel = ac.ausfuehren(self.profilpfad, self.basis, plattformen=plattformen, alle_plattformen=alle,
                                  von=von, log=self.meldungen.put)
            self.meldungen.put(("fertig", excel))
        except Exception:  # noqa: BLE001
            self.meldungen.put(("fehler", traceback.format_exc()))

    def meldungen_abholen(self) -> None:
        try:
            while True:
                meldung = self.meldungen.get_nowait()
                if isinstance(meldung, tuple):
                    art, wert = meldung
                    self.fortschritt.stop()
                    self.start_knopf.configure(state="normal")
                    if art == "fertig":
                        self.letzte_liste = wert
                        self.oeffnen_knopf.configure(state="normal")
                        self.schreiben("Fertig.")
                        ac.oeffnen(wert)
                    else:
                        self.schreiben(wert)
                        messagebox.showerror("AusschreibungsCheck", "Bei der Suche ist ein Fehler aufgetreten. "
                                             "Details stehen im Protokoll.")
                else:
                    self.schreiben(str(meldung))
        except queue.Empty:
            pass
        self.root.after(100, self.meldungen_abholen)

    def schreiben(self, text: str) -> None:
        self.protokoll.configure(state="normal")
        self.protokoll.insert("end", text + "\n")
        self.protokoll.see("end")
        self.protokoll.configure(state="disabled")

    def liste_oeffnen(self) -> None:
        if self.letzte_liste:
            ac.oeffnen(self.letzte_liste)

    def ordner_oeffnen(self) -> None:
        ordner = self.basis / "ergebnisse"
        ordner.mkdir(exist_ok=True)
        ac.oeffnen(ordner)


def main() -> None:
    basis = ac.basisordner()
    root = tk.Tk()
    profilpfad = ac.profil_finden(basis)
    if profilpfad is None:
        root.withdraw()
        messagebox.showerror("AusschreibungsCheck", f"Keine Profil-Datei (profil_*.yaml) in\n{basis}\ngefunden.")
        return
    try:
        Fenster(root, basis, profilpfad)
    except Exception as fehler:  # noqa: BLE001
        root.withdraw()
        messagebox.showerror("AusschreibungsCheck", f"Das Profil {profilpfad.name} konnte nicht gelesen werden:\n{fehler}")
        return
    root.mainloop()


if __name__ == "__main__":
    main()

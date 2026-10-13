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

ERSTLAUF_HINWEIS = "Seit dem letzten Lauf (beim ersten Mal 30 Tage)"


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


class Fenster:
    def __init__(self, root: tk.Tk, basis: Path, profilpfad: Path):
        self.root = root
        self.basis = basis
        self.profilpfad = profilpfad
        self.profil = ac.Profil.laden(profilpfad)
        self.einstellungen = Einstellungen(basis / "daten" / "einstellungen.json")
        self.meldungen: queue.Queue = queue.Queue()
        self.letzte_liste: Path | None = None

        root.title(f"AusschreibungsCheck – {self.profil.name}")
        root.geometry("760x640")
        root.minsize(620, 520)
        haupt = ttk.Frame(root, padding=12)
        haupt.pack(fill="both", expand=True)
        haupt.columnconfigure(0, weight=1)
        haupt.columnconfigure(1, weight=1)

        # Plattformen
        rahmen_p = ttk.LabelFrame(haupt, text="Plattformen", padding=8)
        rahmen_p.grid(row=0, column=0, sticky="nsew", padx=(0, 6))
        gewaehlt = self.einstellungen.daten.get("plattformen")
        self.plattform_vars: dict[str, tk.BooleanVar] = {}
        for name in self.profil.plattformen:
            var = tk.BooleanVar(value=gewaehlt is None or name in gewaehlt)
            ttk.Checkbutton(rahmen_p, text=name, variable=var).pack(anchor="w")
            self.plattform_vars[name] = var
        ttk.Separator(rahmen_p).pack(fill="x", pady=6)
        self.alle_var = tk.BooleanVar(value=self.einstellungen.daten.get("alle_plattformen", False))
        ttk.Checkbutton(rahmen_p, text="Alle Plattformen (auch nicht aufgeführte)",
                        variable=self.alle_var).pack(anchor="w")

        # Zeitraum
        rahmen_z = ttk.LabelFrame(haupt, text="Zeitraum", padding=8)
        rahmen_z.grid(row=0, column=1, sticky="nsew", padx=(6, 0))
        self.zeitraum_var = tk.StringVar(value="neu")
        ttk.Radiobutton(rahmen_z, text=ERSTLAUF_HINWEIS, value="neu",
                        variable=self.zeitraum_var).pack(anchor="w")
        ttk.Radiobutton(rahmen_z, text="Letzte 30 Tage erneut prüfen", value="30",
                        variable=self.zeitraum_var).pack(anchor="w")
        ttk.Label(rahmen_z, text="Erneut prüfen lohnt sich nach geänderten\nSchlagwörtern oder Plattformen.",
                  foreground="#555").pack(anchor="w", pady=(6, 0))

        # Schlagwörter
        rahmen_s = ttk.LabelFrame(haupt, text="Zusätzliche Schlagwörter", padding=8)
        rahmen_s.grid(row=1, column=0, columnspan=2, sticky="nsew", pady=10)
        rahmen_s.columnconfigure(0, weight=1)
        ttk.Label(rahmen_s, foreground="#555", text=(
            f"Ein Schlagwort pro Zeile. Ohne Angabe zählt es {self.profil.mindestpunkte} Punkte "
            "(reicht allein für einen Treffer).\nMit Gewicht: „Starkregen: 5“. "
            f"Das Profil enthält bereits {len(self.profil.begriffe)} Begriffe.")).grid(row=0, column=0, sticky="w")
        self.begriffe_text = tk.Text(rahmen_s, height=6, wrap="none", font=("Segoe UI", 10))
        self.begriffe_text.grid(row=1, column=0, sticky="nsew", pady=(6, 0))
        self.begriffe_text.insert("1.0", self.einstellungen.daten.get("zusatzbegriffe", ""))

        # Knöpfe
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

        # Protokoll
        self.protokoll = scrolledtext.ScrolledText(haupt, height=12, state="disabled", font=("Consolas", 9))
        self.protokoll.grid(row=3, column=0, columnspan=2, sticky="nsew", pady=(10, 0))
        haupt.rowconfigure(3, weight=1)

        self.root.after(100, self.meldungen_abholen)

    # ------------------------------------------------------------------ Ablauf

    def auswahl(self) -> tuple[list[str], bool, str]:
        plattformen = [n for n, v in self.plattform_vars.items() if v.get()]
        return plattformen, self.alle_var.get(), self.begriffe_text.get("1.0", "end").strip()

    def starten(self) -> None:
        plattformen, alle, begriffe = self.auswahl()
        if not plattformen and not alle:
            messagebox.showwarning("AusschreibungsCheck", "Bitte mindestens eine Plattform auswählen.")
            return
        self.einstellungen.daten.update(plattformen=plattformen, alle_plattformen=alle, zusatzbegriffe=begriffe)
        self.einstellungen.speichern()
        zusatz = ac.zusatzbegriffe_lesen(begriffe, self.profil.mindestpunkte)
        von = dt.date.today() - dt.timedelta(days=ac.ERSTLAUF_TAGE) if self.zeitraum_var.get() == "30" else None

        self.start_knopf.configure(state="disabled")
        self.fortschritt.start(12)
        self.schreiben("—" * 40)
        if zusatz:
            self.schreiben("Zusätzliche Schlagwörter: " + ", ".join(f"{b} ({p})" for b, p in zusatz.items()))
        threading.Thread(target=self.suche, args=(plattformen, alle, zusatz, von), daemon=True).start()

    def suche(self, plattformen, alle, zusatz, von) -> None:
        try:
            excel = ac.ausfuehren(self.profilpfad, self.basis, plattformen=plattformen, alle_plattformen=alle,
                                  zusatzbegriffe=zusatz, von=von, log=self.meldungen.put)
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

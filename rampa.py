import re
import io
import pandas as pd
import streamlit as st
from openpyxl import load_workbook

st.set_page_config(page_title="Rampenverteilung für Touren", page_icon="🚚")
st.title("🚚 Rampenverteilung für Touren")
st.write("Laden Sie Ihre Excel-Datei mit den Touren hoch, und die App verteilt sie automatisch auf die Rampen.")

# ==== EINSTELLUNGEN (in der Seitenleiste) ====
st.sidebar.header("Einstellungen")
BROJ_RAMPI = st.sidebar.number_input("Anzahl der Rampen", min_value=1, max_value=20, value=5)
MAX_TURA_PO_RAMPI = st.sidebar.number_input("Max. Touren pro Rampe", min_value=1, max_value=10, value=2)
TEZINA_UDIO = st.sidebar.slider(
    "Gewichtungsanteil beim Ausgleich (0 = nur Anzahl der Bestellungen, 1 = nur Gewicht)",
    min_value=0.0, max_value=1.0, value=0.5, step=0.1
)
trazeni_brojevi_input = st.sidebar.text_input(
    "Zahlen im Blattnamen, die eine Schicht kennzeichnen (durch Komma getrennt)",
    value="6, 10, 14, 16"
)
TRAZENI_BROJEVI = [b.strip() for b in trazeni_brojevi_input.split(",") if b.strip()]

uploaded_file = st.file_uploader("Excel-Datei hochladen (.xlsx)", type=["xlsx"])


# ==== HILFSFUNKTIONEN ====

def sheet_odgovara(sheet_name, brojevi):
    for broj in brojevi:
        pattern = r'(?<!\d)' + re.escape(broj) + r'(?!\d)'
        if re.search(pattern, sheet_name):
            return True
    return False


def pronadji_header_redak(raw_df):
    for i, row in raw_df.iterrows():
        if row.astype(str).str.contains("Tour", case=False, na=False).any():
            return i
    return None


def pronadji_kolonu(kolone, kljucna_rijec):
    for kol in kolone:
        if kljucna_rijec.lower() in kol.lower():
            return kol
    return None


def obradi_list(file_bytes, sheet_name, broj_rampi, max_tura_po_rampi, tezina_udio, log):
    raw = pd.read_excel(file_bytes, sheet_name=sheet_name, header=None)
    header_idx = pronadji_header_redak(raw)

    if header_idx is None:
        log.append(f"⏭️ Übersprungen '{sheet_name}': keine Spalte 'Tour' gefunden, vermutlich kein relevantes Blatt.")
        return None

    file_bytes.seek(0)
    df = pd.read_excel(file_bytes, sheet_name=sheet_name, header=header_idx)
    df.columns = df.columns.astype(str).str.strip()

    kol_isell = pronadji_kolonu(df.columns, "iSell")
    kol_gew = pronadji_kolonu(df.columns, "Gew")
    kol_tour = pronadji_kolonu(df.columns, "Tour")

    nedostaje = [naziv for naziv, kol in
                 [("iSell", kol_isell), ("Gew", kol_gew), ("Tour", kol_tour)]
                 if kol is None]

    if nedostaje:
        log.append(f"⏭️ Übersprungen '{sheet_name}': fehlende Spalten mit {nedostaje}. "
                    f"Gefundene Spalten: {df.columns.tolist()}")
        return None

    df = df.rename(columns={kol_isell: "iSell", kol_gew: "Gew", kol_tour: "Tour"})
    df = df.dropna(subset=["Tour"]).copy()

    if df.empty:
        log.append(f"⏭️ Übersprungen '{sheet_name}': keine Daten nach der Bereinigung.")
        return None

    # Tournummern als Text, ohne Dezimalstellen
    df["Tour"] = df["Tour"].astype(int).astype(str)

    tour_summary = (
        df.groupby("Tour")
          .agg(broj_narudzbi=("iSell", "count"), ukupna_tezina=("Gew", "sum"))
          .reset_index()
    )

    broj_tura = len(tour_summary)
    max_tura_po_rampi_lokalno = max_tura_po_rampi
    if broj_tura > broj_rampi * max_tura_po_rampi:
        max_tura_po_rampi_lokalno = -(-broj_tura // broj_rampi)  # ceil
        log.append(f"⚠️ '{sheet_name}': {broj_tura} Touren, Limit wird auf "
                    f"{max_tura_po_rampi_lokalno} Touren/Rampe erhöht.")

    max_broj = tour_summary["broj_narudzbi"].max()
    max_tez = tour_summary["ukupna_tezina"].max()

    def skor(broj, tez):
        b = broj / max_broj if max_broj else 0
        t = tez / max_tez if max_tez else 0
        return tezina_udio * t + (1 - tezina_udio) * b

    tour_summary["skor"] = tour_summary.apply(
        lambda r: skor(r["broj_narudzbi"], r["ukupna_tezina"]), axis=1
    )
    tour_summary = tour_summary.sort_values("skor", ascending=False).reset_index(drop=True)

    rampe = {
        i: {"ture": [], "broj_narudzbi": 0, "ukupna_tezina": 0.0}
        for i in range(1, broj_rampi + 1)
    }

    for _, red in tour_summary.iterrows():
        kandidati = [r for r in rampe if len(rampe[r]["ture"]) < max_tura_po_rampi_lokalno]
        if not kandidati:
            kandidati = list(rampe.keys())

        najbolja_rampa = min(
            kandidati,
            key=lambda r: skor(rampe[r]["broj_narudzbi"], rampe[r]["ukupna_tezina"])
        )

        rampe[najbolja_rampa]["ture"].append(red["Tour"])
        rampe[najbolja_rampa]["broj_narudzbi"] += red["broj_narudzbi"]
        rampe[najbolja_rampa]["ukupna_tezina"] += red["ukupna_tezina"]

    # Neuanordnung: meiste Bestellungen (dann höchstes Gewicht) -> Rampe 1, wenigste -> letzte Rampe
    poredak = sorted(
        rampe.items(),
        key=lambda item: (item[1]["broj_narudzbi"], item[1]["ukupna_tezina"]),
        reverse=True
    )

    rampe_preslozene = {}
    for novi_broj, (_, podaci) in enumerate(poredak, start=1):
        rampe_preslozene[novi_broj] = podaci

    log.append(f"✅ '{sheet_name}' verarbeitet ({broj_tura} Touren).")
    return rampe_preslozene


def obradi_excel(uploaded_file, broj_rampi, max_tura_po_rampi, tezina_udio, trazeni_brojevi):
    log = []
    file_bytes = io.BytesIO(uploaded_file.getvalue())
    xls = pd.ExcelFile(file_bytes)

    listovi = [s for s in xls.sheet_names if sheet_odgovara(s, trazeni_brojevi)]
    log.append(f"Gefundene Blätter zur Verarbeitung: {listovi}")

    rezultati = {}
    for sheet in listovi:
        file_bytes.seek(0)
        rampe = obradi_list(file_bytes, sheet, broj_rampi, max_tura_po_rampi, tezina_udio, log)
        if rampe is not None:
            rezultati[sheet] = rampe

    # Ergebnisse zurück in die Datei speichern (im Arbeitsspeicher), auf neuen Blättern
    output = io.BytesIO(uploaded_file.getvalue())
    wb = load_workbook(output)

    for sheet, rampe in rezultati.items():
        novi_naziv = f"{sheet} - Rampe"[:31]
        if novi_naziv in wb.sheetnames:
            del wb[novi_naziv]
        ws = wb.create_sheet(novi_naziv)
        ws.append([sheet.upper()])
        for r in range(1, broj_rampi + 1):
            ture_str = ", ".join(str(t) for t in rampe[r]["ture"])
            ws.append([f"RAMPE {r} -> {ture_str}"])

    final_output = io.BytesIO()
    wb.save(final_output)
    final_output.seek(0)

    return final_output, rezultati, log


# ==== HAUPTTEIL DER OBERFLÄCHE ====

if uploaded_file is not None:
    if st.button("Datei verarbeiten"):
        with st.spinner("Verarbeitung läuft..."):
            try:
                rezultat_file, rezultati, log = obradi_excel(
                    uploaded_file, BROJ_RAMPI, MAX_TURA_PO_RAMPI, TEZINA_UDIO, TRAZENI_BROJEVI
                )
            except Exception as e:
                st.error(f"Es ist ein Fehler aufgetreten: {e}")
                st.stop()

        st.subheader("Verarbeitungsprotokoll")
        for linija in log:
            st.write(linija)

        if rezultati:
            st.subheader("Ergebnisübersicht")
            for sheet, rampe in rezultati.items():
                st.markdown(f"**{sheet.upper()}**")
                for r in range(1, BROJ_RAMPI + 1):
                    ture_str = ", ".join(str(t) for t in rampe[r]["ture"])
                    st.write(f"RAMPE {r} -> {ture_str}")

            st.download_button(
                label="📥 Verarbeitete Excel-Datei herunterladen",
                data=rezultat_file,
                file_name="ergebnis_rampen.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("Kein Blatt konnte erfolgreich verarbeitet werden. Bitte Protokoll oben prüfen.")
else:
    st.info("Warte auf Upload der Excel-Datei.")

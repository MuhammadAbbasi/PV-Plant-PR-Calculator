# 📘 Manuale Utente & Guida Operativa - PR Calculator (GET SRL)

> **Documento generato:** 2026-05-15 · **Ultimo aggiornamento:** 2026-09-02 (v15.0 · Architettura Modulare & Supporto Avanzato VCOM)

> [!IMPORTANT]
> **Azienda:** GET SRL  
> **Impianto:** Mazara 01 (36 Inverter fotovoltaici su 3 Trasformatori TX1, TX2, TX3)  
> **Scopo del Software:** Calcolo automatizzato del Performance Ratio (PR) giornaliero e mensile, compensazione perdite di rete/curtailment, conversione e integrazione telemetria VCOM, e generazione automatica dei report Excel e file Madre.

---

## 🚀 Novità e Miglioramenti Versione 15.0

La **versione 15.0** introduce un'architettura completamente rinnovata e risolve specifiche esigenze sui report mensili e sull'integrazione VCOM:

1. **Nuova Colonna "Meter Reading [MWh]" nel File Madre (`00 PR_recalculation_{MESE}.xlsx`)**:
   - Inserita automaticamente la colonna contenente l'ultima lettura cumulativa del contatore SATAC di ciascun giorno (collegata alla cella `$L$110` del file ricalcolato giornaliero), subito prima della colonna *"Energy (day)"*.
   - Nella riga dei **Totali Mensili**, la colonna della lettura contatore viene calcolata con la formula `=MAX(E5:E35)`.

2. **Evidenziazione Arancione Chiaro e Note per Giorni VCOM**:
   - I giorni elaborati con dati **VCOM** vengono evidenziati graficamente con uno sfondo **Arancione Chiaro** (`#FFE0B2` / RGB `255, 224, 178`).
   - Sulla cella della data (Colonna A) viene inserita una nota/commento Excel esplicativo:
     > *"Nota: Giorno elaborato con dati VCOM. È presente una differenza di circa 300 kW (kWh/giorno) nei dati di energia giornalieri tra SCADA e VCOM."*
   - Se un giorno viene ricalcolato successivamente con dati SCADA, l'evidenziazione e la nota vengono rimosse automaticamente.

3. **Integrazione Lettura Contatore VCOM (M_prev + E_today)**:
   - Per i giorni elaborati tramite VCOM, la lettura di fine giornata viene calcolata sommando la lettura finale del giorno precedente con l'energia prodotta nella giornata corrente.

4. **Allineamento Impaginazione Giorno 31 (Riga 35) e Totali (Riga 36)**:
   - Per i mesi di 31 giorni (es. Agosto), il giorno 31 occupa stabilmente la **Riga 35**, mentre la riga dei Totali mensili occupa la **Riga 36**.

5. **Architettura Modulare a Microservizi (`PR_Calculator_v15_App`)**:
   - Il codice sorgente è stato suddiviso in un pacchetto Python modulare professionale a 8 componenti (`config/`, `utils/`, `core/`, `data_converters/`, `excel_engine/`, `gui/`), ed è disponibile in versione distribuibile **`dist/PR_Calculator_v15.exe`**.

---

## 🖥️ Panoramica dell'Interfaccia Grafica (GUI)

L'interfaccia è stata progettata con un design chiaro, moderno ed elegante ispirato a Google Material Design, suddiviso in aree funzionali per guidare l'utente passo dopo passo.

---

## 📌 Mappa delle Aree e dei Comandi

### Area 1: Intestazione e Pulsanti di Accesso Rapido
- **Descrizione:** Mostra il brand GET SRL e il pulsante **[ ? Guida d'Uso ]** per consultare questo manuale in qualsiasi momento.

---

### Area 2: Selezione Cartella di Input
- **Pulsante [ Sfoglia... ]**: Apre la finestra di dialogo per selezionare la cartella contenente i file SCADA o la cartella del mese.
  - **Giorno Singolo:** Selezionare la cartella del giorno (es. `2026 08/19`).
  - **Modalità Mensile (Batch):** Selezionare la cartella del mese (es. `2026 08`). Il software individuerà ed elaborerà in sequenza tutti i giorni.

---

### Area 3: Parametri di Calcolo
| Campo di Input | Funzione e Significato | Valore Consigliato |
| :--- | :--- | :--- |
| **Data (AAAA-MM-GG)** | Data del giorno in analisi. Rilevata automaticamente dai file selezionati. | Formato ISO (es. `2026-08-19`) |
| **PR Mensile PVSyst** | Target teorico mensile, compilato in automatico in base alla data. | Compilato in automatico |
| **Irraggiamento Min (W/m²)** | Soglia minima di irraggiamento per il calcolo delle perdite. | `50` W/m² |
| **Tolleranza Diff. Irraggiamento (%)** | Scostamento massimo consentito prima di applicare il Conditional MAX. | `10` % (Default) |
| **Sorgente Dati (SCADA / VCOM / Misto)** | Consente di scegliere la sorgente dei dati per ciascun giorno (SCADA, VCOM, o Misto). | `SCADA` o `Misto` |

---

### Area 4: Opzione Ricalcolo Forzato
- **Casella [ Ricalcola forzatamente i giorni già elaborati ]**: Se selezionata, sovrascrive ed ricalcola anche i giorni già elaborati precedentemente.

---

### Area 5: Avvio dell'Elaborazione
- **Pulsante [ Calcola Performance Ratio ]**: Esegue il calcolo completo:
  1. Caricamento e pulizia dei file dati (SCADA o VCOM).
  2. Riparazione della serie storica del contatore SATAC.
  3. Calcolo dell'irraggiamento POA, del PR Compensato e delle perdite per inverter e trasformatore.
  4. Scrittura nativa in background sul file giornaliero `PR_recalculation_DD_mmm.xlsx`.
  5. Sincronizzazione automatica sul File Madre `00 PR_recalculation_MESE.xlsx` (inserendo la lettura del contatore, le note VCOM e i colori di evidenziazione).

- **Pulsante [ Interrompi ]**: Consente l'arresto sicuro al termine del giorno in corso.

---

## 🌐 Conversione e Gestione Dati VCOM (v15.0)

Quando i dati SCADA di uno o più giorni sono assenti o incompleti, è possibile attivare la modalità VCOM:

1. **Conversione automatici 3-File VCOM**:
   - I file `Energia_YYYY_MM_DD.csv`, `Potenza_AC_YYYY_MM_DD.csv` e `Potenza_attiva_YYYY_MM_DD.csv` esportati da VCOM vengono convertiti nei 7 file pseudo-SCADA a 15 minuti.
2. **Aggregazione Energia Inverter**:
   - L'energia giornaliera prodotta viene calcolata sommando i dati di tutti i 36 inverter presenti nel file VCOM.
3. **Contatore Cumulativo**:
   - La serie del contatore viene ricostruita partendo dalla lettura finale del giorno precedente ed incrementando la produzione accumulata della giornata.
4. **Evidenziazione e Nota nel File Madre**:
   - La riga del giorno viene colorata in **Arancione Chiaro** (`#FFE0B2`) e viene allegato il commento esplicativo sulla differenza tipica di circa 300 kW tra SCADA e VCOM.

---

## 📑 Struttura del File Madre (`00 PR_recalculation_MESE.xlsx`)

Il file Madre mensile ha la seguente disposizione delle colonne:

| Colonna | Intestazione | Sorgente / Formula |
| :---: | :--- | :--- |
| **A (1)** | Data | `YYYY-MM-DD` |
| **B (2)** | Irradiance TX1 | `='[PR_recalculation_DD_mmm.xlsx]PR_Calc'!$D$111` |
| **C (3)** | Irradiance TX3 | `='[PR_recalculation_DD_mmm.xlsx]PR_Calc'!$F$111` |
| **D (4)** | Irradiance Cond MAX | `='[PR_recalculation_DD_mmm.xlsx]PR_Calc'!$I$111` |
| **E (5)** | **Meter Reading [MWh]** *(Novità v15)* | `='[PR_recalculation_DD_mmm.xlsx]PR_Calc'!$L$110` |
| **F (6)** | Energy (day) | `='[PR_recalculation_DD_mmm.xlsx]PR_Calc'!$M$111` |
| **G (7)** | PR Total | `='[PR_recalculation_DD_mmm.xlsx]PR_Calc'!$BA$5*100` |
| **H (8)** | PR VCOM | Inserito da Sync VCOM |
| **I (9)** | PR Compensated | `='[PR_recalculation_DD_mmm.xlsx]PR_Calc'!$BH$11` |
| **J (10)**| External Availability [%] | `=IF(F{r}="",0,(F{r}/(F{r}+K{r}+L{r}+M{r}))*100)` |
| **K-M (11-13)**| Loss TX1, TX2, TX3 | `='[PR_recalculation_DD_mmm.xlsx]PR_Calc'!$AA$111` / `$AN$111` / `$BA$111` |

- **Riga Totali Mensili (Riga 36 per mesi da 31 giorni)**:
  - `Col E (Meter Reading)`: `=MAX(E5:E35)`
  - `Col F (Energy day)`: `=SUM(F5:F35)`
  - `Col G (PR Total)`: `=AVERAGE(G5:G35)`
  - `Col I (PR Compensated)`: `=AVERAGE(I5:I35)`
  - `Col J (External Availability)`: `=SUMIF(J5:J35,"<>0")/COUNTIF(J5:J35,"<>0")`
  - `Cols K-M (Losses)`: `=SUM(K5:K35)`, `=SUM(L5:L35)`, `=SUM(M5:M35)`

---

## 🛠️ Risoluzione dei Problemi (Troubleshooting)

- **Errore "Colonna2 Not Found"**: Risolto nella v15.0 con la nuova funzione `normalize_columns()` che accetta file contatore o meteo con qualsiasi numero di colonne.
- **Giorni VCOM evidenziati in Arancione**: È il comportamento normale e desiderato della v15.0 per segnalare a colpo d'occhio che i dati provengono da VCOM anziché da SCADA.
- **Eseguibile standalone**: L'applicazione completa è compilata nel file autonomo [`dist/PR_Calculator_v15.exe`](file:///S01/get/2025.01%20Mazara%2001%20A2A/03%20-%20REPORT/Report/09%20Testing/PR%20Calculation%20automation/dist/PR_Calculator_v15.exe), che non richiede l'installazione di Python sul computer dell'operatore.

# 📘 Manuale Utente & Guida Operativa - PR Calculator (GET SRL)

> [!IMPORTANT]
> **Azienda:** GET SRL  
> **Impianto:** Mazara 01 (36 Inverter fotovoltaici su 3 Trasformatori TX1, TX2, TX3)  
> **Scopo del Software:** Calcolo automatizzato del Performance Ratio (PR) giornaliero e mensile, compensazione perdite di rete/curtailment, e generazione automatica dei report Excel e file Madre.

---

## 🖥️ Panoramica dell'Interfaccia Grafica (GUI)

L'interfaccia è stata progettata con un design scuro moderno ed ergonomico, suddivisa in aree funzionali per guidare l'utente junior passo dopo passo.

![Guida Interfaccia Grafica con Aree Numerate](file:///\\s01\get\2025.01%20Mazara%2001%20A2A\03%20-%20REPORT\Report\09%20Testing\PR%20Calculation%20automation\gui_annotated_guide.png)

---

## 📌 Mappa delle Aree e dei Comandi

### Area 1: Intestazione e Logo GET SRL
- **Descrizione:** Mostra il brand aziendale e il titolo del software di automazione per il calcolo delle prestazioni fotovoltaiche.

---

### Area 2: Selezione Cartella di Input
- **Pulsante [ Browse... ]**: Apre la finestra di dialogo di Windows per selezionare la cartella contenente i file grezzi SCADA.
  - **Modalità Giorno Singolo:** Selezionare la cartella del giorno specifico (es. `2026 05/01`).
  - **Modalità Elaborazione Mensile (Batch):** Selezionare la cartella del mese (es. `2026 05`). Il software individuerà automaticamente tutte le sottocartelle numeriche (`01`, `02`, ..., `31`) ed elaborerà l'intero mese in sequenza in modo completamente automatizzato.

---

### Area 3: Parametri di Calcolo
| Campo di Input | Significato e Funzione | Valore Consigliato |
| :--- | :--- | :--- |
| **Date (YYYY-MM-DD)** | Data del giorno in analisi. In modalità batch, viene rilevata e aggiornata automaticamente durante il ciclo. | Formato ISO (es. `2026-05-01`) |
| **PVSyst Monthly PR** | Obiettivo mensile di Performance Ratio stimato dal software di simulazione di progetto PVSyst. | Es. `0.897` (89.7%) |
| **Min Irradiance (W/m²)** | Soglia minima di irraggiamento solare oltre la quale i calcoli di perdita di energia entrano in funzione. | `50` W/m² |

---

### Area 4: Opzione Ricalcolo Forzato
- **Casella di Spunta [ Force reprocess already calculated days ]**: 
  - *Se deselezionata (Default):* In modalità elaborazione mensile batch, il software salta automaticamente i file giornalieri già calcolati in precedenza, rendendo l'esecuzione pressoché istantanea.
  - *Se selezionata:* Il software sovrascrive e ricalcola da zero ogni singolo giorno del mese rigenerando tutte le formule.

---

### Area 5: Avvio dell'Elaborazione
- **Pulsante [ Calculate Performance Ratio ]**: È il motore di calcolo principale del software. Cliccando questo pulsante il sistema esegue:
  1. Caricamento e normalizzazione di tutti i file SCADA (Inverter, Meteo, Contatori SATAC, Regolazione Potenza).
  2. Pulizia automatica dei valori numerici da virgole italiane e caratteri anomali.
  3. Calcolo delle perdite per indisponibilità e curtailment per ciascun quarto d'ora.
  4. Scrittura nativa in background (tramite Excel COM) sul file di calcolo giornaliero `PR_recalculation_DD_mmm.xlsx`.
  5. Aggiornamento in tempo reale del file Madre mensile `00 PR_recalculation_MESE.xlsx`.

---

## 📊 Cruscotto dei Risultati (Area 6)

Una volta terminato il calcolo, l'area di destra mostra i 3 indicatori chiave dell'impianto:

> [!TIP]
> - **AVG INVERTER PR (Verde):** La media aritmetica dei Performance Ratio compensati dei 36 inverter. Rappresenta lo stato di salute e di efficienza pura delle macchine.
> - **COMPENSATED RAW PR (Ciano):** Il Performance Ratio totale dell'impianto in cui l'energia prodotta è stata sommata alle perdite stimate. È il valore ufficiale di riferimento contrattuale.
> - **UNCOMPENSATED PR (Arancione):** Il Performance Ratio grezzo, calcolato esclusivamente sull'energia attiva immessa in rete, senza tenere conto delle perdite per guasti di rete o limitazioni del gestore.

---

## 📑 Tabella di Dettaglio dei 36 Inverter (Area 7)

Questa tabella interattiva espone il comportamento di ogni singolo inverter (da `TX1-INV-1` a `TX3-INV-12`):
- **Nominal DC (kW):** La potenza nominale di targa dell'inverter (328.125, 343.75 o 359.375 kW a seconda della stringa).
- **Actual Energy (kWh):** L'energia reale prodotta dall'inverter nella giornata in esame.
- **Energy Loss (kWh):** La stima dell'energia perduta per indisponibilità della rete o regolazione attiva.
- **Compensated PR (%):** Il Performance Ratio specifico della singola macchina.

---

## 💾 Esportazione dei Dati (Area 8)

- **Pulsante [ Export Complete Data to Excel... ]**: Genera un report Excel autonomo multi-foglio contenente il riassunto esecutivo, l'elenco dei 36 inverter e tutti i 96 quarti d'ora della giornata con ogni singola variabile calcolata.

---

## 📜 Console di Log in Tempo Reale (Area 9)

- **Console Live Log:** Visualizza lo stato di avanzamento delle operazioni, l'apertura dei file Excel in memoria e la notifica di eventuali anomalie o file mancanti.

---

## 🛠️ Guida alla Risoluzione dei Problemi (Troubleshooting per Junior)

> [!WARNING]
> - **Errore "File Not Found":** Verifica che nella cartella selezionata siano presenti tutti e 7 i file SCADA richiesti (`TS_01_Inverter`, `SATAC_Meter`, ecc.).
> - **Errore "#DIV/0!" nei file generati:** È stato eliminato grazie alla funzione `IFERROR`. Se apri un file grezzo e vedi divisioni per zero nelle ore notturne, avvia il calcolo tramite questo software per ripristinare le formule corrette.
> - **Blocco di Rete SMB / Errore OLE:** Il software gestisce automaticamente i percorsi di rete UNC convertendo gli slash in backslash (`\`).

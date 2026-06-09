# 📘 Manuale Utente & Guida Operativa - PR Calculator (GET SRL)

> [!IMPORTANT]
> **Azienda:** GET SRL  
> **Impianto:** Mazara 01 (36 Inverter fotovoltaici su 3 Trasformatori TX1, TX2, TX3)  
> **Scopo del Software:** Calcolo automatizzato del Performance Ratio (PR) giornaliero e mensile, compensazione perdite di rete/curtailment, e generazione automatica dei report Excel e file Madre.

---

## 🖥️ Panoramica dell'Interfaccia Grafica (GUI)

L'interfaccia è stata progettata con un design chiaro moderno, pulito ed elegante ispirato a Google Material Design, suddiviso in aree funzionali per guidare l'utente junior passo dopo passo.

![Guida Interfaccia Grafica con Aree Numerate](file:///\\s01\get\2025.01%20Mazara%2001%20A2A\03%20-%20REPORT\Report\09%20Testing\PR%20Calculation%20automation\gui_annotated_guide.png)

---

## 📌 Mappa delle Aree e dei Comandi

### Area 1: Intestazione, Logo GET SRL, e pulsante ? Guida d'Uso
- **Descrizione:** Mostra il brand aziendale, il titolo del software di automazione per il calcolo delle prestazioni fotovoltaiche, e include un pulsante di accesso rapido **[ ? Guida d'Uso ]** per aprire questo manuale nel browser in qualsiasi momento.

---

### Area 2: Selezione Cartella di Input
- **Pulsante [ Sfoglia... ]**: Apre la finestra di dialogo di Windows per selezionare la cartella contenente i file grezzi SCADA.
  - **Modalità Giorno Singolo:** Selezionare la cartella del giorno specifico (es. `2026 05/01`).
  - **Modalità Elaborazione Mensile (Batch):** Selezionare la cartella del mese (es. `2026 05`). Il software individuerà automaticamente tutte le sottocartelle numeriche (`01`, `02`, ..., `31`) ed elaborerà l'intero mese in sequenza in modo completamente automatizzato.

---

### Area 3: Parametri di Calcolo
| Campo di Input | Significato e Funzione | Valore Consigliato |
| :--- | :--- | :--- |
| **Data (AAAA-MM-GG)** | Data del giorno in analisi. Viene rilevata e compilata automaticamente leggendo i dati dai file Excel selezionati. | Formato ISO (es. `2026-05-01`) |
| **PR Mensile PVSyst** | Obiettivo mensile di Performance Ratio teorico, rilevato e compilato automaticamente dalla tabella di riferimento quando viene aggiornata la data. | Compilato in automatico |
| **Irraggiamento Min (W/m²)** | Soglia minima di irraggiamento solare oltre la quale i calcoli di perdita di energia entrano in funzione. | `50` W/m² |

---

### Area 4: Opzione Ricalcolo Forzato
- **Casella di Spunta [ Ricalcola forzatamente i giorni già elaborati (Modalità Batch) ]**: 
  - *Se deselezionata (Default):* In modalità elaborazione mensile batch, il software salta automaticamente i file giornalieri già calcolati in precedenza, rendendo l'esecuzione pressoché istantanea.
  - *Se selezionata:* Il software sovrascrive e ricalcola da zero ogni singolo giorno del mese rigenerando tutte le formule.

---

### Area 5: Avvio dell'Elaborazione
- **Pulsante [ Calcola Performance Ratio ]**: È il motore di calcolo principale del software. Cliccando questo pulsante il sistema esegue:
  1. Caricamento e normalizzazione di tutti i file SCADA (Inverter, Meteo, Contatori SATAC, Regolazione Potenza).
  2. Pulizia automatica dei valori numerici da virgole italiane e caratteri anomali.
  3. Calcolo delle perdite per indisponibilità e curtailment per ciascun quarto d'ora.
  4. Scrittura nativa in background (tramite Excel COM con supporto Formula2 per evitare errori @) sul file di calcolo giornaliero `PR_recalculation_DD_mmm.xlsx`.
  5. Aggiornamento in tempo reale del file Madre mensile `00 PR_recalculation_MESE.xlsx` (inserendo le formule protette contro i valori vuoti per l'External Availability).

---

## 📊 Tabella Riferimento Target PVSyst (Area 6)

Mostra la tabella dei Target PR previsti da PVSyst per ciascun mese dell'anno. Quando viene inserita o rilevata una data, la riga del mese corrispondente viene automaticamente selezionata ed evidenziata nella tabella, e il valore del Target PR viene copiato nel campo di calcolo.

---

## 🚦 Stato dell'Elaborazione (Area 7)

- **Descrizione:** Mostra i messaggi di progresso e l'esito finale dell'elaborazione (es. *"Calcolo completato con successo!"* o messaggi di errore) in corrispondenza del pulsante principale di calcolo.

---

## 📑 Titolo e Intestazione Console (Area 8)

- **Descrizione:** Intestazione della sezione diagnostica con titolo *"Console Live Log di Esecuzione"* e indicazioni sui messaggi diagnostici in tempo reale relativi all'elaborazione dei file Excel e del motore di calcolo.

---

## 📜 Console Live Log in Tempo Reale (Area 9)

- **Console Live Log:** La finestra di testo che visualizza i log operativi riga per riga (notifiche di avanzamento, caricamenti di file SCADA, ricalcoli batch dei giorni e scrittura sul file Madre).

---

## 📑 Dettaglio e Calcolo Excel (Generato Automaticamente)

Anche se non visualizzati direttamente nell'interfaccia principale della v10, il software genera e popola automaticamente:
- **Dettagli dei 36 Inverter:** Codice inverter, trasformatore (TX1/2/3), potenza nominale, energia prodotta, perdite stimate e PR Compensato scritti direttamente nelle schede del file Excel.
- **Formule Attive nel Foglio Giornaliero:** Formula per il PR Compensato (cella `BH11`) e per le perdite per trasformatore (riga 111).

---

## 🧮 Formula di Dettaglio del PR Compensato (v10.0)

A partire dalla versione 10.0, il software scrive nel foglio `PR_Calc` (cella `BH11`) del file giornaliero la formula attiva per il calcolo del **PR Compensato**:

$$\text{PR Compensato} = \left( \frac{\text{Energia Inverter Total} + \text{Loss TX1} + \text{Loss TX2} + \text{Loss TX3}}{\text{Irradiance Sum} \times \text{Plant CC Power}} \right) \times 100$$

Nello specifico, la formula Excel inserita è:
`=((SUM(Inverter_data!C15:N110, Inverter_data!R15:AC110, Inverter_data!AG15:AR110)*0.25 + AA111 + AN111 + BA111) / (12625 * (SUM($H$15:$H$110)*0.25/1000))) * 100`

Questa formula viene poi sincronizzata nel file **Madre** mensile:
* **Colonna H (PR VCOM / PR Total)**: Collegato alla cella `$BA$5*100` (PR Raw).
* **Colonna I (PR Compensated)**: Collegato alla cella `$BH$11` (PR Compensato).
* **Colonna J (External Availability %)**: Calcolata tramite formula `=IF(E{r}="",0,(E{r}/(E{r}+K{r}+L{r}+M{r}))*100)`.
* **Colonne K, L, M (TX1, TX2, TX3 Energy Loss)**: Collegate alle celle `$AA$111`, `$AN$111`, `$BA$111` dei file giornalieri.


---

## 🛠️ Guida alla Risoluzione dei Problemi (Troubleshooting per Junior)

> [!WARNING]
> - **Errore "File Not Found":** Verifica che nella cartella selezionata siano presenti tutti e 7 i file SCADA richiesti (`TS_01_Inverter`, `SATAC_Meter`, ecc.).
> - **Formato Numerico Italiano:** Nella GUI i numeri decimali vengono inseriti e visualizzati usando la virgola (es. `0,897` o `50,0`), in conformità con i requisiti locali. Il software converte automaticamente i valori in formato corretto per l'esportazione su Excel.
> - **Errore "#DIV/0!" nei file generati:** È stato eliminato grazie alla funzione `IFERROR`. Se apri un file grezzo e vedi divisioni per zero nelle ore notturne, avvia il calcolo tramite questo software per ripristinare le formule corrette.
> - **Blocco di Rete SMB / Errore OLE:** Il software gestisce automaticamente i percorsi di rete UNC condivisi convertendo gli slash in backslash (`\`).


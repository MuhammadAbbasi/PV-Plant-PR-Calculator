# 📘 Manuale Utente & Guida Operativa - PR Calculator (GET SRL)

> **Documento generato:** 2026-05-15 · **Ultimo aggiornamento:** 2026-07-24 (v12.0)

> [!IMPORTANT]
> **Azienda:** GET SRL  
> **Impianto:** Mazara 01 (36 Inverter fotovoltaici su 3 Trasformatori TX1, TX2, TX3)  
> **Scopo del Software:** Calcolo automatizzato del Performance Ratio (PR) giornaliero e mensile, compensazione perdite di rete/curtailment, e generazione automatica dei report Excel e file Madre.

---

## 🖥️ Panoramica dell'Interfaccia Grafica (GUI)

L'interfaccia è stata progettata con un design chiaro moderno, pulito ed elegante ispirato a Google Material Design, suddiviso in aree funzionali per guidare l'utente junior passo dopo passo.

![Guida Interfaccia Grafica con Aree Numerate](file:///\\s01\get\2025.01%20Mazara%2001%20A2A\03%20-%20REPORT\Report\09%20Testing\PR%20Calculation%20automation\archive\gui_annotated_guide.png)

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
| **Tolleranza Diff. Irraggiamento (%)** | Soglia di scostamento consentito (tra 0% e 100%) tra i sensori per il calcolo del Conditional MAX. Se superata, viene preso il valore massimo. (Usata solo con metodo *Conditional MAX*.) | `10` % (Default) |
| **Riferimento POA per il PR** (selettore) | Metodo con cui si ricava l'irraggiamento di riferimento del PR dai due piranometri TX1/TX3. **Media (Average):** media aritmetica dei due sensori (standard IEC); la tolleranza differenziale non viene usata e il relativo campo è disattivato. **Conditional MAX:** usa il sensore maggiore quando i due divergono oltre la tolleranza (più conservativo). La scelta si applica a tutti i valori di PR. Il pulsante attivo è evidenziato in blu. | `Media (Average)` (Default v12.0) |

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

- **Pulsante [ Interrompi (arresto sicuro) ] (v12.0)**: Attivo solo durante l'elaborazione. Alla pressione **non** interrompe il lavoro a metà: imposta una richiesta di arresto che il motore verifica ai punti di controllo sicuri (fra un giorno e l'altro, fra un download VCOM e il successivo). Il giorno in corso viene quindi **completato e salvato**, poi l'esecuzione si ferma nel minor tempo possibile. Il file Madre viene comunque sincronizzato per i giorni già elaborati, così i dati restano coerenti. Lo stato finale riporta *"Elaborazione interrotta dall'utente. Completati: N giorni"*.

---

## 🌐 Dati VCOM automatici per i giorni mancanti (v12.0)

Se in modalità batch alcuni giorni non hanno tutti i 7 file SCADA richiesti, il software non si limita più a saltarli:

1. **Rilevamento:** cerca nella cartella del giorno una sottocartella `vcom/` (o `VCOM/`) contenente `Potenza_AC_*.csv` e `Produzione_energetica_*.csv`.
2. **Download automatico:** se la cartella non esiste, propone di scaricare i dati direttamente dal portale **meteocontrol VCOM** tramite browser automatizzato (Playwright). Il browser viene aperto **in modalità visibile**, così è possibile seguire l'avanzamento dell'estrazione.
3. **Conversione:** i due CSV VCOM a 5 minuti vengono convertiti in file pseudo-SCADA a 15 minuti (`SATAC_Meter_15Min.xlsx`, `TS_01/03_Weather_15Min.xlsx`, `TS_01/02/03_Inverter_15Min.xlsx`), dopodiché il giorno risulta completo e il PR viene calcolato normalmente.

> [!NOTE]
> La conversione genera **6** dei 7 file richiesti. Il file `Regolazione_della_potenza_attiva_*.xlsx` non deriva dai dati VCOM di produzione e va scaricato separatamente (estrattore *potenza attiva*). Senza di esso il giorno resta incompleto.

> [!IMPORTANT]
> Le credenziali VCOM vengono lette da `VCOM Automation/config.json`. Se l'eseguibile compilato segnala che il browser non esiste (percorso `..\Temp\_MEIxxxxx\...`), significa che la versione compilata non trova i browser Playwright: il software ora imposta automaticamente `PLAYWRIGHT_BROWSERS_PATH` sulla cartella utente `%LOCALAPPDATA%\ms-playwright`. Se i browser non fossero mai stati installati, eseguire una volta `playwright install chromium`.

---

## 📊 Tabella Riferimento Target PVSyst (Area 6)

Mostra la tabella dei Target PR previsti da PVSyst per ciascun mese dell'anno. Quando viene inserita o rilevata una data, la riga del mese corrispondente viene automaticamente selezionata ed evidenziata nella tabella, e il valore del Target PR viene copiato nel campo di calcolo.

La colonna **Target Corretto** mostra il target effettivamente usato dal motore dopo la **degradazione contrattuale dello 0,4%/anno** (Allegato 9.1), composta annualmente dall'avvio impianto (Febbraio 2025). L'anno contrattuale decorre da Febbraio a Gennaio: l'Anno 1 (Feb 2025 – Gen 2026) non è degradato, l'Anno 2 (Feb 2026 – Gen 2027) applica il fattore (1−0,4%)¹ = 0,996. La colonna si aggiorna automaticamente in base all'anno della data selezionata.

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

Anche se non visualizzati direttamente nell'interfaccia principale della v11, il software genera e popola automaticamente:
- **Dettagli dei 36 Inverter:** Codice inverter, trasformatore (TX1/2/3), potenza nominale, energia prodotta, perdite stimate e PR Compensato scritti direttamente nelle schede del file Excel.
- **Formule Attive nel Foglio Giornaliero:** Formula per il PR Compensato (cella `BH11`) e per le perdite per trasformatore (riga 111).
- **Riparazione Letture Contatore (v11):** Le letture del contatore SATAC mancanti, nulle o decrescenti (delta negativo) vengono rilevate e ricostruite per interpolazione tra i valori validi adiacenti. Le celle corrette vengono evidenziate in **arancione** con una nota esplicativa, così l'anomalia resta tracciabile.

---

## 🧮 Formula di Dettaglio del PR Compensato (v11.0)

Il software scrive nel foglio `PR_Calc` (cella `BH11`) del file giornaliero la formula attiva per il calcolo del **PR Compensato**. A partire dalla v11.0 il denominatore usa l'irraggiamento di riferimento della **Colonna I** (`SUM($I$15:$I$110)`), coerente con il metodo POA scelto (Conditional MAX o Media) e con la soglia minima di irraggiamento:

$$\text{PR Compensato} = \left( \frac{\text{Energia Inverter Total} + \text{Loss TX1} + \text{Loss TX2} + \text{Loss TX3}}{\text{Plant CC Power} \times \text{Irradiance Sum (Col. I)}} \right) \times 100$$

Nello specifico, la formula Excel inserita è:
`=((SUM(Inverter_data!C15:N110, Inverter_data!R15:AC110, Inverter_data!AG15:AR110)*0.25 + AA111 + AN111 + BA111) / (12625 * SUM($I$15:$I$110))) * 100`

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
> - **File Excel aperto / bloccato (v11):** Se un file giornaliero o il file Madre è già aperto in un'altra finestra di Excel, il software mostra una finestra di conferma e, se si accetta, lo chiude automaticamente per proseguire. Attenzione: eventuali modifiche non salvate in quel file andranno perse.
> - **Celle arancioni nel file giornaliero (v11):** Segnalano letture del contatore SATAC mancanti o anomale ricostruite per interpolazione. Passare il mouse sulla cella per leggere la nota esplicativa. I valori di Energia e Disponibilità del giorno risultano così corretti.
> - **Blocco di Rete SMB / Errore OLE:** Il software gestisce automaticamente i percorsi di rete UNC condivisi convertendo gli slash in backslash (`\`).
> - **"BrowserType.launch: Executable doesn't exist ...\Temp\_MEIxxxxx\..." (v12):** Si verifica solo con l'eseguibile compilato: la versione impacchettata cercava i browser Playwright nella cartella temporanea di PyInstaller anziché in quella utente. Risolto impostando automaticamente `PLAYWRIGHT_BROWSERS_PATH` su `%LOCALAPPDATA%\ms-playwright`. Se compare ancora, i browser non sono installati: eseguire `playwright install chromium`.
> - **L'elaborazione non si ferma subito dopo [ Interrompi ] (v12):** È il comportamento previsto. L'arresto è *sicuro*: il giorno in corso viene completato e salvato prima di fermarsi, per non lasciare file Excel scritti a metà. L'attesa massima corrisponde al tempo di un singolo giorno (o di un singolo download VCOM).
> - **Un giorno con dati VCOM resta "incompleto" (v12):** La conversione VCOM genera 6 dei 7 file richiesti. Verificare che sia presente anche `Regolazione_della_potenza_attiva_*.xlsx`, che va scaricato con l'estrattore della potenza attiva.


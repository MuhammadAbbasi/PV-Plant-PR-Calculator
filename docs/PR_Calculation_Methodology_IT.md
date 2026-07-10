# Calcolo del Performance Ratio (PR) — Metodologia

**Impianto:** impianto fotovoltaico Mazara 01 (GET S.R.L.) · **Documento generato:** 2026-07-10 · **Ultimo aggiornamento:** 2026-07-10

> Questo documento spiega, in linguaggio semplice, come viene calcolato il Performance Ratio dell'impianto Mazara 01. È scritto per **lettori tecnici e non tecnici**. Ogni termine tecnico è definito alla prima comparsa, ogni fonte dati è elencata, ogni formula è scomposta passo passo ed esempi pratici mostrano come si combinano i numeri. Una versione inglese identica è disponibile in [`PR_Calculation_Methodology_EN.md`](PR_Calculation_Methodology_EN.md).

---

## Indice
1. [Introduzione al calcolo del PR](#1-introduzione-al-calcolo-del-pr)
2. [Fonti dati utilizzate](#2-fonti-dati-utilizzate)
3. [Formule principali con annotazioni](#3-formule-principali-con-annotazioni)
4. [Esempio di calcolo passo passo](#4-esempio-di-calcolo-passo-passo)
5. [Domande frequenti](#5-domande-frequenti)
6. [Glossario](#6-glossario)

---

## 1. Introduzione al calcolo del PR

### Che cos'è il Performance Ratio?

Il **Performance Ratio (PR)** è una singola percentuale che risponde a una domanda:

> *"Di tutta l'energia che l'impianto **avrebbe potuto** produrre data la luce solare ricevuta, quanta ne ha **effettivamente** prodotta?"*

Un PR del **100%** significherebbe che l'impianto ha prodotto esattamente la quantità ideale per quella luce (nessuna perdita). Gli impianti reali stanno sempre sotto il 100% a causa della fisica inevitabile (calore, cablaggi, efficienza degli inverter) e di problemi occasionali (guasti, limiti di rete). Un impianto in salute come Mazara 01 si attesta tipicamente intorno all'**82–83%**.

Il PR si calcola così:

```
                Energia effettivamente prodotta
PR (%)  =  --------------------------------------------  × 100
            Energia che la luce solare poteva sostenere
```

- **Energia effettivamente prodotta** — misurata direttamente dai contatori e dagli inverter.
- **Energia che la luce solare poteva sostenere** — la potenza nominale dell'impianto moltiplicata per la quantità di sole (irraggiamento) effettivamente caduta sui pannelli.

### Perché è importante

L'Operatore ha una **garanzia contrattuale** (Allegato 9.1) di mantenere il PR sopra un valore obiettivo ogni anno. Se il PR reale scende sotto il PR garantito, è dovuta una **penale**. Poiché ci sono soldi in gioco, il calcolo deve essere trasparente, ripetibile e difendibile — ed è lo scopo di questo documento.

### I quattro valori di PR che troverai

I report mostrano più di un valore di PR, perché ci sono modi diversi di guardare le prestazioni:

| Nome | Significato semplice |
|------|----------------------|
| **PR Raw** (detto "PR Total") | Rapporto diretto reale-vs-ideale, **senza correzioni**. Cala ogni volta che l'impianto è fermo, anche se il fermo non è colpa dell'operatore. |
| **PR SCADA** | Il PR riportato dal sistema di controllo dell'impianto (SCADA), preso dall'export KPI giornaliero del fornitore. |
| **PR VCOM** | Il PR riportato dalla piattaforma di monitoraggio indipendente **VCOM** (verifica di terza parte). |
| **PR Compensato** | Il PR Raw **più** l'energia dimostrabilmente persa per fermi e limitazioni di rete, **riaggiunta**. Isola la qualità *tecnica* dell'impianto dagli eventi fuori dal controllo dell'apparecchiatura, ed è il valore più rilevante per la garanzia. |

> **SCADA** = *Supervisory Control And Data Acquisition*, il sistema di controllo e monitoraggio in campo dell'impianto.

---

## 2. Fonti dati utilizzate

Ogni giornata di calcolo legge un insieme di file grezzi (campionati ogni **15 minuti**, quindi **96 intervalli al giorno**). Di seguito ogni dataset utilizzato e il suo contributo.

### 2.1 File di misura grezzi (una cartella per giorno)

| File | Cosa contiene | Ruolo nel PR |
|------|---------------|--------------|
| `TS_01_Inverter_15Min.xlsx`<br>`TS_02_Inverter_15Min.xlsx`<br>`TS_03_Inverter_15Min.xlsx` | **Potenza attiva (kW)** di ogni inverter, per le tre stazioni di trasformazione **TX1, TX2, TX3** (12 inverter ciascuna = **36 inverter**). | Fonte dell'**energia effettiva** e del **rilevamento dei fermi** (un inverter che legge ≈0 mentre c'è sole è "fermo"). |
| `TS_01_Weather_15Min.xlsx`<br>`TS_03_Weather_15Min.xlsx` | **Irraggiamento (W/m²)** dai due **piranometri** montati sul piano dei pannelli (POA), alle stazioni TX1 e TX3. | Fonte del **sole ricevuto** — il denominatore del PR. |
| `SATAC_Meter_15Min.xlsx` | La lettura cumulativa (kWh) del **contatore fiscale di energia** al punto di connessione alla rete. | Misura indipendente dell'**energia effettivamente immessa**; usata anche per la disponibilità. |
| `Regolazione_della_potenza_attiva_YYYY_MM_DD.xlsx` | Il **rapporto di limitazione della potenza attiva** (0–1): quanto il gestore di rete ha permesso all'impianto di produrre. | Rileva la **limitazione (curtailment)** imposta dalla rete. |

> **Piranometro** = strumento che misura l'irraggiamento solare (potenza della luce) in watt per metro quadro (W/m²).
> **POA** = *Plane Of Array*, cioè l'irraggiamento misurato con la stessa inclinazione/orientamento dei pannelli — la luce che i pannelli effettivamente "vedono".
> **Curtailment (limitazione)** = il gestore di rete che ordina all'impianto di produrre **meno** di quanto potrebbe, per motivi di rete.

### 2.2 File PR esterni dei fornitori (opzionali, uno al mese)

Se presenti nella cartella del mese, forniscono valori di PR giornalieri già pronti, che vengono scritti nel file di riepilogo mensile ("file Madre"):

| File | Formato | Fornisce |
|------|---------|----------|
| `KPI_Report_Daily.xls` | Excel | **PR SCADA giornaliero (%)** dal sistema di controllo dell'impianto. |
| `Performance_ratio_vcom.csv` | UTF-16, separato da tabulazioni | **PR VCOM giornaliero (%)** dalla piattaforma di monitoraggio indipendente. |

### 2.3 Parametri di progetto e costanti fisse

| Input | Valore | Significato |
|-------|--------|-------------|
| **Potenza nominale totale** (P_nominale) | **12.625 kWp** | La somma della potenza nominale DC di tutti i 36 inverter. "kWp" = kilowatt-picco, la potenza sotto luce solare standard. |
| **Potenza DC per inverter** | 328,125 – 359,375 kWp | La dimensione del campo di ciascun inverter (varia leggermente). |
| **Limite AC per inverter** | 320 kW (utile ≈ 280,3 kW) | Il massimo che un inverter può immettere in rete; il fattore utile 0,876 riflette i limiti reali. |
| **Soglia minima di irraggiamento** | **50 W/m²** | Sotto questa luce debole gli intervalli vengono ignorati (letture troppo rumorose per essere significative). |
| **Tolleranza di scostamento POA** | **10%** | Di quanto i due piranometri possono discordare prima di fidarsi del più alto (vedi §3.1). |
| **Target PR mensile PVSyst** | 82,0% – 90,4% (per mese) | Il PR atteso da progetto per ogni mese, dalla simulazione PVSyst. Usato come riferimento/obiettivo e nelle stime di perdita. |
| **Degradazione annua** | **0,4% all'anno** | Il target PR garantito si riduce dello 0,4% ogni anno (composto) con l'invecchiamento dei pannelli — vedi §3.8. |
| **Intervallo di campionamento** | 15 minuti = 0,25 h | Ogni lettura rappresenta un quarto d'ora di funzionamento. |

**Valori base del target PR mensile PVSyst (Anno 1):**

| Gen | Feb | Mar | Apr | Mag | Giu | Lug | Ago | Set | Ott | Nov | Dic |
|----|----|----|----|----|----|----|----|----|----|----|----|
|90,4|89,6|89,7|86,8|83,2|83,3|82,0|82,8|85,2|87,6|89,4|90,0|

> **PVSyst** = software standard di settore usato per simulare la produzione attesa di un impianto solare. I suoi valori mensili di PR sono il "target di progetto".

---

## 3. Formule principali con annotazioni

Questa sezione costruisce il calcolo dalle fondamenta. Ogni formula è seguita dalla spiegazione di ciascun componente.

### 3.1 Passo 1 — Trasformare la luce in un "irraggiamento di riferimento" (Colonna I)

L'irraggiamento arriva in **W/m²** (una potenza istantanea). Per confrontarlo con l'energia, convertiamo ogni lettura da 15 minuti in **kWh/m²** (una densità di energia):

```
irraggiamento_kWh/m²  =  irraggiamento_W/m²  ÷  4000
```

- **÷ 1000** converte i watt in kilowatt.
- **÷ 4** converte un'ora intera in un quarto d'ora (15 minuti).
- Combinato: **÷ 4000**.

Ci sono **due** piranometri (TX1 e TX3). Per ogni intervallo dobbiamo scegliere **un** valore di riferimento tra i due. La regola (il metodo "**MAX Condizionale**", predefinito) è:

```
se entrambi i sensori leggono 0:            riferimento = 0
altrimenti se uno solo legge 0:             riferimento = il sensore funzionante
altrimenti se |POA1 − POA3| / media  > 10% (tolleranza):   riferimento = il sensore PIÙ ALTO
altrimenti:                                 riferimento = la media dei due
```

- **Perché preferire il sensore più alto quando discordano?** Un piranometro sporco o in ombra legge **troppo basso**. Fare la media con una lettura errata bassa sottostimerebbe la luce e *gonfierebbe* il PR. Fidarsi del sensore più alto (pulito) è la scelta prudente.
- Nello strumento è selezionabile anche un metodo "**Media**" (media semplice dei due); l'impianto usa un metodo in modo coerente.

Infine si applica una **soglia minima di sole**:

```
se riferimento × 4000  <  50 W/m²:   riferimento = 0   (intervallo scartato)
```

I valori per-intervallo sopravvissuti formano la **Colonna I** nel file giornaliero. La loro somma giornaliera è il principale motore del denominatore:

```
Σ I  =  somma della Colonna I su tutti i 96 intervalli   (unità: kWh/m²)
```

> **Colonna I / Σ I** è la densità totale di energia solare utile del giorno — il "carburante" con cui l'impianto ha potuto lavorare.

### 3.2 Passo 2 — Misurare l'energia effettivamente prodotta

L'energia è rilevata in due modi:

**(a) Dagli inverter** — la potenza (kW) di ciascun inverter su un intervallo di 15 minuti diventa energia:

```
energia_inverter_kWh  =  potenza_inverter_kW  ×  0,25       (0,25 h = 15 min)
```

Sommata su tutti i 36 inverter e i 96 intervalli si ottiene l'**energia da inverter** del giorno.

**(b) Dal contatore** — la differenza tra letture cumulative consecutive del contatore:

```
energia_contatore_kWh  =  (lettura_attuale − lettura_precedente)  ×  1000
```

> Letture del contatore mancanti, a zero o "all'indietro" (un contatore cumulativo può solo salire) vengono automaticamente **riparate** per interpolazione tra le letture valide più vicine, ed evidenziate in arancione nel file giornaliero.

### 3.3 Passo 3 — PR Raw (non compensato)

L'energia teorica che la luce poteva sostenere è:

```
Energia attesa  =  P_nominale  ×  Σ I  =  12.625 kWp  ×  Σ I (kWh/m²)
```

> Intuizione: con luce standard (1000 W/m² = 1 kWh/m² all'ora), ogni kWp di pannello produce 1 kWh. Quindi nominale × densità-di-energia-solare = l'energia ideale.

```
                 Energia effettiva (inverter)
PR Raw (%)  =  --------------------------------  × 100
                    12.625  ×  Σ I
```

È onesto ma **impietoso**: se l'impianto è stato spento per ore, l'energia effettiva è bassa e il PR Raw crolla — anche se il fermo è stato ordinato dalla rete e non è un guasto dell'apparecchiatura.

### 3.4 Passo 4 — Perdite recuperabili

Per giudicare la qualità dell'*apparecchiatura*, stimiamo l'energia persa per eventi fuori dal normale funzionamento e la **riaggiungiamo**. In ogni intervallo di 15 minuti, per ogni inverter, **al massimo una** delle seguenti si applica (sono mutuamente esclusive):

**(a) Perdita per fermo (downtime)** — l'inverter è spento (< 1 kW) mentre c'è sole:

```
se altri inverter sullo stesso trasformatore stanno ancora lavorando:
      perdita_fermo = (potenza media degli inverter attivi) × 0,25
altrimenti (l'intero trasformatore è fermo):
      perdita_fermo = (POA_media ÷ 1000) × DC_inverter × target_PVSyst × 0,25
```

- Se i vicini funzionano, sono la miglior stima di ciò che l'inverter spento *avrebbe* prodotto.
- Se è tutto fermo, si ricorre alla stima fisica (sole × dimensione pannelli × PR di progetto).

**(b) Perdita per limitazione (curtailment)** — l'inverter **sta** producendo (≥ 1 kW) ma la rete l'ha limitato (rapporto limite < 0,875):

```
perdita_limitazione = max( 0,  min(attesa, limite_AC) − limite_AC × rapporto_limite )  × 0,25
```

- Cattura il divario tra quanto l'inverter avrebbe potuto produrre e il livello ridotto consentito dalla rete.
- **Importante:** la limitazione si conteggia solo quando l'inverter sta effettivamente producendo. Durante un fermo totale anche il segnale di rete legge ~0, ma quell'energia è già catturata come *fermo* — conteggiarla di nuovo sarebbe un doppio conteggio.

**(c) Perdita di rampa / recupero** — l'inverter sta producendo (≥ 1 kW), non limitato, in un intervallo **immediatamente prima o dopo un fermo totale dell'impianto**:

```
perdita_rampa = max( 0,  potenza_attesa − potenza_effettiva )  × 0,25
```

- Quando un impianto scatta e torna, l'intervallo di recupero mostra l'inverter già in rampa ma ancora sotto quanto la luce sosteneva. Questo registra tale ammanco, che il semplice test di fermo (che richiede una lettura quasi zero) altrimenti mancherebbe.

I tre si sommano nella perdita dell'inverter per quell'intervallo:

```
perdita_inverter = perdita_fermo + perdita_limitazione + perdita_rampa
```

> **potenza_attesa** = (POA_media ÷ 1000) × DC_inverter × target_PVSyst, limitata al limite AC utile dell'inverter (≈ 280,3 kW).

### 3.5 Passo 5 — PR Compensato

```
                  Energia effettiva  +  Σ perdite recuperabili
PR Compensato (%) =  --------------------------------------------  × 100
                            12.625  ×  Σ I
```

Riaggiungendo le perdite recuperabili al numeratore, il PR Compensato riflette come l'impianto rende **quando gli è permesso funzionare** — la misura più equa della qualità tecnica, e quella allineata alla garanzia contrattuale.

> **Regola di controllo:** un PR Compensato corretto non può superare in modo significativo il **100%** — non si può recuperare più di quanto la luce potesse sostenere. Un valore sopra il 100% segnala un problema di dati o di doppio conteggio (vedi FAQ).

### 3.6 PR per inverter

La stessa idea applicata a un singolo inverter, per individuare i sotto-performanti:

```
                        energia_inverter  +  perdita_inverter
PR per inverter (%) =  ---------------------------------------  × 100
                              DC_inverter  ×  Σ I
```

### 3.7 Disponibilità esterna

Quanta dell'energia potenziale è stata effettivamente erogata, in percentuale:

```
                         E
Disponibilità (%) =  -----------------  × 100
                      E + L1 + L2 + L3
```

- **E** = energia contabilizzata al contatore per il giorno.
- **L1, L2, L3** = perdite totali attribuite ai trasformatori TX1, TX2, TX3.
- Un giorno senza fermi dà 100%; un giorno con forti fermi (come l'esempio in §4.2) scende ben sotto.

### 3.8 Il target e la sua degradazione annua

I pannelli solari perdono lentamente efficienza con l'età. Il contratto (Allegato 9.1) lo riflette **riducendo il target PR garantito dello 0,4% ogni anno**, in modo composto:

```
Target(mese, anno)  =  base_PVSyst(mese)  ×  (1 − 0,4%) ^ n
```

- **base_PVSyst(mese)** — il target di progetto per quel mese (tabella in §2.3).
- **n** — il numero di anni contrattuali completati dall'avvio dell'impianto (**febbraio 2025**). Gli anni contrattuali vanno **febbraio → gennaio**.
  - **Anno 1** (feb 2025 – gen 2026): n = 0 → nessuna riduzione.
  - **Anno 2** (feb 2026 – gen 2027): n = 1 → fattore 0,996.
  - **Anno 3**: n = 2 → fattore 0,996² = 0,99202, e così via.

Lo strumento legge automaticamente anno e mese dal percorso della cartella del report (`…/YYYY MM/DD`) e applica il fattore corretto. Il target degradato alimenta sia il target visualizzato sia le stime di perdita basate sulla fisica in §3.4.

---

## 4. Esempio di calcolo passo passo

### 4.1 Un singolo intervallo di 15 minuti

Supponiamo che alle **15:15** i due piranometri leggano **POA1 = 990 W/m²** e **POA3 = 968 W/m²**, e sia giugno 2026 (Anno contrattuale 2).

**1. Conversione in densità di energia:**
```
POA1 = 990 ÷ 4000 = 0,24750 kWh/m²
POA3 = 968 ÷ 4000 = 0,24200 kWh/m²
```

**2. Scelta del riferimento (MAX Condizionale):**
```
media       = (0,24750 + 0,24200) / 2 = 0,24475
scostamento = |0,24750 − 0,24200| / 0,24475 = 2,2%   →  sotto la tolleranza del 10%
riferimento (Colonna I) = media = 0,24475 kWh/m²
```
Verifica soglia: 0,24475 × 4000 = 979 W/m² ≥ 50 → **mantenuto**.

**3. Target degradato per giugno 2026:**
```
n = 1  →  fattore = 0,996
Target = 83,3% × 0,996 = 82,97%   (base giugno PVSyst = 0,833)
```

**4. Se tutti i 36 inverter producessero normalmente**, questo intervallo contribuisce semplicemente con la sua energia da inverter al numeratore e con `0,24475` a Σ I. Nessuna perdita registrata.

**5. Se l'intero impianto fosse SPENTO in questo intervallo** (fermo totale) con lo stesso sole, la perdita per fermo di un inverter da 343,75 kWp sarebbe:
```
perdita_fermo = (979 ÷ 1000) × 343,75 × 0,8297 × 0,25
              = 0,979 × 343,75 × 0,8297 × 0,25
              ≈ 69,8 kWh
```
Su tutti i 36 inverter questo recupera ≈ 2.500 kWh per quel singolo intervallo — l'energia che l'impianto *avrebbe* prodotto.

### 4.2 Un giorno intero — giorno normale vs. giorno con fermo

**Giorno normale (es. 20 giugno 2026):** l'impianto funziona tutto il giorno.
- Σ I ≈ 10,13 kWh/m² · Energia effettiva ≈ 104.400 kWh · Perdite ≈ 0.
- PR Raw ≈ PR Compensato ≈ **83,0%** · Disponibilità = 100%.

**Giorno con fermo (12 giugno 2026):** l'impianto è scattato per ~5 ore (21 degli intervalli soleggiati avevano tutti gli inverter a 0).
- Energia effettiva ≈ 46.900 kWh (circa metà di un giorno normale) → **PR Raw ≈ 39,6%** (colpito duramente dal fermo).
- Perdite recuperabili per fermo ≈ 49.900 kWh vengono riaggiunte.
- **PR Compensato ≈ 81,7%** — mostra che l'*apparecchiatura* era a posto; la perdita era un fermo, non un difetto.
- Disponibilità ≈ 36% (l'impianto ha erogato solo circa un terzo del suo potenziale quel giorno).

> Questo giorno è anche l'esempio da manuale della protezione contro il doppio conteggio: durante il fermo il segnale di regolazione di rete leggeva ~0, che — prima della correzione — aggiungeva erroneamente una perdita di *limitazione* sopra la perdita di *fermo*, gonfiando il PR Compensato a un impossibile **108,5%**. Conteggiando la limitazione solo mentre l'inverter sta effettivamente producendo (§3.4b) si torna al corretto 81,7%.

---

## 5. Domande frequenti

**D1. Qual è la differenza tra PR Raw, SCADA, VCOM e Compensato?**
Il PR Raw è reale-vs-ideale senza correzioni. PR SCADA e PR VCOM sono i valori riportati dal sistema di controllo dell'impianto e dalla piattaforma indipendente VCOM (usati come controlli incrociati). Il PR Compensato riaggiunge l'energia persa per fermi e limitazioni di rete, isolando la qualità tecnica dell'impianto. La garanzia si valuta sulla vista compensata.

**D2. Perché si ignora l'energia sotto i 50 W/m²?**
All'alba/tramonto la luce è così debole che il rumore dei sensori e i valori minimi di produzione rendono il PR privo di significato (e si rischia di dividere per quasi zero). La soglia dei 50 W/m² rimuove questi intervalli inaffidabili.

**D3. Perché fidarsi del piranometro più alto quando i due discordano?**
Un sensore sporco o in ombra legge **troppo basso**. Includerlo nella media sottostimerebbe la luce e farebbe apparire il PR artificialmente alto. Usare la lettura più alta (pulita) è la scelta prudente e difendibile.

**D4. Il PR Compensato può superare il 100%?**
No — non legittimamente. Non si può recuperare più energia di quanta la luce potesse sostenere. Un valore sopra il 100% indica sempre un problema di dati o una perdita conteggiata due volte (es. la limitazione durante un fermo totale). Tali casi sono trattati come errori e corretti.

**D5. Perché "compensare" i fermi — il fermo non è un problema dell'impianto?**
La compensazione separa la *prestazione tecnica* (oggetto della garanzia dell'apparecchiatura) dagli *eventi di disponibilità* (limitazioni di rete, fermi esterni). Entrambi contano, ma sono riportati separatamente: PR Compensato per la qualità, Disponibilità per il tempo di funzionamento.

**D6. Perché il target scende ogni anno?**
I pannelli si degradano con l'età, quindi il contratto abbassa il PR garantito dello 0,4% all'anno. L'Anno 1 (feb 2025–gen 2026) usa i valori pieni di progetto; ogni febbraio successivo il target scende di un gradino.

**D7. Da dove vengono i numeri SCADA e VCOM nel file Madre?**
Quando i file mensili `KPI_Report_Daily.xls` (SCADA) e `Performance_ratio_vcom.csv` (VCOM) sono presenti, i loro valori di PR giornalieri vengono scritti direttamente nelle colonne "PR SCADA" e "PR VCOM" del riepilogo mensile, sostituendo le stime da formula.

**D8. Perché ci sono 96 valori al giorno?**
I dati sono campionati ogni 15 minuti; 24 ore × 4 = 96 intervalli. Ognuno rappresenta un quarto d'ora (0,25 h) di funzionamento.

---

## 6. Glossario

| Termine | Definizione |
|---------|-------------|
| **Performance Ratio (PR)** | Energia effettiva ÷ energia ideale per la luce ricevuta, in percentuale. |
| **POA (Plane Of Array)** | Irraggiamento misurato sul piano dei pannelli — la luce che essi effettivamente ricevono. |
| **Piranometro** | Sensore che misura l'irraggiamento solare in W/m². |
| **Irraggiamento** | Potenza solare istantanea per area (W/m²). |
| **kWp (kilowatt-picco)** | Potenza nominale del pannello/impianto sotto luce standard (1000 W/m²). |
| **SCADA** | Il sistema di controllo e acquisizione dati in campo dell'impianto. |
| **VCOM** | Una piattaforma di monitoraggio indipendente di terza parte. |
| **Curtailment (limitazione)** | Riduzione della produzione dell'impianto ordinata dalla rete. |
| **Fermo (downtime)** | Un inverter o trasformatore che non produce mentre c'è sole. |
| **PR Compensato** | PR Raw con riaggiunte le perdite recuperabili per fermo/limitazione. |
| **Σ I (somma irraggiamento di riferimento)** | Totale giornaliero dell'irraggiamento di riferimento per-intervallo, filtrato dalla soglia (kWh/m²). |
| **PVSyst** | Software di simulazione che fornisce i target PR mensili di progetto. |
| **Degradazione (0,4%/anno)** | Riduzione annua del target PR garantito con l'invecchiamento dei pannelli. |
| **Disponibilità** | Quota di energia potenziale effettivamente erogata (%). |

---

*Questo documento descrive la metodologia implementata in `PR_Calculator_GUI_v11.py` per l'impianto Mazara 01. È tarato sulla configurazione di quell'impianto ma i principi valgono per gli impianti FV in generale.*

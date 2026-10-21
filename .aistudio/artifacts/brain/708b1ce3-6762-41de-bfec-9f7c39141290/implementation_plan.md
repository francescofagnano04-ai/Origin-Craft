# OriginStudio Lab — Analisi Dati Scientifica, Plotting Multi-Serie e Curve Fitting

Applicazione web desktop-grade ispirata a **OriginLab**, progettata per importare tabelle di dati sperimentali (`.txt` / `.csv`), assegnare dinamicamente i ruoli delle colonne (`X`, `Y`), sovrapporre più curve nello stesso grafico, eliminare interattivamente i punti anomali (*Remove Bad Data Points*) direttamente dal grafico e dalla tabella, ed eseguire **Curve Fitting e Interpolazione** (Lineare, Polinomiale di grado $k$, Spline Cubica) generando nuove tabelle e curve campionate con $n$ punti su un intervallo $[C, D]$ definito dall'utente.

---

## User Review & Critical Decisions

> [!IMPORTANT]
> Le scelte progettuali confermate durante la fase preliminare sono integrate come requisiti vincolanti dell'architettura:

- **Modelli di Curve Fitting & Interpolazione Confermati**: Supporto dedicato e rigoroso per **Fit Lineare** ($y = a + bx$ con $R^2$, errore standard e intercetta/pendenza), **Fit Polinomiale di grado $k$** ($1 \le k \le 8$ tramite risoluzione ai minimi quadrati con eliminazione di Gauss-Jordan) e **Spline Cubica Naturale** (interpolazione $C^2$ continua per tratti), tutti dotati di generatore di curva ricalcolata su $n$ punti equidistanti nell'intervallo personalizzato $[C, D]$.
- **Spazio di Lavoro Multi-Finestra Desktop**: Interfaccia a pannelli flessibili e affiancabili (*Workbench Multi-Window*) con **Project Explorer** a sinistra, **Fogli Dati (Workbook)** e **Finestre Grafico (Graph Layer)** visualizzabili in parallelo (split-screen orizzontale/verticale o massimizzazione singola finestra) più **Pannello Ispettore (Proprietà Stile, Assi e Fitting)** a destra.
- **Esportazione Completa**: Esportazione delle tabelle dati pulite o generate dal fit in formato `.txt` (delimitato da tabulazioni in stile Origin) e `.csv`, esportazione vettoriale `.svg` e raster `.png` ad alta risoluzione dei grafici, e salvataggio/ripristino dell'intera sessione di lavoro in `.json`.

---

## 1. Overview & Core Concept

- **What It Does**: Offre un ambiente di laboratorio numerico immediato in cui il ricercatore o lo studente carica un file `.txt` (con separatori automatici: tab, spazio, virgola, punto e virgola e gestione sia del punto `.` che della virgola `,` come separatore decimale), definisce quali colonne fungono da asse $X$ e quali da asse $Y$, traccia istantaneamente una o più serie sullo stesso grafico, pulisce gli outlier con un clic sul grafico (aggiornando in tempo reale il foglio dati e i fit attivi) e calcola curve fittate su domini $[C, D]$ arbitrari con numero di campioni $n$ a scelta.
- **Target Audience / Persona**: Studenti universitari, fisici, chimici, ingegneri e ricercatori di laboratorio che necessitano delle funzioni cardine di OriginLab senza installazioni pesanti, con controllo matematico esatto sui dati esportati.
- **Key Value**: Sincronizzazione bidirezionale immediata tra **Tabella Dati $\leftrightarrow$ Grafico Interattivo $\leftrightarrow$ Tabella Valori Fittati $[C, D]$**, eliminando il passaggio manuale tra software di calcolo e software di plotting.

---

## 2. User Experience & Visual Design

### Key User Flows

1. **Importazione & Ruolo Colonne (`X` / `Y` / `Ignora`)**:
   - L'utente trascina un file `.txt`/`.csv`, incolla dati grezzi dagli appunti, oppure carica un dataset scientifico di esempio pre-configurato.
   - Nell'intestazione di ciascuna colonna del Workbook (es. `A(X)`, `B(Y)`, `C(Y)`), l'utente può commutare con un clic il ruolo della colonna tra **X**, **Y** o **Ignora**, modificare il **Nome (Long Name)** e l'**Unità di Misura (Units)**, che aggiornano automaticamente le etichette dinamiche degli assi del grafico.
2. **Plotting Multi-Colonna & Personalizzazione Grafica**:
   - Selezionando una o più colonne `Y` (associate automaticamente alla colonna `X` di riferimento), l'utente le invia a un grafico esistente o a un nuovo grafico.
   - Dal pannello **Stile & Assi**, modifica per ogni singola curva: colore, spessore del tratto (`1px`–`6px`), stile linea (continua, tratteggiata, punteggiata), modalità di connessione (**Nessuna / Solo Punti**, **Lineare punto-punto**, **Spline morbida**), forma e dimensione dei marker (cerchio, quadrato, triangolo, rombo, croce) e testo della legenda.
   - Personalizza gli assi $X$ e $Y$: titoli dinamici o manuali, limiti `[Min, Max]` automatici o fissati, griglia principale/secondaria e scala lineare.
3. **Strumento "Remove Bad Data Points" (Rimozione Punti Errati)**:
   - Attivando il cursore **Remove Bad Points** (icona mirino rosso nella barra strumenti del grafico), il cursore diventa un mirino di precisione (`crosshair`) che evidenzia il punto sperimentale più vicino con un anello di ispezione e mostra in tempo reale `(Indice Riga, X, Y)`.
   - Un clic sul punto anomalo (o una selezione rettangolare *Box Mask/Delete* su più punti errati) **elimina immediatamente il valore dalla tabella dati sottostante** (e aggiorna all'istante il grafico e qualsiasi curva di fit collegata). Include una cronologia **Annulla (Undo)** e un registro dei punti rimossi per ripristinare eventuali clic accidentali.
4. **Curve Fitting & Generazione Curva tra $[C, D]$ con $n$ Punti**:
   - Nel pannello **Curve Fitting (Stile Origin)**, l'utente seleziona la serie da fittare e il modello:
     - **Lineare**: $y = a + b x$
     - **Polinomiale**: $y = c_0 + c_1 x + c_2 x^2 + \dots + c_k x^k$ (con selettore del grado $k$)
     - **Spline Cubica**: Interpolazione cubica naturale esatta attraverso i nodi sperimentali
   - L'utente imposta i parametri di generazione della curva fittata:
     - **Estremo iniziale $C$ ($X_{\min}$)** (pre-compilato con il minimo $X$ della serie, ma liberamente modificabile per interpolare un sotto-intervallo o estrapolare oltre i dati originali $[A, B]$).
     - **Estremo finale $D$ ($X_{\max}$)** (pre-compilato con il massimo $X$ della serie, liberamente modificabile).
     - **Numero di punti $n$** (es. da $100$ punti originali tra $A$ e $B$ a $n = 250$ o $n = 50$ punti equidistanti tra $C$ e $D$).
   - Cliccando **Esegui Fit & Genera Curva**:
     1. La curva fittata viene sovrapposta al grafico con stile personalizzabile.
     2. Viene mostrata la scheda dei parametri statistici (equazione esplicita, coefficienti calcolati, $R^2$, RMSE).
     3. Viene generata automaticamente nel Workbook una **nuova coppia di colonne (o un nuovo foglio `Fit_Result`) contenente esattamente gli $n$ valori $(X_{\text{fit}}, Y_{\text{fit}})$ tra $C$ e $D$**, pronta per essere ispezionata o esportata in `.txt`.

### Visual Identity & Theme

- **Aesthetic Direction**: *Precision Laboratory Workbench* — estetica tecnica, pulita e ad alta densità informativa, ispirata agli strumenti scientifici da banco e ai moderni IDE di analisi dati.
- **Color Palette & Mood (60-30-10)**:
  - **60% Dominant Neutral Canvas**: Grigio ardesia chiarissimo da laboratorio (`#F8FAFC`) per l'area di lavoro e bianco ottico (`#FFFFFF`) per il foglio di calcolo e il canvas del grafico, garantendo leggibilità ottimale per la pubblicazione scientifica.
  - **30% Structural Surfaces**: Pannelli laterali e barre degli strumenti in grigio tecnico (`#F1F5F9`), bordi sottili da `1px` (`#E2E8F0`) e intestazioni tabella differenziate per ruolo colonna.
  - **10% Accent Budget**: Blu Cobalto Calibrato (`#2563EB`) per i controlli attivi e le selezioni; Rosso Cremisi (`#DC2626`) per la modalità *Remove Bad Data Points* e le curve di fit principali; Verde Smeraldo (`#16A34A`) per la conferma di convergenza del fit ($R^2$).
- **Typography & Hierarchy**:
  - **UI & Intestazioni**: `Plus Jakarta Sans` (pesi 500 e 600) per un'interfaccia nitida e compatta.
  - **Dati Numerici, Tabelle, Coordinate & Coefficienti di Fit**: `JetBrains Mono` con `font-variant-numeric: tabular-nums` obbligatorio su tutte le celle del foglio dati, i cursori del grafico e le equazioni di regressione per un allineamento decimale perfetto.
- **Component Styling & Layout**:
  - Viewport desktop full-screen (`100vw x 100vh`, ottimizzato per `1440px+`) con layout multi-finestra ancorato: barra superiore singola conforme al contratto a 3 zone, barra degli strumenti operativa del workspace, e griglia centrale flessibile (Modalità *Affiancata Tabella + Grafico*, *Griglia Multi-Finestra*, o *Vista Massimizzata*) affiancata dall'Ispettore contestuale.

---

## 3. Key Product Decisions & Trade-Offs

- **Decision 1: Motore Grafico SVG ad Alta Precisione Custom (Zero Dipendenze Opache)**
  - *Chosen Approach*: Costruzione di un motore di rendering grafico SVG scientifico su misura con coordinate matematiche esatte, trasformazioni lineari/zoom, calcolo delle tangenti Spline, mirino interattivo *Data Reader / Bad Point Remover*, selezione rettangolare (*rubber-band box*) ed esportazione diretta del nodo SVG/Canvas in `.svg` e `.png`.
  - *Why*: Le librerie di charting generiche (come Recharts o Chart.js) non consentono di cliccare con precisione sub-pixel su un singolo punto di una serie multi-colonna per rimuoverlo dalla riga specifica della matrice dati originale, né permettono di sovrapporre un dataset sperimentale di $100$ punti su $[A, B]$ con una curva analitica campionata di $n$ punti su un dominio diverso $[C, D]$ mantenendo un'esportazione vettoriale `.svg` pulita da pubblicazione.
- **Decision 2: Sincronizzazione Reattiva tra Foglio Dati, Rimozione Punti e Curve di Fit**
  - *Chosen Approach*: Architettura basata su stato unificato per cui ogni punto eliminato con lo strumento *Remove Bad Data Points* modifica direttamente il foglio dati (con opzione tra *Elimina cella/riga* o *Svuota valore Y della serie selezionata* e stack *Undo* completo) e ricalcola istantaneamente i coefficienti del fit attivo se l'aggiornamento automatico è abilitato.
  - *Why*: Rispecchia fedelmente il flusso di lavoro di OriginLab, in cui la pulizia degli outlier sul grafico serve a migliorare il fit e a ottenere contemporaneamente la tabella dati ripulita.
- **Decision 3: Parser Intelligente per File `.txt` e `.csv` Scientifici**
  - *Chosen Approach*: Rilevamento automatico del delimitatore (Tab `\t`, spazi multipli, `;`, `,`), salto automatico di eventuali righe di commento iniziali (`#`, `%`, `//`), riconoscimento automatico delle intestazioni e conversione trasparente dei numeri con virgola decimale europea (`3,1415` $\rightarrow$ `3.1415`) o notazione scientifica (`1.25e-4`).
  - *Why*: I file `.txt` generati da oscilloscopi, spettrometri e banchi prova in Italia e in Europa mescolano spesso tabulazioni, spazi e virgole decimali.

---

## 4. Technical Architecture & Data Strategy

### Architecture & Component Diagram

```
┌─────────────────────────────────────────────────────────────────────────────────────┐
│                          TOP BAR & WORKSPACE TOOLBAR                                │
│  [OriginStudio Lab]   [Fogli Dati · Grafici · Curve Fit · Esporta]   [Importa TXT]  │
│  Toolbar: [Cursore Standard] [Lettore Coordinate] [Remove Bad Points] [Undo/Redo]   │
└───────────────┬─────────────────────────────────────────┬───────────────────────────┘
                │                                         │
                ▼                                         ▼
┌───────────────────────────────┐         ┌───────────────────────────────────────────┐
│   PROJECT & WINDOW MANAGER    │         │       CENTRAL DOCKABLE WORKSPACE          │
│                               │         │  (Split-View / Multi-Window / Maximized)  │
│  ├─ Workbooks (Fogli Dati)    │         │                                           │
│  │   ├─ Sheet 1: Raw Data     │◄───────►│  ┌─────────────────┐ ┌─────────────────┐  │
│  │   └─ Sheet 2: Fit [C, D]   │         │  │ WORKBOOK GRID   │ │ GRAPH CANVAS    │  │
│  │                            │         │  │ Col A(X) | B(Y) │ │ - Serie Multi-Y │  │
│  ├─ Graphs (Finestre Grafico) │         │  │ Units / Labels  │ │ - Bad Pt Click  │  │
│  │   └─ Graph 1: Multi-Curve  │         │  │ Tabular Cells   │ │ - Fit Overlay   │  │
│  │                            │         │  └────────┬────────┘ └────────┬────────┘  │
│  └─ Fit Reports & History     │         └───────────┼───────────────────┼───────────┘
└───────────────────────────────┘                     │                   │
                                                      ▼                   ▼
┌─────────────────────────────────────────────────────────────────────────────────────┐
│                   RIGHT CONTEXTUAL INSPECTOR & MATHEMATICAL ENGINE                  │
│                                                                                     │
│  ┌─────────────────────────┐  ┌─────────────────────────┐  ┌─────────────────────┐  │
│  │ 1. SERIE & ASSI         │  │ 2. CURVE FITTING [C, D] │  │ 3. BAD POINTS LOG   │  │
│  │ - Seleziona Colonne X/Y │  │ - Lineare / Polin. (k)  │  │ - Modalità: Click o │  │
│  │ - Etichette Dinamiche   │  │   / Spline Cubica       │  │   Box Selection     │  │
│  │ - Colori, Spessore,     │  │ - Dominio: da C a D     │  │ - Elenco punti      │  │
│  │   Stile Linea e Marker  │  │ - Punti generati: n     │  │   rimossi + Riprist.│  │
│  │ - Connessione: Lineare  │  │ - Output: Curva + Nuovo │  │                     │  │
│  │   vs Spline vs Punti    │  │   Foglio/Colonne (n pt) │  │                     │  │
│  └─────────────────────────┘  └─────────────────────────┘  └─────────────────────┘  │
└─────────────────────────────────────────────────────────────────────────────────────┘
```

### Data Model & State

- **`WorkbookSheet`**:
  - `id`, `name`
  - `columns`: Array di `DataColumn` (`id`, `name`, `longName`, `unit`, `role: 'X' | 'Y' | 'IGNORE'`, `values: (number | null)[]`)
  - Supporta un numero arbitrario di colonne `X` e `Y` (ogni colonna `Y` si associa automaticamente alla colonna `X` precedente nel foglio, esattamente come lo standard `XYY` o `XYXY` di OriginLab, oppure a una colonna `X` scelta esplicitamente dall'utente).
- **`GraphWindow`**:
  - `id`, `title`, `subtitle`
  - `xAxis`, `yAxis`: `{ label, autoLabel: boolean, min: number | 'auto', max: number | 'auto', showGrid: boolean, gridSubdivisions: number }`
  - `series`: Array di `PlotSeries` (`id`, `sheetId`, `xColId`, `yColId`, `label`, `color`, `lineWidth`, `lineStyle: 'solid' | 'dashed' | 'dotted'`, `connection: 'none' | 'linear' | 'spline'`, `markerShape: 'circle' | 'square' | 'triangle' | 'diamond' | 'none'`, `markerSize`, `visible`)
  - `fits`: Array di `FittedCurve` (`id`, `sourceSeriesId`, `model: 'linear' | 'polynomial' | 'spline'`, `polyDegree`, `rangeC`, `rangeD`, `numPointsN`, `coefficients`, `rSquared`, `rmse`, `equationText`, `generatedPoints: {x, y}[]`, `outputSheetId`, `color`, `lineWidth`, `lineStyle`)
- **`RemovedPointRecord`**:
  - Traccia `{ id, timestamp, sheetId, rowIndex, xColId, yColId, oldX, oldY, seriesLabel }` per consentire l'annullamento singolo o multiplo dei punti rimossi tramite *Remove Bad Data Points*.

### Interactive Component & State Mapping

- **Parser & Editor Tabellare (`WorkbookGrid`)**:
  - Modifica inline di qualsiasi cella numerica, aggiunta/rimozione di righe e colonne, cambio rapido del ruolo `X` / `Y` / `Ignora` nell'intestazione, modifica di `Nome` e `Unità`, e pulsante rapido **"Traccia Colonne Selezionate nel Grafico"**.
- **Strumento `Remove Bad Data Points` sul Canvas (`GraphCanvas`)**:
  - Quando attivo, il movimento del mouse calcola la distanza euclidea normalizzata in pixel verso tutti i punti visibili delle serie attive.
  - Il punto sotto il mirino viene evidenziato in rosso con anteprima della riga corrispondente nella tabella dati.
  - Al clic (o al rilascio del rettangolo di selezione multipla), il valore viene rimosso dalla colonna `Y` (o l'intera riga viene eliminata se la serie è l'unica del foglio, in base all'opzione scelta), lo storico `removedPoints` registra l'operazione e i fit collegati si aggiornano automaticamente.
- **Motore Matematico di Curve Fitting (`FittingEngine`)**:
  - **Fit Lineare**: Calcolo analitico esatto di pendenza $b$, intercetta $a$, coefficiente di determinazione $R^2$ e deviazione standard dei residui.
  - **Fit Polinomiale di grado $k$**: Costruzione della matrice di Vandermonde $V^T V c = V^T y$ e risoluzione numerica con pivoting parziale per stabilità numerica, restituendo i coefficienti $c_0, \dots, c_k$, l'equazione formattata e $R^2$.
  - **Spline Cubica Naturale**: Risoluzione del sistema tridiagonale per le derivate seconde $M_i$ sui nodi ordinati $(x_i, y_i)$ con condizioni al contorno naturali ($S''(x_0) = S''(x_m) = 0$) ed estrapolazione tangente lineare fuori dal dominio dei nodi.
  - **Campionamento su $[C, D]$ con $n$ punti**: Genera il vettore $X_{\text{fit}}[j] = C + j \cdot \frac{D - C}{n - 1}$ per $j = 0, \dots, n-1$, valuta $Y_{\text{fit}}[j] = f(X_{\text{fit}}[j])$, disegna la curva continua sul grafico e crea/aggiorna un foglio dati dedicato con gli $n$ punti esportabili in `.txt` o `.csv`.

# Analisi Regressione Multipla

Questo è uno strumento per eseguire analisi di regressione multipla su dati contenuti in file Excel. Il programma offre un'interfaccia grafica che consente di selezionare un file Excel, scegliere le colonne da utilizzare per la regressione e visualizzare i risultati, inclusi i modelli matematici, i parametri stimati, e varie metriche di performance per ogni modello.

## Funzionalità

- **Selezione del Separatore Decimale:** Consente di scegliere se utilizzare il punto (`.`) o la virgola (`,`) come separatore decimale.
- **Caricamento di File Excel:** Permette di caricare file Excel e selezionare le colonne X e Y per la regressione.
- **Esecuzione della Regressione:** Supporta quattro modelli di regressione:
  - Lineare
  - Esponenziale
  - Logaritmico
  - Legge di potenza
- **Visualizzazione dei Risultati:** Mostra il miglior modello selezionato, la formula matematica, i parametri stimati, e le formule Excel per applicare i risultati direttamente in un foglio Excel. Vengono inoltre visualizzate le metriche di performance come R², MAE, MSE, e RMSE per ogni modello.

## Requisiti

- **Python 3.x**
- **Librerie Python:**
  - `tkinter` (per l'interfaccia grafica)
  - `pandas` (per la gestione dei dati)
  - `numpy` (per i calcoli numerici)
  - `scipy` (per l'ottimizzazione e il fitting dei modelli)
  - `sklearn` (per le metriche di valutazione)

### Installazione delle Dipendenze

Per installare automaticamente tutte le dipendenze necessarie, puoi utilizzare il file `requirements.txt`. Esegui i seguenti comandi:

1. Clona o scarica il repository.
2. Apri una terminale nella cartella del progetto.
3. Esegui il comando:

```bash
pip install -r requirements.txt

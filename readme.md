# Excel Regression Analysis Tool 🔍 

[English](#english) | [Italiano](#italiano)

---

<a name="english"></a>
## 🇬🇧 English

### Description
Excel Regression Analysis Tool is a user-friendly desktop application that performs various types of regression analysis on Excel data. The tool supports multiple regression models and provides detailed statistical metrics to help you choose the best fit for your data.

### Features
- Support for multiple regression models:
  - Linear regression
  - Exponential regression
  - Logarithmic regression
  - Power Law regression
- Easy Excel file import
- Automatic calculation of key metrics (R², MAE, MSE, RMSE)
- Excel formula generation for the best-fit model
- Support for different decimal separators (point/comma)
- Clean and intuitive graphical interface

### Installation
1. Clone this repository or download the source code
2. Create a virtual environment (recommended):
   ```bash
   python -m venv venv
   source venv/bin/activate  # On Windows: venv\Scripts\activate
   ```
3. Install the required packages:
   ```bash
   pip install -r requirements.txt
   ```

### Usage
1. Run the application:
   ```bash
   python regression_tool.py
   ```
2. Click "Select Excel File" to load your data
3. Choose the decimal separator according to your Excel file format
4. Select the X (independent) and Y (dependent) columns from your data
5. Click "Perform Regression" to run the analysis
6. Review the results, including:
   - Best model selection
   - R² value
   - Mathematical formula
   - Excel formula
   - Comparison metrics for all models

### Requirements
- Python 3.7 or higher
- See requirements.txt for detailed package dependencies

---

<a name="italiano"></a>
## 🇮🇹 Italiano

### Descrizione
Excel Regression Analysis Tool è un'applicazione desktop user-friendly che esegue vari tipi di analisi di regressione su dati Excel. Lo strumento supporta diversi modelli di regressione e fornisce metriche statistiche dettagliate per aiutarti a scegliere il modello migliore per i tuoi dati.

### Funzionalità
- Supporto per diversi modelli di regressione:
  - Regressione lineare
  - Regressione esponenziale
  - Regressione logaritmica
  - Regressione con legge di potenza
- Facile importazione di file Excel
- Calcolo automatico delle metriche chiave (R², MAE, MSE, RMSE)
- Generazione della formula Excel per il modello migliore
- Supporto per diversi separatori decimali (punto/virgola)
- Interfaccia grafica pulita e intuitiva

### Installazione
1. Clona questo repository o scarica il codice sorgente
2. Crea un ambiente virtuale (consigliato):
   ```bash
   python -m venv venv
   source venv/bin/activate  # Su Windows: venv\Scripts\activate
   ```
3. Installa i pacchetti richiesti:
   ```bash
   pip install -r requirements.txt
   ```

### Utilizzo
1. Avvia l'applicazione:
   ```bash
   python regression_tool.py
   ```
2. Clicca su "Select Excel File" per caricare i tuoi dati
3. Scegli il separatore decimale in base al formato del tuo file Excel
4. Seleziona le colonne X (indipendente) e Y (dipendente) dai tuoi dati
5. Clicca su "Perform Regression" per eseguire l'analisi
6. Esamina i risultati, tra cui:
   - Selezione del modello migliore
   - Valore R²
   - Formula matematica
   - Formula Excel
   - Metriche di confronto per tutti i modelli

### Requisiti
- Python 3.7 o superiore
- Vedi requirements.txt per le dipendenze dettagliate dei pacchetti
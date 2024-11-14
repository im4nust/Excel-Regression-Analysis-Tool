import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import numpy as np
from scipy.optimize import curve_fit
from sklearn.metrics import r2_score, mean_absolute_error, mean_squared_error
import math

# Creazione della finestra principale
app = tk.Tk()
app.title("Analisi Regressione Multipla")
app.geometry("600x500")

# Menu a discesa per selezionare il separatore decimale
decimal_separator_var = tk.StringVar(value="Point")  # Variabile per il separatore decimale

decimal_separator_label = tk.Label(app, text="Seleziona separatore decimale:")
decimal_separator_label.pack(pady=5)
decimal_separator_menu = tk.OptionMenu(app, decimal_separator_var, "Point", "Comma")
decimal_separator_menu.pack(pady=5)

# Dictionario dei modelli in markdown
model_formulas = {
    "Linear": "y = a * x + b",
    "Exponential": "y = a * e^(b * x)",
    "Logarithmic": "y = a * ln(x) + b",
    "Power Law": "y = a * x^b"
}

# Funzioni di regressione
def linear_model(x, a, b):
    return a * x + b

def exponential_model(x, a, b):
    return a * np.exp(b * x)

def logarithmic_model(x, a, b):
    return a * np.log(x) + b

def power_law_model(x, a, b):
    return a * np.power(x, b)

models = {
    "Linear": linear_model,
    "Exponential": exponential_model,
    "Logarithmic": logarithmic_model,
    "Power Law": power_law_model
}

# Funzione per formattare i numeri in base al separatore decimale scelto
def format_number(value):
    if decimal_separator_var.get() == "Comma":
        return str(value).replace('.', ',')  # Usa la virgola per il separatore
    else:
        return str(value)  # Usa il punto per il separatore

# Funzione per ottenere la formula Excel con il formato giusto
def get_excel_formula(model, params):
    if model == "Linear":
        return f"=({format_number(params[0])} * A1) + {format_number(params[1])}"
    elif model == "Exponential":
        return f"=({format_number(params[0])} * EXP({format_number(params[1])} * A1))"
    elif model == "Logarithmic":
        return f"=({format_number(params[0])} * LN(A1)) + {format_number(params[1])}"
    elif model == "Power Law":
        return f"=({format_number(params[0])} * (A1^{format_number(params[1])}))"

# Funzione per eseguire la regressione e determinare il miglior modello
def perform_regression(df, x_column, y_column):
    x = df[x_column].values
    y = df[y_column].values
    best_model = None
    best_r2 = -float('inf')
    best_params = None
    best_formula = ""

    results = []
    
    for name, model in models.items():
        try:
            # Stima i parametri
            params, _ = curve_fit(model, x, y, maxfev=10000)
            y_pred = model(x, *params)
            r2 = r2_score(y, y_pred)
            mae = mean_absolute_error(y, y_pred)
            mse = mean_squared_error(y, y_pred)
            rmse = math.sqrt(mse)

            results.append((name, r2, params, mae, mse, rmse))
            
            # Aggiorna se è il miglior modello trovato
            if r2 > best_r2:
                best_r2 = r2
                best_model = name
                best_params = params
                if name == "Linear":
                    best_formula = f"y = {params[0]:.4f} * x + {params[1]:.4f}"
                elif name == "Exponential":
                    best_formula = f"y = {params[0]:.4f} * e^({params[1]:.4f} * x)"
                elif name == "Logarithmic":
                    best_formula = f"y = {params[0]:.4f} * ln(x) + {params[1]:.4f}"
                elif name == "Power Law":
                    best_formula = f"y = {params[0]:.4f} * x^{params[1]:.4f}"
        
        except Exception as e:
            print(f"Errore nel modello {name}: {e}")
    
    return best_model, best_r2, best_params, best_formula, results

# Funzione per aprire e caricare il file Excel
def load_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    if file_path:
        df = pd.read_excel(file_path)
        x_column_menu['menu'].delete(0, 'end')
        y_column_menu['menu'].delete(0, 'end')
        
        # Aggiorna le opzioni dei menu a discesa
        for column in df.columns:
            x_column_menu['menu'].add_command(label=column, command=tk._setit(x_column_var, column))
            y_column_menu['menu'].add_command(label=column, command=tk._setit(y_column_var, column))
        
        app.df = df
        messagebox.showinfo("File caricato", "Il file è stato caricato con successo!")

# Funzione per eseguire la regressione e visualizzare il risultato
def run_regression():
    x_column = x_column_var.get()
    y_column = y_column_var.get()

    if not x_column or not y_column:
        messagebox.showerror("Errore", "Seleziona sia la colonna X che Y.")
        return
    
    best_model, best_r2, best_params, best_formula, results = perform_regression(app.df, x_column, y_column)
    
    # Testo dei risultati
    result_text = f"Miglior modello selezionato: {best_model}\n\n"
    result_text += f"R²: {format_number(best_r2)}\n"
    result_text += f"Formula matematica: {model_formulas[best_model]}\n"
    result_text += f"Coefficiente/i: {', '.join([format_number(param) for param in best_params])}\n"
    result_text += f"Formula Excel: {get_excel_formula(best_model, best_params)}\n\n"
    
    result_text += "Metriche per tutti i modelli:\n"
    for name, r2, params, mae, mse, rmse in results:
        result_text += f"{name}: R^2={format_number(r2)}, MAE={format_number(mae)}, MSE={format_number(mse)}, RMSE={format_number(rmse)}\n"
    
    # Visualizzazione dei risultati in un Text widget
    result_text_widget.delete(1.0, tk.END)  # Pulisce il Text widget
    result_text_widget.insert(tk.END, result_text)

# Pulsante per selezionare il file
load_button = tk.Button(app, text="Seleziona File Excel", command=load_file)
load_button.pack(pady=10)

# Menu a discesa per la selezione delle colonne X e Y
x_column_var = tk.StringVar()
y_column_var = tk.StringVar()

x_column_label = tk.Label(app, text="Seleziona colonna per X:")
x_column_label.pack()
x_column_menu = ttk.OptionMenu(app, x_column_var, "")
x_column_menu.pack()

y_column_label = tk.Label(app, text="Seleziona colonna per Y:")
y_column_label.pack()
y_column_menu = ttk.OptionMenu(app, y_column_var, "")
y_column_menu.pack()

# Pulsante per eseguire la regressione
run_button = tk.Button(app, text="Calcola Regressione", command=run_regression)
run_button.pack(pady=20)

# Text widget per mostrare i risultati
result_text_widget = tk.Text(app, width=70, height=15, wrap=tk.WORD)
result_text_widget.pack(pady=10)
result_text_widget.config(state=tk.NORMAL)  # Rende il widget modificabile

app.mainloop()

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import numpy as np
from scipy.optimize import curve_fit
from sklearn.metrics import r2_score, mean_absolute_error, mean_squared_error
import math

# Creazione della finestra principale
app = tk.Tk()
app.title("Excel Regression Analysis Tool")
app.geometry("600x500")

# Menu a discesa per selezionare il separatore decimale
decimal_separator_var = tk.StringVar(value="Point")  

decimal_separator_label = tk.Label(app, text="Select decimal separator:")
decimal_separator_label.pack(pady=5)
decimal_separator_menu = tk.OptionMenu(app, decimal_separator_var, "Point", "Comma")
decimal_separator_menu.pack(pady=5)

# Dizionario dei modelli in markdown
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

def preprocess_data(df, x_column, y_column):
    """Preprocessa i dati per la regressione, considerando solo le colonne X e Y selezionate"""
    try:
        # Estrai solo le colonne necessarie
        df_subset = df[[x_column, y_column]].copy()
        
        # Converti le colonne in numeriche, gestendo eventuali errori di formato
        df_subset[x_column] = pd.to_numeric(df_subset[x_column].astype(str).str.replace(',', '.'), errors='coerce')
        df_subset[y_column] = pd.to_numeric(df_subset[y_column].astype(str).str.replace(',', '.'), errors='coerce')
        
        # Rimuovi righe con valori non numerici o mancanti
        df_subset = df_subset.dropna()
        
        # Rimuovi valori zero per modelli logaritmici e power law
        df_subset = df_subset[(df_subset[x_column] > 0) & (df_subset[y_column] > 0)]
        
        # Rimuovi outliers usando IQR
        Q1_x = df_subset[x_column].quantile(0.25)
        Q3_x = df_subset[x_column].quantile(0.75)
        IQR_x = Q3_x - Q1_x
        df_subset = df_subset[~((df_subset[x_column] < (Q1_x - 1.5 * IQR_x)) | 
                               (df_subset[x_column] > (Q3_x + 1.5 * IQR_x)))]
        
        Q1_y = df_subset[y_column].quantile(0.25)
        Q3_y = df_subset[y_column].quantile(0.75)
        IQR_y = Q3_y - Q1_y
        df_subset = df_subset[~((df_subset[y_column] < (Q1_y - 1.5 * IQR_y)) | 
                               (df_subset[y_column] > (Q3_y + 1.5 * IQR_y)))]
        
        if df_subset.empty:
            messagebox.showwarning("Warning", "No valid data points after preprocessing!")
            return None
            
        return df_subset
        
    except Exception as e:
        messagebox.showerror("Error", f"Error in data preprocessing: {str(e)}")
        return None

def clean_data(df):
    if df is None:
        return None
    try:
        # Rimuovere le righe con NaN o inf
        df = df.replace([np.inf, -np.inf], np.nan)
        df = df.dropna()
        return df
    except Exception as e:
        messagebox.showerror("Error", f"Error cleaning data: {str(e)}")
        return None

def format_number(value):
    try:
        if isinstance(value, (int, float)):
            if decimal_separator_var.get() == "Comma":
                return f"{value:.6f}".replace('.', ',')
            return f"{value:.6f}"
        return str(value)
    except:
        return str(value)

def get_excel_formula(model, params):
    try:
        if model == "Linear":
            return f"=({format_number(params[0])} * A1) + {format_number(params[1])}"
        elif model == "Exponential":
            return f"=({format_number(params[0])} * EXP({format_number(params[1])} * A1))"
        elif model == "Logarithmic":
            return f"=({format_number(params[0])} * LN(A1)) + {format_number(params[1])}"
        elif model == "Power Law":
            return f"=({format_number(params[0])} * (A1^{format_number(params[1])}))"
    except Exception as e:
        return f"Error generating Excel formula: {str(e)}"

def load_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    if file_path:
        try:
            # Leggi solo le intestazioni delle colonne senza caricare tutti i dati
            df_headers = pd.read_excel(file_path, nrows=0)
            columns = df_headers.columns.tolist()
            
            # Aggiorna i menu a discesa con le colonne disponibili
            x_column_menu['menu'].delete(0, 'end')
            y_column_menu['menu'].delete(0, 'end')
            
            for column in columns:
                x_column_menu['menu'].add_command(label=column, command=tk._setit(x_column_var, column))
                y_column_menu['menu'].add_command(label=column, command=tk._setit(y_column_var, column))
            
            # Salva il percorso del file per caricarlo completamente solo quando necessario
            app.file_path = file_path
            messagebox.showinfo("Success", "File loaded successfully!")
        except Exception as e:
            messagebox.showerror("Error", f"Error loading file: {str(e)}")

def perform_regression(df, x_column, y_column):
    if df is None or df.empty:
        messagebox.showerror("Error", "No valid data to perform regression.")
        return None, None, None, None, []

    try:
        x = df[x_column].values
        y = df[y_column].values
        
        if len(x) < 2 or len(y) < 2:
            messagebox.showerror("Error", "Not enough valid data points for regression.")
            return None, None, None, None, []

        best_model = None
        best_r2 = -float('inf')
        best_params = None
        best_formula = ""
        results = []

        for name, model in models.items():
            try:
                # Stima i parametri con bounds e better initial guesses
                if name in ["Exponential", "Power Law"]:
                    params, _ = curve_fit(model, x, y, p0=[1.0, 1.0], maxfev=10000)
                else:
                    params, _ = curve_fit(model, x, y, maxfev=10000)
                
                y_pred = model(x, *params)
                r2 = r2_score(y, y_pred)
                mae = mean_absolute_error(y, y_pred)
                mse = mean_squared_error(y, y_pred)
                rmse = math.sqrt(mse)

                results.append((name, r2, params, mae, mse, rmse))
                
                if r2 > best_r2:
                    best_r2 = r2
                    best_model = name
                    best_params = params
                    
                    if name == "Linear":
                        best_formula = f"y = {format_number(params[0])} * x + {format_number(params[1])}"
                    elif name == "Exponential":
                        best_formula = f"y = {format_number(params[0])} * e^({format_number(params[1])} * x)"
                    elif name == "Logarithmic":
                        best_formula = f"y = {format_number(params[0])} * ln(x) + {format_number(params[1])}"
                    elif name == "Power Law":
                        best_formula = f"y = {format_number(params[0])} * x^{format_number(params[1])}"

            except Exception as e:
                print(f"Errore nel modello {name}: {str(e)}")
                continue
        
        if best_model is None:
            messagebox.showerror("Error", "No regression model could be fitted to the data.")
            return None, None, None, None, []
            
        return best_model, best_r2, best_params, best_formula, results

    except Exception as e:
        messagebox.showerror("Error", f"Error in regression analysis: {str(e)}")
        return None, None, None, None, []

def run_regression():
    if not hasattr(app, 'file_path'):
        messagebox.showerror("Error", "Please load a file first.")
        return

    x_column = x_column_var.get()
    y_column = y_column_var.get()

    if not x_column or not y_column:
        messagebox.showerror("Error", "Please select both X and Y columns.")
        return

    try:
        # Carica il file Excel saltando la prima riga (intestazioni)
        df = pd.read_excel(app.file_path, usecols=[x_column, y_column])
        
        # Preprocess and clean data
        df_processed = preprocess_data(df, x_column, y_column)
        if df_processed is None or df_processed.empty:
            messagebox.showerror("Error", "No valid data after preprocessing.")
            return

        best_model, best_r2, best_params, best_formula, results = perform_regression(df_processed, x_column, y_column)
        
        if best_model is None:
            result_text = "No valid regression model could be fitted to the data.\n"
            result_text += "Please check your data format and try again."
        else:
            result_text = f"Best model selected: {best_model}\n\n"
            result_text += f"R²: {format_number(best_r2)}\n"
            result_text += f"Mathematical formula: {best_formula}\n"
            result_text += f"Excel formula: {get_excel_formula(best_model, best_params)}\n\n"
            
            result_text += "Metrics for all models:\n"
            for name, r2, params, mae, mse, rmse in results:
                result_text += f"\n{name}:\n"
                result_text += f"R² = {format_number(r2)}\n"
                result_text += f"MAE = {format_number(mae)}\n"
                result_text += f"MSE = {format_number(mse)}\n"
                result_text += f"RMSE = {format_number(rmse)}\n"
        
        result_text_widget.delete(1.0, tk.END)
        result_text_widget.insert(tk.END, result_text)
    
    except Exception as e:
        messagebox.showerror("Error", f"Error in regression analysis: {str(e)}")
        result_text_widget.delete(1.0, tk.END)
        result_text_widget.insert(tk.END, f"Error occurred: {str(e)}")

# GUI setup
load_button = tk.Button(app, text="Select Excel File", command=load_file)
load_button.pack(pady=10)

x_column_var = tk.StringVar()
y_column_var = tk.StringVar()

x_column_label = tk.Label(app, text="Select X column:")
x_column_label.pack()
x_column_menu = ttk.OptionMenu(app, x_column_var, "")
x_column_menu.pack()

y_column_label = tk.Label(app, text="Select Y column:")
y_column_label.pack()
y_column_menu = ttk.OptionMenu(app, y_column_var, "")
y_column_menu.pack()

run_button = tk.Button(app, text="Perform Regression", command=run_regression)
run_button.pack(pady=20)

result_text_widget = tk.Text(app, width=70, height=15, wrap=tk.WORD)
result_text_widget.pack(pady=10)
result_text_widget.config(state=tk.NORMAL)

app.mainloop()
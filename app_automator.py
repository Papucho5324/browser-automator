import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import threading
import time
import subprocess
import os

# Importaciones de Selenium
from selenium import webdriver
from selenium.webdriver.edge.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
# --- CAMBIO 1: IMPORTAR KEYS PARA PODER DAR ENTER ---
from selenium.webdriver.common.keys import Keys 

class AutomationApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Automatización Edge - Modo Debug")
        self.root.geometry("850x750")

        # --- VARIABLES ---
        self.edge_path_var = tk.StringVar(value=r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe")
        self.user_data_var = tk.StringVar(value=r"C:\EdgeAutomationProfile")
        
        # Variables de Excel y Lógica
        self.file_path = tk.StringVar()
        self.column_var = tk.StringVar()
        self.delay_var = tk.DoubleVar(value=1.0)
        self.selector_var = tk.StringVar(value="input.el-input__inner")
        self.start_row_var = tk.IntVar(value=1)
        
        self.df = None
        self.is_running = False

        self._build_ui()

    def _build_ui(self):
        # --- SECCIÓN 0: INICIAR NAVEGADOR ---
        frame_browser = ttk.LabelFrame(self.root, text="0. Preparación del Navegador", padding=10)
        frame_browser.pack(fill="x", padx=10, pady=5)

        lbl_instruccion = ttk.Label(frame_browser, text="1. Inicia el navegador. 2. Realiza el Login manual. 3. No cierres esa ventana.", foreground="red")
        lbl_instruccion.pack(anchor="w", pady=(0, 5))

        frame_paths = ttk.Frame(frame_browser)
        frame_paths.pack(fill="x", pady=2)
        ttk.Label(frame_paths, text="Ruta Edge:").pack(side="left")
        ttk.Entry(frame_paths, textvariable=self.edge_path_var, width=40).pack(side="left", padx=5)
        
        btn_launch = ttk.Button(frame_browser, text="🚀 ABRIR EDGE EN MODO SEGURO (DEBUG)", command=self.launch_browser_process)
        btn_launch.pack(fill="x", pady=5)

        # --- SECCIÓN 1: SELECCIÓN DE ARCHIVO ---
        frame_file = ttk.LabelFrame(self.root, text="1. Fuente de Datos (Excel)", padding=10)
        frame_file.pack(fill="x", padx=10, pady=5)

        ttk.Button(frame_file, text="📂 Seleccionar Excel", command=self.browse_file).pack(side="left", padx=5)
        ttk.Label(frame_file, textvariable=self.file_path, font=("Arial", 8, "italic")).pack(side="left", padx=5)

        # --- SECCIÓN 2: CONFIGURACIÓN ---
        frame_config = ttk.LabelFrame(self.root, text="2. Configuración de Inserción", padding=10)
        frame_config.pack(fill="x", padx=10, pady=5)

        ttk.Label(frame_config, text="Columna ID/Sno:").grid(row=0, column=0, sticky="w", padx=5)
        self.combo_columns = ttk.Combobox(frame_config, textvariable=self.column_var, state="readonly", width=15)
        self.combo_columns.grid(row=0, column=1, sticky="w", padx=5)
        self.combo_columns.bind("<<ComboboxSelected>>", self.preview_data)

        ttk.Label(frame_config, text="Selector CSS:").grid(row=0, column=2, sticky="w", padx=5)
        ttk.Entry(frame_config, textvariable=self.selector_var, width=30).grid(row=0, column=3, sticky="w", padx=5)

        ttk.Label(frame_config, text="Delay (seg):").grid(row=1, column=0, sticky="w", padx=5, pady=5)
        ttk.Entry(frame_config, textvariable=self.delay_var, width=10).grid(row=1, column=1, sticky="w", padx=5, pady=5)
        
        ttk.Label(frame_config, text="Fila Inicial:").grid(row=1, column=2, sticky="w", padx=5, pady=5)
        ttk.Entry(frame_config, textvariable=self.start_row_var, width=10).grid(row=1, column=3, sticky="w", padx=5, pady=5)

        # --- SECCIÓN 3: PREVISUALIZACIÓN ---
        frame_preview = ttk.LabelFrame(self.root, text="3. Datos a Procesar", padding=10)
        frame_preview.pack(fill="both", expand=True, padx=10, pady=5)

        self.tree = ttk.Treeview(frame_preview, columns=("fila", "valor"), show="headings", height=6)
        self.tree.heading("fila", text="Fila Excel")
        self.tree.heading("valor", text="Dato")
        self.tree.column("fila", width=80, anchor="center")
        self.tree.column("valor", width=300, anchor="w")
        
        scroll = ttk.Scrollbar(frame_preview, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scroll.set)
        self.tree.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")

        # --- SECCIÓN 4: EJECUCIÓN ---
        frame_action = ttk.Frame(self.root, padding=10)
        frame_action.pack(fill="x", padx=10, pady=5)

        self.btn_run = ttk.Button(frame_action, text="▶ EJECUTAR AUTOMATIZACIÓN", command=self.start_thread)
        self.btn_run.pack(fill="x", ipady=5)
        
        self.log_text = tk.Text(self.root, height=6, state="disabled", bg="#1e1e1e", fg="#00ff00", font=("Consolas", 9))
        self.log_text.pack(fill="x", padx=10, pady=5)

    def launch_browser_process(self):
        edge_exe = self.edge_path_var.get()
        user_data = self.user_data_var.get()
        if not os.path.exists(edge_exe):
            messagebox.showerror("Error", f"No se encuentra Edge en: {edge_exe}")
            return
        cmd = [edge_exe, "--remote-debugging-port=9222", f"--user-data-dir={user_data}"]
        try:
            subprocess.Popen(cmd)
            self.log(f"Navegador lanzado en puerto 9222.\nPerfil: {user_data}")
            self.log("Por favor, inicia sesión en la ventana que se abrió antes de continuar.")
        except Exception as e:
            messagebox.showerror("Error al lanzar Edge", str(e))

    def browse_file(self):
        filename = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xls")])
        if filename:
            self.file_path.set(filename)
            try:
                self.df = pd.read_excel(filename)
                self.combo_columns['values'] = self.df.columns.tolist()
                if "sno" in self.df.columns:
                    self.combo_columns.set("sno")
                elif len(self.df.columns) > 0:
                    self.combo_columns.current(0)
                self.preview_data()
                self.log(f"Excel cargado: {len(self.df)} registros.")
            except Exception as e:
                messagebox.showerror("Error", str(e))

    def preview_data(self, event=None):
        if self.df is None: return
        for i in self.tree.get_children(): self.tree.delete(i)
        col = self.column_var.get()
        if col in self.df.columns:
            start_idx = max(0, self.start_row_var.get() - 2)
            subset = self.df.iloc[start_idx:].head(50) 
            for idx, row in subset.iterrows():
                self.tree.insert("", "end", values=(idx + 2, row[col]))

    def log(self, msg):
        self.log_text.config(state="normal")
        self.log_text.insert("end", f"> {msg}\n")
        self.log_text.see("end")
        self.log_text.config(state="disabled")

    def start_thread(self):
        if not self.file_path.get():
            messagebox.showwarning("Atención", "Selecciona un archivo Excel.")
            return
        self.is_running = True
        self.btn_run.config(state="disabled")
        threading.Thread(target=self.run_automation, daemon=True).start()

    def run_automation(self):
        driver = None
        try:
            self.log("Conectando a la sesión del navegador existente (Puerto 9222)...")
            edge_options = Options()
            edge_options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
            driver = webdriver.Edge(options=edge_options)
            self.log("¡Conexión exitosa con el navegador!")

            col_name = self.column_var.get()
            selector = self.selector_var.get()
            delay = self.delay_var.get()
            start_row = self.start_row_var.get()
            
            start_index = max(0, start_row - 2)
            data_subset = self.df.iloc[start_index:]

            for index, row in data_subset.iterrows():
                if not self.is_running: break 
                
                excel_row = index + 2
                valor = str(row[col_name]) 
                
                try:
                    self.log(f"Fila {excel_row}: Insertando '{valor}'...")
                    
                    # 1. Buscar elemento
                    input_element = WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, selector))
                    )
                    
                    # --- LÓGICA MODIFICADA ---
                    
                    # 2. Limpiar (por seguridad antes de escribir)
                    input_element.clear()
                    
                    # 3. Escribir valor
                    input_element.send_keys(valor)
                    
                    # 4. Dar ENTER
                    input_element.send_keys(Keys.ENTER)
                    
                    # 5. Esperar un momento (vital para que el ENTER procese)
                    time.sleep(0.5) 
                    
                    # 6. Limpiar el campo DESPUÉS del enter (si la página lo permite)
                    # Usamos try/except porque si el ENTER recarga la página, 
                    # el 'input_element' viejo dará error (StaleElementReference).
                    try:
                        input_element.clear()
                    except:
                        # Si falla, probablemente la página cambió o recargó, 
                        # así que no pasa nada, se buscará de nuevo en la sgte vuelta.
                        pass 

                    self.log(f"--> OK.")
                    
                except Exception as e_row:
                    self.log(f"ERROR en fila {excel_row}: {str(e_row)}")
                
                time.sleep(delay)

            self.log("--- PROCESO COMPLETADO ---")
            messagebox.showinfo("Fin", "La automatización ha terminado.")

        except Exception as e:
            self.log(f"ERROR CRÍTICO: {str(e)}")
            if "Connection refused" in str(e):
                messagebox.showerror("Error de Conexión", "No se pudo conectar a Edge. ¿Ejecutaste el Paso 0?")
            else:
                messagebox.showerror("Error", str(e))
        
        finally:
            self.is_running = False
            self.btn_run.config(state="normal")

if __name__ == "__main__":
    root = tk.Tk()
    app = AutomationApp(root)
    root.mainloop()

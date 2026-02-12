import os
import re
import tkinter as tk
from docx import Document
from tkinter import filedialog, ttk, scrolledtext
from collections import defaultdict
from tables_extractor import extract_data_from_tables
from export_data import *

class WordTableExtractor:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Extractor de Horarios Académicos - V2.0 (Multifile)")
        self.root.geometry("1400x800")
        
        self.loaded_files = []
        self.files_tables = {}
        self.current_file_path = None
        self.selector_frame = None
        
        self.setup_gui()
        
    def setup_gui(self):
        # Frame principal
        main_frame = tk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Frame superior para controles
        control_frame = tk.Frame(main_frame)
        control_frame.pack(fill=tk.X, pady=(0, 10))
        
        # ----- Fila 1: selector de archivos -----
        file_frame = tk.Frame(control_frame)
        file_frame.pack(fill=tk.X, pady=5)
        
        tk.Label(file_frame, text="Archivos cargados:").pack(side=tk.LEFT, padx=(0,5))
        
        self.file_selector_var = tk.StringVar()
        self.file_selector = ttk.Combobox(file_frame, textvariable=self.file_selector_var,
                                          state="readonly", width=60)
        self.file_selector.pack(side=tk.LEFT, padx=5)
        self.file_selector.bind("<<ComboboxSelected>>", self.on_file_select)
        
        btn_load = tk.Button(file_frame, text="📂 Cargar archivos", 
                           command=self.load_files, width=20)
        btn_load.pack(side=tk.LEFT, padx=5)
        
        # ----- Fila 2: botones de acción -----
        action_frame = tk.Frame(control_frame)
        action_frame.pack(fill=tk.X, pady=5)
        
        btn_extract = tk.Button(action_frame, text="🔍 Extraer datos", 
                               command=self.extract_all_files, width=20)
        btn_extract.pack(side=tk.LEFT, padx=2)
        
        btn_export = tk.Button(action_frame, text="💾 Exportar a Markdown", 
                              command=export_to_md, width=20)
        btn_export.pack(side=tk.LEFT, padx=2)
        
        # btn_export_txt = tk.Button(action_frame, text="📄 Exportar TXT", 
        #                           command=export_txt, width=15)
        # btn_export_txt.pack(side=tk.LEFT, padx=2)
        
        btn_clear = tk.Button(action_frame, text="🗑️ Limpiar todo", 
                             command=self.clear_data, width=20)
        btn_clear.pack(side=tk.LEFT, padx=2)
        
        # Etiqueta de estado
        self.status_label = tk.Label(control_frame, text="🟢 Listo", fg="green", 
                                    font=("Arial", 10))
        self.status_label.pack(side=tk.RIGHT, padx=10)
        
        # ----- Notebook con pestañas -----
        self.notebook = ttk.Notebook(main_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True)
        
        # Pestaña 1: Vista de datos extraídos
        self.table_frame = tk.Frame(self.notebook)
        self.notebook.add(self.table_frame, text="📊 Datos Extraídos")
        
        # Treeview con scrollbars
        tree_frame = tk.Frame(self.table_frame)
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        self.tree = ttk.Treeview(tree_frame, show="headings", height=20)
        
        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)
        
        # Pestaña 2: Resumen
        self.summary_frame = tk.Frame(self.notebook)
        self.notebook.add(self.summary_frame, text="📋 Resumen")
        
        self.summary_text = scrolledtext.ScrolledText(self.summary_frame, 
                                                     wrap=tk.WORD, 
                                                     font=("Courier", 10))
        self.summary_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Pestaña 3: Vista del archivo
        self.file_view_frame = tk.Frame(self.notebook)
        self.notebook.add(self.file_view_frame, text="📄 Vista del Archivo")
        
        self.file_text = scrolledtext.ScrolledText(self.file_view_frame, 
                                                  wrap=tk.WORD, 
                                                  font=("Courier", 9))
        self.file_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Pestaña 4: Estadísticas
        self.stats_frame = tk.Frame(self.notebook)
        self.notebook.add(self.stats_frame, text="📈 Estadísticas")
        
        self.stats_text = scrolledtext.ScrolledText(self.stats_frame, 
                                                   wrap=tk.WORD, 
                                                   font=("Courier", 10))
        self.stats_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
    
    # ------------------------------------------------------------
    #  Gestión de archivos
    # ------------------------------------------------------------
    def load_files(self):
        """Carga uno o varios archivos .docx y los añade a la lista"""
        file_paths = filedialog.askopenfilenames(
            filetypes=[("Word documents", "*.docx"), ("Todos los archivos", "*.*")],
            title="Seleccionar archivos"
        )
        
        if not file_paths:
            return
        
        nuevos = 0
        for fp in file_paths:
            if fp not in self.loaded_files:
                self.loaded_files.append(fp)
                nuevos += 1
        
        if nuevos == 0:
            self.status_label.config(text="🟡 Archivos ya cargados", fg="orange")
            return
        
        # Actualizar combobox con los nombres de archivo
        nombres = [os.path.basename(fp) for fp in self.loaded_files]
        self.file_selector['values'] = nombres
        
        # Seleccionar el primer archivo de la nueva carga
        if self.current_file_path is None:
            idx = self.loaded_files.index(file_paths[0])
            self.file_selector.current(idx)
            self.current_file_path = self.loaded_files[idx]
            self.display_file_content(self.current_file_path)
        
        self.status_label.config(text=f"🟢 {nuevos} archivo(s) cargado(s)", fg="green")
    
    def on_file_select(self, event=None):
        """Cambia el archivo activo y actualiza las vistas"""
        if not self.loaded_files or not self.file_selector_var.get():
            return
        
        # Obtener la ruta completa a partir del nombre visible
        selected_name = self.file_selector_var.get()
        for fp in self.loaded_files:
            if os.path.basename(fp) == selected_name:
                self.current_file_path = fp
                break
        
        # Actualizar vista del archivo
        self.display_file_content(self.current_file_path)
        
        # Si ya hay tablas extraídas para este archivo, mostrarlas
        if self.current_file_path in self.files_tables:
            # Actualizar self.tables_data (lo esperan los métodos existentes)
            self.tables_data = self.files_tables[self.current_file_path]
            self.display_data()
            self.update_summary()
            self.update_statistics()
        else:
            # Limpiar áreas de datos
            for item in self.tree.get_children():
                self.tree.delete(item)
            self.summary_text.delete(1.0, tk.END)
            self.stats_text.delete(1.0, tk.END)
            if self.selector_frame:
                self.selector_frame.destroy()
                self.selector_frame = None
    
    def display_file_content(self, file_path):
        """Muestra el contenido de un archivo Word en la pestaña correspondiente"""
        try:
            doc = Document(file_path)
            content_lines = []
            for para in doc.paragraphs:
                if para.text.strip():
                    content_lines.append(para.text)
            for table in doc.tables:
                content_lines.append("\n" + "="*50 + " TABLA " + "="*50)
                for row in table.rows:
                    row_text = ' | '.join([cell.text.strip() for cell in row.cells])
                    content_lines.append(row_text)
            content = '\n'.join(content_lines)
            
            self.file_text.delete(1.0, tk.END)
            self.file_text.insert(1.0, content)
        except Exception as e:
            self.file_text.delete(1.0, tk.END)
            self.file_text.insert(1.0, f"Error al leer el archivo: {str(e)}")
    
    # ------------------------------------------------------------
    #  Extracción y visualización de tablas
    # ------------------------------------------------------------
    def extract_all_files(self):
        """Extrae tablas de todos los archivos cargados"""
        if not self.loaded_files:
            self.status_label.config(text="🔴 No hay archivos cargados", fg="red")
            return
        
        self.status_label.config(text="🟡 Extrayendo tablas...", fg="orange")
        self.root.update_idletasks()
        
        total_tablas = 0
        for fp in self.loaded_files:
            try:
                tablas = extract_data_from_tables(fp)
                if tablas:
                    self.files_tables[fp] = tablas
                    total_tablas += len(tablas)
                else:
                    self.files_tables[fp] = []
            except Exception as e:
                print(f"Error extrayendo {fp}: {e}")
                self.files_tables[fp] = []
        
        # Seleccionar el primer archivo si no hay ninguno activo
        if self.current_file_path is None and self.loaded_files:
            self.current_file_path = self.loaded_files[0]
            self.file_selector.current(0)
        
        # Mostrar las tablas del archivo activo
        if self.current_file_path in self.files_tables:
            self.tables_data = self.files_tables[self.current_file_path]
            self.display_data()
            self.update_summary()
            self.update_statistics()
        
        self.status_label.config(text=f"🟢 Extracción completa ({total_tablas} tablas)", fg="green")
    
    # ------------------------------------------------------------
    #  Visualización de datos (heredado con pequeñas adaptaciones)
    # ------------------------------------------------------------
    def limpiar_texto(self, texto):
        if not texto:
            return ""
        texto = re.sub(r'\s+', ' ', texto.strip())
        texto = re.sub(r'[\x00-\x1F\x7F]', '', texto)
        return texto

    def display_data(self):
        """Muestra los datos del archivo activo en el Treeview"""
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.tree["columns"] = ()
        
        if not hasattr(self, 'tables_data') or not self.tables_data:
            return
        
        tabla = self.tabletables_datas_data[0]
        if not tabla or len(tabla) < 2:
            return
        
        # Encabezados desde la primera fila
        headers = tabla[0]
        datos = tabla[1:] if len(tabla) > 1 else []
        
        clean_headers = []
        for i, header in enumerate(headers):
            header_str = str(header).strip()
            if not header_str:
                clean_headers.append("Turno/Hora" if i == 0 else f"Día {i}")
            else:
                clean_headers.append(header_str)
        
        self.tree["columns"] = clean_headers
        
        for i, header in enumerate(clean_headers):
            col_id = f"#{i+1}"
            self.tree.heading(col_id, text=header, anchor="w")
            if i == 0:
                self.tree.column(col_id, width=180, minwidth=120, stretch=False, anchor="w")
            else:
                self.tree.column(col_id, width=300, minwidth=200, stretch=True, anchor="w")
        
        # Insertar filas de datos
        for fila in datos:
            if len(fila) < len(clean_headers):
                fila_completa = list(fila) + [""] * (len(clean_headers) - len(fila))
            elif len(fila) > len(clean_headers):
                fila_completa = fila[:len(clean_headers)]
            else:
                fila_completa = list(fila)
            
            fila_str = [str(val).strip() if val is not None else "" for val in fila_completa]
            self.tree.insert("", "end", values=fila_str)
        
        # Selector de tablas si hay más de una
        if self.selector_frame:
            self.selector_frame.destroy()
            self.selector_frame = None
        
        if len(self.tables_data) > 1:
            self.create_table_selector()
    
    def create_table_selector(self):
        """Crea un Combobox para elegir tabla dentro del archivo activo"""
        self.selector_frame = tk.Frame(self.table_frame)
        self.selector_frame.pack(fill=tk.X, padx=5, pady=5)
        
        tk.Label(self.selector_frame, text="Seleccionar tabla:").pack(side=tk.LEFT, padx=5)
        
        self.table_var = tk.StringVar()
        table_options = [f"Tabla {i+1}" for i in range(len(self.tables_data))]
        table_dropdown = ttk.Combobox(self.selector_frame, textvariable=self.table_var,
                                     values=table_options, state="readonly", width=15)
        table_dropdown.pack(side=tk.LEFT, padx=5)
        table_dropdown.current(0)
        table_dropdown.bind("<<ComboboxSelected>>", self.on_table_select)
    
    def on_table_select(self, event):
        """Cambia la tabla mostrada dentro del archivo activo"""
        if not hasattr(self, 'table_var'):
            return
        idx = int(self.table_var.get().split()[-1]) - 1
        if 0 <= idx < len(self.tables_data):
            tabla = self.tables_data[idx]
            # Reconstruir treeview con esta tabla (mismo código que en display_data)
            self._display_specific_table(tabla)
    
    def _display_specific_table(self, tabla):
        """Muestra una tabla específica en el Treeview"""
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.tree["columns"] = ()
        
        if not tabla or len(tabla) < 2:
            return
        
        headers = tabla[0]
        datos = tabla[1:]
        
        clean_headers = []
        for i, header in enumerate(headers):
            header_str = str(header).strip()
            if not header_str:
                clean_headers.append("Turno/Hora" if i == 0 else f"Día {i}")
            else:
                clean_headers.append(header_str)
        
        self.tree["columns"] = clean_headers
        for i, header in enumerate(clean_headers):
            col_id = f"#{i+1}"
            self.tree.heading(col_id, text=header, anchor="w")
            if i == 0:
                self.tree.column(col_id, width=180, minwidth=120, stretch=False, anchor="w")
            else:
                self.tree.column(col_id, width=300, minwidth=200, stretch=True, anchor="w")
        
        for fila in datos:
            if len(fila) < len(clean_headers):
                fila_completa = list(fila) + [""] * (len(clean_headers) - len(fila))
            elif len(fila) > len(clean_headers):
                fila_completa = fila[:len(clean_headers)]
            else:
                fila_completa = list(fila)
            fila_str = [str(val).strip() if val is not None else "" for val in fila_completa]
            self.tree.insert("", "end", values=fila_str)
    
    # ------------------------------------------------------------
    #  Resumen y estadísticas (sobre el archivo activo)
    # ------------------------------------------------------------
    def update_summary(self):
        self.summary_text.delete(1.0, tk.END)
        if not hasattr(self, 'tables_data') or not self.tables_data:
            self.summary_text.insert(tk.END, "No hay datos extraídos para el archivo activo.")
            return
        
        nombre = os.path.basename(self.current_file_path) if self.current_file_path else "Sin archivo"
        self.summary_text.insert(tk.END, f"ARCHIVO: {nombre}\n")
        self.summary_text.insert(tk.END, "="*50 + "\n\n")
        
        for i, tabla in enumerate(self.tables_data):
            self.summary_text.insert(tk.END, f"TABLA {i+1}:\n")
            self.summary_text.insert(tk.END, f"- Filas: {len(tabla)}\n")
            if tabla:
                self.summary_text.insert(tk.END, f"- Columnas: {len(tabla[0])}\n")
                self.summary_text.insert(tk.END, f"- Encabezados: {', '.join(tabla[0])}\n")
            self.summary_text.insert(tk.END, "\n")
    
    def update_statistics(self):
        self.stats_text.delete(1.0, tk.END)
        if not hasattr(self, 'tables_data') or not self.tables_data:
            self.stats_text.insert(tk.END, "No hay datos para analizar.")
            return
        
        nombre = os.path.basename(self.current_file_path) if self.current_file_path else "Sin archivo"
        self.stats_text.insert(tk.END, f"ESTADÍSTICAS - {nombre}\n")
        self.stats_text.insert(tk.END, "="*50 + "\n\n")
        
        total_filas = 0
        total_columnas = 0
        for i, tabla in enumerate(self.tables_data):
            filas = len(tabla) - 1 if len(tabla) > 1 else 0
            columnas = len(tabla[0]) if tabla else 0
            total_filas += filas
            total_columnas += columnas
            
            self.stats_text.insert(tk.END, f"Tabla {i+1}:\n")
            self.stats_text.insert(tk.END, f"  • Filas de datos: {filas}\n")
            self.stats_text.insert(tk.END, f"  • Columnas: {columnas}\n")
            self.stats_text.insert(tk.END, f"  • Celdas de datos: {filas * columnas}\n")
            
            horarios = defaultdict(list)
            for fila_idx, fila in enumerate(tabla[1:], start=1):
                if fila and len(fila) > 0:
                    horario = fila[0].strip()
                    if horario:
                        horarios[horario].append(fila_idx)
            duplicados = {h: f for h, f in horarios.items() if len(f) > 1}
            if duplicados:
                self.stats_text.insert(tk.END, f"  • Horarios duplicados: {len(duplicados)}\n")
                for horario, filas in list(duplicados.items())[:3]:
                    self.stats_text.insert(tk.END, f"    - {horario}: {len(filas)} filas\n")
                if len(duplicados) > 3:
                    self.stats_text.insert(tk.END, f"    ... y {len(duplicados)-3} más\n")
            self.stats_text.insert(tk.END, "\n")
        
        self.stats_text.insert(tk.END, "TOTALES DEL ARCHIVO:\n")
        self.stats_text.insert(tk.END, f"• Tablas: {len(self.tables_data)}\n")
        self.stats_text.insert(tk.END, f"• Filas totales: {total_filas}\n")
        if self.tables_data:
            self.stats_text.insert(tk.END, f"• Columnas promedio: {total_columnas/len(self.tables_data):.1f}\n")
    
    # ------------------------------------------------------------
    #  Limpieza general
    # ------------------------------------------------------------
    def clear_data(self):
        """Elimina todos los archivos cargados y las tablas extraídas"""
        self.loaded_files = []
        self.files_tables = {}
        self.current_file_path = None
        self.file_selector['values'] = []
        self.file_selector_var.set("")
        
        # Limpiar áreas visuales
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.summary_text.delete(1.0, tk.END)
        self.stats_text.delete(1.0, tk.END)
        self.file_text.delete(1.0, tk.END)
        
        if self.selector_frame:
            self.selector_frame.destroy()
            self.selector_frame = None
        
        self.status_label.config(text="🟢 Todo limpio", fg="green")
    
    def run(self):
        self.root.mainloop()

def main():
    app = WordTableExtractor()
    app.run()

if __name__ == "__main__":
    main()
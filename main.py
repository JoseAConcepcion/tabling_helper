import tkinter as tk
from tkinter import filedialog, ttk, messagebox, scrolledtext
from extract_data import *
from collections import defaultdict
import os

class WordTableExtractor:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Extractor de Horarios Académicos - V2.0")
        self.root.geometry("1400x800")
        
        self.data = {}
        self.tables_data = []
        self.current_file_path = ""
        
        self.setup_gui()
        
    def setup_gui(self):
        # Frame principal
        main_frame = tk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Frame superior para controles
        control_frame = tk.Frame(main_frame)
        control_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Botón para cargar archivo
        btn_frame = tk.Frame(control_frame)
        btn_frame.pack(fill=tk.X, pady=5)
        
        tk.Label(btn_frame, text="Archivo:").pack(side=tk.LEFT, padx=(0, 5))
        
        self.file_label = tk.Label(btn_frame, text="Ningún archivo cargado", 
                                  fg="gray", width=50, anchor="w")
        self.file_label.pack(side=tk.LEFT, padx=5)
        
        btn_load = tk.Button(btn_frame, text="📂 Seleccionar archivo", 
                           command=self.load_file, width=20)
        btn_load.pack(side=tk.LEFT, padx=5)
        
        # Frame para botones de acción
        action_frame = tk.Frame(control_frame)
        action_frame.pack(fill=tk.X, pady=5)
        
        btn_extract = tk.Button(action_frame, text="🔍 Extraer datos", 
                               command=self.extract_data, width=15)
        btn_extract.pack(side=tk.LEFT, padx=2)
        
        btn_export = tk.Button(action_frame, text="💾 Exportar Excel", 
                              command=(), width=15)
        btn_export.pack(side=tk.LEFT, padx=2)
        
        btn_export_txt = tk.Button(action_frame, text="📄 Exportar TXT", 
                                  command=(), width=15)
        btn_export_txt.pack(side=tk.LEFT, padx=2)
        
        btn_clear = tk.Button(action_frame, text="🗑️ Limpiar", 
                             command=self.clear_data, width=15)
        btn_clear.pack(side=tk.LEFT, padx=2)
        
        # Etiqueta de estado
        self.status_label = tk.Label(control_frame, text="🟢 Listo", fg="green", 
                                    font=("Arial", 10))
        self.status_label.pack(side=tk.RIGHT, padx=10)
        
        # Notebook para pestañas
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
        
    def load_file(self):
        """Carga un archivo para procesar"""
        file_path = filedialog.askopenfilename(
            filetypes=[
                ("Todos los archivos", "*.*"),
                ("Word documents", "*.docx"),
                ("Text files", "*.txt"),
                ("Old Word files", "*.doc")
            ],
            title="Seleccionar archivo"
        )
        
        if file_path:
            self.current_file_path = file_path
            self.file_label.config(text=os.path.basename(file_path))
            self.display_file_content(file_path)
            
            if self.auto_extract_var.get():
                self.extract_data()
    
    def display_file_content(self, file_path):
        """Muestra el contenido del archivo en la pestaña correspondiente"""
        try:
            content = self.read_file_content(file_path)
            self.file_text.delete(1.0, tk.END)
            self.file_text.insert(1.0, content[:10000])  # Mostrar primeros 10000 caracteres
            
            # Resaltar líneas con tablas
            self.highlight_table_lines()
            
        except Exception as e:
            self.file_text.delete(1.0, tk.END)
            self.file_text.insert(1.0, f"Error al leer el archivo: {str(e)}")
    
    def read_file_content(self, file_path):
        """Lee el contenido del archivo usando diferentes métodos"""
        method = self.read_method.get()
        
        if method == "text" or file_path.endswith('.txt'):
            # Leer como texto plano
            with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                return f.read()
        
        elif method == "docx" or file_path.endswith('.docx'):
            # Intentar leer como .docx
            try:
                return self.read_docx_file(file_path)
            except:
                # Si falla, intentar como texto
                with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                    return f.read()
        
        else:  # Auto-detección
            try:
                # Primero intentar como .docx
                return self.read_docx_file(file_path)
            except:
                # Si falla, leer como texto
                with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                    return f.read()
    
    def read_docx_file(self, file_path):
        """Lee un archivo .docx"""
        if not file_path.endswith('.docx'):
            # Intentar leer como si fuera .docx aunque no tenga la extensión
            pass
        
        try:
            doc = Document(file_path)
            content = []
            for para in doc.paragraphs:
                content.append(para.text)
            
            for table in doc.tables:
                for row in table.rows:
                    row_text = ' | '.join([cell.text for cell in row.cells])
                    content.append(row_text)
            
            return '\n'.join(content)
        
        except Exception as e:
            # Si falla, verificar si es un archivo zip/XML corrupto
            try:
                with open(file_path, 'rb') as f:
                    # Intentar leer como texto binario
                    content = f.read().decode('utf-8', errors='ignore')
                    return content
            except:
                raise Exception(f"No se pudo leer como .docx: {str(e)}")
    
    def highlight_table_lines(self):
        """Resalta líneas que parecen contener tablas"""
        content = self.file_text.get(1.0, tk.END)
        lines = content.split('\n')
        
        self.file_text.tag_configure("table_line", background="lightyellow")
        self.file_text.tag_configure("header", background="lightblue")
        
        start_index = "1.0"
        for line in lines:
            # Buscar patrones de tabla
            if '|' in line and ('+' in line or any(x in line for x in ['LUNES', 'MARTES', 'MIERCOLES'])):
                line_end = f"{start_index}+{len(line)}c"
                self.file_text.tag_add("table_line", start_index, line_end)
            
            # Resaltar encabezados
            if any(x in line.upper() for x in ['GRUPO:', 'CARRERA:', 'AÑO:', 'HORARIO']):
                line_end = f"{start_index}+{len(line)}c"
                self.file_text.tag_add("header", start_index, line_end)
            
            start_index = f"{start_index}+{len(line)+1}c"
    
    def extract_data(self):
        pass
    
    def display_data(self):
        pass
    def update_summary(self):
        pass

# Función principal
def main():
    app = WordTableExtractor()
    app.run()


if __name__ == "__main__":
    main()
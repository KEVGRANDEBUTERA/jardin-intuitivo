import pandas as pd
import math
import os
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from openpyxl import load_workbook
from openpyxl.chart import BarChart, LineChart, Reference, PieChart
import threading

class SimuladorGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("🎯 SISTEMA DE SIMULACIÓN MULTI-MODELO")
        self.root.geometry("900x700")
        self.root.configure(bg="#f0f0f0")
        
        # Datos base
        self.RI_BASE = [0.213, 0.345, 0.021, 0.987, 0.543, 0.4885, 0.3122, 0.3577, 0.2361, 0.6323, 
                        0.6394, 0.6084, 0.6829, 0.0651, 0.0246, 0.0376, 0.7462, 0.5346, 0.3833, 0.3408, 
                        0.1365, 0.1250, 0.0876, 0.1321, 0.7300, 0.3545, 0.5596, 0.0898, 0.5765, 0.2129]
        
        self.simulaciones_realizadas = []
        
        # Estilo
        self.style = ttk.Style()
        self.style.theme_use('clam')
        
        self.crear_interfaz()
    
    def crear_interfaz(self):
        """Crea la interfaz gráfica principal"""
        
        # Header
        header_frame = tk.Frame(self.root, bg="#2c3e50", height=100)
        header_frame.pack(fill=tk.X)
        
        titulo = tk.Label(header_frame, text="★ SISTEMA DE SIMULACIÓN MULTI-MODELO ★", 
                         font=("Arial", 18, "bold"), bg="#2c3e50", fg="white")
        titulo.pack(pady=15)
        
        # Frame principal con notebook (tabs)
        notebook = ttk.Notebook(self.root)
        notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Tab 1: Simulaciones
        tab_simulaciones = ttk.Frame(notebook)
        notebook.add(tab_simulaciones, text="📊 Simulaciones")
        self.crear_tab_simulaciones(tab_simulaciones)
        
        # Tab 2: Resultados
        tab_resultados = ttk.Frame(notebook)
        notebook.add(tab_resultados, text="📈 Resultados")
        self.crear_tab_resultados(tab_resultados)
        
        # Tab 3: Archivo
        tab_archivo = ttk.Frame(notebook)
        notebook.add(tab_archivo, text="📁 Archivo")
        self.crear_tab_archivo(tab_archivo)
    
    def crear_tab_simulaciones(self, parent):
        """Tab con botones para cada simulación"""
        
        label_titulo = tk.Label(parent, text="Selecciona una simulación para ejecutar:", 
                               font=("Arial", 14, "bold"), bg="#f0f0f0")
        label_titulo.pack(pady=20)
        
        # Frame para los botones
        botones_frame = tk.Frame(parent, bg="#f0f0f0")
        botones_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        botones = [
            ("1️⃣ Transformada Inversa\n(Demanda de Cepillos)", self.simular_transformada_inversa, "#3498db"),
            ("2️⃣ Distribución Uniforme\n(Temperatura Estufa)", self.simular_uniforme, "#e74c3c"),
            ("3️⃣ Distribución Poisson\n(Piezas por Hora)", self.simular_poisson, "#2ecc71"),
            ("4️⃣ Distribución Bernoulli\n(Probabilidad de Fallas)", self.simular_bernoulli, "#f39c12"),
            ("5️⃣ Distribución Exponencial\n(Tiempo en Banco)", self.simular_exponencial, "#9b59b6"),
        ]
        
        for i, (texto, comando, color) in enumerate(botones):
            fila = i // 2
            columna = i % 2
            
            btn = tk.Button(botones_frame, text=texto, command=comando, 
                           font=("Arial", 11, "bold"), bg=color, fg="white",
                           relief=tk.RAISED, bd=3, cursor="hand2", padx=20, pady=20)
            btn.grid(row=fila, column=columna, padx=10, pady=10, sticky="nsew")
        
        botones_frame.grid_rowconfigure(0, weight=1)
        botones_frame.grid_rowconfigure(1, weight=1)
        botones_frame.grid_rowconfigure(2, weight=1)
        botones_frame.grid_columnconfigure(0, weight=1)
        botones_frame.grid_columnconfigure(1, weight=1)
        
        # Barra de estado
        status_frame = tk.Frame(parent, bg="#ecf0f1", relief=tk.SUNKEN, bd=1)
        status_frame.pack(fill=tk.X, side=tk.BOTTOM, padx=10, pady=5)
        
        self.status_label = tk.Label(status_frame, text="Listo para ejecutar simulaciones", 
                                     font=("Arial", 10), bg="#ecf0f1", fg="#2c3e50")
        self.status_label.pack(pady=5)
    
    def crear_tab_resultados(self, parent):
        """Tab para ver resultados y estadísticas"""
        
        label_titulo = tk.Label(parent, text="Simulaciones Realizadas:", 
                               font=("Arial", 14, "bold"), bg="#f0f0f0")
        label_titulo.pack(pady=15)
        
        # Frame con scrollbar
        frame_scroll = tk.Frame(parent)
        frame_scroll.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
        
        scrollbar = ttk.Scrollbar(frame_scroll)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.listbox_resultados = tk.Listbox(frame_scroll, yscrollcommand=scrollbar.set,
                                            font=("Arial", 11), bg="white", relief=tk.SUNKEN, bd=1)
        self.listbox_resultados.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.listbox_resultados.yview)
        
        # Botones
        botones_frame = tk.Frame(parent, bg="#f0f0f0")
        botones_frame.pack(fill=tk.X, padx=20, pady=10)
        
        btn_actualizar = tk.Button(botones_frame, text="🔄 Actualizar", 
                                  command=self.actualizar_resultados, 
                                  font=("Arial", 10, "bold"), bg="#3498db", fg="white",
                                  relief=tk.RAISED, bd=2, cursor="hand2", padx=15, pady=8)
        btn_actualizar.pack(side=tk.LEFT, padx=5)
        
        btn_limpiar = tk.Button(botones_frame, text="🗑️ Limpiar Datos", 
                               command=self.limpiar_datos, 
                               font=("Arial", 10, "bold"), bg="#e74c3c", fg="white",
                               relief=tk.RAISED, bd=2, cursor="hand2", padx=15, pady=8)
        btn_limpiar.pack(side=tk.LEFT, padx=5)
        
        self.actualizar_resultados()
    
    def crear_tab_archivo(self, parent):
        """Tab para gestionar archivos"""
        
        frame_contenido = tk.Frame(parent, bg="#f0f0f0")
        frame_contenido.pack(fill=tk.BOTH, expand=True, padx=30, pady=30)
        
        # Información
        info_text = """
📊 GESTIÓN DE ARCHIVOS

✓ Todos tus datos se guardan automáticamente en:
  → Resultados_Simulacion.xlsx

✓ El archivo contiene:
  → 5 hojas (una por cada simulación)
  → Gráficas embebidas en cada hoja
  → Datos organizados y legibles
  → Estadísticas completas

✓ Funciones disponibles:
  → Generar archivo: Crea el Excel con todas las simulaciones
  → Abrir carpeta: Abre la carpeta del proyecto
  → Abrir Excel: Abre directamente el archivo Excel
        """
        
        label_info = tk.Label(frame_contenido, text=info_text, 
                             font=("Arial", 11), bg="#f0f0f0", justify=tk.LEFT,
                             relief=tk.SUNKEN, bd=2, padx=20, pady=20)
        label_info.pack(fill=tk.BOTH, expand=True)
        
        # Botones
        botones_frame = tk.Frame(parent, bg="#f0f0f0")
        botones_frame.pack(fill=tk.X, padx=20, pady=10)
        
        btn_generar = tk.Button(botones_frame, text="📊 Generar Archivo Excel", 
                               command=self.generar_archivo_final, 
                               font=("Arial", 11, "bold"), bg="#2ecc71", fg="white",
                               relief=tk.RAISED, bd=2, cursor="hand2", padx=20, pady=10)
        btn_generar.pack(side=tk.LEFT, padx=5)
        
        btn_abrir_carpeta = tk.Button(botones_frame, text="📁 Abrir Carpeta", 
                                     command=self.abrir_carpeta, 
                                     font=("Arial", 11, "bold"), bg="#3498db", fg="white",
                                     relief=tk.RAISED, bd=2, cursor="hand2", padx=20, pady=10)
        btn_abrir_carpeta.pack(side=tk.LEFT, padx=5)
        
        btn_abrir_excel = tk.Button(botones_frame, text="📈 Abrir Excel", 
                                   command=self.abrir_excel, 
                                   font=("Arial", 11, "bold"), bg="#9b59b6", fg="white",
                                   relief=tk.RAISED, bd=2, cursor="hand2", padx=20, pady=10)
        btn_abrir_excel.pack(side=tk.LEFT, padx=5)
    
    def guardar_en_excel(self, df, nombre_hoja, tipo_grafica='bar'):
        """Guarda datos en Excel con gráfica"""
        archivo = "Resultados_Simulacion.xlsx"
        
        if not os.path.isfile(archivo):
            df.to_excel(archivo, sheet_name=nombre_hoja, index=False)
        else:
            with pd.ExcelWriter(archivo, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                df.to_excel(writer, sheet_name=nombre_hoja, index=False)
        
        self.agregar_grafica(archivo, nombre_hoja, df, tipo_grafica)
    
    def agregar_grafica(self, archivo, nombre_hoja, df, tipo_grafica):
        """Agrega gráfica al Excel"""
        wb = load_workbook(archivo)
        ws = wb[nombre_hoja]
        
        columnas_numericas = df.select_dtypes(include=['number']).columns.tolist()
        
        if len(columnas_numericas) < 1:
            return
        
        max_row = len(df) + 1
        col_datos = 2 if len(df.columns) > 1 else 1
        
        if tipo_grafica == 'bar':
            chart = BarChart()
            chart.title = f"Gráfica: {nombre_hoja}"
            chart.x_axis.title = df.columns[0]
            chart.y_axis.title = columnas_numericas[0]
            
            data = Reference(ws, min_col=col_datos, min_row=1, max_row=max_row)
            cats = Reference(ws, min_col=1, min_row=2, max_row=max_row)
            chart.add_data(data, titles_from_data=True)
            chart.set_categories(cats)
            
        elif tipo_grafica == 'line':
            chart = LineChart()
            chart.title = f"Gráfica: {nombre_hoja}"
            chart.x_axis.title = df.columns[0]
            chart.y_axis.title = columnas_numericas[0]
            
            data = Reference(ws, min_col=col_datos, min_row=1, max_row=max_row)
            cats = Reference(ws, min_col=1, min_row=2, max_row=max_row)
            chart.add_data(data, titles_from_data=True)
            chart.set_categories(cats)
            
        elif tipo_grafica == 'pie':
            chart = PieChart()
            chart.title = f"Gráfica: {nombre_hoja}"
            
            labels = Reference(ws, min_col=1, min_row=2, max_row=max_row)
            data = Reference(ws, min_col=col_datos, min_row=1, max_row=max_row)
            chart.add_data(data, titles_from_data=True)
            chart.set_categories(labels)
        
        ws.add_chart(chart, "D2")
        wb.save(archivo)
    
    def mostrar_estadisticas(self, df, nombre_modelo):
        """Muestra estadísticas en un popup"""
        stats_text = f"{'='*40}\nESTADÍSTICAS: {nombre_modelo}\n{'='*40}\n"
        stats_text += f"Total de registros: {len(df)}\n\n"
        
        columnas_numericas = df.select_dtypes(include=['number']).columns.tolist()
        for col in columnas_numericas:
            if col != 'ri':
                stats_text += f"{col}:\n"
                stats_text += f"  Promedio: {df[col].mean():.4f}\n"
                stats_text += f"  Máximo: {df[col].max():.4f}\n"
                stats_text += f"  Mínimo: {df[col].min():.4f}\n"
                stats_text += f"  Desv. Est.: {df[col].std():.4f}\n\n"
        
        # Crear ventana de estadísticas
        ventana_stats = tk.Toplevel(self.root)
        ventana_stats.title(f"Estadísticas - {nombre_modelo}")
        ventana_stats.geometry("500x400")
        
        text_widget = tk.Text(ventana_stats, font=("Courier", 10), bg="white", relief=tk.SUNKEN, bd=1)
        text_widget.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        text_widget.insert(tk.END, stats_text)
        text_widget.config(state=tk.DISABLED)
        
        scrollbar = ttk.Scrollbar(text_widget, command=text_widget.yview)
        text_widget.config(yscrollcommand=scrollbar.set)
    
    def simular_transformada_inversa(self):
        """Simula Transformada Inversa"""
        def ejecutar():
            self.status_label.config(text="⏳ Ejecutando Transformada Inversa...")
            self.root.update()
            
            res = [0 if r < 0.1111 else 1 if r < 0.4444 else 2 if r < 0.6666 else 3 for r in self.RI_BASE]
            df = pd.DataFrame({
                'Día': range(1, 31),
                'ri': self.RI_BASE,
                'Demanda': res
            })
            
            self.guardar_en_excel(df, "Transformada_Inversa", 'bar')
            self.mostrar_estadisticas(df, "Demanda de Cepillos")
            self.simulaciones_realizadas.append("✓ Transformada Inversa")
            self.status_label.config(text="✓ Transformada Inversa completada exitosamente")
            messagebox.showinfo("✓ Éxito", "Simulación de Transformada Inversa completada\nDatos guardados en Excel")
        
        threading.Thread(target=ejecutar, daemon=True).start()
    
    def simular_uniforme(self):
        """Simula Distribución Uniforme"""
        def ejecutar():
            self.status_label.config(text="⏳ Ejecutando Distribución Uniforme...")
            self.root.update()
            
            temp = [round(95 + 5*r, 2) for r in self.RI_BASE]
            df = pd.DataFrame({
                'Medición': range(1, 31),
                'ri': self.RI_BASE,
                'Temp_C': temp
            })
            
            self.guardar_en_excel(df, "Uniforme", 'line')
            self.mostrar_estadisticas(df, "Temperatura en Estufa")
            self.simulaciones_realizadas.append("✓ Distribución Uniforme")
            self.status_label.config(text="✓ Distribución Uniforme completada exitosamente")
            messagebox.showinfo("✓ Éxito", "Simulación de Distribución Uniforme completada\nDatos guardados en Excel")
        
        threading.Thread(target=ejecutar, daemon=True).start()
    
    def simular_poisson(self):
        """Simula Distribución Poisson"""
        def ejecutar():
            self.status_label.config(text="⏳ Ejecutando Distribución Poisson...")
            self.root.update()
            
            def p_val(r):
                if r < 0.1353: return 0
                if r < 0.4060: return 1
                if r < 0.6767: return 2
                if r < 0.8571: return 3
                if r < 0.9473: return 4
                return 5
            
            res = [p_val(r) for r in self.RI_BASE[:20]]
            df = pd.DataFrame({
                'Hora': range(1, 21),
                'ri': self.RI_BASE[:20],
                'Piezas': res
            })
            
            self.guardar_en_excel(df, "Poisson", 'bar')
            self.mostrar_estadisticas(df, "Producción de Piezas")
            self.simulaciones_realizadas.append("✓ Distribución Poisson")
            self.status_label.config(text="✓ Distribución Poisson completada exitosamente")
            messagebox.showinfo("✓ Éxito", "Simulación de Distribución Poisson completada\nDatos guardados en Excel")
        
        threading.Thread(target=ejecutar, daemon=True).start()
    
    def simular_bernoulli(self):
        """Simula Distribución Bernoulli"""
        def ejecutar():
            self.status_label.config(text="⏳ Ejecutando Distribución Bernoulli...")
            self.root.update()
            
            res = [1 if r <= 0.2 else 0 for r in self.RI_BASE]
            df = pd.DataFrame({
                'Día': range(1, 31),
                'ri': self.RI_BASE,
                'Falla': res
            })
            
            self.guardar_en_excel(df, "Bernoulli", 'pie')
            self.mostrar_estadisticas(df, "Probabilidad de Fallas")
            self.simulaciones_realizadas.append("✓ Distribución Bernoulli")
            self.status_label.config(text="✓ Distribución Bernoulli completada exitosamente")
            messagebox.showinfo("✓ Éxito", "Simulación de Distribución Bernoulli completada\nDatos guardados en Excel")
        
        threading.Thread(target=ejecutar, daemon=True).start()
    
    def simular_exponencial(self):
        """Simula Distribución Exponencial"""
        def ejecutar():
            self.status_label.config(text="⏳ Ejecutando Distribución Exponencial...")
            self.root.update()
            
            tiempos = [round(-3 * math.log(1 - r), 2) for r in self.RI_BASE[:15]]
            df = pd.DataFrame({
                'Cliente': range(1, 16),
                'ri': self.RI_BASE[:15],
                'Tiempo_min': tiempos
            })
            
            self.guardar_en_excel(df, "Exponencial", 'line')
            self.mostrar_estadisticas(df, "Tiempo de Atención en Banco")
            self.simulaciones_realizadas.append("✓ Distribución Exponencial")
            self.status_label.config(text="✓ Distribución Exponencial completada exitosamente")
            messagebox.showinfo("✓ Éxito", "Simulación de Distribución Exponencial completada\nDatos guardados en Excel")
        
        threading.Thread(target=ejecutar, daemon=True).start()
    
    def actualizar_resultados(self):
        """Actualiza la lista de resultados"""
        self.listbox_resultados.delete(0, tk.END)
        
        if self.simulaciones_realizadas:
            for sim in self.simulaciones_realizadas:
                self.listbox_resultados.insert(tk.END, sim)
            self.listbox_resultados.insert(tk.END, "")
            self.listbox_resultados.insert(tk.END, f"Total: {len(self.simulaciones_realizadas)} simulación(es)")
        else:
            self.listbox_resultados.insert(tk.END, "⚠️ Aún no hay simulaciones realizadas")
            self.listbox_resultados.insert(tk.END, "")
            self.listbox_resultados.insert(tk.END, "Ejecuta una simulación para ver resultados aquí")
    
    def limpiar_datos(self):
        """Limpia todos los datos"""
        if messagebox.askyesno("Confirmación", "¿Deseas eliminar todos los datos?"):
            if os.path.isfile("Resultados_Simulacion.xlsx"):
                os.remove("Resultados_Simulacion.xlsx")
                self.simulaciones_realizadas = []
                self.actualizar_resultados()
                messagebox.showinfo("✓ Éxito", "Datos eliminados correctamente")
            else:
                messagebox.showwarning("⚠️ Advertencia", "No hay archivos para eliminar")
    
    def generar_archivo_final(self):
        """Genera el archivo final"""
        if os.path.isfile("Resultados_Simulacion.xlsx"):
            messagebox.showinfo("✓ Éxito", f"Archivo 'Resultados_Simulacion.xlsx' listo\n\nSimulaciones incluidas: {len(self.simulaciones_realizadas)}")
        else:
            messagebox.showwarning("⚠️ Advertencia", "No hay datos guardados. Ejecuta primero algunas simulaciones.")
    
    def abrir_carpeta(self):
        """Abre la carpeta del proyecto"""
        carpeta = os.getcwd()
        if os.name == 'nt':  # Windows
            os.startfile(carpeta)
        elif os.name == 'posix':  # Mac/Linux
            os.system(f'open "{carpeta}"')
    
    def abrir_excel(self):
        """Abre el archivo Excel"""
        archivo = "Resultados_Simulacion.xlsx"
        if os.path.isfile(archivo):
            if os.name == 'nt':  # Windows
                os.startfile(archivo)
            elif os.name == 'posix':  # Mac
                os.system(f'open "{archivo}"')
            else:  # Linux
                os.system(f'xdg-open "{archivo}"')
        else:
            messagebox.showwarning("⚠️ Advertencia", f"El archivo '{archivo}' no existe.\nEjecuta primero algunas simulaciones.")

if __name__ == "__main__":
    root = tk.Tk()
    app = SimuladorGUI(root)
    root.mainloop()

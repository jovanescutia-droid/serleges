import tkinter as tk
from tkinter import ttk, messagebox
from docxtpl import DocxTemplate
from num2words import num2words
from datetime import datetime
import os, sys

class SISEReleger:
    def __init__(self, root):
        self.root = root
        self.root.title("SISE - SERLEGES S.A.S. de C.V.")
        self.root.geometry("900x700")
        
        # Ruta de red exacta del servidor según tu reporte
        self.ruta_red = r'\\DESKTOP-EI28A74\Carpeta comparida - Respaldo\Respaldo Serleges\Escritorio\respaldo\FORMATOS\PROMOCIONES, PREVENCIONES Y AUTORICACIONES'
        self.ruta_local = os.path.join(os.path.dirname(os.path.abspath(__file__)), "plantillas")

        # Priorizar red si existe acceso desde esta PC
        self.ruta_activa = self.ruta_red if os.path.exists(self.ruta_red) else self.ruta_local
        self.setup_ui()

    def setup_ui(self):
        frame = tk.Frame(self.root, padx=30, pady=20)
        frame.pack(fill="both", expand=True)

        tk.Label(frame, text="SISTEMA SERLEGES - VERSIÓN WINDOWS 7", font=("Arial", 16, "bold")).pack(pady=10)
        
        # Selector de archivos dinámico
        try:
            archivos = [f for f in os.listdir(self.ruta_activa) if f.endswith('.docx')]
        except:
            archivos = ["Error de conexión a la red"]

        self.combo = ttk.Combobox(frame, values=archivos, state="readonly", width=70)
        self.combo.pack(pady=10)
        if archivos and "Error" not in archivos[0]: self.combo.current(0)

        # Campos de entrada
        self.exp = self.crear_campo(frame, "NÚMERO DE EXPEDIENTE:")
        self.juz = self.crear_campo(frame, "JUZGADO:")
        self.act = self.crear_campo(frame, "ACTUACIÓN (PARA EL NOMBRE DEL ARCHIVO):")
        self.mon = self.crear_campo(frame, "CUANTÍA $ (SOLO NÚMEROS):")

        btn = tk.Button(frame, text="GENERAR E IMPRIMIR LEGAL (OFICIO)", bg="#003366", fg="white", 
                        font=("Arial", 12, "bold"), command=self.procesar, height=2)
        btn.pack(pady=20, fill="x")

    def crear_campo(self, master, txt):
        tk.Label(master, text=txt, font=("Arial", 10, "bold")).pack()
        e = tk.Entry(master, font=("Arial", 11), bd=2)
        e.pack(pady=5, fill="x")
        return e

    def procesar(self):
        try:
            if not self.combo.get(): raise Exception("Seleccione una plantilla")
            
            doc = DocxTemplate(os.path.join(self.ruta_activa, self.combo.get()))
            m = float(self.mon.get() or 0)
            
            ctx = {
                'ACTOR': "JOVAN OCTAVIO ESCUTIA OCAMPO",
                'EXPEDIENTE': self.exp.get().upper(),
                'DEMANDADO': "POR DEFINIR",
                'CUANTIA_NUM': "{:,.2f}".format(m),
                'CUANTIA_LETRA': num2words(m, lang='es').upper() + " PESOS 00/100 M.N.",
                'FECHA': datetime.now().strftime("%d de %B de %Y")
            }
            doc.render(ctx)
            
            # Forzar tamaño LEGAL (Oficio)
            from docx.shared import Inches
            sec = doc.sections[0]
            sec.page_height, sec.page_width = Inches(14), Inches(8.5)
            
            # Guardar en Escritorio
            nombre_file = f"SERLEGES_{self.exp.get()}_{self.act.get()}".replace("/", "-").replace(" ", "_").upper()
            out = os.path.join(os.path.expanduser('~'), 'Desktop', f"{nombre_file}.docx")
            
            doc.save(out)
            messagebox.showinfo("Éxito", f"Archivo guardado en Escritorio:\n{nombre_file}")
            os.startfile(out, "print") # Manda a la impresora predeterminada
            
        except Exception as e: 
            messagebox.showerror("Error", f"Verifique los datos o la red:\n{str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    SISEReleger(root)
    root.mainloop()

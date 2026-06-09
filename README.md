import tkinter as tk
from tkinter import filedialog, messagebox
from PyPDF2 import PdfReader, PdfWriter, PdfMerger
from PIL import Image
from pdf2docx import Converter
from docx import Document
from docx.shared import Inches
from tkinter import filedialog, messagebox
import os

# ============================================================
# FUNCIONES
# ============================================================

def separar_pdf():
 archivo = filedialog.askopenfilename(
title="Seleccionar PDF",
filetypes=[("PDF Files", "*.pdf")]
 )

 if not archivo:
 return

 carpeta_salida = filedialog.askdirectory(
title="Seleccionar carpeta de destino"
 )

 if not carpeta_salida:
 return

 try:
 reader = PdfReader(archivo)

 for i, pagina in enumerate(reader.pages):
 writer = PdfWriter()
 writer.add_page(pagina)

 salida = os.path.join(
 carpeta_salida,
 f"Pagina_{i+1}.pdf"
 )

 with open(salida, "wb") as f:
 writer.write(f)

 messagebox.showinfo(
 "Proceso completado",
 f"Se generaron {len(reader.pages)} archivos PDF."
 )

 except Exception as e:
 messagebox.showerror("Error", str(e))

def unir_pdfs():
 archivos = filedialog.askopenfilenames(
title="Seleccionar PDFs",
filetypes=[("PDF Files", "*.pdf")]
 )

 if not archivos:
 return

 salida = filedialog.asksaveasfilename(
title="Guardar PDF",
defaultextension=".pdf",
filetypes=[("PDF Files", "*.pdf")]
 )

 if not salida:
 return

 try:
 merger = PdfMerger()

 for pdf in archivos:
 merger.append(pdf)

 merger.write(salida)
 merger.close()

 messagebox.showinfo(
 "Proceso completado",
 "Los PDFs fueron unidos correctamente."
 )

 except Exception as e:
 messagebox.showerror("Error", str(e))

def imagen_a_pdf():
 archivos = filedialog.askopenfilenames(
title="Seleccionar imágenes",
filetypes=[
 ("Imágenes", "*.jpg *.jpeg *.png *.bmp *.webp")
 ]
 )

 if not archivos:
 return

 salida = filedialog.asksaveasfilename(
title="Guardar PDF",
defaultextension=".pdf",
filetypes=[("PDF", "*.pdf")]
 )

 if not salida:
 return

 try:
 imagenes = []

 for archivo in archivos:
 try:
 img = Image.open(archivo)

 # Convertir cualquier formato a RGB
 img = img.convert("RGB")

 imagenes.append(img)

 except Exception as e:
 messagebox.showwarning(
 "Advertencia",
 f"No se pudo abrir:\n{archivo}\n\n{e}"
 )

 if len(imagenes) == 0:
 raise Exception("No se encontraron imágenes válidas.")

 primera = imagenes.pop(0)

 primera.save(
 salida,
 "PDF",
resolution=100.0,
save_all=True,
append_images=imagenes
 )

 messagebox.showinfo(
 "Éxito",
 "PDF generado correctamente."
 )

 except Exception as e:
 messagebox.showerror(
 "Error",
 f"No fue posible generar el PDF:\n\n{e}"
 )

def pdf_a_word():
 archivo_pdf = filedialog.askopenfilename(
title="Seleccionar PDF",
filetypes=[("PDF Files", "*.pdf")]
 )

 if not archivo_pdf:
 return

 archivo_docx = filedialog.asksaveasfilename(
title="Guardar Word",
defaultextension=".docx",
filetypes=[("Word", "*.docx")]
 )

 if not archivo_docx:
 return

 try:
 # Convertir PDF a Word
 cv = Converter(archivo_pdf)
 cv.convert(archivo_docx)
 cv.close()

 # Ajustar tamaño Carta
 doc = Document(archivo_docx)

 for section in doc.sections:
 section.page_width = Inches(8.5)
 section.page_height = Inches(11)

 # Márgenes opcionales
 section.top_margin = Inches(1)
 section.bottom_margin = Inches(1)
 section.left_margin = Inches(1)
 section.right_margin = Inches(1)

 doc.save(archivo_docx)

 messagebox.showinfo(
 "Éxito",
 "PDF convertido a Word en tamaño Carta."
 )

 except Exception as e:
 messagebox.showerror(
 "Error",
 f"No fue posible convertir el PDF:\n\n{e}"
 )

# ============================================================
# EFECTOS HOVER
# ============================================================

def hover_enter(event):
event.widget.config(bg="#1e40af")

def hover_leave(event):
 colores = {
 btn_separar: "#2563eb",
 btn_unir: "#16a34a",
 btn_imagen: "#f59e0b",
 btn_word : "#7c3aed",
 btn_salir: "#dc2626"

 }

event.widget.config(bg=colores[event.widget])

# ============================================================
# VENTANA PRINCIPAL
# ============================================================

root = tk.Tk()
root.title("Gestor de PDF")
root.geometry("750x650")
root.resizable(False, False)
root.configure(bg="#eef2f7")

# Centrar ventana
ancho = 550
alto = 500

x = (root.winfo_screenwidth() // 2) - (ancho // 2)
y = (root.winfo_screenheight() // 2) - (alto // 2)

root.geometry(f"{ancho}x{alto}+{x}+{y}")

# ============================================================
# TÍTULO
# ============================================================

titulo = tk.Label(
 root,
text="📄 Gestor de PDF",
font=("Segoe UI", 24, "bold"),
bg="#f5f7fa",
fg="#1f2937"
)
titulo.pack(pady=30)

# ============================================================
# FRAME
# ============================================================

frame = tk.Frame(root, bg="#f5f7fa")
frame.pack(expand=True)

# ============================================================
# BOTONES
# ============================================================

btn_separar = tk.Button(
 frame,
text="📑 Separar PDF",
font=("Segoe UI", 12, "bold"),
bg="#2563eb",
fg="white",
width=25,
height=2,
relief="flat",
cursor="hand2",
command=separar_pdf
)
btn_separar.pack(pady=10)

btn_unir = tk.Button(
 frame,
text="📂 Unir PDFs",
font=("Segoe UI", 12, "bold"),
bg="#16a34a",
fg="white",
width=25,
height=2,
relief="flat",
cursor="hand2",
command=unir_pdfs
)
btn_unir.pack(pady=10)

btn_imagen = tk.Button(
 frame,
text="🖼️ Imagen a PDF",
font=("Segoe UI", 12, "bold"),
bg="#f59e0b",
fg="white",
width=25,
height=2,
relief="flat",
cursor="hand2",
command=imagen_a_pdf
)
btn_imagen.pack(pady=10)

btn_word = tk.Button(
 frame,
text="📝 PDF a Word",
font=("Segoe UI", 12, "bold"),
bg="#7c3aed",
fg="white",
width=25,
height=2,
relief="flat",
cursor="hand2",
command=pdf_a_word
)
btn_word.pack(pady=10)

btn_salir = tk.Button(
 frame,
text="❌ Salir",
font=("Segoe UI", 12, "bold"),
bg="#dc2626",
fg="white",
width=25,
height=2,
relief="flat",
cursor="hand2",
command=root.destroy
)
btn_salir.pack(pady=10)

# ============================================================
# EFECTO HOVER
# ============================================================

for boton in [btn_separar, btn_unir, btn_imagen, btn_salir]:
 boton.bind("<Enter>", hover_enter)
 boton.bind("<Leave>", hover_leave)

# ============================================================
# FOOTER
# ============================================================

footer = tk.Label(
 root,
text="Versión 1.0 | Gestor de PDF",
font=("Segoe UI", 9),
bg="#f5f7fa",
fg="gray"
)
footer.pack(pady=15)

# ============================================================
# EJECUTAR
# ============================================================

root.mainloop()
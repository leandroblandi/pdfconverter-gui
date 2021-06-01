# -*- coding:utf-8 -*-

import tkinter as tk #Tk, Entry, Label, Button, NORMAL
import tkinter.filedialog
import tkinter.messagebox
import docx2pdf, time, os
import sys

""" Funcion que convierte desde
	word a pdf generando uno nuevo con
	nombre 'documento_hora-minuto-segundo.pdf' como resultado """
def convert_pdf(word):
	if word != "" and ".docx" in word:
		desktop_route = os.path.join(os.path.join(os.environ["USERPROFILE"]), "Desktop")
		name = desktop_route + "\\documento_" + time.strftime("%H-%M-%S") + ".pdf"
		return docx2pdf.convert(word, name)
	else:
		sys.exit(1)

# Wrapper con widgets necesarios para el programa
class GUIPDF(object):
	root = None
	file_open = False
	def __init__(self, root):
		# Configuracion basica de la ventana
		self.root = root
		self.root.title("Conversor PDF")
		self.root.geometry("300x150")
		self.root.minsize(300,200)
		self.root.maxsize(300,250)
		self.root.resizable(False, False)
		self.root.rowconfigure(0, weight=1)
		self.root.columnconfigure(0, weight=1)
		self.root.config(bg="#ededed")

	def open_file(self):
		# Abrir un archivo y retornar su ruta
		self.file_open = tkinter.filedialog.askopenfilename(title="Seleccione un archivo .pdf", 
													filetypes=[('DOCX','*.docx')])
		self.entry_route.config(state=tk.NORMAL)
		self.entry_route.insert(0, self.file_open)
		return self.file_open

	def start(self, word):
		# Si no se selecciono un archivo
		if not self.file_open:
			tkinter.messagebox.showerror("Error", "Primero debe seleccionar un archivo")
			return False	
		# Cambiar el estado de la Label y convertir self.file_open
		convert_pdf(word)
		self.state_conversion.config(text="ESTADO: ¡CONVERSION EXITOSA!", 
								     fg="green")

	def help(self):
		tkinter.messagebox.showinfo("Como empezar", 
							"1. Haga click en el botón de archivo, y seleccione el que desee convertir\n2. Una vez seleccionó el archivo tipo Word, presione en el botón de convertir.\n3. Aguarde unos instantes (0.5mins-1min) y será creado el pdf y guardado en el Escritorio de su PC.")

	# Inicializa los widgets de root
	def initialize(self):
		if self.root != None:
			# Titulo
			self.title_label = tk.Label(self.root, 
								     text="Conversor PDF", 
								     bg="#ededed")
			self.title_label.config(font=("Calibri", 18, "bold"))

			self.title_label.grid(row=0, 
								  column=0, 
								  columnspan=3, 
								  padx=5, 
								  pady=5, 
								  sticky="NSEW")

			# Ruta del archivo
			self.entry_route = tk.Entry(self.root, 
									 relief="flat", 
									 background="#ffffff", 
									 disabledbackground="#ffffff")

			self.entry_route.grid(row=1, 
				   				  column=0, 
				   				  padx=5, 
				   				  pady=5, 
				   				  sticky="NSEW")

			# Boton para abrir archivo
			self.open_button = tk.Button(self.root,
									  text="Abrir archivo", 
									  command=self.open_file, 
									  bg="#f09e1a", 
									  relief="flat", 
									  fg="#ffffff", 
									  cursor="hand2")

			self.open_button.grid(row=1, 
								  column=2, 
								  padx=5, 
								  pady=5, 
								  sticky="NSEW")

			# Estado de la conversion
			self.state_conversion = tk.Label(self.root, 
										  text="ESTADO: NO CONVERTIDO", 
										  fg="red", 
										  bg="#ededed")

			self.state_conversion.grid(columnspan=3, 
									   row=2, 
									   column=0, 
									   padx=5, 
									   pady=5, 
									   sticky="NSEW")

			# Boton para convertir
			self.convert_button = tk.Button(self.root, 
										 text="Convertir", 
										 command=lambda: self.start(self.file_open), 
										 relief="flat", 
										 bg="#31b046", 
										 fg="#ffffff", 
										 cursor="hand2", 
										 activebackground="#2d963f")

			self.convert_button.grid(columnspan=3,
									 row=3, 
									 column=0, 
									 padx=5, 
									 pady=5, 
									 sticky="NSEW")

			self.about_label = tk.Label(self.root, 
								     text="Creado por Leandro Blandi", 
								     bg="#ededed")

			self.about_label.config(font=("Helvetica", 8))

			self.about_label.grid(row=4, 
								  column=0, 
								  columnspan=2,
								  padx=5, 
								  pady=10, 
								  sticky="NSEW")
			self.help_button = tk.Button(self.root, 
									  text="?",
									  font=("Helvetica", 18), 
									  command=self.help,
									  relief="flat",
									  bg="#cfe6e5",
									  cursor="hand2")
			self.help_button.grid(row=4,
								  column=1,
								  columnspan=2,
								  padx=5,
								  pady=5,
								  sticky="NSEW")

if __name__ == "__main__":
	root = tk.Tk()
	gui = GUIPDF(root)
	gui.initialize()
	root.mainloop()

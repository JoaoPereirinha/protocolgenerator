import pandas as pd
from docxtpl import DocxTemplate
import os
import tkinter as tk
from tkinter import filedialog, messagebox

def escolher_excel():
    path = filedialog.askopenfilename(filetypes=[("Ficheiros Excel", "*.xlsx")])
    entrada_excel.delete(0, tk.END)
    entrada_excel.insert(0, path)

def escolher_modelo():
    path = filedialog.askopenfilename(filetypes=[("Ficheiros Word", "*.docx")])
    entrada_modelo.delete(0, tk.END)
    entrada_modelo.insert(0, path)

def escolher_destino():
    path = filedialog.askdirectory()
    entrada_destino.delete(0, tk.END)
    entrada_destino.insert(0, path)

def gerar_documentos():
    excel_path = entrada_excel.get()
    template_path = entrada_modelo.get()
    output_dir = entrada_destino.get()

    if not all([excel_path, template_path, output_dir]):
        messagebox.showerror("Erro", "Por favor preencha todos os campos.")
        return

    try:
        df = pd.read_excel(excel_path)
        for _, row in df.iterrows():
            doc = DocxTemplate(template_path)
            contexto = row.to_dict()
            nome_ficheiro = f"{contexto['nomeAluno'].replace(' ', '_')}_protocolo.docx"
            caminho_saida = os.path.join(output_dir, nome_ficheiro)
            doc.render(contexto)
            doc.save(caminho_saida)
        messagebox.showinfo("Sucesso", "Documentos gerados com sucesso!")
    except Exception as e:
        messagebox.showerror("Erro ao gerar documentos", str(e))

# GUI
janela = tk.Tk()
janela.title("Gerador de Protocolos de Estágio")
janela.geometry("600x300")

tk.Label(janela, text="Ficheiro Excel com dados:").pack()
entrada_excel = tk.Entry(janela, width=70)
entrada_excel.pack()
tk.Button(janela, text="Escolher ficheiro", command=escolher_excel).pack()

tk.Label(janela, text="Modelo Word:").pack()
entrada_modelo = tk.Entry(janela, width=70)
entrada_modelo.pack()
tk.Button(janela, text="Escolher ficheiro", command=escolher_modelo).pack()

tk.Label(janela, text="Pasta de destino:").pack()
entrada_destino = tk.Entry(janela, width=70)
entrada_destino.pack()
tk.Button(janela, text="Escolher pasta", command=escolher_destino).pack()

tk.Button(janela, text="Gerar Documentos", bg="green", fg="white", command=gerar_documentos).pack(pady=10)

janela.mainloop()

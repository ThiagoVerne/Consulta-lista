import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

df = None


def carregar_planilha():
    global df
    caminho = filedialog.askopenfilename(
        title="Selecione a planilha Excel",
        filetypes=[("Arquivos Excel", "*.xlsx")]
    )
    if caminho:
        try:
            df = pd.read_excel(caminho)
            messagebox.showinfo("Sucesso", "Planilha carregada com sucesso!")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao carregar planilha:\n{e}")


def consultar():
    global df
    if df is None:
        messagebox.showwarning("Aviso", "Carregue a planilha primeiro.")
        return

    valor_a = entrada_a.get().strip()
    valor_b = entrada_b.get().strip()

    if not valor_a or not valor_b:
        messagebox.showwarning("Aviso", "Preencha os dois parâmetros.")
        return

    
    resultado = df[
        (df.iloc[:, 0].astype(str) == valor_a) &
        (df.iloc[:, 1].astype(str) == valor_b)
    ]

    if not resultado.empty:
        resposta = str(resultado.iloc[0, 2])  
      
        if len(resposta) < 1300:
            resposta = resposta.ljust(1300, " ")
        campo_resposta.config(state=tk.NORMAL)
        campo_resposta.delete(1.0, tk.END)
        campo_resposta.insert(tk.END, resposta)
        campo_resposta.config(state=tk.DISABLED)
    else:
        messagebox.showinfo("Resultado", "Nenhum registro encontrado.")

# tela principal
janela = tk.Tk()
janela.title("Martriz de Acesso")
janela.geometry("800x600")
janela.resizable(False, False)

# Botão para carregar planilha
btn_carregar = tk.Button(janela, text="Carregar Planilha", command=carregar_planilha)
btn_carregar.pack(pady=5)

# Campo de entrada
frame_inputs = tk.Frame(janela)
frame_inputs.pack(pady=10)

tk.Label(frame_inputs, text="Área:").grid(row=0, column=0, padx=5, pady=5)
entrada_a = tk.Entry(frame_inputs, width=30)
entrada_a.grid(row=0, column=1, padx=5, pady=5)

tk.Label(frame_inputs, text="Cargo:").grid(row=1, column=0, padx=5, pady=5)
entrada_b = tk.Entry(frame_inputs, width=30)
entrada_b.grid(row=1, column=1, padx=5, pady=5)

# Botão de consulta
btn_consultar = tk.Button(janela, text="Consultar", command=consultar)
btn_consultar.pack(pady=5)

# Campo de resposta
tk.Label(janela, text="Pastas").pack()
campo_resposta = tk.Text(janela, height=7, width=80, wrap=tk.WORD)
campo_resposta.pack(pady=5)
campo_resposta.config(state=tk.DISABLED)

tk.Label(janela, text="Aplicativos").pack()
campo_resposta = tk.Text(janela, height=7, width=80, wrap=tk.WORD)
campo_resposta.pack(pady=5)
campo_resposta.config(state=tk.DISABLED)

tk.Label(janela, text="Acessos a Ambientes").pack()
campo_resposta = tk.Text(janela, height=7, width=80, wrap=tk.WORD)
campo_resposta.pack(pady=5)
campo_resposta.config(state=tk.DISABLED)
# Rodar programa
janela.mainloop()
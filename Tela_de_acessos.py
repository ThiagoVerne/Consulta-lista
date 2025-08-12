import pandas as pd
import tkinter as tk
from tkinter import messagebox, filedialog

# colocar o caminho da planilha para puxar os acessos
CAMINHO_PLANILHA = r"Insira o caminho da planilha c:/pastaxpto "


df = None
valores_a = []
valores_b = []



def carregar_planilha():
    global df, valores_a, valores_b

    try:
        df = pd.read_excel(CAMINHO_PLANILHA, engine='openpyxl')

        valores_a = sorted(df.iloc[:, 0].dropna().astype(str).unique())
        valores_b = sorted(df.iloc[:, 1].dropna().astype(str).unique())

        atualizar_dropdown(menu_a, selecionado_a, valores_a)
        atualizar_dropdown(menu_b, selecionado_b, valores_b)

        messagebox.showinfo("Sucesso", "Planilha carregada com sucesso!")

    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao carregar a planilha:\n{e}")


def atualizar_dropdown(menu_widget, var_control, lista_valores):
    menu_widget['menu'].delete(0, 'end')
    if lista_valores:
        var_control.set(lista_valores[0])
        for val in lista_valores:
            menu_widget['menu'].add_command(label=val, command=tk._setit(var_control, val))
    else:
        var_control.set("")


def consultar():
    global df

    if df is None:
        messagebox.showwarning("Aviso", "A planilha ainda não foi carregada.")
        return

    valor_a = selecionado_a.get().strip()
    valor_b = selecionado_b.get().strip()

    if not valor_a or not valor_b:
        messagebox.showwarning("Aviso", "Selecione ambos os parâmetros.")
        return

    resultado = df[
        (df.iloc[:, 0].astype(str) == valor_a) &
        (df.iloc[:, 1].astype(str) == valor_b)
    ]

    if not resultado.empty:
        campo1.config(state=tk.NORMAL)
        campo2.config(state=tk.NORMAL)
        campo3.config(state=tk.NORMAL)

        campo1.delete(1.0, tk.END)
        campo2.delete(1.0, tk.END)
        campo3.delete(1.0, tk.END)

        campo1.insert(tk.END, str(resultado.iloc[0, 2]))  # Coluna C
        campo2.insert(tk.END, str(resultado.iloc[0, 3]))  # Coluna D
        campo3.insert(tk.END, str(resultado.iloc[0, 4]))  # Coluna E

        campo1.config(state=tk.DISABLED)
        campo2.config(state=tk.DISABLED)
        campo3.config(state=tk.DISABLED)
    else:
        messagebox.showinfo("Resultado", "Nenhum registro encontrado.")


def exportar_txt():
    if df is None:
        messagebox.showwarning("Aviso", "A planilha ainda não foi carregada.")
        return

    valor_a = selecionado_a.get().strip()
    valor_b = selecionado_b.get().strip()

    resultado = df[
        (df.iloc[:, 0].astype(str) == valor_a) &
        (df.iloc[:, 1].astype(str) == valor_b)
    ]

    if resultado.empty:
        messagebox.showinfo("Exportar", "Nenhum registro para exportar.")
        return

    caminho_arquivo = filedialog.asksaveasfilename(
        defaultextension=".txt",
        filetypes=[("Arquivo de Texto", "*.txt")],
        title="Salvar como"
    )

    if not caminho_arquivo:
        return

    try:
        with open(caminho_arquivo, "w", encoding="utf-8") as f:
            for _, row in resultado.iterrows():
                linha = f"{row[0]} | {row[1]} | {row[2]} | {row[3]} | {row[5]}"
                f.write(linha + "\n")
        messagebox.showinfo("Exportar", "Dados exportados com sucesso!")

    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao exportar:\n{e}")

# colocar o nome da empresa desejada na linha 119
janela = tk.Tk()
janela.title("Matriz de Acessos Empresa TAL") 
janela.geometry("850x600")
janela.resizable(False, False)

selecionado_a = tk.StringVar()
selecionado_b = tk.StringVar()

frame_select = tk.Frame(janela)
frame_select.pack(pady=20)

tk.Label(frame_select, text="Área:").grid(row=0, column=0, padx=5, pady=5)
menu_a = tk.OptionMenu(frame_select, selecionado_a, "")
menu_a.config(width=40)
menu_a.grid(row=0, column=1, padx=5, pady=5)

tk.Label(frame_select, text="Cargo:").grid(row=1, column=0, padx=5, pady=5)
menu_b = tk.OptionMenu(frame_select, selecionado_b, "")
menu_b.config(width=40)
menu_b.grid(row=1, column=1, padx=5, pady=5)

btn_consultar = tk.Button(janela, text="Consultar", command=consultar)
btn_consultar.pack(pady=10)

def criar_campo_rotulado(texto, altura):
    tk.Label(janela, text=texto).pack()
    campo = tk.Text(janela, height=altura, width=100, wrap=tk.WORD)
    campo.pack(pady=5)
    campo.config(state=tk.DISABLED)
    return campo

campo1 = criar_campo_rotulado("Pasta Rede:", 5)
campo2 = criar_campo_rotulado("Aplicativos ou Sistemas:", 5)
campo3 = criar_campo_rotulado("Acesso a Ambiente:", 5)

btn_exportar = tk.Button(janela, text="Exportar Resultado em TXT", command=exportar_txt)
btn_exportar.pack(pady=20)


carregar_planilha()


janela.mainloop()

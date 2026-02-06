import tkinter as tk
from tkinter import filedialog, messagebox

# ===== IMPORTA SEU CÓDIGO =====
from core import (
    read_with_fallback,
    main_itau,
    main_santander,
    main_bradesco
)

arquivo_selecionado = None


def escolher_arquivo():
    global arquivo_selecionado

    arquivo = filedialog.askopenfilename(
        title="Selecione o extrato",
        filetypes=[("Excel", "*.xls *.xlsx")]
    )

    if arquivo:
        arquivo_selecionado = arquivo
        lbl_arquivo["text"] = arquivo


def gerar():
    if not arquivo_selecionado:
        messagebox.showwarning("Atenção", "Selecione um arquivo primeiro!")
        return

    banco = banco_var.get()

    try:
        # ===== EXATAMENTE SEU MÉTODO =====
        df = read_with_fallback(arquivo_selecionado)
        df.attrs["arquivo_origem"] = arquivo_selecionado  # Armazena o caminho do arquivo no DataFrame

        if banco == "Itaú":

            main_itau(df)

            messagebox.showinfo(
                "Sucesso",
                "Relatório Itaú gerado com sucesso!"
            )

        elif banco == "Santander":

            main_santander(df)

            messagebox.showinfo(
                "Sucesso",
                "Relatório Santander gerado com sucesso!"
            )

        elif banco == "Bradesco":

            main_bradesco(df)
            messagebox.showinfo(
                "sucesso",
                "Relatório Bradesco gerado com sucesso!"
            )

        else:
            messagebox.showwarning(
                "Atenção",
                "Selecione um banco"
            )

    except Exception as e:
        messagebox.showerror(
            "Erro",
            f"Erro ao processar:\n{str(e)}"
        )


# ============ TELA ============

app = tk.Tk()
app.title("Gerador de Relatório Bancário")
app.geometry("520x260")

tk.Label(app, text="Selecione o Banco:").pack(pady=5)

banco_var = tk.StringVar()

combo = tk.OptionMenu(
    app,
    banco_var,
    "Itaú",
    "Santander",
    "Bradesco"
)
combo.pack()

tk.Button(
    app,
    text="Selecionar Arquivo",
    command=escolher_arquivo
).pack(pady=10)

lbl_arquivo = tk.Label(app, text="Nenhum arquivo selecionado")
lbl_arquivo.pack()

tk.Button(
    app,
    text="GERAR RELATÓRIO",
    command=gerar,
    bg="#003399",
    fg="white",
    width=20
).pack(pady=15)

app.mainloop()

import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import os
import threading


def selecionar_arquivo():
    caminho = filedialog.askopenfilename(
        title="Selecione a planilha",
        filetypes=[("Arquivos Excel", "*.xls *.xlsx")]
    )
    if caminho:
        entrada_arquivo.delete(0, tk.END)
        entrada_arquivo.insert(0, caminho)


def selecionar_pasta():
    caminho = filedialog.askdirectory(title="Selecione a pasta de destino")
    if caminho:
        entrada_pasta.delete(0, tk.END)
        entrada_pasta.insert(0, caminho)


def dividir_planilha_thread():
    try:
        arquivo = entrada_arquivo.get()
        pasta_saida = entrada_pasta.get()

        if not arquivo:
            messagebox.showwarning("Aviso", "Selecione um arquivo.")
            return

        if not pasta_saida:
            messagebox.showwarning("Aviso", "Selecione a pasta de destino.")
            return

        # 🔄 fase 1: leitura (feedback visual)
        progresso.config(mode="indeterminate")
        progresso.start(10)
        status_label.config(text="Lendo planilha...")
        janela.update_idletasks()

        df = pd.read_excel(arquivo, dtype=str)

        progresso.stop()
        progresso.config(mode="determinate")

        limite = 20000
        total = len(df)
        total_partes = (total - 1) // limite + 1

        base_nome = os.path.splitext(os.path.basename(arquivo))[0]

        progresso["maximum"] = total_partes
        progresso["value"] = 0

        for i in range(0, total, limite):
            parte_num = i // limite + 1
            inicio = i + 1
            fim = min(i + limite, total)

            percentual = int((parte_num / total_partes) * 100)

            status_label.config(
                text=f"Parte {parte_num}/{total_partes} | Registros {inicio}-{fim} de {total} | {percentual}%"
            )

            parte = df.iloc[i:i+limite]

            nome_saida = os.path.join(
                pasta_saida,
                f"{base_nome}_parte_{parte_num}.xlsx"
            )

            with pd.ExcelWriter(nome_saida, engine="xlsxwriter") as writer:
                parte.to_excel(writer, index=False, sheet_name="Dados")

                workbook = writer.book
                worksheet = writer.sheets["Dados"]

                formato_texto = workbook.add_format({'num_format': '@'})

                for col_idx in range(len(parte.columns)):
                    col_letter = chr(65 + col_idx)
                    worksheet.set_column(f'{col_letter}:{col_letter}', 20, formato_texto)

            progresso["value"] = parte_num
            janela.update_idletasks()

        status_label.config(text="Concluído!")

        messagebox.showinfo(
            "Sucesso",
            f"{total_partes} arquivos gerados.\nTotal de registros: {total}"
        )

    except Exception as e:
        messagebox.showerror("Erro", str(e))


def iniciar_processo():
    thread = threading.Thread(target=dividir_planilha_thread)
    thread.start()


# Interface
janela = tk.Tk()
janela.title("Divisor de Planilhas")
janela.geometry("560x280")

tk.Label(janela, text="Arquivo Excel:").pack(pady=5)

entrada_arquivo = tk.Entry(janela, width=70)
entrada_arquivo.pack(pady=5)

tk.Button(janela, text="Selecionar Arquivo", command=selecionar_arquivo).pack(pady=5)

tk.Label(janela, text="Pasta de destino:").pack(pady=5)

entrada_pasta = tk.Entry(janela, width=70)
entrada_pasta.pack(pady=5)

tk.Button(janela, text="Selecionar Pasta", command=selecionar_pasta).pack(pady=5)

tk.Button(janela, text="Dividir Planilha", command=iniciar_processo).pack(pady=10)

progresso = ttk.Progressbar(janela, length=420, mode="determinate")
progresso.pack(pady=10)

status_label = tk.Label(janela, text="Aguardando...")
status_label.pack()

janela.mainloop()
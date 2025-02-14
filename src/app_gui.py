import os
import tkinter as tk
from datetime import datetime
from tkinter import filedialog, messagebox

from tim.black.consolidate import consolidate as consolidate_black
from tim.empresa_nacional.consolidate import \
    consolidate as consolidate_empresa_nacional


def init_consolidate():
    planilha_func = entry_func.get()
    planilha_fatura_tim = entry_fatura_tim.get()
    mes_ano = entry_mes_ano.get()
    
    if not os.path.isfile(planilha_func):
        messagebox.showerror("Erro", "Planilha de Funcionários não encontrada.")
        return
    if not os.path.isfile(planilha_fatura_tim):
        messagebox.showerror("Erro", "Planilha TIM não encontrada.")
        return
    
    try:
        mes_str, ano_str = mes_ano.split('/')
        mes = int(mes_str)
        ano = int(ano_str)
        ano_atual = datetime.now().year
        if not (1 <= mes <= 12 and ano <= ano_atual):
            messagebox.showerror("Erro", "Mês deve ser entre 01 e 12 e ano deve ser até o ano atual.")
            return
    except Exception:
        messagebox.showerror("Erro", "Formato de mês/ano inválido. Use MM/AAAA.")
        return

    plano = plano_var.get()
    try:
        if plano == "TIM BLACK":
            consolidate_black(planilha_func, planilha_fatura_tim, mes, ano)
        elif plano == "TIM EMPRESA NACIONAL":
            consolidate_empresa_nacional(planilha_func, planilha_fatura_tim, mes, ano)
        else:
            messagebox.showerror("Erro", "Plano inválido selecionado.")
    except Exception as e:
        messagebox.showerror("Erro", f"Erro durante consolidação: {str(e)}")

def select_funcionario():
    file_path = filedialog.askopenfilename(
        filetypes=[("Excel Files", "*.xlsx")],
        title="Selecione a planilha de Funcionários"
    )
    if file_path:
        entry_func.delete(0, tk.END)
        entry_func.insert(0, file_path)

def select_fatura_tim():
    file_path = filedialog.askopenfilename(
        filetypes=[("Excel Files", "*.xlsx")],
        title="Selecione a planilha da Fatura TIM"
    )
    if file_path:
        entry_fatura_tim.delete(0, tk.END)
        entry_fatura_tim.insert(0, file_path)

root = tk.Tk()
root.title("Consolidação de Fatura TIM")
root.geometry("580x250")

main_frame = tk.Frame(root, bg=root["bg"])
main_frame.pack(fill="both", expand=True, padx=20, pady=20)

plano_var = tk.StringVar(value="TIM BLACK")
label_plano = tk.Label(main_frame, text="Selecione o plano:")
label_plano.grid(row=0, column=0, pady=5, padx=5, sticky="e")
frame_planos = tk.Frame(main_frame)
frame_planos.grid(row=0, column=1, pady=5, padx=5, sticky="w")

radio_black = tk.Radiobutton(frame_planos, text="TIM BLACK", variable=plano_var, value="TIM BLACK")
radio_black.pack(side="left", padx=5)

radio_empresa = tk.Radiobutton(frame_planos, text="TIM EMPRESA NACIONAL", variable=plano_var, value="TIM EMPRESA NACIONAL")
radio_empresa.pack(side="left", padx=5)

label_func = tk.Label(main_frame, text="Planilha Funcionários:")
label_func.grid(row=1, column=0, pady=5, padx=5, sticky="e")
entry_func = tk.Entry(main_frame, width=50)
entry_func.grid(row=1, column=1, pady=5, padx=5)
btn_func = tk.Button(main_frame, text="Selecionar", command=select_funcionario)
btn_func.grid(row=1, column=2, pady=5, padx=5)

label_fat = tk.Label(main_frame, text="Planilha Fatura TIM:")
label_fat.grid(row=2, column=0, pady=5, padx=5, sticky="e")
entry_fatura_tim = tk.Entry(main_frame, width=50)
entry_fatura_tim.grid(row=2, column=1, pady=5, padx=5)
btn_fat = tk.Button(main_frame, text="Selecionar", command=select_fatura_tim)
btn_fat.grid(row=2, column=2, pady=5, padx=5)

label_mes_ano = tk.Label(main_frame, text="Mês/Ano (MM/AAAA):")
label_mes_ano.grid(row=3, column=0, pady=5, padx=5, sticky="e")
entry_mes_ano = tk.Entry(main_frame, width=50)
entry_mes_ano.grid(row=3, column=1, pady=5, padx=5)

btn_consolidar = tk.Button(
    main_frame,
    text="Consolidar",
    command=init_consolidate,
    width=20,
    height=2,
    font=("Arial", 10, "bold")
)
btn_consolidar.grid(row=4, column=1, pady=15)

root.mainloop()
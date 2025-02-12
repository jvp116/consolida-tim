import tkinter as tk
from tkinter import filedialog, messagebox
import os
from datetime import datetime
from utils import (
    extrair_dados_planilha,
    determinar_conta_contabil,
    determinar_subconta_entidade,
    determinar_codigo_departamento,
    determinar_restricao,
    determinar_envio_aviso,
    determinar_historico
)

def consolidate():
    planilha_func = entry_func.get()
    planilha_tb = entry_tb.get()
    mes_ano = entry_mes_ano.get()
    
    # Valida os caminhos dos arquivos
    if not os.path.isfile(planilha_func):
        messagebox.showerror("Erro", "Planilha de Funcionários não encontrada.")
        return
    if not os.path.isfile(planilha_tb):
        messagebox.showerror("Erro", "Planilha TIM Black não encontrada.")
        return
    
    # Valida o formato do mês/ano e extrai mes e ano
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

    try:
        # Carrega dados da planilha de Funcionários
        df_func = extrair_dados_planilha(
            planilha_func, "TIM BLACK", ["Cód", "Nome", "Número", "Auxílio"], "Número"
        )
        # Tratar valores nulos na coluna "Auxílio"
        df_func["Auxílio"] = df_func["Auxílio"].fillna(0)
        
        # Carrega dados da planilha TIM Black
        df_tb = extrair_dados_planilha(
            planilha_tb, "Resumo Detalhamento", ["Acesso", "Valor"]
        )

        # Agrupa os dados da planilha TIM Black por "Acesso"
        df_grouped = df_tb.groupby("Acesso", as_index=False).sum()

        # Junção com os dados de Funcionários (df_func) utilizando "Acesso" e o índice "Número" dos funcionários
        df_pre_consolida = df_grouped.merge(
            df_func, left_on="Acesso", right_index=True, how="left"
        )

        # Renomear colunas conforme a lógica de consolidação
        df_pre_consolida.rename(
            columns={
                "Cód": "Subconta/Entidade",
                "Acesso": "Número",
                "Valor": "Valor Bruto",
                "Auxílio": "Valor Auxílio"
            },
            inplace=True
        )

        # Preencher valores ausentes
        df_pre_consolida.fillna({
            "Subconta/Entidade": 0,
            "Nome": "Funcionário não encontrado",
            "Valor Auxílio": 0
        }, inplace=True)

        # Calcular valor líquido
        df_pre_consolida["Valor Líquido"] = df_pre_consolida["Valor Bruto"] - df_pre_consolida["Valor Auxílio"]

        # Filtrar apenas os registros com valor líquido positivo
        df_pre_consolida = df_pre_consolida[df_pre_consolida["Valor Líquido"] > 0]

        # Converter Subconta/Entidade para inteiro e ordenar
        df_pre_consolida["Subconta/Entidade"] = df_pre_consolida["Subconta/Entidade"].astype(int)
        df_pre_consolida = df_pre_consolida.sort_values(by="Subconta/Entidade")

        # Aplicar as funções para criar as colunas adicionais
        df_pre_consolida["Conta Contábil"] = df_pre_consolida["Subconta/Entidade"].apply(determinar_conta_contabil)
        df_pre_consolida["Código Departamento"] = df_pre_consolida["Subconta/Entidade"].apply(determinar_codigo_departamento)
        df_pre_consolida["Histórico"] = df_pre_consolida.apply(
            lambda row: determinar_historico(row["Subconta/Entidade"], row["Número"], row["Nome"], mes, ano), axis=1
        )
        df_pre_consolida["Subconta/Entidade"] = df_pre_consolida.apply(
            lambda row: determinar_subconta_entidade(row["Conta Contábil"], row["Subconta/Entidade"]), axis=1
        )
        df_pre_consolida["Fundo"] = 10
        df_pre_consolida["Restrição"] = df_pre_consolida["Conta Contábil"].apply(determinar_restricao)
        df_pre_consolida["Envio de Aviso"] = df_pre_consolida["Conta Contábil"].apply(determinar_envio_aviso)

        # Selecionar colunas na ordem correta
        df_final = df_pre_consolida[[
            "Conta Contábil", "Subconta/Entidade", "Fundo", "Código Departamento",
            "Restrição", "Valor Líquido", "Envio de Aviso", "Histórico"
        ]]

        # Perguntar ao usuário onde deseja salvar o arquivo consolidado
        output_file = filedialog.asksaveasfilename(
            defaultextension='.xlsx', 
            filetypes=[("Excel Files", "*.xlsx")], 
            title="Salvar Consolidação"
        )
        if output_file:
            df_final.to_excel(output_file, index=False, engine="openpyxl")
            messagebox.showinfo("Sucesso", f"Consolidação realizada com sucesso e salva em: {output_file}")
    except Exception as e:
        messagebox.showerror("Erro", f"Erro durante consolidação: {str(e)}")

def select_func():
    file_path = filedialog.askopenfilename(
        filetypes=[("Excel Files", "*.xlsx")],
        title="Selecione a planilha de Funcionários"
    )
    if file_path:
        entry_func.delete(0, tk.END)
        entry_func.insert(0, file_path)

def select_tb():
    file_path = filedialog.askopenfilename(
        filetypes=[("Excel Files", "*.xlsx")],
        title="Selecione a planilha TIM Black"
    )
    if file_path:
        entry_tb.delete(0, tk.END)
        entry_tb.insert(0, file_path)

# Configuração da janela principal com TKinterModernThemes
root = tk.Tk()

root.title("Consolidação de Fatura TIM")
root.geometry("600x200")  # Adjusted size for modern layout

# Create a main container frame with padding
main_frame = tk.Frame(root, bg=root["bg"])
main_frame.pack(fill="both", expand=True, padx=20, pady=20)

# Modern interface with same components:
label_func = tk.Label(main_frame, text="Planilha Funcionários:")
label_func.grid(row=0, column=0, pady=5, padx=5, sticky="e")
entry_func = tk.Entry(main_frame, width=50)
entry_func.grid(row=0, column=1, pady=5, padx=5)
btn_func = tk.Button(main_frame, text="Selecionar", command=select_func)
btn_func.grid(row=0, column=2, pady=5, padx=5)

label_tb = tk.Label(main_frame, text="Planilha TIM Black:")
label_tb.grid(row=1, column=0, pady=5, padx=5, sticky="e")
entry_tb = tk.Entry(main_frame, width=50)
entry_tb.grid(row=1, column=1, pady=5, padx=5)
btn_tb = tk.Button(main_frame, text="Selecionar", command=select_tb)
btn_tb.grid(row=1, column=2, pady=5, padx=5)

label_mes_ano = tk.Label(main_frame, text="Mês/Ano (MM/AAAA):")
label_mes_ano.grid(row=2, column=0, pady=5, padx=5, sticky="e")
entry_mes_ano = tk.Entry(main_frame, width=50)
entry_mes_ano.grid(row=2, column=1, pady=5, padx=5)

btn_consolidar = tk.Button(main_frame, text="Consolidar", command=consolidate)
btn_consolidar.grid(row=3, column=1, pady=15)

root.mainloop()
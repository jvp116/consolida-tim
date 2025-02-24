from tkinter import filedialog, messagebox

from tim.black.utils import (determinar_codigo_departamento,
                             determinar_conta_contabil, determinar_envio_aviso,
                             determinar_historico, determinar_restricao,
                             determinar_subconta_entidade,
                             extrair_dados_planilha)


def consolidate(planilha_func, planilha_fatura_tim, mes, ano):
    try:
        df_func = extrair_dados_planilha(
            planilha_func, "TIM BLACK", ["Cód", "Nome/Depto", "Número", "Auxílio"], "Número"
        )
        # Tratar valores nulos na coluna "Auxílio"
        df_func["Auxílio"] = df_func["Auxílio"].fillna(0)
        
        df_fat = extrair_dados_planilha(
            planilha_fatura_tim, "Resumo Detalhamento", ["Acesso", "Valor"]
        )

        # Agrupa os dados da planilha por "Acesso"
        df_grouped = df_fat.groupby("Acesso", as_index=False).sum()

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
            "Nome/Depto": "Funcionário não encontrado",
            "Valor Auxílio": 0
        }, inplace=True)

        # Convert "Valor Bruto" from text with comma separator to float
        df_pre_consolida["Valor Bruto"] = df_pre_consolida["Valor Bruto"].astype(str).str.replace(",", ".").astype(float)
        
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
            lambda row: determinar_historico(row["Subconta/Entidade"], row["Número"], row["Nome/Depto"], mes, ano), axis=1
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

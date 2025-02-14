from tkinter import filedialog, messagebox

from tim.empresa_nacional.utils import (determinar_codigo_departamento,
                                        determinar_conta_contabil,
                                        determinar_envio_aviso,
                                        determinar_historico,
                                        determinar_restricao,
                                        determinar_subconta_entidade,
                                        extrair_dados_planilha)


def consolidate(planilha_func, planilha_fatura_tim, mes, ano):
    try:
        df_func = extrair_dados_planilha(
            planilha_func, "TIM EMPRESA NACIONAL", ["COD/DPTO", "NOME", "NUMERO"], "NUMERO"
        )
        
        df_tb = extrair_dados_planilha(
            planilha_fatura_tim, "Resumo Detalhamento", ["Acesso", "Valor"]
        )

        # Agrupa os dados da planilha por "Acesso"
        df_grouped = df_tb.groupby("Acesso", as_index=False).sum()

        # Junção com os dados de Funcionários (df_func) utilizando "Acesso" e o índice "NUMERO" dos funcionários
        df_pre_consolida = df_grouped.merge(
            df_func, left_on="Acesso", right_index=True, how="left"
        )
        
        # Renomear colunas conforme a lógica de consolidação
        df_pre_consolida.rename(
            columns={
                "COD/DPTO": "Subconta/Entidade",
                "Acesso": "Número",
            },
            inplace=True
        )
        
        # Preencher valores ausentes
        df_pre_consolida.fillna({
            "Subconta/Entidade": 0,
            "Nome": "Funcionário não encontrado",
        }, inplace=True)

        # Converter Subconta/Entidade para inteiro e ordenar
        df_pre_consolida["Subconta/Entidade"] = df_pre_consolida["Subconta/Entidade"].astype(int)
        df_pre_consolida = df_pre_consolida.sort_values(by="Subconta/Entidade")

        # Aplicar as funções para criar as colunas adicionais
        df_pre_consolida["Conta Contábil"] = df_pre_consolida["Subconta/Entidade"].apply(determinar_conta_contabil)
        df_pre_consolida["Código Departamento"] = df_pre_consolida["Subconta/Entidade"].apply(determinar_codigo_departamento)
        df_pre_consolida["Histórico"] = df_pre_consolida.apply(
            lambda row: determinar_historico(row["Subconta/Entidade"], row["Número"], row["NOME"], mes, ano), axis=1
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
            "Restrição", "Valor", "Envio de Aviso", "Histórico"
        ]]

        print(df_final)
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
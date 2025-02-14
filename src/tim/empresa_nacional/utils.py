import pandas as pd


def extrair_dados_planilha(caminho_arquivo, nome_aba, colunas_desejadas, indice=None):
    df = pd.read_excel(caminho_arquivo, sheet_name=nome_aba, usecols=colunas_desejadas)
    if indice and indice in df.columns:
        df.set_index(indice, inplace=True)
    return df

def determinar_conta_contabil(codigo):
    if codigo == 0:
        return "Funcionário não encontrado"
    elif codigo <= 999:
        return 1135001
    elif len(str(codigo)) in [4, 6]:
        return 2136001
    elif codigo >= 600000000:
        return 4121090
    return None

def determinar_subconta_entidade(conta_contabil, subconta_entidade):
    if conta_contabil == 2136001:
        return 101321
    elif conta_contabil == 4121090:
        return ""
    return subconta_entidade

def determinar_codigo_departamento(codigo):
    if len(str(codigo)) in [1, 2, 3, 4, 6]:
        return 600000101
    return codigo

def determinar_restricao(conta_contabil):
    if str(conta_contabil).startswith(('1', '2')):
        return "0a"
    elif str(conta_contabil).startswith(('3', '4')):
        return "0e"
    return ""

def determinar_envio_aviso(conta_contabil):
    if conta_contabil == 2136001:
        return "true"
    return "false"

def determinar_historico(codigo, numero, nome, mes, ano):
    if len(str(codigo)) <= 3:
        return f"Adto. Salário Fatura TIM {mes}/{ano} - {numero}"
    elif len(str(codigo)) in [4, 6]:
        return f"{codigo} - Fatura TIM {mes}/{ano} - {nome} - {numero}"
    elif codigo >= 600000000:
        return f"Despesas Fatura TIM {mes}/{ano} - {numero}"
    return ""

import pandas as pd
import numpy as np
from datetime import datetime
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import PatternFill, Font, Alignment
import warnings
warnings.filterwarnings('ignore')

# ======================
# CONFIGURAÇÕES INICIAIS
# ======================

# Dicionário com as divisões, GOs e UFs
DIVISOES = {
    "DIV 01": {"GO": ["GO 01"], "UFs": ["AL", "CE", "PB", "PE", "RN"]},
    "DIV 02": {"GO": ["GO 01", "GO 02"], "UFs": ["BA", "SE"]},
    "DIV 03": {"GO": ["GO 01"], "UFs": ["CE", "MA", "PI"]},
    "DIV 04": {"GO": ["GO 02"], "UFs": ["AP", "PA"]},
    "DIV 05": {"GO": ["GO 02"], "UFs": ["AM", "RO", "TO"]},
    "DIV 06": {"GO": ["GO 02"], "UFs": ["DF", "GO"]},
    "DIV 07": {"GO": ["GO 02"], "UFs": ["MT"]},
    "DIV 10": {"GO": ["GO 03"], "UFs": ["ES"]}
}

# Lista das colunas desejadas no resultado final
COLUNAS_IMPORTANTES = [
    "Divisão", "Gerência Operacional", "UF", "Base", "Origem", "Classificacao",
    "Solicitante", "Data_Criacao", "Data_Chegada", "Data_Previsao_Conclusao",
    "Data_Previsao_Chegada", "Data_Conclusao", "Data_de_Fechamento", "Fila",
    "Grupo", "Numero_Chamado", "Prioridade", "Substatus", "Status", "SubTipo",
    "Tipo", "Valor_Total", "Negocio", "Local_Nome", "Uniorg_Comercial",
    "Tempo_de_Custo", "Data_do_Primeiro_Encaminhamento", "Fornecedor",
    "Nota_Inicial", "originador", "regional", "Responsavel", "prazo_inicio",
    "prazo_conclusao", "rede", "modulo"
]

# Caminhos dos arquivos
FILE_PATH = r"c:\Users\Rodri\Downloads\STD - Avanço\Base Geral STD.xlsx"
OUTPUT_BASE_TRATADA = r"c:\Users\Rodri\Downloads\STD - Avanço\Base_Tratada.xlsx"
OUTPUT_ANALISE_COMPLETA = "Analise_Chamados_Completa.xlsx"

# ======================
# FUNÇÕES AUXILIARES
# ======================

def classificar_prazo_conclusao(row):
    """Classifica o status de prazo de conclusão"""
    if pd.isna(row["Data_Conclusao"]) and pd.notna(row["Data_Previsao_Conclusao"]):
        if pd.Timestamp.today() > row["Data_Previsao_Conclusao"]:
            return "Vencido"
        else:
            return "Em Aberto"
    elif pd.notna(row["Data_Conclusao"]) and pd.notna(row["Data_Previsao_Conclusao"]):
        if row["Data_Conclusao"] <= row["Data_Previsao_Conclusao"]:
            return "NP"
        else:
            return "FP"
    else:
        return "Sem Informação"

def calcular_status_prazo_inicio(row):
    """Calcula status de prazo de início"""
    if pd.isna(row['Data_Previsao_Chegada']):
        return 'Não Definido'
    
    # Usar a data de chegada ou primeiro encaminhamento para verificar início
    data_inicio = row['Data_Chegada']
    if not pd.isna(row['Data_do_Primeiro_Encaminhamento']):
        data_inicio = row['Data_do_Primeiro_Encaminhamento']
    
    if pd.isna(data_inicio):
        return 'Não Definido'
    
    if data_inicio <= row['Data_Previsao_Chegada']:
        return 'NP'
    else:
        return 'FP'

def calcular_status_prazo_conclusao(row):
    """Calcula status de prazo de conclusão"""
    if pd.isna(row['Data_Previsao_Conclusao']) or pd.isna(row['Data_Conclusao']):
        return 'Não Definido'
    
    if row['Data_Conclusao'] <= row['Data_Previsao_Conclusao']:
        return 'NP'
    else:
        return 'FP'

# ======================
# PROCESSAMENTO PRINCIPAL
# ======================

def main():
    # 1. TRATAMENTO INICIAL DA BASE
    print("🔍 Lendo base de dados...")
    df = pd.read_excel(FILE_PATH)
    
    # Verificar e criar colunas faltantes
    for coluna in COLUNAS_IMPORTANTES:
        if coluna not in df.columns:
            df[coluna] = None
    
    # Selecionar apenas as colunas desejadas
    df_filtrado = df[COLUNAS_IMPORTANTES].copy()
    
    # Criar dicionário reverso para mapeamento de UF para Divisão e GO
    uf_para_divisao = {}
    uf_para_go = {}
    
    for divisao, info in DIVISOES.items():
        for uf in info["UFs"]:
            uf_para_divisao[uf] = divisao
            # Para UFs com múltiplas GOs, usar a primeira
            uf_para_go[uf] = info["GO"][0]
    
    # Preencher automaticamente as colunas com base na UF
    if "UF" in df_filtrado.columns:
        df_filtrado["Divisão"] = df_filtrado["UF"].map(uf_para_divisao)
        df_filtrado["Gerência Operacional"] = df_filtrado["UF"].map(uf_para_go)
    
    # Salvar base tratada
    df_filtrado.to_excel(OUTPUT_BASE_TRATADA, index=False, engine="openpyxl")
    print("✅ Base tratada salva em:", OUTPUT_BASE_TRATADA)
    
    # 2. ANÁLISE DOS DADOS
    print("📊 Iniciando análise dos dados...")
    
    # Converter colunas de data para datetime
    date_columns = ['Data_Criacao', 'Data_Chegada', 'Data_Previsao_Conclusao', 
                    'Data_Previsao_Chegada', 'Data_Conclusao', 'Data_de_Fechamento', 
                    'Data_do_Primeiro_Encaminhamento']
    
    for col in date_columns:
        if col in df_filtrado.columns:
            df_filtrado[col] = pd.to_datetime(df_filtrado[col], errors='coerce')
    
    # Aplicar classificação de prazos
    df_filtrado['Classificacao_Prazo'] = df_filtrado.apply(classificar_prazo_conclusao, axis=1)
    df_filtrado['Status_Prazo_Inicio'] = df_filtrado.apply(calcular_status_prazo_inicio, axis=1)
    df_filtrado['Status_Prazo_Conclusao'] = df_filtrado.apply(calcular_status_prazo_conclusao, axis=1)
    
    # Criar coluna de mês/ano para análise temporal
    df_filtrado['Mes_Ano'] = df_filtrado['Data_Criacao'].dt.to_period('M')
    
    # ANÁLISE 1: Estatísticas gerais de prazos
    estatisticas_gerais = pd.DataFrame({
        'Metrica': ['Total_Chamados', 'NP_Inicio', 'FP_Inicio', 'NP_Conclusao', 'FP_Conclusao'],
        'Valor': [
            len(df_filtrado),
            len(df_filtrado[df_filtrado['Status_Prazo_Inicio'] == 'NP']),
            len(df_filtrado[df_filtrado['Status_Prazo_Inicio'] == 'FP']),
            len(df_filtrado[df_filtrado['Status_Prazo_Conclusao'] == 'NP']),
            len(df_filtrado[df_filtrado['Status_Prazo_Conclusao'] == 'FP'])
        ]
    })
    
    # ANÁLISE 2: Evolução mensal de prazos (início e conclusão)
    evolucao_mensal = df_filtrado.groupby('Mes_Ano').agg({
        'Numero_Chamado': 'count',
        'Status_Prazo_Inicio': lambda x: (x == 'NP').sum(),
        'Status_Prazo_Conclusao': lambda x: (x == 'NP').sum()
    }).rename(columns={
        'Numero_Chamado': 'Total_Chamados', 
        'Status_Prazo_Inicio': 'NP_Inicio',
        'Status_Prazo_Conclusao': 'NP_Conclusao'
    })
    
    evolucao_mensal['FP_Inicio'] = evolucao_mensal['Total_Chamados'] - evolucao_mensal['NP_Inicio']
    evolucao_mensal['FP_Conclusao'] = evolucao_mensal['Total_Chamados'] - evolucao_mensal['NP_Conclusao']
    evolucao_mensal['Percentual_NP_Inicio'] = (evolucao_mensal['NP_Inicio'] / evolucao_mensal['Total_Chamados'] * 100).round(2)
    evolucao_mensal['Percentual_NP_Conclusao'] = (evolucao_mensal['NP_Conclusao'] / evolucao_mensal['Total_Chamados'] * 100).round(2)
    evolucao_mensal['Meta'] = 96
    evolucao_mensal['Diferenca_Meta_Inicio'] = evolucao_mensal['Percentual_NP_Inicio'] - evolucao_mensal['Meta']
    evolucao_mensal['Diferenca_Meta_Conclusao'] = evolucao_mensal['Percentual_NP_Conclusao'] - evolucao_mensal['Meta']
    
    # ANÁLISE 3: Por Divisão
    analise_divisao = df_filtrado.groupby('Divisão').agg({
        'Numero_Chamado': 'count',
        'Status_Prazo_Inicio': lambda x: (x == 'NP').sum(),
        'Status_Prazo_Conclusao': lambda x: (x == 'NP').sum()
    }).rename(columns={
        'Numero_Chamado': 'Total_Chamados', 
        'Status_Prazo_Inicio': 'NP_Inicio',
        'Status_Prazo_Conclusao': 'NP_Conclusao'
    })
    
    analise_divisao['FP_Inicio'] = analise_divisao['Total_Chamados'] - analise_divisao['NP_Inicio']
    analise_divisao['FP_Conclusao'] = analise_divisao['Total_Chamados'] - analise_divisao['NP_Conclusao']
    analise_divisao['Percentual_NP_Inicio'] = (analise_divisao['NP_Inicio'] / analise_divisao['Total_Chamados'] * 100).round(2)
    analise_divisao['Percentual_NP_Conclusao'] = (analise_divisao['NP_Conclusao'] / analise_divisao['Total_Chamados'] * 100).round(2)
    analise_divisao = analise_divisao.sort_values('Percentual_NP_Conclusao', ascending=False)
    
    # ANÁLISE 4: Por Regional
    analise_regional = df_filtrado.groupby('regional').agg({
        'Numero_Chamado': 'count',
        'Status_Prazo_Inicio': lambda x: (x == 'NP').sum(),
        'Status_Prazo_Conclusao': lambda x: (x == 'NP').sum()
    }).rename(columns={
        'Numero_Chamado': 'Total_Chamados', 
        'Status_Prazo_Inicio': 'NP_Inicio',
        'Status_Prazo_Conclusao': 'NP_Conclusao'
    })
    
    analise_regional['FP_Inicio'] = analise_regional['Total_Chamados'] - analise_regional['NP_Inicio']
    analise_regional['FP_Conclusao'] = analise_regional['Total_Chamados'] - analise_regional['NP_Conclusao']
    analise_regional['Percentual_NP_Inicio'] = (analise_regional['NP_Inicio'] / analise_regional['Total_Chamados'] * 100).round(2)
    analise_regional['Percentual_NP_Conclusao'] = (analise_regional['NP_Conclusao'] / analise_regional['Total_Chamados'] * 100).round(2)
    analise_regional = analise_regional.sort_values('Percentual_NP_Conclusao', ascending=False)
    
    # ANÁLISE 5: Por Tipo de Chamado
    analise_tipo = df_filtrado.groupby('Tipo').agg({
        'Numero_Chamado': 'count',
        'Status_Prazo_Inicio': lambda x: (x == 'NP').sum(),
        'Status_Prazo_Conclusao': lambda x: (x == 'NP').sum()
    }).rename(columns={
        'Numero_Chamado': 'Total_Chamados', 
        'Status_Prazo_Inicio': 'NP_Inicio',
        'Status_Prazo_Conclusao': 'NP_Conclusao'
    })
    
    analise_tipo['FP_Inicio'] = analise_tipo['Total_Chamados'] - analise_tipo['NP_Inicio']
    analise_tipo['FP_Conclusao'] = analise_tipo['Total_Chamados'] - analise_tipo['NP_Conclusao']
    analise_tipo['Percentual_NP_Inicio'] = (analise_tipo['NP_Inicio'] / analise_tipo['Total_Chamados'] * 100).round(2)
    analise_tipo['Percentual_NP_Conclusao'] = (analise_tipo['NP_Conclusao'] / analise_tipo['Total_Chamados'] * 100).round(2)
    analise_tipo = analise_tipo.sort_values('Percentual_NP_Conclusao', ascending=False)
    
    # ANÁLISE 6: Tempo médio de resolução
    df_filtrado['Tempo_Resolucao'] = (df_filtrado['Data_Conclusao'] - df_filtrado['Data_Criacao']).dt.total_seconds() / 3600  # em horas
    tempo_resolucao = df_filtrado.groupby('Status_Prazo_Conclusao')['Tempo_Resolucao'].agg(['mean', 'median', 'std']).round(2)
    
    # ANÁLISE 7: Top 10 responsáveis com mais chamados
    top_responsaveis = df_filtrado['Responsavel'].value_counts().head(10).reset_index()
    top_responsaveis.columns = ['Responsavel', 'Total_Chamados']
    
    # ANÁLISE 8: Chamados por prioridade
    analise_prioridade = df_filtrado.groupby('Prioridade').agg({
        'Numero_Chamado': 'count',
        'Status_Prazo_Inicio': lambda x: (x == 'NP').sum(),
        'Status_Prazo_Conclusao': lambda x: (x == 'NP').sum()
    }).rename(columns={
        'Numero_Chamado': 'Total_Chamados', 
        'Status_Prazo_Inicio': 'NP_Inicio',
        'Status_Prazo_Conclusao': 'NP_Conclusao'
    })
    
    analise_prioridade['FP_Inicio'] = analise_prioridade['Total_Chamados'] - analise_prioridade['NP_Inicio']
    analise_prioridade['FP_Conclusao'] = analise_prioridade['Total_Chamados'] - analise_prioridade['NP_Conclusao']
    analise_prioridade['Percentual_NP_Inicio'] = (analise_prioridade['NP_Inicio'] / analise_prioridade['Total_Chamados'] * 100).round(2)
    analise_prioridade['Percentual_NP_Conclusao'] = (analise_prioridade['NP_Conclusao'] / analise_prioridade['Total_Chamados'] * 100).round(2)
    analise_prioridade = analise_prioridade.sort_values('Percentual_NP_Conclusao', ascending=False)
    
    # ANÁLISE 9: Detalhamento dos chamados FP (fora do prazo)
    chamados_fp_inicio = df_filtrado[df_filtrado['Status_Prazo_Inicio'] == 'FP'].copy()
    chamados_fp_inicio['Dias_Atraso_Inicio'] = (chamados_fp_inicio['Data_Chegada'] - chamados_fp_inicio['Data_Previsao_Chegada']).dt.days
    
    chamados_fp_conclusao = df_filtrado[df_filtrado['Status_Prazo_Conclusao'] == 'FP'].copy()
    chamados_fp_conclusao['Dias_Atraso_Conclusao'] = (chamados_fp_conclusao['Data_Conclusao'] - chamados_fp_conclusao['Data_Previsao_Conclusao']).dt.days
    
    # 3. EXPORTAÇÃO DOS RESULTADOS
    print("💾 Salvando resultados da análise...")
    
    # Criar nova planilha com as análises
    with pd.ExcelWriter(OUTPUT_ANALISE_COMPLETA, engine='openpyxl') as writer:
        # Aba com dados originais
        df_filtrado.to_excel(writer, sheet_name='Dados_Originais', index=False)
        
        # Abas com análises
        estatisticas_gerais.to_excel(writer, sheet_name='Estatisticas_Gerais', index=False)
        evolucao_mensal.to_excel(writer, sheet_name='Evolucao_Mensal')
        analise_divisao.to_excel(writer, sheet_name='Por_Divisao')
        analise_regional.to_excel(writer, sheet_name='Por_Regional')
        analise_tipo.to_excel(writer, sheet_name='Por_Tipo')
        tempo_resolucao.to_excel(writer, sheet_name='Tempo_Resolucao')
        top_responsaveis.to_excel(writer, sheet_name='Top_Responsaveis', index=False)
        analise_prioridade.to_excel(writer, sheet_name='Por_Prioridade')
        chamados_fp_inicio.to_excel(writer, sheet_name='Chamados_FP_Inicio', index=False)
        chamados_fp_conclusao.to_excel(writer, sheet_name='Chamados_FP_Conclusao', index=False)

        # Formatação condicional para destacar valores abaixo da meta
        workbook = writer.book
        
        # Formatar Evolução Mensal
        worksheet = workbook['Evolucao_Mensal']
        red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
        green_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
        
        for row in range(2, worksheet.max_row + 1):
            # Destacar células onde Percentual_NP_Inicio < 96
            if worksheet.cell(row=row, column=7).value < 96:  # Coluna G é Percentual_NP_Inicio
                worksheet.cell(row=row, column=7).fill = red_fill
            else:
                worksheet.cell(row=row, column=7).fill = green_fill
                
            # Destacar células onde Percentual_NP_Conclusao < 96
            if worksheet.cell(row=row, column=8).value < 96:  # Coluna H é Percentual_NP_Conclusao
                worksheet.cell(row=row, column=8).fill = red_fill
            else:
                worksheet.cell(row=row, column=8).fill = green_fill

        # Adicionar resumo executivo
        resumo_executivo = pd.DataFrame({
            'Metrica': [
                'Total de Chamados',
                'Chamados NP Início',
                'Chamados FP Início',
                '% NP Início',
                'Chamados NP Conclusão',
                'Chamados FP Conclusão',
                '% NP Conclusão'
            ],
            'Valor': [
                len(df_filtrado),
                len(df_filtrado[df_filtrado['Status_Prazo_Inicio'] == 'NP']),
                len(df_filtrado[df_filtrado['Status_Prazo_Inicio'] == 'FP']),
                f"{(len(df_filtrado[df_filtrado['Status_Prazo_Inicio'] == 'NP']) / len(df_filtrado) * 100):.2f}%",
                len(df_filtrado[df_filtrado['Status_Prazo_Conclusao'] == 'NP']),
                len(df_filtrado[df_filtrado['Status_Prazo_Conclusao'] == 'FP']),
                f"{(len(df_filtrado[df_filtrado['Status_Prazo_Conclusao'] == 'NP']) / len(df_filtrado) * 100):.2f}%"
            ]
        })
        
        resumo_executivo.to_excel(writer, sheet_name='Resumo_Executivo', index=False)
        
        # Formatar Resumo Executivo
        worksheet = workbook['Resumo_Executivo']
        for row in range(2, worksheet.max_row + 1):
            if row in [5, 8]:  # Linhas com percentuais
                if float(worksheet.cell(row=row, column=2).value.replace('%', '')) < 96:
                    worksheet.cell(row=row, column=2).fill = red_fill
                else:
                    worksheet.cell(row=row, column=2).fill = green_fill

    print("✅ Análise concluída! Arquivo salvo em:", OUTPUT_ANALISE_COMPLETA)
    print(f"📊 Total de chamados analisados: {len(df_filtrado)}")
    print(f"📈 Chamados NP Início: {len(df_filtrado[df_filtrado['Status_Prazo_Inicio'] == 'NP'])} ({(len(df_filtrado[df_filtrado['Status_Prazo_Inicio'] == 'NP']) / len(df_filtrado) * 100):.2f}%)")
    print(f"📉 Chamados FP Início: {len(df_filtrado[df_filtrado['Status_Prazo_Inicio'] == 'FP'])} ({(len(df_filtrado[df_filtrado['Status_Prazo_Inicio'] == 'FP']) / len(df_filtrado) * 100):.2f}%)")
    print(f"📈 Chamados NP Conclusão: {len(df_filtrado[df_filtrado['Status_Prazo_Conclusao'] == 'NP'])} ({(len(df_filtrado[df_filtrado['Status_Prazo_Conclusao'] == 'NP']) / len(df_filtrado) * 100):.2f}%)")
    print(f"📉 Chamados FP Conclusão: {len(df_filtrado[df_filtrado['Status_Prazo_Conclusao'] == 'FP'])} ({(len(df_filtrado[df_filtrado['Status_Prazo_Conclusao'] == 'FP']) / len(df_filtrado) * 100):.2f}%)")

if __name__ == "__main__":
    main()
# 📊✨ Análise STD - Processamento de Dados

Script completo para tratamento, análise e geração de métricas de desempenho de chamados, replicando em Python todas as medidas DAX utilizadas no Power BI.

---

## 🚀 Visão Geral

Este projeto automatiza o processamento de dados operacionais de chamados (OS), aplicando regras de negócio, cálculos de SLA e indicadores de performance, entregando uma base tratada pronta para análise estratégica.

Desenvolvido para rodar no **Google Colab**, com upload de arquivo direto na interface e download automático dos resultados ao final do processamento.

---

## ☁️ Ambiente de Execução

- Plataforma: **Google Colab**
- O script solicita o upload do arquivo CSV na inicialização
- Os resultados são disponibilizados para download automaticamente ao final
- Caminhos de trabalho: `/content/uploads` (entrada) e `/content/output` (saída)

---

## 📂 Processamento de Dados

- **Carregamento Inteligente:** Importação automática de arquivos **CSV** com tentativas automáticas de encoding (`utf-8`, `latin1`, `cp1252`) e separador (`,` `;` `\t` `|`)
- **Mapeamento Automático:** Associação de UFs às Divisões e Gerências Operacionais
- **Conversão de Datas:** Padronização de todas as colunas de data para formato `datetime`
- **Criação de Calendário:** Geração automática de calendário completo com base nas datas dos chamados

---

## 📈 Métricas de SLA (Equivalentes DAX)

- **SLA Início e Término:** Cálculo de percentuais de chamados dentro do prazo
- **Comparação com Metas:** Diferença em pontos percentuais versus metas configuráveis (`META_SLA = 96%` e `META_LIMPEZA = 98%`)
- **Prazo Ajustado:** Conversão inteligente de valores `"NA"` para `"NP"`

---

## ⚡ Indicadores de Performance

- **Tempo de Atendimento:** Dias entre criação, conclusão e fechamento
- **Dias de Atraso:** Cálculo preciso de dias em atraso
- **Duração do Chamado:** Tempo total desde a criação até a data atual
- **Faixas de Tempo:** Classificação automática:
  - -30 dias *(padrão)*
  - +30 dias
  - +60 dias
  - +90 dias

---

## 🔍 Análises de Status

- **Status Completo:** Monitoramento de status do chamado, fechamento e financeiro
- **Estoque Atual:** Identificação de chamados pendentes
- **Fechamento Pendente:** Alertas para chamados concluídos há mais de 30 dias sem fechamento
- **À Vencer WTM:** Alertas para chamados entre 20 e 29 dias de duração

---

## 💰 Indicadores Financeiros

- **Valor Total de OS:** Soma total das ordens de serviço
- **Média de Valor OS:** Ticket médio das ordens
- **Quantidade de Agências:** Contagem de Uniorgs Comerciais

---

## 👥 Análises Dimensionais

Métricas segmentadas por seis dimensões:

- Divisão
- Regional
- Tipo
- Prioridade
- Fornecedor
- UF

Além das visões:

- **Top Responsáveis:** Ranking por volume de chamados atribuídos
- **Evolução Mensal:** Análise temporal de performance
- **Métricas Acumuladas:** Visão histórica consolidada

---

## 📤 Exportação Avançada

- **Base Tratada Completa:** Dataset final totalmente processado (`Base_Tratada.xlsx`)
- **Planilha Analítica Multi-aba:** Organização estruturada das análises (`Analise_Chamados_Completa.xlsx`)
- **Formatação Condicional:** Destaques visuais automáticos por faixas de SLA e status de atraso
- **Tabela de Medidas DAX Equivalentes:** Comparativo técnico entre DAX e Python

---

## 🛠 Tecnologias Utilizadas

- **Python 🐍** – Linguagem principal
- **Pandas 📊** – Manipulação e análise de dados
- **NumPy 🔢** – Operações vetorizadas e cálculos de dias úteis
- **OpenPyXL 📁** – Integração e exportação para Excel

---

## ⚙️ Configurações Principais (`Config`)

| Parâmetro | Valor padrão | Descrição |
|---|---|---|
| `CSV_FILENAME` | `PreventivasFornecedor.csv` | Nome do arquivo de entrada |
| `META_SLA` | `96.0` | Meta de SLA (%) |
| `META_LIMPEZA` | `98.0` | Meta de limpeza/SLA estendido (%) |
| `DIAS_FECHAMENTO_PENDENTE` | `30` | Dias até alertar fechamento pendente |
| `TOP_RESPONSABLES` | `15` | Quantidade de responsáveis no ranking |

---

## 🎯 Objetivo Estratégico

Reduzir dependência exclusiva do Power BI para cálculos complexos, garantindo:

- Reprodutibilidade
- Automação de métricas
- Padronização de regras de negócio
- Escalabilidade analítica

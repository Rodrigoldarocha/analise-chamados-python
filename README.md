# 📊✨ Análise STD - Processamento de Dados

Script completo para tratamento, análise e geração de métricas de desempenho de chamados, replicando em Python todas as medidas DAX utilizadas no Power BI.

---

## 🚀 Visão Geral

Este projeto automatiza o processamento de dados operacionais de chamados (OS), aplicando regras de negócio, cálculos de SLA e indicadores de performance, entregando uma base tratada pronta para análise estratégica.

---

## 📂 Processamento de Dados

- **Carregamento Inteligente:** Importação automática de arquivos Excel com tratamento robusto de erros  
- **Mapeamento Automático:** Associação de UFs às Divisões e Gerências Operacionais  
- **Conversão de Datas:** Padronização de todas as colunas de data para formato `datetime`  
- **Criação de Calendário:** Geração automática de calendário completo com base nas datas dos chamados  

---

## 📈 Métricas de SLA (Equivalentes DAX)

- **SLA Início e Término:** Cálculo de percentuais de chamados dentro do prazo  
- **Comparação com Metas:** Diferença em pontos percentuais versus metas definidas (96% e 98%)  
- **Prazo Ajustado:** Conversão inteligente de valores "NA" para "NP"  

---

## ⚡ Indicadores de Performance

- **Tempo de Atendimento:** Dias entre criação, conclusão e fechamento  
- **Dias de Atraso:** Cálculo preciso de dias em atraso  
- **Duração do Chamado:** Tempo total desde a criação até a data atual  
- **Faixas de Tempo:** Classificação automática:
  - +30 dias  
  - +60 dias  
  - +90 dias  

---

## 🔍 Análises de Status

- **Status Completo:** Monitoramento de status do chamado, fechamento e financeiro  
- **Estoque Atual:** Identificação de chamados pendentes  
- **Fechamento Pendente:** Alertas para chamados concluídos há mais de 30 dias sem fechamento  
- **À Vencer WTM:** Alertas para chamados entre 20 e 29 dias  

---

## 💰 Indicadores Financeiros

- **Valor Total de OS:** Soma total das ordens de serviço  
- **Média de Valor OS:** Ticket médio das ordens  
- **Quantidade de Agências:** Contagem de Uniorgs Comerciais  

---

## 👥 Análises Dimensionais

- Métricas segmentadas por:
  - Divisão  
  - Tipo  
  - Prioridade  

- **Top Responsáveis:** Ranking de desempenho  
- **Evolução Mensal:** Análise temporal de performance  
- **Métricas Acumuladas:** Visão histórica consolidada  

---

## 📤 Exportação Avançada

- **Base Tratada Completa:** Dataset final totalmente processado  
- **Planilha Analítica Multi-aba:** Organização estruturada das análises  
- **Formatação Condicional:** Destaques visuais automáticos  
- **Tabela de Medidas DAX Equivalentes:** Comparativo técnico entre DAX e Python  

---

## 🛠 Tecnologias Utilizadas

- **Python 🐍** – Linguagem principal  
- **Pandas 📊** – Manipulação e análise de dados  
- **OpenPyXL 📁** – Integração e exportação para Excel  

---

## 🎯 Objetivo Estratégico

Reduzir dependência exclusiva do Power BI para cálculos complexos, garantindo:

- Reprodutibilidade  
- Automação de métricas  
- Padronização de regras de negócio  
- Escalabilidade analítica  

---

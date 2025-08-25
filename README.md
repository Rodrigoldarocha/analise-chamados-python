📊✨ Análise STD - Processamento de Dados

📝 Script completo para tratamento, análise e geração de métricas de desempenho de chamados, replicando todas as medidas DAX do Power BI em Python.

🚀 Funcionalidades Principais
📂 Processamento de Dados
Carregamento Inteligente: Importa automaticamente dados de arquivos Excel com tratamento robusto de erros

Mapeamento Automático: Associa UFs às Divisões e Gerências Operacionais correspondentes

Conversão de Datas: Padroniza todas as colunas de data para formato datetime

Criação de Calendário: Gera calendário completo baseado nas datas dos chamados

📈 Métricas de SLA (Equivalentes DAX)
SLA Início/Término: Calcula percentuais de chamados dentro do prazo

Comparações com Meta: Diferença em pontos percentuais vs metas (96% e 98%)

Prazo Ajustado: Conversão inteligente de "NA" para "NP" nos prazos

⚡ Indicadores de Performance
Tempos de Atendimento: Dias entre criação, conclusão e fechamento

Dias de Atraso: Calcula precisamente dias em atraso

Duração do Chamado: Tempo total desde criação até momento atual

Faixas de Tempo: Classificação em "+90 dias", "+60 dias", "+30 dias"

🔍 Análises de Status
Status Completo: Chamado, Fechamento e Financeiro

Estoque Atual: Identifica chamados pendentes

Fechamento Pendente: Alertas para concluídos há +30 dias sem fechamento

À Vencer WTM: Alertas para chamados entre 20-29 dias

💰 Indicadores Financeiros
Valor Total de OS: Soma do valor das ordens de serviço

Média Valor OS: Valor médio das OS

Quantidade de Agências: Contagem de Uniorgs Comerciais

👥 Análises Dimensionais
Por Divisão, Tipo, Prioridade: Métricas segmentadas

Top Responsáveis: Ranking dos responsáveis

Evolução Mensal: Performance temporal

Métricas Acumuladas: Visão histórica completa

📤 Exportação Avançada
Base Tratada Completa: Todos dados processados

Planilha Analítica Multi-aba: Análises organizadas

Formatação Condicional: Destaques visuais automáticos

Medidas DAX Equivalentes: Tabela comparativa completa

🛠 Tecnologias Utilizadas
Python 🐍 – linguagem principal

Pandas 📊 – manipulação e análise de dados

OpenPyXL 📁 – integração com Excel
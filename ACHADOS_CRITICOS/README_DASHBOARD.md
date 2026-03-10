# 🏥 Dashboard de Achados Críticos - CDI

## 📋 Visão Geral

Dashboard moderno e interativo para análise de achados críticos em Centro de Diagnóstico por Imagem (CDI). Sistema completo com interface dark mode, análises em tempo real e relatórios automatizados.

## ✨ Características Principais

### 🎨 Interface Moderna
- **Dark Mode** com design glassmorphism
- **Responsiva** e otimizada para diferentes telas
- **Visualizações interativas** com Plotly
- **UX intuitiva** com Streamlit

### 📊 Funcionalidades Analíticas
- **Correlação automática** de planilhas usando múltiplos critérios
- **Cálculo preciso** de tempos de comunicação
- **Métricas de compliance** em tempo real
- **Rankings** de médicos e achados críticos
- **Distribuição temporal** dos comunicados

### 📈 Visualizações Incluídas
1. **Gráfico de Compliance** - Taxa de comunicação dentro do prazo
2. **Performance dos Médicos** - Top comunicadores e suas performances
3. **Achados Críticos** - Frequência e tempos por tipo de achado
4. **Distribuição Temporal** - Análise por faixas de tempo
5. **Tabela Detalhada** - Dados completos para análise

### 📥 Export de Relatórios
- **Excel multisheets** com dados correlacionados
- **Estatísticas por médico** organizadas
- **Estatísticas por achados** detalhadas
- **Download instantâneo** de relatórios

## 🚀 Como Usar

### 1. Instalação das Dependências
```bash
pip install -r requirements_dashboard.txt
```

### 2. Inicializar o Dashboard
```bash
# Opção 1: Script automático
python run_dashboard.py

# Opção 2: Comando direto
streamlit run dashboard_achados_criticos.py
```

### 2.1 Deploy no Render (mesma UI/UX)
Este projeto já está pronto para deploy com os arquivos:
- `render.yaml`
- `start_render.sh`
- `.streamlit/config.toml`

Passos:
```bash
# 1) Subir este projeto no GitHub
# 2) No Render, criar "Blueprint" apontando para o repositório
# 3) O Render detectará o render.yaml automaticamente
```

Configuração aplicada:
- Streamlit com tema dark idêntico ao ambiente local
- Porta dinâmica via variável `PORT` do Render
- Compatível com upload de planilhas grandes (até 400 MB)

### 3. Usar o Dashboard
1. **Acesse:** http://localhost:8501
2. **Upload:** Carregue as duas planilhas na sidebar:
   - Planilha de Achados Críticos (.xlsx/.xls)
   - Planilha de Status dos Exames (.xlsx/.xls)
3. **Processar:** Clique em "🔄 Processar Planilhas"
4. **Analisar:** Visualize os gráficos e métricas gerados
5. **Exportar:** Baixe relatórios em Excel

## 📁 Estrutura de Arquivos

```
📦 Sistema Completo
├── 🎛️ dashboard_achados_criticos.py    # Dashboard principal
├── 🚀 run_dashboard.py                 # Script de inicialização
├── 📊 analisador_achados_criticos.py   # Motor de análise (CLI)
├── 📈 visualizador_achados_criticos.py # Gerador de relatórios estáticos
├── 🔧 debug_linha3.py                  # Ferramenta de debug
├── 📋 requirements_dashboard.txt       # Dependências do dashboard
└── 📄 README_DASHBOARD.md             # Documentação completa
```

## 🎯 Métricas Disponíveis

### 📊 Visão Geral
- **Taxa de Compliance** - Percentual de comunicados no prazo
- **Total no Prazo** - Quantidade de achados comunicados ≤ 1h
- **Total Fora do Prazo** - Quantidade de achados comunicados > 1h
- **Tempo Mediano** - Tempo mediano de comunicação

### 👨‍⚕️ Por Médicos
- Total de comunicados por médico
- Percentual de comunicados fora do prazo
- Tempo médio de comunicação
- Ranking de performance

### 🔍 Por Achados Críticos
- Frequência de cada tipo de achado
- Taxa de atraso por tipo de achado
- Tempo médio por categoria
- Identificação de padrões problemáticos

## 💡 Tecnologias Utilizadas

- **Frontend:** Streamlit + CSS personalizado
- **Visualizações:** Plotly.js (gráficos interativos)
- **Processamento:** Pandas + NumPy
- **Export:** OpenPyXL para Excel
- **Design:** Dark theme + Glassmorphism

## 🔧 Configurações Avançadas

### Personalizar Tema
```python
# Editar cores no dashboard_achados_criticos.py
st.set_page_config(
    page_title="Seu Título",
    page_icon="🏥",
    layout="wide"
)
```

### Modificar Métricas
```python
# Ajustar limite de prazo (padrão: 1 hora)
df_com_status['fora_do_prazo'] = df_com_status['tempo_comunicacao_horas'] > 1
```

## 📝 Formato das Planilhas

### Planilha de Achados Críticos
**Colunas obrigatórias:**
- `SAME` - ID do paciente
- `Nome_Paciente` - Nome completo
- `Data_Exame` - Data do exame
- `Descrição_Procedimento` - Tipo de exame
- `Data_Sinalização` - Quando foi comunicado
- `Medico Laudo` - Médico responsável
- `Achado Crítico` - Tipo do achado

### Planilha de Status
**Colunas obrigatórias:**
- `SAME` - ID do paciente
- `NOME_PACIENTE` - Nome completo
- `DATA_HORA_PRESCRICAO` - Data do exame
- `DESCRICAO_PROCEDIMENTO` - Tipo de exame
- `STATUS_ALAUDAR` - Quando entrou status "a laudar"

## 🎭 Preview das Telas

### 🏠 Tela Principal
- Header com gradiente azul-roxo
- Cards com métricas principais em glassmorphism
- Sidebar com controles de upload

### 📊 Gráficos Interativos
- **Donut Chart** para compliance geral
- **Barras Horizontais** para ranking de médicos
- **Heatmap** para achados críticos
- **Distribuição** por faixas temporais

### 📋 Tabela de Dados
- Expansível com todos os registros
- Ordenação por tempo de comunicação
- Filtros e busca integrados

## 🚨 Troubleshooting

### Erro ao Carregar Planilhas
```
❌ Formato não suportado ou planilha corrompida
✅ Verificar: .xlsx ou .xls, estrutura de colunas
```

### Erro na Correlação
```
❌ Nenhum registro correlacionado
✅ Verificar: correspondência SAME, nomes, datas
```

### Dashboard Não Carrega
```
❌ Porta 8501 ocupada
✅ Usar: streamlit run dashboard.py --server.port 8502
```

## 📞 Suporte

Sistema desenvolvido especificamente para análise de achados críticos em CDI.

**Recursos inclusos:**
- 🔄 Processamento automático de dados
- 📊 Visualizações em tempo real
- 📁 Export de relatórios completos
- 🎨 Interface moderna e intuitiva

---

**🏥 Sistema de Análise de Achados Críticos - CDI**
*Desenvolvido com ❤️ para otimizar a comunicação médica*

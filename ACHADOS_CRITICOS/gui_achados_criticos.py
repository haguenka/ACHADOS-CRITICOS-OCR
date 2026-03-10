#!/usr/bin/env python3
"""
Interface GUI Moderna para Análise de Achados Críticos
Sistema completo com dark theme usando Tkinter
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from ttkthemes import ThemedTk
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure
import seaborn as sns
from datetime import datetime
import threading
from pathlib import Path
import io
import unicodedata

# Configurar estilo matplotlib para dark theme
plt.style.use('dark_background')
sns.set_palette("husl")

class ModernGUI:
    def __init__(self):
        """Inicializa a interface GUI moderna"""
        self.root = ThemedTk(theme="equilux")  # Dark theme
        self.setup_window()
        self.setup_variables()
        self.setup_styles()
        self.create_widgets()

        # Dados
        self.df_achados = None
        self.df_status = None
        self.df_correlacionado = None

    def setup_window(self):
        """Configura a janela principal"""
        self.root.title("🏥 CDI - Análise de Achados Críticos")
        self.root.geometry("1400x900")
        self.root.minsize(1200, 700)

        # Ícone personalizado
        try:
            self.root.iconbitmap("hospital.ico")
        except:
            pass

        # Centralizar janela
        self.root.update_idletasks()
        x = (self.root.winfo_screenwidth() - self.root.winfo_reqwidth()) // 2
        y = (self.root.winfo_screenheight() - self.root.winfo_reqheight()) // 2
        self.root.geometry(f"+{x}+{y}")

    def setup_variables(self):
        """Configura as variáveis do tkinter"""
        self.progress_var = tk.DoubleVar()
        self.status_var = tk.StringVar(value="Pronto para análise")
        self.achados_file = tk.StringVar()
        self.status_file = tk.StringVar()

        # Variáveis para filtro por mês
        self.filter_enabled = tk.BooleanVar(value=False)
        self.selected_month = tk.StringVar(value="Todos os Meses")
        self.selected_year = tk.StringVar(value="2025")

    def setup_styles(self):
        """Configura estilos personalizados"""
        style = ttk.Style()

        # Cores do tema dark
        self.colors = {
            'bg_primary': '#2C3E50',
            'bg_secondary': '#34495E',
            'accent': '#3498DB',
            'success': '#27AE60',
            'warning': '#F39C12',
            'danger': '#E74C3C',
            'text': '#FFFFFF',
            'text_secondary': '#BDC3C7'
        }

        # Configurar estilos customizados
        style.configure('Header.TLabel',
                       font=('Helvetica', 18, 'bold'),
                       foreground=self.colors['text'])

        style.configure('Subheader.TLabel',
                       font=('Helvetica', 12, 'bold'),
                       foreground=self.colors['accent'])

        style.configure('Info.TLabel',
                       font=('Helvetica', 10),
                       foreground=self.colors['text_secondary'])

        style.configure('Success.TButton',
                       font=('Helvetica', 10, 'bold'))

        style.configure('Primary.TButton',
                       font=('Helvetica', 10, 'bold'))

    def create_widgets(self):
        """Cria todos os widgets da interface"""
        # Frame principal
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Configurar grid
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(1, weight=1)

        # Header
        self.create_header(main_frame)

        # Conteúdo principal
        self.create_main_content(main_frame)

        # Sidebar
        self.create_sidebar(main_frame)

        # Footer
        self.create_footer(main_frame)

    def create_header(self, parent):
        """Cria o cabeçalho"""
        header_frame = ttk.Frame(parent, padding="10")
        header_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))

        # Título principal
        title = ttk.Label(header_frame, text="🏥 CDI - Dashboard de Achados Críticos", style='Header.TLabel')
        title.pack(anchor=tk.W)

        # Subtítulo
        subtitle = ttk.Label(header_frame, text="Sistema Avançado de Análise e Monitoramento", style='Info.TLabel')
        subtitle.pack(anchor=tk.W, pady=(5, 0))

        # Separador
        separator = ttk.Separator(header_frame, orient='horizontal')
        separator.pack(fill=tk.X, pady=(10, 0))

    def create_sidebar(self, parent):
        """Cria a sidebar com controles"""
        sidebar_frame = ttk.LabelFrame(parent, text="⚙️ Controles", padding="15")
        sidebar_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(0, 10))
        sidebar_frame.configure(width=300)

        # Upload de arquivos
        self.create_file_upload_section(sidebar_frame)

        # Filtro por mês
        self.create_month_filter_section(sidebar_frame)

        # Controles de processamento
        self.create_processing_section(sidebar_frame)

        # Métricas resumidas
        self.create_metrics_section(sidebar_frame)

    def create_file_upload_section(self, parent):
        """Seção de upload de arquivos"""
        # Título da seção
        upload_label = ttk.Label(parent, text="📁 Upload de Planilhas", style='Subheader.TLabel')
        upload_label.pack(anchor=tk.W, pady=(0, 10))

        # Achados críticos
        achados_frame = ttk.Frame(parent)
        achados_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Label(achados_frame, text="Achados Críticos:", style='Info.TLabel').pack(anchor=tk.W)

        entry_frame1 = ttk.Frame(achados_frame)
        entry_frame1.pack(fill=tk.X, pady=(5, 0))

        self.achados_entry = ttk.Entry(entry_frame1, textvariable=self.achados_file, state='readonly')
        self.achados_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))

        achados_btn = ttk.Button(entry_frame1, text="📂", width=3,
                                command=lambda: self.select_file('achados'))
        achados_btn.pack(side=tk.RIGHT)

        # Status dos exames
        status_frame = ttk.Frame(parent)
        status_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Label(status_frame, text="Status dos Exames:", style='Info.TLabel').pack(anchor=tk.W)

        entry_frame2 = ttk.Frame(status_frame)
        entry_frame2.pack(fill=tk.X, pady=(5, 0))

        self.status_entry = ttk.Entry(entry_frame2, textvariable=self.status_file, state='readonly')
        self.status_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))

        status_btn = ttk.Button(entry_frame2, text="📂", width=3,
                               command=lambda: self.select_file('status'))
        status_btn.pack(side=tk.RIGHT)

    def create_month_filter_section(self, parent):
        """Seção de filtro por mês"""
        # Separador
        ttk.Separator(parent, orient='horizontal').pack(fill=tk.X, pady=(10, 15))

        # Título da seção
        filter_label = ttk.Label(parent, text="📅 Filtro por Período", style='Subheader.TLabel')
        filter_label.pack(anchor=tk.W, pady=(0, 10))

        # Checkbox para habilitar filtro
        filter_check = ttk.Checkbutton(parent, text="Filtrar por mês específico",
                                      variable=self.filter_enabled,
                                      command=self.on_filter_toggle)
        filter_check.pack(anchor=tk.W, pady=(0, 10))

        # Frame para seletores (inicialmente desabilitado)
        self.filter_frame = ttk.Frame(parent)
        self.filter_frame.pack(fill=tk.X, pady=(0, 10))

        # Seletor de ano
        year_frame = ttk.Frame(self.filter_frame)
        year_frame.pack(fill=tk.X, pady=(0, 5))

        ttk.Label(year_frame, text="Ano:", style='Info.TLabel').pack(side=tk.LEFT)
        self.year_combo = ttk.Combobox(year_frame, textvariable=self.selected_year,
                                      values=["2024", "2025", "2026"],
                                      state='readonly', width=10)
        self.year_combo.pack(side=tk.RIGHT)

        # Seletor de mês
        month_frame = ttk.Frame(self.filter_frame)
        month_frame.pack(fill=tk.X, pady=(0, 5))

        ttk.Label(month_frame, text="Mês:", style='Info.TLabel').pack(side=tk.LEFT)
        self.month_combo = ttk.Combobox(month_frame, textvariable=self.selected_month,
                                       values=[
                                           "Todos os Meses",
                                           "Janeiro", "Fevereiro", "Março", "Abril",
                                           "Maio", "Junho", "Julho", "Agosto",
                                           "Setembro", "Outubro", "Novembro", "Dezembro"
                                       ], state='readonly', width=15)
        self.month_combo.pack(side=tk.RIGHT)

        # Inicialmente desabilitado
        self.on_filter_toggle()

    def on_filter_toggle(self):
        """Habilita/desabilita os controles de filtro"""
        state = 'normal' if self.filter_enabled.get() else 'disabled'
        self.year_combo.configure(state='readonly' if self.filter_enabled.get() else 'disabled')
        self.month_combo.configure(state='readonly' if self.filter_enabled.get() else 'disabled')

        # Se desabilitado, resetar para "Todos os Meses"
        if not self.filter_enabled.get():
            self.selected_month.set("Todos os Meses")

    def create_processing_section(self, parent):
        """Seção de controles de processamento"""
        # Separador
        ttk.Separator(parent, orient='horizontal').pack(fill=tk.X, pady=(10, 15))

        # Título
        process_label = ttk.Label(parent, text="🚀 Processamento", style='Subheader.TLabel')
        process_label.pack(anchor=tk.W, pady=(0, 10))

        # Botão processar
        self.process_btn = ttk.Button(parent, text="🔄 Processar Dados", style='Primary.TButton',
                                     command=self.process_data)
        self.process_btn.pack(fill=tk.X, pady=(0, 10))

        # Barra de progresso
        self.progress_bar = ttk.Progressbar(parent, variable=self.progress_var, mode='determinate')
        self.progress_bar.pack(fill=tk.X, pady=(0, 5))

        # Status
        self.status_label = ttk.Label(parent, textvariable=self.status_var, style='Info.TLabel')
        self.status_label.pack(anchor=tk.W)

        # Botões de ação
        buttons_frame = ttk.Frame(parent)
        buttons_frame.pack(fill=tk.X, pady=(15, 0))

        self.export_btn = ttk.Button(buttons_frame, text="📊 Exportar Excel",
                                    command=self.export_excel, state='disabled')
        self.export_btn.pack(fill=tk.X, pady=(0, 5))

        self.clear_btn = ttk.Button(buttons_frame, text="🗑️ Limpar Dados",
                                   command=self.clear_data)
        self.clear_btn.pack(fill=tk.X)

    def create_metrics_section(self, parent):
        """Seção de métricas resumidas"""
        # Separador
        ttk.Separator(parent, orient='horizontal').pack(fill=tk.X, pady=(15, 15))

        # Título
        metrics_label = ttk.Label(parent, text="📊 Métricas Principais", style='Subheader.TLabel')
        metrics_label.pack(anchor=tk.W, pady=(0, 10))

        # Frame para métricas
        self.metrics_frame = ttk.Frame(parent)
        self.metrics_frame.pack(fill=tk.X)

        # Inicialmente vazio
        no_data_label = ttk.Label(self.metrics_frame, text="Carregue os dados para ver as métricas",
                                 style='Info.TLabel')
        no_data_label.pack(anchor=tk.W)

    def create_main_content(self, parent):
        """Cria a área principal de conteúdo"""
        # Notebook para abas
        self.notebook = ttk.Notebook(parent)
        self.notebook.grid(row=1, column=1, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(0, 0))

        # Aba: Dashboard
        self.create_dashboard_tab()

        # Aba: Dados
        self.create_data_tab()

        # Aba: Relatórios
        self.create_reports_tab()

    def create_dashboard_tab(self):
        """Cria a aba do dashboard com gráficos"""
        dashboard_frame = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(dashboard_frame, text="📈 Dashboard")

        # Frame para gráficos
        self.charts_frame = ttk.Frame(dashboard_frame)
        self.charts_frame.pack(fill=tk.BOTH, expand=True)

        # Mensagem inicial
        welcome_frame = ttk.Frame(self.charts_frame)
        welcome_frame.pack(expand=True)

        welcome_label = ttk.Label(welcome_frame, text="📊 Dashboard de Análises", style='Header.TLabel')
        welcome_label.pack(pady=(100, 20))

        info_label = ttk.Label(welcome_frame,
                              text="Carregue as planilhas e processe os dados\npara visualizar gráficos interativos",
                              style='Info.TLabel', justify=tk.CENTER)
        info_label.pack()

    def create_data_tab(self):
        """Cria a aba de visualização de dados"""
        data_frame = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(data_frame, text="📋 Dados")

        # Título
        data_title = ttk.Label(data_frame, text="📋 Dados Correlacionados", style='Subheader.TLabel')
        data_title.pack(anchor=tk.W, pady=(0, 10))

        # Frame para Treeview com scrollbar
        tree_frame = ttk.Frame(data_frame)
        tree_frame.pack(fill=tk.BOTH, expand=True)

        # Treeview
        self.data_tree = ttk.Treeview(tree_frame)
        self.data_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Scrollbars
        v_scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=self.data_tree.yview)
        v_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.data_tree.configure(yscrollcommand=v_scrollbar.set)

        h_scrollbar = ttk.Scrollbar(data_frame, orient="horizontal", command=self.data_tree.xview)
        h_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)
        self.data_tree.configure(xscrollcommand=h_scrollbar.set)

    def create_reports_tab(self):
        """Cria a aba de relatórios"""
        reports_frame = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(reports_frame, text="📄 Relatórios")

        # Título
        reports_title = ttk.Label(reports_frame, text="📄 Relatório Executivo", style='Subheader.TLabel')
        reports_title.pack(anchor=tk.W, pady=(0, 10))

        # Área de texto com scroll
        self.report_text = scrolledtext.ScrolledText(reports_frame, wrap=tk.WORD, width=80, height=30,
                                                    font=('Consolas', 10), bg='#2C3E50', fg='white',
                                                    insertbackground='white')
        self.report_text.pack(fill=tk.BOTH, expand=True)

    def create_footer(self, parent):
        """Cria o rodapé"""
        footer_frame = ttk.Frame(parent, padding="10")
        footer_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E))

        # Separador
        ttk.Separator(footer_frame, orient='horizontal').pack(fill=tk.X, pady=(0, 10))

        # Info
        footer_text = "🏥 Sistema de Análise de Achados Críticos - CDI | Desenvolvido para otimização da comunicação médica"
        footer_label = ttk.Label(footer_frame, text=footer_text, style='Info.TLabel')
        footer_label.pack()

    def select_file(self, file_type):
        """Seleciona arquivo via dialog"""
        filetypes = [
            ("Arquivos Excel", "*.xlsx *.xls"),
            ("Todos os arquivos", "*.*")
        ]

        filename = filedialog.askopenfilename(
            title=f"Selecionar planilha de {file_type}",
            filetypes=filetypes
        )

        if filename:
            if file_type == 'achados':
                self.achados_file.set(filename)
            else:
                self.status_file.set(filename)

    def _normalize_text(self, text):
        """Normaliza texto para comparação de nomes de coluna."""
        if not isinstance(text, str):
            return ""
        normalized = unicodedata.normalize('NFKD', text)
        normalized = normalized.encode('ascii', 'ignore').decode('ascii')
        return normalized.lower().strip()

    def _find_column(self, df, required_terms, any_terms=None, exclude_terms=None):
        """Busca coluna por termos obrigatórios e opcionais."""
        for col in df.columns:
            col_norm = self._normalize_text(col)
            has_required = all(term in col_norm for term in required_terms)
            has_any = True if not any_terms else any(term in col_norm for term in any_terms)
            has_excluded = False if not exclude_terms else any(term in col_norm for term in exclude_terms)

            if has_required and has_any and not has_excluded:
                return col
        return None

    def _parse_datetime_series(self, series):
        """Converte série para datetime priorizando padrão brasileiro."""
        parsed = pd.to_datetime(series, format='%d/%m/%Y %H:%M', errors='coerce')
        parsed = parsed.fillna(pd.to_datetime(series, format='%d-%m-%Y %H:%M', errors='coerce'))
        parsed = parsed.fillna(pd.to_datetime(series, format='%d/%m/%Y', errors='coerce'))
        parsed = parsed.fillna(pd.to_datetime(series, format='%d-%m-%Y', errors='coerce'))
        parsed = parsed.fillna(pd.to_datetime(series, errors='coerce', dayfirst=True))
        return parsed

    def _parse_datetime_value(self, value):
        """Converte valor único para datetime priorizando padrão brasileiro."""
        return self._parse_datetime_series(pd.Series([value])).iloc[0]

    def _format_date_columns(self, df):
        """Formata colunas de data para DD/MM/AAAA apenas para exibição/export."""
        df_fmt = df.copy()
        date_cols = {
            'Data_Exame',
            'Data_Sinalização',
            'DATA_HORA_PRESCRICAO',
            'STATUS_ALAUDAR',
            'data_sinalizacao_dt',
            'status_laudar_dt',
        }
        for col in date_cols:
            if col in df_fmt.columns:
                parsed = self._parse_datetime_series(df_fmt[col])
                formatted = parsed.dt.strftime('%d/%m/%Y')
                df_fmt[col] = formatted.where(parsed.notna(), '')
        return df_fmt

    def _load_status_dataframe(self):
        """Carrega a planilha de status tentando diferentes linhas de cabeçalho."""
        file_path = self.status_file.get()
        engine = 'openpyxl' if file_path.endswith('.xlsx') else 'xlrd'

        last_df = None
        for header in (1, 0):
            df = pd.read_excel(file_path, engine=engine, header=header)
            df = df.dropna(how='all')
            last_df = df

            if self._find_column(df, ['same']):
                return df

        return last_df

    def _standardize_input_columns(self):
        """Padroniza nomes de coluna para evitar falhas por variação de cabeçalho."""
        achados_map = {
            'SAME': self._find_column(self.df_achados, ['same']),
            'Nome_Paciente': self._find_column(self.df_achados, ['nome', 'paciente']),
            'Data_Exame': self._find_column(self.df_achados, ['data', 'exame']),
            'Descrição_Procedimento': self._find_column(self.df_achados, ['descr', 'proced']),
            'Data_Sinalização': self._find_column(self.df_achados, ['data', 'sinal']),
            'Medico Laudo': self._find_column(self.df_achados, ['medico', 'laudo']),
            'Achado Crítico': self._find_column(self.df_achados, ['achado', 'critico']),
        }

        status_map = {
            'SAME': self._find_column(self.df_status, ['same']),
            'NOME_PACIENTE': self._find_column(self.df_status, ['nome', 'paciente']),
            'DATA_HORA_PRESCRICAO': self._find_column(self.df_status, ['data'], any_terms=['prescri', 'hora']),
            'DESCRICAO_PROCEDIMENTO': self._find_column(self.df_status, ['descr', 'proced']),
            'STATUS_ALAUDAR': self._find_column(self.df_status, ['status', 'laudar']),
        }

        if not achados_map['SAME']:
            raise KeyError(f"Coluna SAME não encontrada na planilha de achados. Colunas: {list(self.df_achados.columns)}")
        if not status_map['SAME']:
            raise KeyError(f"Coluna SAME não encontrada na planilha de status. Colunas: {list(self.df_status.columns)}")

        rename_achados = {original: canonical for canonical, original in achados_map.items()
                          if original and original != canonical}
        rename_status = {original: canonical for canonical, original in status_map.items()
                         if original and original != canonical}

        if rename_achados:
            self.df_achados = self.df_achados.rename(columns=rename_achados)
        if rename_status:
            self.df_status = self.df_status.rename(columns=rename_status)

        # Colunas opcionais ausentes viram vazias para manter o fluxo sem crash.
        for col in ['Nome_Paciente', 'Data_Exame', 'Descrição_Procedimento',
                    'Data_Sinalização', 'Medico Laudo', 'Achado Crítico']:
            if col not in self.df_achados.columns:
                self.df_achados[col] = pd.NA

        for col in ['NOME_PACIENTE', 'DATA_HORA_PRESCRICAO', 'DESCRICAO_PROCEDIMENTO', 'STATUS_ALAUDAR']:
            if col not in self.df_status.columns:
                self.df_status[col] = pd.NA

    def process_data(self):
        """Processa os dados em thread separada"""
        if not self.achados_file.get() or not self.status_file.get():
            messagebox.showerror("Erro", "Por favor, selecione ambas as planilhas")
            return

        self.process_btn.configure(state='disabled')
        self.progress_var.set(0)

        # Executar em thread separada para não travar a interface
        thread = threading.Thread(target=self._process_data_thread)
        thread.daemon = True
        thread.start()

    def _process_data_thread(self):
        """Thread de processamento dos dados"""
        try:
            # Etapa 1: Carregar planilhas
            self.root.after(0, self.update_status, "📂 Carregando planilhas...")
            self.root.after(0, self.progress_var.set, 10)

            # Carregar achados
            engine = 'openpyxl' if self.achados_file.get().endswith('.xlsx') else 'xlrd'
            self.df_achados = pd.read_excel(self.achados_file.get(), engine=engine)

            # Carregar status
            self.df_status = self._load_status_dataframe()

            # Padronizar nomes de colunas para evitar KeyError por variação de cabeçalho
            self._standardize_input_columns()

            self.root.after(0, self.progress_var.set, 30)

            # Etapa 2: Correlacionar dados
            self.root.after(0, self.update_status, "🔗 Correlacionando dados...")
            self.correlate_data()
            self.root.after(0, self.progress_var.set, 60)

            # Etapa 3: Calcular tempos
            self.root.after(0, self.update_status, "⏰ Calculando tempos...")
            self.calculate_times()
            self.root.after(0, self.progress_var.set, 80)

            # Etapa 4: Gerar visualizações
            self.root.after(0, self.update_status, "📊 Gerando gráficos...")
            self.root.after(0, self.create_charts)
            self.root.after(0, self.progress_var.set, 90)

            # Etapa 5: Atualizar interface
            self.root.after(0, self.update_interface)
            self.root.after(0, self.progress_var.set, 100)
            # Mensagem personalizada baseada no filtro
            if self.filter_enabled.get() and self.selected_month.get() != "Todos os Meses":
                filter_msg = f" ({self.selected_month.get()}/{self.selected_year.get()})"
            else:
                filter_msg = ""

            self.root.after(0, self.update_status, f"✅ Análise concluída! {len(self.df_correlacionado)} registros processados{filter_msg}")

        except Exception as e:
            self.root.after(0, messagebox.showerror, "Erro", f"Erro no processamento: {str(e)}")
            self.root.after(0, self.update_status, "❌ Erro no processamento")

        finally:
            self.root.after(0, lambda: self.process_btn.configure(state='normal'))

    def apply_month_filter(self, df):
        """Aplica filtro por mês nos dados se habilitado"""
        if not self.filter_enabled.get() or self.selected_month.get() == "Todos os Meses":
            return df

        # Mapeamento de mês para número
        month_map = {
            "Janeiro": 1, "Fevereiro": 2, "Março": 3, "Abril": 4,
            "Maio": 5, "Junho": 6, "Julho": 7, "Agosto": 8,
            "Setembro": 9, "Outubro": 10, "Novembro": 11, "Dezembro": 12
        }

        selected_month_num = month_map.get(self.selected_month.get())
        selected_year = int(self.selected_year.get())

        if selected_month_num is None:
            return df

        # Filtrar por mês na coluna Data_Sinalização
        df_filtered = df.copy()
        if 'Data_Sinalização' in df_filtered.columns:
            df_filtered['Data_Sinalização'] = self._parse_datetime_series(df_filtered['Data_Sinalização'])
            mask = (df_filtered['Data_Sinalização'].dt.month == selected_month_num) & \
                   (df_filtered['Data_Sinalização'].dt.year == selected_year)
            df_filtered = df_filtered[mask]

        return df_filtered

    def correlate_data(self):
        """Correlaciona as planilhas usando critérios rigorosos"""
        # Aplicar filtro por mês nos achados críticos
        df_achados_filtered = self.apply_month_filter(self.df_achados)

        correlacoes = []

        for idx, achado in df_achados_filtered.iterrows():
            same_achado = achado.get('SAME')
            nome_achado = achado.get('Nome_Paciente', '')
            data_exame_achado = achado.get('Data_Exame')
            descricao_achado = achado.get('Descrição_Procedimento', '')

            if pd.isna(same_achado):
                continue

            # Buscar por SAME
            exames_same = self.df_status[self.df_status['SAME'] == same_achado]
            if len(exames_same) == 0:
                continue

            # Filtrar por nome (básico)
            if pd.notna(nome_achado) and nome_achado:
                exames_nome = exames_same[
                    exames_same['NOME_PACIENTE'].fillna('').astype(str).str.contains(nome_achado.split()[0], case=False, na=False) |
                    exames_same['NOME_PACIENTE'].fillna('').astype(str).str.contains(nome_achado.split()[-1], case=False, na=False)
                ]
            else:
                exames_nome = exames_same

            # FILTRO RIGOROSO POR DATA EXATA
            if pd.notna(data_exame_achado):
                try:
                    data_achado_dt = self._parse_datetime_value(data_exame_achado)
                    if pd.isna(data_achado_dt):
                        continue
                    data_achado_str = data_achado_dt.strftime('%d/%m/%Y')

                    exames_data = []
                    for _, exame in exames_nome.iterrows():
                        if pd.notna(exame.get('DATA_HORA_PRESCRICAO')):
                            try:
                                data_exame_dt = self._parse_datetime_value(exame['DATA_HORA_PRESCRICAO'])
                                if pd.isna(data_exame_dt):
                                    continue
                                data_exame_str = data_exame_dt.strftime('%d/%m/%Y')
                                if data_achado_str == data_exame_str:
                                    exames_data.append(exame)
                            except:
                                continue

                    if len(exames_data) == 0:
                        continue  # Se não encontrar na data exata, pular

                    exames_data = pd.DataFrame(exames_data)
                except:
                    continue
            else:
                continue  # Se não tiver data, pular

            # FILTRO RIGOROSO POR PROCEDIMENTO
            if pd.notna(descricao_achado):
                # Tentar match exato primeiro
                exames_match_exato = []
                for _, exame in exames_data.iterrows():
                    if pd.notna(exame.get('DESCRICAO_PROCEDIMENTO')):
                        desc_achado_clean = descricao_achado.strip().upper()
                        desc_exame_clean = str(exame['DESCRICAO_PROCEDIMENTO']).strip().upper()

                        if desc_achado_clean == desc_exame_clean:
                            exames_match_exato.append(exame)

                if len(exames_match_exato) > 0:
                    exames_finais = pd.DataFrame(exames_match_exato)
                else:
                    # Match por palavras-chave rigoroso
                    exames_procedimento = []
                    descricao_lower = descricao_achado.lower()

                    for _, exame in exames_data.iterrows():
                        if pd.notna(exame.get('DESCRICAO_PROCEDIMENTO')):
                            descricao_exame_lower = str(exame['DESCRICAO_PROCEDIMENTO']).lower()

                            # Verificar palavras importantes
                            palavras_importantes = ['tc', 'rm', 'rx', 'angio', 'abdome', 'torax', 'cranio']
                            matches = sum(1 for palavra in palavras_importantes
                                        if palavra in descricao_lower and palavra in descricao_exame_lower)
                            total = sum(1 for palavra in palavras_importantes if palavra in descricao_lower)

                            if total > 0 and (matches / total) >= 0.7:
                                exames_procedimento.append(exame)

                    if len(exames_procedimento) == 0:
                        continue  # Se não encontrar procedimento compatível, pular

                    exames_finais = pd.DataFrame(exames_procedimento)
            else:
                continue  # Se não tiver procedimento, pular

            # Selecionar o melhor exame com critérios rigorosos
            if len(exames_finais) == 1:
                # Se há apenas 1 exame, é o correto
                melhor_exame = exames_finais.iloc[0]
            else:
                # Se há múltiplos exames, priorizar exames com STATUS_ALAUDAR
                exames_com_status = exames_finais[exames_finais['STATUS_ALAUDAR'].notna()]
                melhor_exame = exames_com_status.iloc[0] if len(exames_com_status) > 0 else exames_finais.iloc[0]

            # Combinar dados (FORA do if/else)
            correlacao = achado.copy()
            for col in melhor_exame.index:
                if col != 'SAME':
                    correlacao[col] = melhor_exame[col]

            correlacoes.append(correlacao)

        self.df_correlacionado = pd.DataFrame(correlacoes)

    def calculate_times(self):
        """Calcula tempos de comunicação"""
        if self.df_correlacionado is None or len(self.df_correlacionado) == 0:
            return

        # Filtrar registros com STATUS_ALAUDAR válido
        df_com_status = self.df_correlacionado[
            self.df_correlacionado['STATUS_ALAUDAR'].notna()
        ].copy()

        if len(df_com_status) == 0:
            return

        # Converter datas
        df_com_status['data_sinalizacao_dt'] = self._parse_datetime_series(
            df_com_status['Data_Sinalização']
        )
        df_com_status['status_laudar_dt'] = self._parse_datetime_series(
            df_com_status['STATUS_ALAUDAR']
        )

        # Calcular tempos
        df_com_status['tempo_comunicacao_horas'] = (
            df_com_status['data_sinalizacao_dt'] - df_com_status['status_laudar_dt']
        ).dt.total_seconds() / 3600

        df_com_status['fora_do_prazo'] = df_com_status['tempo_comunicacao_horas'] > 1

        # Filtrar registros válidos
        registros_validos = (
            df_com_status['data_sinalizacao_dt'].notna() &
            df_com_status['status_laudar_dt'].notna() &
            (df_com_status['tempo_comunicacao_horas'] >= 0)
        )

        self.df_correlacionado = df_com_status[registros_validos].copy()

    def create_charts(self):
        """Cria os gráficos na interface"""
        if self.df_correlacionado is None or len(self.df_correlacionado) == 0:
            return

        # Limpar frame anterior
        for widget in self.charts_frame.winfo_children():
            widget.destroy()

        # Configurar grid 2x2 para gráficos
        for i in range(2):
            self.charts_frame.rowconfigure(i, weight=1)
            self.charts_frame.columnconfigure(i, weight=1)

        # Gráfico 1: Compliance (pie chart)
        self.create_compliance_chart()

        # Gráfico 2: Top médicos (bar chart)
        self.create_doctors_chart()

        # Gráfico 3: Distribuição temporal
        self.create_time_distribution_chart()

        # Gráfico 4: Achados críticos
        self.create_findings_chart()

    def create_compliance_chart(self):
        """Cria gráfico de compliance"""
        frame = ttk.Frame(self.charts_frame)
        frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=5, pady=5)

        fig = Figure(figsize=(6, 4), dpi=100, facecolor='#2C3E50')
        ax = fig.add_subplot(111)

        no_prazo = (~self.df_correlacionado['fora_do_prazo']).sum()
        fora_prazo = self.df_correlacionado['fora_do_prazo'].sum()

        sizes = [no_prazo, fora_prazo]
        labels = ['No Prazo', 'Fora do Prazo']
        colors = ['#27AE60', '#E74C3C']

        wedges, texts, autotexts = ax.pie(sizes, labels=labels, colors=colors, autopct='%1.1f%%',
                                         startangle=90, textprops={'color': 'white', 'fontsize': 10})

        ax.set_title('Compliance de Comunicação', color='white', fontsize=12, fontweight='bold', pad=20)
        fig.patch.set_facecolor('#2C3E50')

        canvas = FigureCanvasTkAgg(fig, frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

    def create_doctors_chart(self):
        """Cria gráfico dos médicos"""
        frame = ttk.Frame(self.charts_frame)
        frame.grid(row=0, column=1, sticky=(tk.W, tk.E, tk.N, tk.S), padx=5, pady=5)

        fig = Figure(figsize=(6, 4), dpi=100, facecolor='#2C3E50')
        ax = fig.add_subplot(111)

        # Estatísticas por médico
        stats_medicos = self.df_correlacionado.groupby('Medico Laudo').agg({
            'fora_do_prazo': ['count', 'sum']
        })
        stats_medicos.columns = ['total_comunicados', 'fora_prazo']
        stats_medicos['percentual_fora_prazo'] = (
            (stats_medicos['fora_prazo'] / stats_medicos['total_comunicados']) * 100
        ).round(1)

        # Top 8 médicos
        top_medicos = stats_medicos.nlargest(8, 'total_comunicados')

        # Cores baseadas na performance
        colors = ['#27AE60' if pct <= 20 else '#F39C12' if pct <= 50 else '#E74C3C'
                 for pct in top_medicos['percentual_fora_prazo']]

        # Nomes simplificados
        nomes = [nome.split()[-1] for nome in top_medicos.index]

        bars = ax.barh(nomes, top_medicos['total_comunicados'], color=colors, alpha=0.8)

        # Adicionar valores nas barras
        for i, bar in enumerate(bars):
            width = bar.get_width()
            pct = top_medicos['percentual_fora_prazo'].iloc[i]
            ax.text(width + 0.1, bar.get_y() + bar.get_height()/2,
                   f'{int(width)} ({pct:.1f}%)',
                   ha='left', va='center', color='white', fontsize=8)

        ax.set_title('Top Médicos - Comunicados', color='white', fontsize=12, fontweight='bold', pad=20)
        ax.set_xlabel('Total de Comunicados', color='white')
        ax.tick_params(colors='white')
        fig.patch.set_facecolor('#2C3E50')

        canvas = FigureCanvasTkAgg(fig, frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

    def create_time_distribution_chart(self):
        """Cria gráfico de distribuição temporal"""
        frame = ttk.Frame(self.charts_frame)
        frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=5, pady=5)

        fig = Figure(figsize=(6, 4), dpi=100, facecolor='#2C3E50')
        ax = fig.add_subplot(111)

        # Criar bins
        bins = [0, 1, 6, 24, 168, float('inf')]
        labels = ['≤ 1h', '1-6h', '6-24h', '1-7 dias', '> 7 dias']

        self.df_correlacionado['faixa_tempo'] = pd.cut(
            self.df_correlacionado['tempo_comunicacao_horas'],
            bins=bins, labels=labels, right=False
        )

        faixa_counts = self.df_correlacionado['faixa_tempo'].value_counts()

        colors = ['#27AE60', '#F1C40F', '#E67E22', '#E74C3C', '#8E44AD']

        bars = ax.bar(faixa_counts.index, faixa_counts.values,
                     color=colors[:len(faixa_counts)], alpha=0.8)

        # Adicionar valores nas barras
        for bar in bars:
            height = bar.get_height()
            ax.text(bar.get_x() + bar.get_width()/2., height + 0.1,
                   f'{int(height)}', ha='center', va='bottom', color='white', fontsize=9)

        ax.set_title('Distribuição por Faixa de Tempo', color='white', fontsize=12, fontweight='bold', pad=20)
        ax.set_ylabel('Número de Casos', color='white')
        ax.tick_params(colors='white')
        fig.patch.set_facecolor('#2C3E50')

        canvas = FigureCanvasTkAgg(fig, frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

    def create_findings_chart(self):
        """Cria gráfico dos achados críticos"""
        frame = ttk.Frame(self.charts_frame)
        frame.grid(row=1, column=1, sticky=(tk.W, tk.E, tk.N, tk.S), padx=5, pady=5)

        fig = Figure(figsize=(6, 4), dpi=100, facecolor='#2C3E50')
        ax = fig.add_subplot(111)

        # Top achados
        achados_counts = self.df_correlacionado['Achado Crítico'].value_counts().head(6)

        # Simplificar nomes
        nomes_simp = []
        for nome in achados_counts.index:
            if ':' in nome:
                partes = nome.split(':')
                especialidade = partes[0].replace('Especialista em ', '').replace('Especialidade ', '')
                achado = partes[1].strip()[:25]
                nome_simp = f"{especialidade}:\n{achado}"
            else:
                nome_simp = nome[:30]
            nomes_simp.append(nome_simp)

        bars = ax.bar(range(len(achados_counts)), achados_counts.values,
                     color='#3498DB', alpha=0.8)

        # Adicionar valores nas barras
        for bar in bars:
            height = bar.get_height()
            ax.text(bar.get_x() + bar.get_width()/2., height + 0.1,
                   f'{int(height)}', ha='center', va='bottom', color='white', fontsize=9)

        ax.set_title('Top Achados Críticos', color='white', fontsize=12, fontweight='bold', pad=20)
        ax.set_ylabel('Frequência', color='white')
        ax.set_xticks(range(len(achados_counts)))
        ax.set_xticklabels(nomes_simp, rotation=45, ha='right', fontsize=8, color='white')
        ax.tick_params(colors='white')
        fig.patch.set_facecolor('#2C3E50')

        plt.tight_layout()

        canvas = FigureCanvasTkAgg(fig, frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

    def update_interface(self):
        """Atualiza toda a interface com os novos dados"""
        self.update_metrics()
        self.update_data_table()
        self.update_report()
        self.export_btn.configure(state='normal')

    def update_metrics(self):
        """Atualiza as métricas na sidebar"""
        if self.df_correlacionado is None:
            return

        # Limpar métricas anteriores
        for widget in self.metrics_frame.winfo_children():
            widget.destroy()

        # Calcular métricas
        total_casos = len(self.df_correlacionado)
        no_prazo = (~self.df_correlacionado['fora_do_prazo']).sum()
        fora_prazo = self.df_correlacionado['fora_do_prazo'].sum()
        percentual_compliance = (no_prazo / total_casos) * 100
        tempo_mediano = self.df_correlacionado['tempo_comunicacao_horas'].median()

        # Compliance
        compliance_frame = ttk.Frame(self.metrics_frame)
        compliance_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Label(compliance_frame, text="Taxa de Compliance:", style='Info.TLabel').pack(anchor=tk.W)

        compliance_color = 'success' if percentual_compliance >= 80 else 'warning' if percentual_compliance >= 60 else 'danger'
        compliance_text = f"{percentual_compliance:.1f}% ({no_prazo}/{total_casos})"

        if compliance_color == 'success':
            color_code = '#27AE60'
        elif compliance_color == 'warning':
            color_code = '#F39C12'
        else:
            color_code = '#E74C3C'

        compliance_label = tk.Label(compliance_frame, text=compliance_text,
                                  fg=color_code, bg='#34495E', font=('Helvetica', 12, 'bold'))
        compliance_label.pack(anchor=tk.W, pady=(2, 0))

        # Tempo mediano
        time_frame = ttk.Frame(self.metrics_frame)
        time_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Label(time_frame, text="Tempo Mediano:", style='Info.TLabel').pack(anchor=tk.W)
        time_label = tk.Label(time_frame, text=f"{tempo_mediano:.1f} horas",
                             fg='#3498DB', bg='#34495E', font=('Helvetica', 12, 'bold'))
        time_label.pack(anchor=tk.W, pady=(2, 0))

        # Casos fora do prazo
        late_frame = ttk.Frame(self.metrics_frame)
        late_frame.pack(fill=tk.X)

        ttk.Label(late_frame, text="Casos Fora do Prazo:", style='Info.TLabel').pack(anchor=tk.W)
        late_label = tk.Label(late_frame, text=f"{fora_prazo} casos",
                             fg='#E74C3C', bg='#34495E', font=('Helvetica', 12, 'bold'))
        late_label.pack(anchor=tk.W, pady=(2, 0))

    def update_data_table(self):
        """Atualiza a tabela de dados"""
        if self.df_correlacionado is None:
            return

        # Limpar dados anteriores
        for item in self.data_tree.get_children():
            self.data_tree.delete(item)

        # Colunas principais
        cols = ['SAME', 'Nome_Paciente', 'Medico Laudo', 'Achado Crítico',
                'Data_Sinalização', 'STATUS_ALAUDAR', 'tempo_comunicacao_horas', 'fora_do_prazo']

        # Configurar colunas
        self.data_tree['columns'] = cols
        self.data_tree['show'] = 'headings'

        for col in cols:
            self.data_tree.heading(col, text=col)
            self.data_tree.column(col, width=120, minwidth=80)

        # Adicionar dados (primeiros 100 registros)
        df_display = self.df_correlacionado[cols].head(100)

        for _, row in df_display.iterrows():
            values = []
            for col in cols:
                value = row[col]
                if col == 'tempo_comunicacao_horas':
                    values.append(f"{value:.2f}h" if pd.notna(value) else "N/A")
                elif col == 'fora_do_prazo':
                    values.append("SIM" if value else "NÃO")
                elif col in {'Data_Sinalização', 'STATUS_ALAUDAR'}:
                    dt_value = self._parse_datetime_value(value)
                    values.append(dt_value.strftime('%d/%m/%Y') if pd.notna(dt_value) else "")
                else:
                    values.append(str(value) if pd.notna(value) else "")

            # Colorir linhas fora do prazo
            tags = ('late',) if row['fora_do_prazo'] else ('normal',)
            self.data_tree.insert('', 'end', values=values, tags=tags)

        # Configurar tags de cores
        self.data_tree.tag_configure('late', background='#E74C3C', foreground='white')
        self.data_tree.tag_configure('normal', background='#34495E', foreground='white')

    def update_report(self):
        """Atualiza o relatório executivo"""
        if self.df_correlacionado is None:
            return

        # Limpar relatório anterior
        self.report_text.delete(1.0, tk.END)

        # Gerar relatório
        total_comunicados = len(self.df_correlacionado)
        fora_prazo = self.df_correlacionado['fora_do_prazo'].sum()
        no_prazo = total_comunicados - fora_prazo
        percentual_fora_prazo = (fora_prazo / total_comunicados) * 100

        tempo_medio = self.df_correlacionado['tempo_comunicacao_horas'].mean()
        tempo_mediano = self.df_correlacionado['tempo_comunicacao_horas'].median()
        tempo_max = self.df_correlacionado['tempo_comunicacao_horas'].max()

        # Top 5 médicos
        stats_medicos = self.df_correlacionado.groupby('Medico Laudo').agg({
            'fora_do_prazo': ['count', 'sum']
        })
        stats_medicos.columns = ['total_comunicados', 'fora_prazo']
        stats_medicos['percentual_fora_prazo'] = (
            (stats_medicos['fora_prazo'] / stats_medicos['total_comunicados']) * 100
        ).round(2)
        top5_medicos = stats_medicos.nlargest(5, 'total_comunicados')

        # Top 5 achados
        top5_achados = self.df_correlacionado['Achado Crítico'].value_counts().head(5)

        # Montar relatório
        relatorio = f"""
RELATÓRIO EXECUTIVO - ANÁLISE DE ACHADOS CRÍTICOS
================================================================
Data da Análise: {datetime.now().strftime('%d/%m/%Y %H:%M')}

📊 RESUMO EXECUTIVO
================================================================
• Total de comunicados analisados: {total_comunicados:,}
• Comunicados no prazo (≤ 1h): {no_prazo:,} ({100-percentual_fora_prazo:.1f}%)
• Comunicados fora do prazo (> 1h): {fora_prazo:,} ({percentual_fora_prazo:.1f}%)

⏱️ ANÁLISE TEMPORAL
================================================================
• Tempo médio de comunicação: {tempo_medio:.1f} horas ({tempo_medio/24:.1f} dias)
• Tempo mediano de comunicação: {tempo_mediano:.1f} horas ({tempo_mediano/24:.1f} dias)
• Tempo máximo registrado: {tempo_max:.1f} horas ({tempo_max/24:.1f} dias)

👨‍⚕️ TOP 5 MÉDICOS COM MAIS COMUNICADOS
================================================================
"""

        for i, (medico, dados) in enumerate(top5_medicos.iterrows(), 1):
            relatorio += f"{i}. {medico.split()[-1]}: {dados['total_comunicados']} comunicados ({dados['percentual_fora_prazo']:.1f}% fora do prazo)\n"

        relatorio += f"""
🔍 TOP 5 ACHADOS CRÍTICOS MAIS FREQUENTES
================================================================
"""

        for i, (achado, freq) in enumerate(top5_achados.items(), 1):
            achado_simp = achado.split(':')[-1].strip()[:50] if ':' in achado else achado[:50]
            relatorio += f"{i}. {achado_simp}: {freq} ocorrências\n"

        relatorio += f"""
📈 RECOMENDAÇÕES PRIORITÁRIAS
================================================================
1. IMPLEMENTAR MONITORAMENTO EM TEMPO REAL
   - Sistema de alertas automáticos para achados críticos
   - Dashboard de acompanhamento em tempo real

2. TREINAMENTO E CAPACITAÇÃO
   - Foco nos médicos com maior volume de comunicados
   - Revisão dos protocolos de comunicação de emergência

3. REVISÃO DE PROCESSOS
   - Investigar causas dos atrasos sistemáticos
   - Otimizar fluxo de trabalho para achados críticos

4. METAS DE MELHORIA
   - Meta imediata: 80% dos comunicados dentro de 1 hora
   - Meta de longo prazo: 95% dos comunicados dentro de 1 hora

================================================================
Relatório gerado automaticamente pelo Sistema de Análise de
Achados Críticos desenvolvido especificamente para o CDI.
================================================================
"""

        # Inserir relatório
        self.report_text.insert(tk.END, relatorio)

    def export_excel(self):
        """Exporta dados para Excel"""
        if self.df_correlacionado is None:
            messagebox.showerror("Erro", "Nenhum dado para exportar")
            return

        filename = filedialog.asksaveasfilename(
            title="Salvar relatório Excel",
            defaultextension=".xlsx",
            filetypes=[("Arquivos Excel", "*.xlsx"), ("Todos os arquivos", "*.*")]
        )

        if filename:
            try:
                with pd.ExcelWriter(filename, engine='openpyxl') as writer:
                    export_df = self._format_date_columns(self.df_correlacionado)
                    # Sheet principal
                    export_df.to_excel(writer, sheet_name='Dados_Correlacionados', index=False)

                    # Estatísticas médicos
                    stats_medicos = self.df_correlacionado.groupby('Medico Laudo').agg({
                        'fora_do_prazo': ['count', 'sum'],
                        'tempo_comunicacao_horas': ['mean', 'median', 'max']
                    }).round(2)
                    stats_medicos.columns = ['total_comunicados', 'fora_prazo', 'tempo_medio_h', 'tempo_mediano_h', 'tempo_max_h']
                    stats_medicos['percentual_fora_prazo'] = (
                        (stats_medicos['fora_prazo'] / stats_medicos['total_comunicados']) * 100
                    ).round(2)
                    stats_medicos.to_excel(writer, sheet_name='Estatisticas_Medicos')

                    # Estatísticas achados
                    stats_achados = self.df_correlacionado.groupby('Achado Crítico').agg({
                        'fora_do_prazo': ['count', 'sum'],
                        'tempo_comunicacao_horas': ['mean', 'median', 'max']
                    }).round(2)
                    stats_achados.columns = ['total_ocorrencias', 'fora_prazo', 'tempo_medio_h', 'tempo_mediano_h', 'tempo_max_h']
                    stats_achados['percentual_fora_prazo'] = (
                        (stats_achados['fora_prazo'] / stats_achados['total_ocorrencias']) * 100
                    ).round(2)
                    stats_achados.to_excel(writer, sheet_name='Estatisticas_Achados')

                messagebox.showinfo("Sucesso", f"Relatório exportado com sucesso!\n{filename}")

            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao exportar: {str(e)}")

    def clear_data(self):
        """Limpa todos os dados"""
        if messagebox.askyesno("Confirmar", "Deseja limpar todos os dados carregados?"):
            self.df_achados = None
            self.df_status = None
            self.df_correlacionado = None

            # Limpar interface
            self.achados_file.set("")
            self.status_file.set("")
            self.progress_var.set(0)
            self.update_status("Dados limpos. Pronto para nova análise.")

            # Limpar gráficos
            for widget in self.charts_frame.winfo_children():
                widget.destroy()

            # Recriar tela inicial
            welcome_frame = ttk.Frame(self.charts_frame)
            welcome_frame.pack(expand=True)

            welcome_label = ttk.Label(welcome_frame, text="📊 Dashboard de Análises", style='Header.TLabel')
            welcome_label.pack(pady=(100, 20))

            info_label = ttk.Label(welcome_frame,
                                  text="Carregue as planilhas e processe os dados\npara visualizar gráficos interativos",
                                  style='Info.TLabel', justify=tk.CENTER)
            info_label.pack()

            # Limpar métricas
            for widget in self.metrics_frame.winfo_children():
                widget.destroy()
            no_data_label = ttk.Label(self.metrics_frame, text="Carregue os dados para ver as métricas",
                                     style='Info.TLabel')
            no_data_label.pack(anchor=tk.W)

            # Limpar tabela
            for item in self.data_tree.get_children():
                self.data_tree.delete(item)

            # Limpar relatório
            self.report_text.delete(1.0, tk.END)

            # Desabilitar export
            self.export_btn.configure(state='disabled')

    def update_status(self, message):
        """Atualiza a mensagem de status"""
        self.status_var.set(message)

    def run(self):
        """Executa a interface"""
        self.root.mainloop()

def main():
    """Função principal"""
    try:
        app = ModernGUI()
        app.run()
    except Exception as e:
        messagebox.showerror("Erro Fatal", f"Erro ao inicializar aplicação: {str(e)}")

if __name__ == "__main__":
    main()

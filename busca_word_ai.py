#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Sistema de Busca Avançada em Arquivos Word com IA
Utiliza Google Gemini para identificação inteligente de nomes e empresas
Autor: Gerado automaticamente
Versão: 2.0
"""

import os
import sys
import json
import time
import threading
import queue
from pathlib import Path
from datetime import datetime
from typing import List, Dict, Tuple
import tkinter as tk
from tkinter import messagebox
import customtkinter as ctk

# Bibliotecas para manipulação de Word
try:
    from docx import Document
    from docx.oxml.text.paragraph import CT_P
    from docx.oxml.table import CT_Tbl
    from docx.table import _Cell, Table
    from docx.text.paragraph import Paragraph
except ImportError:
    print("ERRO: python-docx não instalado. Execute: pip install python-docx")
    sys.exit(1)

# Biblioteca Google Gemini
try:
    import google.generativeai as genai
except ImportError:
    print("ERRO: google-generativeai não instalado. Execute: pip install google-generativeai")
    sys.exit(1)


# ===========================
# CONFIGURAÇÕES GLOBAIS
# ===========================
PASTA_BASE = r"\\192.168.20.100\trabalho\Transcrições"
CONFIG_FILE = "config.json"
EXTENSIONS = ['.docx', '.doc']
NUM_THREADS = 10

# Cores (Tema Claro Agradável)
COLORS = {
    "bg_main": "#F5F7FA",           # Fundo principal (azul acinzentado muito claro)
    "bg_card": "#FFFFFF",            # Cards e frames (branco puro)
    "bg_header": "#4A90E2",          # Cabeçalho (azul suave)
    "bg_input": "#ECF0F1",           # Inputs (cinza muito claro)
    "fg_primary": "#2C3E50",         # Texto primário (cinza escuro azulado)
    "fg_secondary": "#7F8C8D",       # Texto secundário (cinza médio)
    "fg_header": "#FFFFFF",          # Texto do cabeçalho (branco)
    "accent": "#5DADE2",             # Botão principal (azul claro)
    "accent_hover": "#3498DB",       # Hover do botão (azul médio)
    "success": "#27AE60",            # Verde suave
    "error": "#E74C3C",              # Vermelho suave
    "warning": "#F39C12",            # Laranja suave
    "info": "#3498DB",               # Azul informação
    "border": "#BDC3C7"              # Bordas sutis
}


# ===========================
# CLASSE DE CONFIGURAÇÃO
# ===========================
class ConfigManager:
    """Gerencia a persistência da API Key"""

    def __init__(self, config_file: str = CONFIG_FILE):
        self.config_file = config_file
        self.config = self.load_config()

    def load_config(self) -> Dict:
        """Carrega configurações do arquivo JSON"""
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except Exception as e:
                print(f"Erro ao carregar config: {e}")
                return {}
        return {}

    def save_config(self, config: Dict) -> bool:
        """Salva configurações no arquivo JSON"""
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(config, f, indent=4, ensure_ascii=False)
            return True
        except Exception as e:
            print(f"Erro ao salvar config: {e}")
            return False

    def get_api_key(self) -> str:
        """Retorna a API Key salva"""
        return self.config.get('gemini_api_key', '')

    def set_api_key(self, api_key: str) -> bool:
        """Define e salva a API Key"""
        self.config['gemini_api_key'] = api_key
        return self.save_config(self.config)


# ===========================
# CLASSE DO MOTOR DE BUSCA
# ===========================
class WordSearchEngine:
    """Motor de busca paralelo para arquivos Word"""

    def __init__(self, base_path: str, num_threads: int = NUM_THREADS):
        self.base_path = base_path
        self.num_threads = num_threads
        self.results_queue = queue.Queue()
        self.progress_queue = queue.Queue()
        self.stop_event = threading.Event()
        self.files_found = []
        self.files_with_errors = []
        self.total_files_processed = 0
        self.lock = threading.Lock()

    def search_in_document(self, file_path: str, search_term: str, use_ai: bool = True,
                          api_key: str = None) -> Tuple[bool, str]:
        """
        Busca um termo em um documento Word
        Prioriza termos em negrito e sublinhado
        Retorna (encontrado: bool, contexto: str)
        """
        try:
            doc = Document(file_path)
            search_term_lower = search_term.lower()
            found_contexts = []

            # Buscar em parágrafos
            for para in doc.paragraphs:
                para_text = para.text

                # Verificar se o termo está no parágrafo
                if search_term_lower in para_text.lower():
                    # Verificar formatação (negrito e sublinhado tem prioridade)
                    has_bold_underline = False

                    for run in para.runs:
                        if search_term_lower in run.text.lower():
                            if run.bold and run.underline:
                                has_bold_underline = True
                                found_contexts.append(f"[DESTAQUE] {para_text[:200]}")
                                break

                    if not has_bold_underline:
                        found_contexts.append(para_text[:200])

            # Buscar em tabelas
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        cell_text = cell.text
                        if search_term_lower in cell_text.lower():
                            found_contexts.append(f"[TABELA] {cell_text[:200]}")

            if found_contexts:
                context = " | ".join(found_contexts[:3])  # Primeiros 3 contextos
                return True, context

            # Se habilitado, usar IA para verificação mais profunda
            if use_ai and api_key:
                full_text = "\n".join([p.text for p in doc.paragraphs])
                ai_result = self.check_with_ai(full_text, search_term, api_key)
                if ai_result:
                    return True, f"[IA] Menção encontrada no contexto do documento"

            return False, ""

        except PermissionError:
            # Arquivo bloqueado
            return False, "LOCKED"
        except Exception as e:
            return False, f"ERROR: {str(e)}"

    def check_with_ai(self, text: str, search_term: str, api_key: str) -> bool:
        """
        Usa Google Gemini para verificar se o texto menciona o termo buscado
        Modelo: gemini-2.0-flash-exp
        """
        try:
            genai.configure(api_key=api_key)
            model = genai.GenerativeModel('gemini-2.0-flash-exp')

            prompt = f"""
            Analise o seguinte texto e determine se há menção à pessoa ou empresa: "{search_term}"

            Considere:
            - Variações de nome (abreviações, apelidos)
            - Menções indiretas
            - Contexto de negócios/jurídico

            Responda APENAS "SIM" ou "NÃO".

            Texto:
            {text[:5000]}
            """

            response = model.generate_content(prompt)
            result = response.text.strip().upper()

            return "SIM" in result

        except Exception as e:
            print(f"Erro na verificação com IA: {e}")
            return False

    def process_files_in_thread(self, file_paths: List[str], search_term: str,
                                use_ai: bool, api_key: str, thread_id: int):
        """Processa uma lista de arquivos em uma thread"""
        for file_path in file_paths:
            if self.stop_event.is_set():
                break

            try:
                found, context = self.search_in_document(file_path, search_term, use_ai, api_key)

                with self.lock:
                    self.total_files_processed += 1

                if found:
                    with self.lock:
                        self.files_found.append((file_path, context))
                    self.results_queue.put(('found', file_path, context))
                elif context == "LOCKED":
                    with self.lock:
                        self.files_with_errors.append((file_path, "Arquivo bloqueado/aberto"))
                    self.results_queue.put(('locked', file_path, None))
                elif context.startswith("ERROR:"):
                    with self.lock:
                        self.files_with_errors.append((file_path, context))
                    self.results_queue.put(('error', file_path, context))

                # Atualizar progresso
                self.progress_queue.put(('progress', self.total_files_processed))

            except Exception as e:
                with self.lock:
                    self.files_with_errors.append((file_path, f"Erro: {str(e)}"))
                self.results_queue.put(('error', file_path, str(e)))

        # Sinalizar que esta thread terminou
        self.progress_queue.put(('thread_done', thread_id))

    def search(self, search_term: str, use_ai: bool = True, api_key: str = None) -> bool:
        """
        Inicia busca paralela
        Retorna True se iniciou com sucesso
        """
        self.stop_event.clear()
        self.files_found = []
        self.files_with_errors = []
        self.total_files_processed = 0

        # Verificar se o caminho existe
        if not os.path.exists(self.base_path):
            self.results_queue.put(('error_path', self.base_path, None))
            return False

        # Listar todos os arquivos Word recursivamente
        all_files = []
        try:
            for ext in EXTENSIONS:
                all_files.extend(Path(self.base_path).rglob(f'*{ext}'))
        except Exception as e:
            self.results_queue.put(('error_path', str(e), None))
            return False

        if not all_files:
            self.results_queue.put(('no_files', None, None))
            return False

        all_files = [str(f) for f in all_files]

        # Dividir arquivos entre threads
        files_per_thread = len(all_files) // self.num_threads
        if files_per_thread == 0:
            files_per_thread = 1

        threads = []
        for i in range(self.num_threads):
            start_idx = i * files_per_thread
            if i == self.num_threads - 1:
                end_idx = len(all_files)
            else:
                end_idx = start_idx + files_per_thread

            thread_files = all_files[start_idx:end_idx]

            if thread_files:
                t = threading.Thread(
                    target=self.process_files_in_thread,
                    args=(thread_files, search_term, use_ai, api_key, i),
                    daemon=True
                )
                threads.append(t)
                t.start()

        # Thread de monitoramento
        def monitor_threads():
            for t in threads:
                t.join()
            self.progress_queue.put(('complete', None))

        monitor_thread = threading.Thread(target=monitor_threads, daemon=True)
        monitor_thread.start()

        return True

    def stop(self):
        """Para a busca"""
        self.stop_event.set()


# ===========================
# INTERFACE GRÁFICA
# ===========================
class SearchApp(ctk.CTk):
    """Aplicação principal com interface gráfica moderna"""

    def __init__(self):
        super().__init__()

        # Configurações da janela
        self.title("🔍 Busca Avançada em Transcrições - IA Integrada")
        self.geometry("1400x900")

        # Configurar tema CLARO
        ctk.set_appearance_mode("light")
        ctk.set_default_color_theme("blue")

        # Gerenciadores
        self.config_manager = ConfigManager()
        self.search_engine = WordSearchEngine(PASTA_BASE, NUM_THREADS)

        # Variáveis
        self.search_in_progress = False
        self.start_time = None
        self.progress_window = None

        # Construir interface
        self.build_ui()

        # Iniciar monitoramento de progresso
        self.after(100, self.check_progress)

    def build_ui(self):
        """Constrói a interface gráfica"""

        # ===== FRAME PRINCIPAL =====
        self.main_frame = ctk.CTkFrame(self, fg_color=COLORS["bg_main"])
        self.main_frame.pack(fill="both", expand=True, padx=0, pady=0)

        # ===== CABEÇALHO =====
        header_frame = ctk.CTkFrame(self.main_frame, fg_color=COLORS["bg_header"], height=100)
        header_frame.pack(fill="x", padx=0, pady=0)
        header_frame.pack_propagate(False)

        # Container do cabeçalho (para alinhar título e botão de config)
        header_container = ctk.CTkFrame(header_frame, fg_color="transparent")
        header_container.pack(fill="both", expand=True)

        # Título centralizado
        title_label = ctk.CTkLabel(
            header_container,
            text="🔍 BUSCA AVANÇADA EM TRANSCRIÇÕES",
            font=ctk.CTkFont(size=32, weight="bold"),
            text_color=COLORS["fg_header"]
        )
        title_label.pack(pady=25)

        subtitle_label = ctk.CTkLabel(
            header_container,
            text="Powered by Google Gemini AI • 10 Threads Paralelas",
            font=ctk.CTkFont(size=12),
            text_color=COLORS["fg_header"]
        )
        subtitle_label.pack(pady=(0, 10))

        # Botão de configuração (canto superior direito)
        config_btn = ctk.CTkButton(
            header_frame,
            text="⚙️",
            width=50,
            height=50,
            font=ctk.CTkFont(size=24),
            command=self.show_config_dialog,
            fg_color=COLORS["bg_card"],
            text_color=COLORS["fg_primary"],
            hover_color=COLORS["bg_input"],
            corner_radius=25
        )
        config_btn.place(relx=0.98, rely=0.5, anchor="e")

        # ===== BUSCA =====
        search_frame = ctk.CTkFrame(self.main_frame, fg_color=COLORS["bg_card"], corner_radius=15)
        search_frame.pack(fill="x", padx=30, pady=(30, 20))

        # Container interno com padding
        search_inner = ctk.CTkFrame(search_frame, fg_color="transparent")
        search_inner.pack(fill="x", padx=20, pady=20)

        search_label = ctk.CTkLabel(
            search_inner,
            text="👤 Nome ou Empresa:",
            font=ctk.CTkFont(size=16, weight="bold"),
            text_color=COLORS["fg_primary"]
        )
        search_label.pack(anchor="w", pady=(0, 10))

        # Container para input e botão
        input_container = ctk.CTkFrame(search_inner, fg_color="transparent")
        input_container.pack(fill="x")

        self.search_entry = ctk.CTkEntry(
            input_container,
            placeholder_text="Digite o nome da pessoa ou empresa para buscar...",
            height=50,
            font=ctk.CTkFont(size=16),
            fg_color=COLORS["bg_input"],
            text_color=COLORS["fg_primary"],
            border_width=2,
            border_color=COLORS["border"],
            corner_radius=10
        )
        self.search_entry.pack(side="left", fill="x", expand=True, padx=(0, 15))
        self.search_entry.bind('<Return>', lambda e: self.start_search())

        search_btn = ctk.CTkButton(
            input_container,
            text="🔍 BUSCAR",
            width=180,
            height=50,
            command=self.start_search,
            font=ctk.CTkFont(size=16, weight="bold"),
            fg_color=COLORS["accent"],
            hover_color=COLORS["accent_hover"],
            text_color="white",
            corner_radius=10
        )
        search_btn.pack(side="right")

        # Info IA sempre ativo
        ai_info_label = ctk.CTkLabel(
            search_inner,
            text="🤖 Busca com IA sempre ativa • Prioriza termos em Negrito e Sublinhado",
            font=ctk.CTkFont(size=12),
            text_color=COLORS["fg_secondary"]
        )
        ai_info_label.pack(anchor="w", pady=(10, 0))

        # ===== INFO PASTA =====
        info_frame = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        info_frame.pack(fill="x", padx=30, pady=(0, 10))

        info_label = ctk.CTkLabel(
            info_frame,
            text=f"📁 Pasta Base: {PASTA_BASE}",
            font=ctk.CTkFont(size=11),
            text_color=COLORS["fg_secondary"]
        )
        info_label.pack(anchor="w")

        # ===== CONTAINER DE RESULTADOS =====
        results_container = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        results_container.pack(fill="both", expand=True, padx=30, pady=(10, 20))

        # Configurar grid
        results_container.grid_columnconfigure(0, weight=1)
        results_container.grid_columnconfigure(1, weight=1)
        results_container.grid_rowconfigure(0, weight=1)

        # ===== PAINEL DE RESULTADOS =====
        results_frame = ctk.CTkFrame(results_container, fg_color=COLORS["bg_card"], corner_radius=15)
        results_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 10))

        results_header = ctk.CTkFrame(results_frame, fg_color=COLORS["success"], corner_radius=10, height=50)
        results_header.pack(fill="x", padx=15, pady=15)
        results_header.pack_propagate(False)

        results_title = ctk.CTkLabel(
            results_header,
            text="📋 RESULTADOS",
            font=ctk.CTkFont(size=18, weight="bold"),
            text_color="white"
        )
        results_title.pack(pady=10)

        # Lista de resultados com scrollbar
        self.results_textbox = ctk.CTkTextbox(
            results_frame,
            font=ctk.CTkFont(size=13),
            fg_color=COLORS["bg_input"],
            text_color=COLORS["fg_primary"],
            wrap="word",
            corner_radius=10
        )
        self.results_textbox.pack(fill="both", expand=True, padx=15, pady=(0, 15))
        self.results_textbox.bind('<Double-Button-1>', self.open_file_from_selection)

        # ===== PAINEL DE ERROS =====
        errors_frame = ctk.CTkFrame(results_container, fg_color=COLORS["bg_card"], corner_radius=15)
        errors_frame.grid(row=0, column=1, sticky="nsew", padx=(10, 0))

        errors_header = ctk.CTkFrame(errors_frame, fg_color=COLORS["warning"], corner_radius=10, height=50)
        errors_header.pack(fill="x", padx=15, pady=15)
        errors_header.pack_propagate(False)

        errors_title = ctk.CTkLabel(
            errors_header,
            text="⚠️ ARQUIVOS NÃO ACESSADOS",
            font=ctk.CTkFont(size=18, weight="bold"),
            text_color="white"
        )
        errors_title.pack(pady=10)

        self.errors_textbox = ctk.CTkTextbox(
            errors_frame,
            font=ctk.CTkFont(size=12),
            fg_color=COLORS["bg_input"],
            text_color=COLORS["fg_secondary"],
            wrap="word",
            corner_radius=10
        )
        self.errors_textbox.pack(fill="both", expand=True, padx=15, pady=(0, 15))

        # ===== STATUS BAR =====
        status_frame = ctk.CTkFrame(self.main_frame, fg_color=COLORS["bg_header"], height=50)
        status_frame.pack(fill="x", padx=0, pady=0)
        status_frame.pack_propagate(False)

        self.status_label = ctk.CTkLabel(
            status_frame,
            text="✅ Sistema pronto para busca...",
            font=ctk.CTkFont(size=13),
            text_color=COLORS["fg_header"]
        )
        self.status_label.pack(side="left", padx=30, pady=15)

    def show_config_dialog(self):
        """Mostra janela de configuração da API Key"""
        # Criar janela modal
        config_window = ctk.CTkToplevel(self)
        config_window.title("⚙️ Configurações")
        config_window.geometry("500x300")
        config_window.resizable(False, False)

        # Centralizar janela
        config_window.transient(self)
        config_window.grab_set()

        # Frame principal
        main_config_frame = ctk.CTkFrame(config_window, fg_color=COLORS["bg_main"])
        main_config_frame.pack(fill="both", expand=True, padx=0, pady=0)

        # Título
        title_config = ctk.CTkLabel(
            main_config_frame,
            text="🔑 Configuração da API Key",
            font=ctk.CTkFont(size=20, weight="bold"),
            text_color=COLORS["fg_primary"]
        )
        title_config.pack(pady=20)

        # Frame do conteúdo
        content_frame = ctk.CTkFrame(main_config_frame, fg_color=COLORS["bg_card"], corner_radius=15)
        content_frame.pack(fill="both", expand=True, padx=30, pady=(10, 20))

        # Instrução
        instruction_label = ctk.CTkLabel(
            content_frame,
            text="Insira sua API Key do Google Gemini:",
            font=ctk.CTkFont(size=14),
            text_color=COLORS["fg_primary"]
        )
        instruction_label.pack(pady=(20, 10), padx=20)

        # Campo de entrada
        api_key_entry = ctk.CTkEntry(
            content_frame,
            placeholder_text="Cole sua API Key aqui...",
            width=400,
            height=40,
            font=ctk.CTkFont(size=13),
            fg_color=COLORS["bg_input"],
            text_color=COLORS["fg_primary"],
            border_width=2,
            border_color=COLORS["border"]
        )
        api_key_entry.pack(pady=10, padx=20)

        # Carregar API Key salva
        saved_key = self.config_manager.get_api_key()
        if saved_key:
            api_key_entry.insert(0, saved_key)

        # Link para obter API Key
        link_label = ctk.CTkLabel(
            content_frame,
            text="🔗 Obter API Key: https://makersuite.google.com/app/apikey",
            font=ctk.CTkFont(size=11),
            text_color=COLORS["info"],
            cursor="hand2"
        )
        link_label.pack(pady=(5, 15), padx=20)

        # Função para salvar
        def save_and_close():
            api_key = api_key_entry.get().strip()
            if not api_key:
                messagebox.showwarning("Atenção", "Por favor, insira a API Key.")
                return

            if self.config_manager.set_api_key(api_key):
                messagebox.showinfo("Sucesso", "API Key salva com sucesso!")
                self.status_label.configure(text="✅ API Key configurada e salva")
                config_window.destroy()
            else:
                messagebox.showerror("Erro", "Erro ao salvar API Key")

        # Botões
        buttons_frame = ctk.CTkFrame(content_frame, fg_color="transparent")
        buttons_frame.pack(pady=(10, 20))

        save_btn = ctk.CTkButton(
            buttons_frame,
            text="💾 Salvar",
            width=150,
            height=40,
            command=save_and_close,
            font=ctk.CTkFont(size=14, weight="bold"),
            fg_color=COLORS["success"],
            hover_color="#229954",
            text_color="white"
        )
        save_btn.pack(side="left", padx=10)

        cancel_btn = ctk.CTkButton(
            buttons_frame,
            text="✕ Cancelar",
            width=150,
            height=40,
            command=config_window.destroy,
            font=ctk.CTkFont(size=14),
            fg_color=COLORS["error"],
            hover_color="#C0392B",
            text_color="white"
        )
        cancel_btn.pack(side="left", padx=10)

    def start_search(self):
        """Inicia o processo de busca"""
        if self.search_in_progress:
            messagebox.showwarning("Atenção", "Uma busca já está em andamento!")
            return

        search_term = self.search_entry.get().strip()
        if not search_term:
            messagebox.showwarning("Atenção", "Por favor, digite um nome ou empresa para buscar.")
            return

        # IA sempre ativa
        use_ai = True
        api_key = self.config_manager.get_api_key()

        if not api_key:
            messagebox.showwarning(
                "Configuração Necessária",
                "Por favor, configure a API Key do Gemini clicando no botão ⚙️ no canto superior direito."
            )
            return

        # Limpar resultados anteriores
        self.results_textbox.delete("1.0", "end")
        self.errors_textbox.delete("1.0", "end")

        # Atualizar status
        self.search_in_progress = True
        self.start_time = time.time()
        self.status_label.configure(text="🔄 Buscando arquivos com IA...")

        # Mostrar janela de progresso
        self.show_progress_window()

        # Iniciar busca em thread separada
        search_thread = threading.Thread(
            target=self.search_engine.search,
            args=(search_term, use_ai, api_key),
            daemon=True
        )
        search_thread.start()

    def show_progress_window(self):
        """Mostra janela de progresso modal"""
        self.progress_window = ctk.CTkToplevel(self)
        self.progress_window.title("Processando...")
        self.progress_window.geometry("450x300")
        self.progress_window.resizable(False, False)

        # Centralizar janela
        self.progress_window.transient(self)
        self.progress_window.grab_set()

        # Frame principal
        progress_main = ctk.CTkFrame(self.progress_window, fg_color=COLORS["bg_card"])
        progress_main.pack(fill="both", expand=True)

        # Ícone animado
        self.progress_icon = ctk.CTkLabel(
            progress_main,
            text="⏳",
            font=ctk.CTkFont(size=60)
        )
        self.progress_icon.pack(pady=(30, 10))

        # Mensagem
        self.progress_message = ctk.CTkLabel(
            progress_main,
            text="Processando arquivos com IA...",
            font=ctk.CTkFont(size=18, weight="bold"),
            text_color=COLORS["fg_primary"]
        )
        self.progress_message.pack(pady=10)

        # Contador de arquivos
        self.progress_count = ctk.CTkLabel(
            progress_main,
            text="Arquivos encontrados: 0",
            font=ctk.CTkFont(size=15),
            text_color=COLORS["fg_secondary"]
        )
        self.progress_count.pack(pady=5)

        # Tempo decorrido
        self.progress_time = ctk.CTkLabel(
            progress_main,
            text="Tempo: 00:00",
            font=ctk.CTkFont(size=15),
            text_color=COLORS["fg_secondary"]
        )
        self.progress_time.pack(pady=5)

        # Barra de progresso indeterminada
        self.progress_bar = ctk.CTkProgressBar(
            progress_main,
            mode="indeterminate",
            width=380,
            height=10,
            progress_color=COLORS["accent"]
        )
        self.progress_bar.pack(pady=(20, 30))
        self.progress_bar.start()

    def check_progress(self):
        """Verifica o progresso da busca periodicamente"""
        try:
            while True:
                msg_type, data, extra = self.search_engine.progress_queue.get_nowait()

                if msg_type == 'complete':
                    self.finish_search()
                    break
                elif msg_type == 'progress':
                    # Atualizar contadores no progress window
                    if self.progress_window and self.progress_window.winfo_exists():
                        elapsed = time.time() - self.start_time
                        minutes = int(elapsed // 60)
                        seconds = int(elapsed % 60)

                        self.progress_count.configure(
                            text=f"Arquivos encontrados: {len(self.search_engine.files_found)}"
                        )
                        self.progress_time.configure(
                            text=f"Tempo: {minutes:02d}:{seconds:02d}"
                        )
        except queue.Empty:
            pass

        # Continuar verificando
        if self.search_in_progress:
            self.after(100, self.check_progress)

    def finish_search(self):
        """Finaliza a busca e exibe resultados"""
        self.search_in_progress = False

        # Fechar janela de progresso
        if self.progress_window and self.progress_window.winfo_exists():
            self.progress_bar.stop()
            self.progress_window.destroy()

        # Calcular tempo total
        elapsed = time.time() - self.start_time
        minutes = int(elapsed // 60)
        seconds = int(elapsed % 60)

        # Exibir resultados
        num_found = len(self.search_engine.files_found)
        num_errors = len(self.search_engine.files_with_errors)

        if num_found > 0:
            self.results_textbox.insert("end", f"✅ {num_found} arquivo(s) encontrado(s)!\n\n", "success")

            for i, (file_path, context) in enumerate(self.search_engine.files_found, 1):
                filename = os.path.basename(file_path)
                self.results_textbox.insert("end", f"{i}. 📄 {filename}\n", "filename")
                self.results_textbox.insert("end", f"   📁 {file_path}\n", "path")
                if context:
                    self.results_textbox.insert("end", f"   💬 {context}\n", "context")
                self.results_textbox.insert("end", "\n")

            self.results_textbox.insert("end", "\n💡 Dica: Clique duas vezes em um arquivo para abrir\n", "tip")
            self.status_label.configure(
                text=f"✅ Busca concluída! {num_found} arquivo(s) encontrado(s) em {minutes:02d}:{seconds:02d}"
            )
        else:
            self.results_textbox.insert("end", "❌ Nenhum arquivo encontrado.\n\n")
            self.status_label.configure(text=f"⚠️ Nenhum resultado encontrado em {minutes:02d}:{seconds:02d}")

        # Exibir erros
        if num_errors > 0:
            self.errors_textbox.insert("end", f"⚠️ {num_errors} arquivo(s) não acessado(s):\n\n")

            for i, (file_path, error_msg) in enumerate(self.search_engine.files_with_errors, 1):
                filename = os.path.basename(file_path)
                self.errors_textbox.insert("end", f"{i}. 🔒 {filename}\n")
                self.errors_textbox.insert("end", f"   Motivo: {error_msg}\n\n")
        else:
            self.errors_textbox.insert("end", "✅ Todos os arquivos foram acessados com sucesso!")

        # Mostrar popup de conclusão
        messagebox.showinfo(
            "Busca Concluída",
            f"Busca finalizada em {minutes:02d}:{seconds:02d}\n\n"
            f"✅ Encontrados: {num_found}\n"
            f"⚠️ Erros: {num_errors}"
        )

    def open_file_from_selection(self, event):
        """Abre arquivo ao clicar duas vezes"""
        try:
            # Pegar a linha clicada
            index = self.results_textbox.index("@%s,%s" % (event.x, event.y))
            line_start = index.split('.')[0]
            line_content = self.results_textbox.get(f"{line_start}.0", f"{line_start}.end")

            # Buscar o arquivo correspondente
            for file_path, _ in self.search_engine.files_found:
                if file_path in line_content or os.path.basename(file_path) in line_content:
                    self.open_file(file_path)
                    break
        except Exception as e:
            print(f"Erro ao abrir arquivo: {e}")

    def open_file(self, file_path: str):
        """Abre um arquivo Word"""
        try:
            if os.path.exists(file_path):
                if sys.platform == "win32":
                    os.startfile(file_path)
                elif sys.platform == "darwin":  # macOS
                    os.system(f'open "{file_path}"')
                else:  # Linux
                    os.system(f'xdg-open "{file_path}"')

                self.status_label.configure(text=f"📄 Arquivo aberto: {os.path.basename(file_path)}")
            else:
                messagebox.showerror("Erro", "Arquivo não encontrado!")
        except Exception as e:
            messagebox.showerror("Erro", f"Não foi possível abrir o arquivo:\n{str(e)}")


# ===========================
# FUNÇÃO PRINCIPAL
# ===========================
def main():
    """Função principal"""
    print("=" * 60)
    print("🔍 Sistema de Busca Avançada em Arquivos Word com IA")
    print("=" * 60)
    print(f"📁 Pasta Base: {PASTA_BASE}")
    print(f"🧵 Threads: {NUM_THREADS}")
    print(f"📝 Extensões: {', '.join(EXTENSIONS)}")
    print(f"🤖 IA: Sempre ativa (Google Gemini)")
    print("=" * 60)
    print()

    # Verificar se a pasta existe
    if not os.path.exists(PASTA_BASE):
        print(f"⚠️ AVISO: Pasta base não encontrada: {PASTA_BASE}")
        print("O programa continuará, mas você pode precisar ajustar o caminho.")
        print()

    # Iniciar aplicação
    app = SearchApp()
    app.mainloop()


if __name__ == "__main__":
    main()

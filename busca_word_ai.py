#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Sistema de Busca Avançada em Arquivos Word com IA
Utiliza Google Gemini para identificação inteligente de nomes e empresas
Autor: Gerado automaticamente
Versão: 1.0
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

# Cores (Tema Neutro Moderno)
COLORS = {
    "bg_dark": "#1a1a1a",
    "bg_medium": "#2d2d2d",
    "bg_light": "#3a3a3a",
    "fg_light": "#f5f5f5",
    "fg_medium": "#cccccc",
    "fg_dark": "#999999",
    "accent": "#4a4a4a",
    "success": "#4CAF50",
    "error": "#f44336",
    "warning": "#ff9800",
    "info": "#2196F3"
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

    def search_in_document(self, file_path: str, search_term: str, use_ai: bool = False,
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
        Modelo: gemini-2.0-flash-lite
        """
        try:
            genai.configure(api_key=api_key)
            model = genai.GenerativeModel('gemini-2.0-flash-exp')  # Usando o modelo especificado

            prompt = f"""
            Analise o seguinte texto e determine se há menção à pessoa ou empresa: "{search_term}"

            Considere:
            - Variações de nome (abreviações, apelidos)
            - Menções indiretas
            - Contexto de negócios/jurídico

            Responda APENAS "SIM" ou "NÃO".

            Texto:
            {text[:5000]}  # Limita a 5000 caracteres para economia
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

    def search(self, search_term: str, use_ai: bool = False, api_key: str = None) -> bool:
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

        # Configurar tema
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("dark-blue")

        # Gerenciadores
        self.config_manager = ConfigManager()
        self.search_engine = WordSearchEngine(PASTA_BASE, NUM_THREADS)

        # Variáveis
        self.search_in_progress = False
        self.start_time = None
        self.progress_window = None

        # Construir interface
        self.build_ui()

        # Carregar API Key salva
        saved_api_key = self.config_manager.get_api_key()
        if saved_api_key:
            self.api_key_entry.insert(0, saved_api_key)
            self.save_key_var.set(True)

        # Iniciar monitoramento de progresso
        self.after(100, self.check_progress)

    def build_ui(self):
        """Constrói a interface gráfica"""

        # ===== FRAME PRINCIPAL =====
        self.main_frame = ctk.CTkFrame(self, fg_color=COLORS["bg_dark"])
        self.main_frame.pack(fill="both", expand=True, padx=0, pady=0)

        # ===== TÍTULO =====
        title_frame = ctk.CTkFrame(self.main_frame, fg_color=COLORS["bg_light"], height=80)
        title_frame.pack(fill="x", padx=0, pady=0)
        title_frame.pack_propagate(False)

        title_label = ctk.CTkLabel(
            title_frame,
            text="🔍 BUSCA AVANÇADA EM TRANSCRIÇÕES",
            font=ctk.CTkFont(size=28, weight="bold"),
            text_color=COLORS["fg_light"]
        )
        title_label.pack(pady=20)

        # ===== CONFIGURAÇÃO API =====
        api_frame = ctk.CTkFrame(self.main_frame, fg_color=COLORS["bg_medium"])
        api_frame.pack(fill="x", padx=20, pady=(20, 10))

        api_label = ctk.CTkLabel(
            api_frame,
            text="🔑 API Key Google Gemini:",
            font=ctk.CTkFont(size=14, weight="bold"),
            text_color=COLORS["fg_light"]
        )
        api_label.pack(side="left", padx=10, pady=15)

        self.api_key_entry = ctk.CTkEntry(
            api_frame,
            placeholder_text="Insira sua API Key do Google Gemini",
            width=400,
            height=35,
            font=ctk.CTkFont(size=12),
            show="*"
        )
        self.api_key_entry.pack(side="left", padx=10, pady=15)

        self.save_key_var = ctk.BooleanVar(value=False)
        save_key_check = ctk.CTkCheckBox(
            api_frame,
            text="Salvar Key",
            variable=self.save_key_var,
            font=ctk.CTkFont(size=12),
            text_color=COLORS["fg_medium"]
        )
        save_key_check.pack(side="left", padx=10, pady=15)

        save_btn = ctk.CTkButton(
            api_frame,
            text="💾 Salvar",
            width=100,
            height=35,
            command=self.save_api_key,
            fg_color=COLORS["accent"],
            hover_color=COLORS["bg_light"]
        )
        save_btn.pack(side="left", padx=5, pady=15)

        # ===== BUSCA =====
        search_frame = ctk.CTkFrame(self.main_frame, fg_color=COLORS["bg_medium"])
        search_frame.pack(fill="x", padx=20, pady=10)

        search_label = ctk.CTkLabel(
            search_frame,
            text="👤 Nome ou Empresa:",
            font=ctk.CTkFont(size=14, weight="bold"),
            text_color=COLORS["fg_light"]
        )
        search_label.pack(side="left", padx=10, pady=15)

        self.search_entry = ctk.CTkEntry(
            search_frame,
            placeholder_text="Digite o nome da pessoa ou empresa para buscar...",
            width=500,
            height=40,
            font=ctk.CTkFont(size=14)
        )
        self.search_entry.pack(side="left", padx=10, pady=15)
        self.search_entry.bind('<Return>', lambda e: self.start_search())

        self.use_ai_var = ctk.BooleanVar(value=False)
        use_ai_check = ctk.CTkCheckBox(
            search_frame,
            text="🤖 Usar IA (Gemini)",
            variable=self.use_ai_var,
            font=ctk.CTkFont(size=12),
            text_color=COLORS["fg_medium"]
        )
        use_ai_check.pack(side="left", padx=10, pady=15)

        search_btn = ctk.CTkButton(
            search_frame,
            text="🔍 BUSCAR",
            width=150,
            height=40,
            command=self.start_search,
            font=ctk.CTkFont(size=14, weight="bold"),
            fg_color=COLORS["info"],
            hover_color="#1976D2"
        )
        search_btn.pack(side="left", padx=10, pady=15)

        # ===== INFO PASTA =====
        info_frame = ctk.CTkFrame(self.main_frame, fg_color=COLORS["bg_dark"])
        info_frame.pack(fill="x", padx=20, pady=5)

        info_label = ctk.CTkLabel(
            info_frame,
            text=f"📁 Pasta Base: {PASTA_BASE}",
            font=ctk.CTkFont(size=11),
            text_color=COLORS["fg_dark"]
        )
        info_label.pack(pady=5)

        # ===== CONTAINER DE RESULTADOS =====
        results_container = ctk.CTkFrame(self.main_frame, fg_color=COLORS["bg_dark"])
        results_container.pack(fill="both", expand=True, padx=20, pady=10)

        # Configurar grid
        results_container.grid_columnconfigure(0, weight=1)
        results_container.grid_columnconfigure(1, weight=1)
        results_container.grid_rowconfigure(0, weight=1)

        # ===== PAINEL DE RESULTADOS =====
        results_frame = ctk.CTkFrame(results_container, fg_color=COLORS["bg_medium"])
        results_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 5))

        results_title = ctk.CTkLabel(
            results_frame,
            text="📋 RESULTADOS",
            font=ctk.CTkFont(size=16, weight="bold"),
            text_color=COLORS["fg_light"]
        )
        results_title.pack(pady=10)

        # Lista de resultados com scrollbar
        self.results_textbox = ctk.CTkTextbox(
            results_frame,
            font=ctk.CTkFont(size=12),
            fg_color=COLORS["bg_dark"],
            text_color=COLORS["fg_light"],
            wrap="word"
        )
        self.results_textbox.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        self.results_textbox.bind('<Double-Button-1>', self.open_file_from_selection)

        # ===== PAINEL DE ERROS =====
        errors_frame = ctk.CTkFrame(results_container, fg_color=COLORS["bg_medium"])
        errors_frame.grid(row=0, column=1, sticky="nsew", padx=(5, 0))

        errors_title = ctk.CTkLabel(
            errors_frame,
            text="⚠️ ARQUIVOS NÃO ACESSADOS",
            font=ctk.CTkFont(size=16, weight="bold"),
            text_color=COLORS["warning"]
        )
        errors_title.pack(pady=10)

        self.errors_textbox = ctk.CTkTextbox(
            errors_frame,
            font=ctk.CTkFont(size=11),
            fg_color=COLORS["bg_dark"],
            text_color=COLORS["fg_medium"],
            wrap="word"
        )
        self.errors_textbox.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        # ===== STATUS BAR =====
        status_frame = ctk.CTkFrame(self.main_frame, fg_color=COLORS["bg_light"], height=40)
        status_frame.pack(fill="x", padx=0, pady=0)
        status_frame.pack_propagate(False)

        self.status_label = ctk.CTkLabel(
            status_frame,
            text="✅ Sistema pronto para busca...",
            font=ctk.CTkFont(size=12),
            text_color=COLORS["fg_medium"]
        )
        self.status_label.pack(side="left", padx=20, pady=10)

    def save_api_key(self):
        """Salva a API Key se o checkbox estiver marcado"""
        api_key = self.api_key_entry.get().strip()

        if not api_key:
            messagebox.showwarning("Atenção", "Por favor, insira a API Key antes de salvar.")
            return

        if self.save_key_var.get():
            if self.config_manager.set_api_key(api_key):
                messagebox.showinfo("Sucesso", "API Key salva com sucesso!")
                self.status_label.configure(text="✅ API Key salva com sucesso")
            else:
                messagebox.showerror("Erro", "Erro ao salvar API Key")
        else:
            # Remover API Key salva
            self.config_manager.set_api_key("")
            messagebox.showinfo("Info", "API Key removida da configuração")
            self.status_label.configure(text="ℹ️ API Key não está mais salva")

    def start_search(self):
        """Inicia o processo de busca"""
        if self.search_in_progress:
            messagebox.showwarning("Atenção", "Uma busca já está em andamento!")
            return

        search_term = self.search_entry.get().strip()
        if not search_term:
            messagebox.showwarning("Atenção", "Por favor, digite um nome ou empresa para buscar.")
            return

        use_ai = self.use_ai_var.get()
        api_key = self.api_key_entry.get().strip()

        if use_ai and not api_key:
            messagebox.showwarning("Atenção", "Por favor, insira a API Key do Gemini para usar IA.")
            return

        # Limpar resultados anteriores
        self.results_textbox.delete("1.0", "end")
        self.errors_textbox.delete("1.0", "end")

        # Atualizar status
        self.search_in_progress = True
        self.start_time = time.time()
        self.status_label.configure(text="🔄 Buscando arquivos...")

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
        self.progress_window.geometry("400x250")
        self.progress_window.resizable(False, False)

        # Centralizar janela
        self.progress_window.transient(self)
        self.progress_window.grab_set()

        # Ícone animado
        self.progress_icon = ctk.CTkLabel(
            self.progress_window,
            text="⏳",
            font=ctk.CTkFont(size=50)
        )
        self.progress_icon.pack(pady=20)

        # Mensagem
        self.progress_message = ctk.CTkLabel(
            self.progress_window,
            text="Processando arquivos...",
            font=ctk.CTkFont(size=16, weight="bold")
        )
        self.progress_message.pack(pady=10)

        # Contador de arquivos
        self.progress_count = ctk.CTkLabel(
            self.progress_window,
            text="Arquivos encontrados: 0",
            font=ctk.CTkFont(size=14)
        )
        self.progress_count.pack(pady=5)

        # Tempo decorrido
        self.progress_time = ctk.CTkLabel(
            self.progress_window,
            text="Tempo: 00:00",
            font=ctk.CTkFont(size=14)
        )
        self.progress_time.pack(pady=5)

        # Barra de progresso indeterminada
        self.progress_bar = ctk.CTkProgressBar(
            self.progress_window,
            mode="indeterminate",
            width=350
        )
        self.progress_bar.pack(pady=20)
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

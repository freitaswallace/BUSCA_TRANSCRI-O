# 🔍 Sistema de Busca Avançada em Arquivos Word com IA

Sistema inteligente de busca local para arquivos Word (.docx, .doc) com integração ao Google Gemini AI para identificação avançada de nomes e empresas.

## 📋 Características

### ✨ Funcionalidades Principais

- **Busca Textual Avançada**: Localiza menções a pessoas ou empresas em documentos Word
- **Priorização Inteligente**: Dá destaque a termos em **negrito** e <u>sublinhado</u>
- **Integração com IA**: Usa Google Gemini 2.0 Flash para identificação contextual
- **Busca Paralela**: 10 threads simultâneas para máxima performance
- **Interface Moderna**: GUI limpa com tema escuro e cores neutras
- **Tratamento de Erros**: Gerencia arquivos bloqueados sem interromper a busca
- **Feedback em Tempo Real**: Pop-up de progresso com contadores e tempo decorrido
- **Abertura Rápida**: Duplo clique para abrir arquivos diretamente

### 🎨 Interface

- **Cores Neutras**: Tema escuro moderno (#1a1a1a, #2d2d2d, #3a3a3a)
- **CustomTkinter**: Interface gráfica moderna e responsiva
- **Painéis Divididos**: Resultados e erros exibidos separadamente
- **Status Bar**: Feedback constante sobre o estado da busca

## 🚀 Instalação

### Pré-requisitos

- Python 3.8 ou superior
- Acesso à rede onde estão os arquivos (\\192.168.20.100\trabalho\Transcrições)
- API Key do Google Gemini (opcional, para busca com IA)

### Passos de Instalação

1. **Clone ou baixe o repositório**

```bash
git clone <repositório>
cd BUSCA_TRANSCRI-O
```

2. **Instale as dependências**

```bash
pip install -r requirements.txt
```

3. **Configure a API Key (opcional)**

Para usar a busca com IA, você precisa de uma API Key do Google Gemini:
- Acesse: https://makersuite.google.com/app/apikey
- Crie uma nova API Key
- Cole a chave na interface do programa

## 📖 Como Usar

### Execução

```bash
python busca_word_ai.py
```

### Passo a Passo

1. **Configure a API Key** (primeira vez)
   - Cole sua API Key do Google Gemini no campo "🔑 API Key"
   - Marque "Salvar Key" para não precisar digitar novamente
   - Clique em "💾 Salvar"

2. **Digite o Nome ou Empresa**
   - No campo "👤 Nome ou Empresa", digite o termo de busca
   - Exemplo: "João Silva", "Empresa XYZ", etc.

3. **Escolha o Modo de Busca**
   - ☑️ **Sem IA**: Busca textual rápida (recomendado para nomes exatos)
   - ☑️ **Com IA**: Busca contextual com Google Gemini (para variações e contexto)

4. **Execute a Busca**
   - Clique em "🔍 BUSCAR" ou pressione Enter
   - Aguarde o processamento (progresso exibido em tempo real)

5. **Visualize os Resultados**
   - Painel esquerdo: Arquivos encontrados
   - Painel direito: Arquivos não acessados (bloqueados/com erro)
   - **Duplo clique** em um arquivo para abri-lo

## ⚙️ Configurações

### Caminho Base

Por padrão, o sistema busca em:
```
\\192.168.20.100\trabalho\Transcrições
```

Para alterar, edite a variável `PASTA_BASE` no arquivo `busca_word_ai.py` (linha 54).

### Número de Threads

Por padrão, o sistema usa **10 threads** paralelas. Para ajustar:

```python
NUM_THREADS = 10  # Altere para o número desejado
```

### Extensões de Arquivo

Por padrão, busca em `.docx` e `.doc`. Para adicionar outras:

```python
EXTENSIONS = ['.docx', '.doc']  # Adicione outras extensões
```

## 🤖 Sobre a IA

### Modelo Utilizado

- **Google Gemini 2.0 Flash Exp**
- Modelo leve e rápido para análise de texto
- Identifica variações de nomes, abreviações e menções indiretas

### Quando Usar IA?

✅ **Use IA quando:**
- Buscar variações de nome (ex: "José" vs "Zé")
- Identificar menções indiretas
- Análise contextual de negócios/jurídica

❌ **Não use IA para:**
- Buscas simples de nomes exatos (mais lento)
- Grande volume de documentos (custo de API)

## 📊 Recursos Técnicos

### Threading Pesado

- **10 threads** processam arquivos simultaneamente
- Divisão inteligente de carga entre threads
- Interface não congela durante processamento

### Tratamento de Erros

- **Arquivos Bloqueados**: Sistema pula e registra
- **Erros de Permissão**: Não interrompem a busca
- **Relatório Completo**: Lista todos os erros ao final

### Priorização de Formatação

O sistema dá **prioridade máxima** para:
1. ✅ Textos em **negrito** + <u>sublinhado</u>
2. ✅ Textos em **negrito**
3. ✅ Textos em <u>sublinhado</u>
4. ✅ Texto normal

## 🔒 Segurança

### API Key

- Armazenada localmente em `config.json`
- Não é compartilhada ou enviada para servidores externos
- Use a opção "Salvar Key" apenas em computadores pessoais

### Privacidade

- Todo processamento é local
- IA (quando ativada) envia apenas trechos do texto para análise
- Limite de 5000 caracteres por requisição

## 🐛 Resolução de Problemas

### Erro: "Pasta base não encontrada"

- Verifique se tem acesso à rede: `\\192.168.20.100`
- Confirme que a pasta existe: `\trabalho\Transcrições`
- Em Linux/Mac, monte o compartilhamento de rede

### Erro: "python-docx não instalado"

```bash
pip install python-docx
```

### Erro: "google-generativeai não instalado"

```bash
pip install google-generativeai
```

### Interface não abre

```bash
# Reinstale customtkinter
pip uninstall customtkinter
pip install customtkinter==5.2.2
```

## 📝 Estrutura de Arquivos

```
BUSCA_TRANSCRI-O/
│
├── busca_word_ai.py          # Script principal
├── requirements.txt          # Dependências
├── README_BUSCA.md          # Este arquivo
├── config.json              # Configurações (criado automaticamente)
└── BuscaFichas_V54.ps1      # Script PowerShell original
```

## 🔄 Atualizações Futuras

- [ ] Exportar resultados para Excel
- [ ] Busca por expressões regulares
- [ ] Filtros avançados (data, tamanho, etc.)
- [ ] Histórico de buscas
- [ ] Preview de documentos na interface

## 👨‍💻 Desenvolvimento

### Tecnologias Utilizadas

- **Python 3.8+**
- **CustomTkinter 5.2.2** - Interface gráfica moderna
- **python-docx 1.1.2** - Manipulação de arquivos Word
- **google-generativeai 0.8.3** - Integração com Gemini AI

### Contribuindo

Sugestões e melhorias são bem-vindas!

## 📄 Licença

Este projeto é de uso interno.

---

**Desenvolvido com ❤️ e ☕**

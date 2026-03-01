# ExcelToPDF

Ferramenta simples e intuitiva para converter arquivos **Excel (.xlsx / .xls)** em **PDF** — sem complicações.

---

## 🎯 O que este programa faz

1. Você abre o programa.
2. Uma janela aparece para você **escolher o arquivo Excel**.
3. O programa converte o arquivo para PDF automaticamente.
4. A **pasta com o PDF é aberta na tela** para você encontrar facilmente.

Pronto! Não é necessário digitar nada ou entender programação.

---

## 📋 Pré-requisitos

### Python
O programa precisa do **Python 3.10 ou superior** instalado.

- Baixe em: https://www.python.org/downloads/
- Durante a instalação, marque a opção **"Add Python to PATH"**.

### Programa de conversão (escolha um)

| Opção | Quando usar | Custo |
|---|---|---|
| **Microsoft Excel** | Se você já tem o Excel instalado | Pago (já incluso no Office) |
| **LibreOffice** | Alternativa gratuita | Gratuito |

> Se você **não tem o Excel**, instale o LibreOffice:
> 👉 https://www.libreoffice.org/download/libreoffice/

---

## ⚙️ Instalação

### 1. Baixe os arquivos do projeto

Clique em **Code → Download ZIP** nesta página e extraia para uma pasta de sua preferência.

### 2. Instale as dependências

Abra o **Terminal** (no Windows: Prompt de Comando ou PowerShell) e navegue até a pasta do projeto:

```
cd caminho\para\a\pasta\ExcelToPDF
```

Em seguida, execute:

```
pip install -r requirements.txt
```

> **Somente se você usa o Microsoft Excel no Windows**, instale também:
> ```
> pip install pywin32
> ```

---

## ▶️ Como executar o programa

No terminal, dentro da pasta do projeto, execute:

```
python converter.py
```

O programa abrirá uma janela e guiará você em cada etapa.

---

## 🖥️ Fluxo do programa (passo a passo)

```
┌─────────────────────────────────────┐
│  1. Programa exibe mensagem de      │
│     boas-vindas com instruções      │
└────────────────┬────────────────────┘
                 ↓
┌─────────────────────────────────────┐
│  2. Janela para SELECIONAR o        │
│     arquivo Excel (.xlsx)           │
└────────────────┬────────────────────┘
                 ↓
┌─────────────────────────────────────┐
│  3. Programa detecta automaticamente│
│     Excel ou LibreOffice e converte │
└────────────────┬────────────────────┘
                 ↓
┌─────────────────────────────────────┐
│  4. PDF salvo na mesma pasta do     │
│     arquivo Excel original          │
└────────────────┬────────────────────┘
                 ↓
┌─────────────────────────────────────┐
│  5. Pasta aberta automaticamente +  │
│     mensagem de sucesso exibida     │
└─────────────────────────────────────┘
```

---

## ❗ Tratamento de erros

O programa exibe mensagens amigáveis nos seguintes casos:

| Situação | Mensagem exibida |
|---|---|
| Nenhum arquivo selecionado | Aviso de cancelamento |
| Arquivo não encontrado | Indica o caminho e sugere verificar |
| Extensão inválida (não é .xlsx/.xls) | Explica o problema e pede para selecionar arquivo correto |
| Excel / LibreOffice não instalado | Instrui como instalar |
| Falha durante a conversão | Exibe detalhes do erro |

---

## ⚠️ Limitações conhecidas

- No **Windows**, a conversão via Excel requer que o Microsoft Excel esteja instalado e ativado.
- O LibreOffice pode apresentar pequenas diferenças de formatação em relação ao Excel original (especialmente em arquivos com gráficos ou formatação avançada).
- Em **macOS e Linux**, apenas a conversão via LibreOffice é suportada.
- O programa converte **apenas a primeira planilha ativa** (comportamento padrão do Excel e LibreOffice).

---

## 📁 Estrutura do projeto

```
ExcelToPDF/
├── converter.py       # Código principal do programa
├── requirements.txt   # Dependências do Python
└── README.md          # Este arquivo de instruções
```
"""
ExcelToPDF - Conversor de Excel para PDF
=========================================
Ferramenta simples para converter arquivos .xlsx em PDF.
Funciona com Microsoft Excel instalado OU com LibreOffice (gratuito).

Fluxo do programa:
  1. Uma janela abre para você escolher o arquivo Excel.
  2. O programa converte automaticamente para PDF.
  3. A pasta com o PDF é aberta automaticamente ao final.
"""

import os
import sys
import platform
import subprocess
import shutil


# ---------------------------------------------------------------------------
# ENTRADA DO USUÁRIO
# ---------------------------------------------------------------------------

def selecionar_arquivo() -> str:
    """
    Abre uma janela para o usuário escolher o arquivo Excel (.xlsx).
    Retorna o caminho completo do arquivo selecionado,
    ou uma string vazia se o usuário cancelar.
    """
    import tkinter as tk
    from tkinter import filedialog

    root = tk.Tk()
    root.withdraw()  # Oculta a janela principal do tkinter
    root.attributes("-topmost", True)  # Garante que a janela fique na frente

    arquivo = filedialog.askopenfilename(
        title="Selecione o arquivo Excel",
        filetypes=[("Arquivos Excel", "*.xlsx *.xls"), ("Todos os arquivos", "*.*")],
    )
    root.destroy()
    return arquivo


# ---------------------------------------------------------------------------
# CONVERSÃO: cenário 1 – Microsoft Excel instalado (Windows)
# ---------------------------------------------------------------------------

def converter_com_excel(caminho_excel: str, caminho_pdf: str) -> None:
    """
    Converte o arquivo Excel para PDF usando o Microsoft Excel (Windows).
    Levanta RuntimeError se algo der errado.
    """
    try:
        import win32com.client  # type: ignore[import]
    except ImportError:
        raise RuntimeError(
            "O módulo 'pywin32' não está instalado.\n"
            "Execute: pip install pywin32"
        )

    excel = None
    pasta = None
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        pasta = excel.Workbooks.Open(caminho_excel)
        # xlTypePDF = 0
        pasta.ExportAsFixedFormat(0, caminho_pdf)
    except Exception as erro:
        raise RuntimeError(f"Erro ao converter com o Excel: {erro}") from erro
    finally:
        if pasta:
            try:
                pasta.Close(False)
            except Exception:
                pass
        if excel:
            try:
                excel.Quit()
            except Exception:
                pass


# ---------------------------------------------------------------------------
# CONVERSÃO: cenário 2 – LibreOffice instalado (Windows, macOS, Linux)
# ---------------------------------------------------------------------------

def _caminho_libreoffice() -> str | None:
    """
    Localiza o executável do LibreOffice no sistema.
    Retorna o caminho se encontrado, caso contrário None.
    """
    # Verificação pelo PATH (funciona em Linux/macOS)
    if shutil.which("libreoffice"):
        return "libreoffice"
    if shutil.which("soffice"):
        return "soffice"

    # Caminhos comuns no Windows
    caminhos_windows = [
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    ]
    for caminho in caminhos_windows:
        if os.path.isfile(caminho):
            return caminho

    # Caminhos comuns no macOS
    caminhos_mac = [
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    ]
    for caminho in caminhos_mac:
        if os.path.isfile(caminho):
            return caminho

    return None


def converter_com_libreoffice(caminho_excel: str, caminho_pdf: str) -> None:
    """
    Converte o arquivo Excel para PDF usando o LibreOffice.
    Levanta RuntimeError se algo der errado.
    """
    soffice = _caminho_libreoffice()
    if soffice is None:
        raise RuntimeError(
            "LibreOffice não encontrado.\n"
            "Instale o LibreOffice gratuitamente em: https://www.libreoffice.org/download/libreoffice/"
        )

    pasta_destino = os.path.dirname(caminho_pdf)

    try:
        resultado = subprocess.run(
            [
                soffice,
                "--headless",
                "--convert-to",
                "pdf",
                "--outdir",
                pasta_destino,
                caminho_excel,
            ],
            capture_output=True,
            text=True,
            timeout=120,
        )
    except subprocess.TimeoutExpired:
        raise RuntimeError("A conversão demorou demais e foi cancelada.") from None
    except Exception as erro:
        raise RuntimeError(f"Erro ao executar o LibreOffice: {erro}") from erro

    if resultado.returncode != 0:
        detalhe = resultado.stderr.strip() or resultado.stdout.strip()
        raise RuntimeError(f"O LibreOffice retornou um erro:\n{detalhe}")

    # O LibreOffice cria o PDF com o mesmo nome do Excel, mas extensão .pdf
    nome_base = os.path.splitext(os.path.basename(caminho_excel))[0]
    pdf_gerado = os.path.join(pasta_destino, nome_base + ".pdf")

    # Se o usuário pediu um nome diferente, renomeia
    if os.path.abspath(pdf_gerado) != os.path.abspath(caminho_pdf):
        if os.path.exists(pdf_gerado):
            os.replace(pdf_gerado, caminho_pdf)


# ---------------------------------------------------------------------------
# DETECÇÃO AUTOMÁTICA DE MÉTODO DE CONVERSÃO
# ---------------------------------------------------------------------------

def _excel_disponivel() -> bool:
    """Verifica se o Microsoft Excel está disponível (somente Windows)."""
    if platform.system() != "Windows":
        return False
    try:
        import win32com.client  # type: ignore[import]  # noqa: F401
        import pythoncom  # type: ignore[import]  # noqa: F401

        pythoncom.CoInitialize()
        app = win32com.client.Dispatch("Excel.Application")
        app.Quit()
        return True
    except Exception:
        return False


def converter(caminho_excel: str, caminho_pdf: str) -> str:
    """
    Escolhe automaticamente o método disponível (Excel ou LibreOffice)
    e realiza a conversão.
    Retorna o nome do método utilizado ('Excel' ou 'LibreOffice').
    Levanta RuntimeError em caso de falha.
    """
    if _excel_disponivel():
        converter_com_excel(caminho_excel, caminho_pdf)
        return "Excel"

    if _caminho_libreoffice():
        converter_com_libreoffice(caminho_excel, caminho_pdf)
        return "LibreOffice"

    raise RuntimeError(
        "Nenhum programa de conversão encontrado.\n\n"
        "Você precisa ter instalado:\n"
        "  • Microsoft Excel  OU\n"
        "  • LibreOffice (gratuito em https://www.libreoffice.org/)\n\n"
        "Após instalar um deles, execute o programa novamente."
    )


# ---------------------------------------------------------------------------
# FEEDBACK VISUAL: abre a pasta do PDF gerado
# ---------------------------------------------------------------------------

def abrir_pasta(caminho_pasta: str) -> None:
    """
    Abre o gerenciador de arquivos na pasta indicada.
    Funciona no Windows, macOS e Linux.
    """
    sistema = platform.system()
    try:
        if sistema == "Windows":
            os.startfile(caminho_pasta)  # type: ignore[attr-defined]
        elif sistema == "Darwin":  # macOS
            subprocess.Popen(["open", caminho_pasta])
        else:  # Linux e outros
            subprocess.Popen(["xdg-open", caminho_pasta])
    except Exception:
        pass  # Falhar ao abrir a pasta não deve interromper o programa


# ---------------------------------------------------------------------------
# VALIDAÇÃO DO ARQUIVO
# ---------------------------------------------------------------------------

def validar_arquivo(caminho: str) -> None:
    """
    Verifica se o arquivo existe e tem extensão suportada.
    Levanta ValueError com mensagem amigável em caso de problema.
    """
    if not caminho:
        raise ValueError("Nenhum arquivo foi selecionado.")

    if not os.path.isfile(caminho):
        raise ValueError(
            f"O arquivo abaixo não foi encontrado:\n{caminho}\n\n"
            "Verifique se o arquivo não foi movido ou renomeado."
        )

    extensao = os.path.splitext(caminho)[1].lower()
    if extensao not in (".xlsx", ".xls"):
        raise ValueError(
            f"O arquivo selecionado não é um Excel válido.\n"
            f"Extensão encontrada: '{extensao}'\n\n"
            "Por favor, selecione um arquivo com extensão .xlsx ou .xls."
        )


# ---------------------------------------------------------------------------
# PONTO DE ENTRADA PRINCIPAL
# ---------------------------------------------------------------------------

def main() -> None:
    """
    Fluxo principal do programa:
      1. Selecionar arquivo Excel via janela de escolha.
      2. Validar o arquivo.
      3. Converter para PDF.
      4. Exibir mensagem de sucesso e abrir a pasta.
    """
    import tkinter as tk
    from tkinter import filedialog, messagebox

    # Inicializa o tkinter uma única vez para as caixas de diálogo
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)

    # Passo 1: Selecionar arquivo
    messagebox.showinfo(
        "ExcelToPDF",
        "Bem-vindo ao ExcelToPDF!\n\n"
        "Na próxima tela, selecione o arquivo Excel (.xlsx) que deseja converter para PDF.\n\n"
        "Clique em OK para continuar.",
    )

    caminho_excel = filedialog.askopenfilename(
        title="Selecione o arquivo Excel",
        filetypes=[("Arquivos Excel", "*.xlsx *.xls"), ("Todos os arquivos", "*.*")],
    )

    if not caminho_excel:
        messagebox.showwarning("Operação cancelada", "Nenhum arquivo foi selecionado. O programa será encerrado.")
        root.destroy()
        return

    # Passo 2: Validar arquivo
    try:
        validar_arquivo(caminho_excel)
    except ValueError as erro:
        messagebox.showerror("Arquivo inválido", str(erro))
        root.destroy()
        return

    # Passo 3: Definir caminho do PDF e converter
    pasta_destino = os.path.dirname(os.path.abspath(caminho_excel))
    nome_base = os.path.splitext(os.path.basename(caminho_excel))[0]
    caminho_pdf = os.path.join(pasta_destino, nome_base + ".pdf")

    try:
        metodo = converter(caminho_excel, caminho_pdf)
    except RuntimeError as erro:
        messagebox.showerror("Erro na conversão", str(erro))
        root.destroy()
        return

    # Passo 4: Sucesso – abrir pasta e notificar usuário
    abrir_pasta(pasta_destino)

    messagebox.showinfo(
        "Conversão concluída! ✔",
        f"O PDF foi gerado com sucesso usando {metodo}!\n\n"
        f"Arquivo salvo em:\n{caminho_pdf}\n\n"
        "A pasta foi aberta automaticamente para você.",
    )

    root.destroy()


if __name__ == "__main__":
    main()

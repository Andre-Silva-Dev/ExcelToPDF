"""
Testes unitários para o módulo converter.py
Execução: python -m pytest tests/test_converter.py -v
"""

import os
import sys
import pytest

# Garante que o módulo converter.py seja encontrado
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import converter


# ---------------------------------------------------------------------------
# Testes de validação de arquivo
# ---------------------------------------------------------------------------

class TestValidarArquivo:
    def test_caminho_vazio_levanta_valueerror(self):
        with pytest.raises(ValueError, match="Nenhum arquivo foi selecionado"):
            converter.validar_arquivo("")

    def test_arquivo_inexistente_levanta_valueerror(self, tmp_path):
        caminho_falso = str(tmp_path / "nao_existe.xlsx")
        with pytest.raises(ValueError, match="não foi encontrado"):
            converter.validar_arquivo(caminho_falso)

    def test_extensao_invalida_levanta_valueerror(self, tmp_path):
        arquivo = tmp_path / "documento.pdf"
        arquivo.write_bytes(b"conteudo qualquer")
        with pytest.raises(ValueError, match="não é um Excel válido"):
            converter.validar_arquivo(str(arquivo))

    def test_extensao_xlsx_aceita(self, tmp_path):
        arquivo = tmp_path / "planilha.xlsx"
        arquivo.write_bytes(b"conteudo qualquer")
        # Não deve levantar exceção
        converter.validar_arquivo(str(arquivo))

    def test_extensao_xls_aceita(self, tmp_path):
        arquivo = tmp_path / "planilha.xls"
        arquivo.write_bytes(b"conteudo qualquer")
        # Não deve levantar exceção
        converter.validar_arquivo(str(arquivo))

    def test_extensao_maiuscula_aceita(self, tmp_path):
        arquivo = tmp_path / "planilha.XLSX"
        arquivo.write_bytes(b"conteudo qualquer")
        # Não deve levantar exceção (normalização para minúsculas)
        converter.validar_arquivo(str(arquivo))

    def test_extensao_txt_levanta_valueerror(self, tmp_path):
        arquivo = tmp_path / "arquivo.txt"
        arquivo.write_bytes(b"conteudo qualquer")
        with pytest.raises(ValueError, match="não é um Excel válido"):
            converter.validar_arquivo(str(arquivo))


# ---------------------------------------------------------------------------
# Testes do localizador do LibreOffice
# ---------------------------------------------------------------------------

class TestCaminhoLibreoffice:
    def test_retorna_string_ou_none(self):
        resultado = converter._caminho_libreoffice()
        assert resultado is None or isinstance(resultado, str)

    def test_retorna_executavel_valido_se_encontrado(self):
        resultado = converter._caminho_libreoffice()
        if resultado is not None:
            # Se retornou algo, deve ser um caminho de arquivo ou um comando
            assert len(resultado) > 0


# ---------------------------------------------------------------------------
# Testes de conversão via LibreOffice (ignorado se não estiver instalado)
# ---------------------------------------------------------------------------

class TestConverterComLibreoffice:
    def test_levanta_runtimeerror_se_libreoffice_ausente(self, tmp_path, monkeypatch):
        # Simula ausência do LibreOffice
        monkeypatch.setattr(converter, "_caminho_libreoffice", lambda: None)

        arquivo_excel = tmp_path / "planilha.xlsx"
        arquivo_excel.write_bytes(b"conteudo falso")
        caminho_pdf = str(tmp_path / "planilha.pdf")

        with pytest.raises(RuntimeError, match="LibreOffice não encontrado"):
            converter.converter_com_libreoffice(str(arquivo_excel), caminho_pdf)

    def test_levanta_runtimeerror_em_falha_de_processo(self, tmp_path, monkeypatch):
        # Simula LibreOffice retornando erro
        import subprocess

        monkeypatch.setattr(converter, "_caminho_libreoffice", lambda: "/usr/bin/soffice")

        fake_result = subprocess.CompletedProcess(
            args=[], returncode=1, stdout="", stderr="Erro simulado"
        )
        monkeypatch.setattr(subprocess, "run", lambda *a, **kw: fake_result)

        arquivo_excel = tmp_path / "planilha.xlsx"
        arquivo_excel.write_bytes(b"conteudo falso")
        caminho_pdf = str(tmp_path / "planilha.pdf")

        with pytest.raises(RuntimeError, match="LibreOffice retornou um erro"):
            converter.converter_com_libreoffice(str(arquivo_excel), caminho_pdf)


# ---------------------------------------------------------------------------
# Testes de detecção automática do método de conversão
# ---------------------------------------------------------------------------

class TestConverter:
    def test_levanta_runtimeerror_sem_excel_nem_libreoffice(self, tmp_path, monkeypatch):
        monkeypatch.setattr(converter, "_excel_disponivel", lambda: False)
        monkeypatch.setattr(converter, "_caminho_libreoffice", lambda: None)

        arquivo_excel = tmp_path / "planilha.xlsx"
        arquivo_excel.write_bytes(b"conteudo falso")
        caminho_pdf = str(tmp_path / "planilha.pdf")

        with pytest.raises(RuntimeError, match="Nenhum programa de conversão encontrado"):
            converter.converter(str(arquivo_excel), caminho_pdf)

    def test_usa_libreoffice_quando_excel_indisponivel(self, tmp_path, monkeypatch):
        monkeypatch.setattr(converter, "_excel_disponivel", lambda: False)
        monkeypatch.setattr(converter, "_caminho_libreoffice", lambda: "/usr/bin/soffice")

        chamadas = []

        def libreoffice_fake(caminho_excel, caminho_pdf):
            chamadas.append((caminho_excel, caminho_pdf))

        monkeypatch.setattr(converter, "converter_com_libreoffice", libreoffice_fake)

        arquivo_excel = tmp_path / "planilha.xlsx"
        arquivo_excel.write_bytes(b"conteudo falso")
        caminho_pdf = str(tmp_path / "planilha.pdf")

        resultado = converter.converter(str(arquivo_excel), caminho_pdf)

        assert resultado == "LibreOffice"
        assert len(chamadas) == 1

    def test_usa_excel_quando_disponivel(self, tmp_path, monkeypatch):
        monkeypatch.setattr(converter, "_excel_disponivel", lambda: True)

        chamadas = []

        def excel_fake(caminho_excel, caminho_pdf):
            chamadas.append((caminho_excel, caminho_pdf))

        monkeypatch.setattr(converter, "converter_com_excel", excel_fake)

        arquivo_excel = tmp_path / "planilha.xlsx"
        arquivo_excel.write_bytes(b"conteudo falso")
        caminho_pdf = str(tmp_path / "planilha.pdf")

        resultado = converter.converter(str(arquivo_excel), caminho_pdf)

        assert resultado == "Excel"
        assert len(chamadas) == 1

# -*- coding: utf-8 -*-
"""
App de extração de "linha digitável" e valor de boletos em PDFs (ZIP),
gerando planilhas de dados e de relatório.

Correções/boas práticas:
- Reutilização de um único genai.Client (thread-safe e recomendado).
- Troca do modelo para 'gemini-2.5-flash'.
- PDFs até ~50MB enviados inline; acima disso via Files API (2GB, retenção 48h).
- Chave de API lida de variável de ambiente (GEMINI_API_KEY/GOOGLE_API_KEY).
- Correção de '>=' no regex e pequenos ajustes.

Requisitos:
    pip install google-genai openpyxl eel

Referências:
- Google Gen AI SDK (uso do Client, generate_content): https://googleapis.github.io/python-genai/
- Limites e métodos de arquivo (inline vs Files API, PDF ~50MB; Files API 2GB): 
  https://ai.google.dev/gemini-api/docs/file-input-methods
  https://ai.google.dev/gemini-api/docs/files
"""

import os
import sys
import subprocess
import concurrent.futures
import re
import time
import tkinter as tk
from tkinter import filedialog
from datetime import date, datetime
import zipfile
import eel
import io
import ctypes
import json

# ======================================================
# 1) CONFIGURAÇÕES INICIAIS
# ======================================================

# >>> IMPORTANTE: leia a chave de API do ambiente (boa prática)
API_KEY = os.getenv("AIzaSyDGLXq7jF0EgE-kEZ30TRb8ikQAFfNrjLQ") 

# Modelo recomendado e estável
GEMINI_MODEL = "gemini-2.5-flash"

# ======================================================
# 2) DEPENDÊNCIAS DO EXCEL E GENAI
# ======================================================
try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
except ImportError:
    print("Erro: Instale a biblioteca openpyxl (pip install openpyxl)")
    sys.exit(1)

try:
    from google import genai
    from google.genai import types
except ImportError:
    print("Erro: Instale a biblioteca google-genai (pip install google-genai)")
    sys.exit(1)

# ======================================================
# 2.1) CLIENTE GLOBAL DO SDK (reutilizado em todas as threads)
# ======================================================
CLIENT = genai.Client(api_key=API_KEY)

# ======================================================
# 3) LÓGICA DE LIMPEZA E EXTRAÇÃO (SOMENTE IA)
# ======================================================
def limpar_valor_brasileiro(valor_str):
    if not valor_str or valor_str == "0.00":
        return "0.00"
    limpo = re.sub(r'[^\d\.,]', '', str(valor_str))
    if '.' in limpo and ',' in limpo:
        limpo = limpo.replace('.', '')
    limpo = limpo.replace(',', '.')
    return limpo

def processar_com_gemini_vision(pdf_bytes: bytes):
    """
    - Até ~50 MB (PDF): envia inline (Part.from_bytes).
    - Acima disso: usa Files API (upload) e referencia o arquivo.
    Retorna (linha_digitavel, valor) ou ("", "Erro ...").
    """
    try:
        prompt = (
            "Analise este boleto bancário. Extraia os dados abaixo:\n"
            "1. LINHA DIGITÁVEL: Código de barras numérico.\n"
            "   - Boletos bancários padrão: EXATAMENTE 47 DÍGITOS.\n"
            "   - Contas de consumo/impostos: EXATAMENTE 48 DÍGITOS.\n"
            "   - Não invente nem adicione zeros extras!\n"
            "2. VALOR DO DOCUMENTO: Valor total.\n"
            "Responda EXATAMENTE neste formato:\n"
            "LINHA: [numeros]\n"
            "VALOR: [valor]\n"
        )

        # Limites: PDF inline ~50 MB; acima disso use Files API (2GB/arquivo; 48h de retenção)
        PDF_INLINE_MAX = 50 * 1024 * 1024  # ~50 MB

        if len(pdf_bytes) <= PDF_INLINE_MAX:
            contents = [
                types.Part.from_bytes(data=pdf_bytes, mime_type="application/pdf"),
                prompt,
            ]
        else:
            uploaded = CLIENT.files.upload(
                file=io.BytesIO(pdf_bytes),
                config=types.UploadFileConfig(
                    mime_type="application/pdf",
                    display_name="boleto.pdf"
                ),
            )
            contents = [uploaded, prompt]

        config = types.GenerateContentConfig(temperature=0.0)
        response = CLIENT.models.generate_content(
            model=GEMINI_MODEL,
            contents=contents,
            config=config
        )

        linha, valor = "", "0.00"
        if response and getattr(response, "text", None):
            texto_resp = response.text

            # Extrai LINHA (permite dígitos, pontos e espaços; depois limpa)
            match_linha = re.search(r'LINHA:\s*([\d\.\s]+)', texto_resp, flags=re.IGNORECASE)
            if match_linha:
                l = re.sub(r'\D', '', match_linha.group(1))
                if len(l) >= 44:  # (corrigido: '>=' em vez de '&gt;=')
                    linha = l

            # Extrai VALOR
            match_valor = re.search(r'VALOR:\s*(.+)', texto_resp, flags=re.IGNORECASE)
            if match_valor:
                valor = limpar_valor_brasileiro(match_valor.group(1).strip())

            # Fallback: caça 47 ou 48 dígitos "puros" em todo o texto
            if not linha:
                texto_limpo = re.sub(r'[\.\-\s]', '', texto_resp)
                bruto = re.search(r'\b\d{47}\b|\b\d{48}\b', texto_limpo)
                if bruto:
                    linha = bruto.group(0)

        return linha, valor

    except Exception as e:
        msg = str(e)
        if "429" in msg:
            return "", "Erro: Cota Excedida"
        return "", f"Erro IA: {type(e).__name__}"

class DummyStop:
    def is_set(self): return False

def processar_pdf(nome_arquivo, conteudo_bytes, stop_event):
    if stop_event.is_set(): return None
    agencia, conta = "0050", "569539-7"

    try:
        linha_final, valor_final = processar_com_gemini_vision(conteudo_bytes)

        if isinstance(valor_final, str) and valor_final.startswith("Erro"):
            status = valor_final
        elif not linha_final and (not valor_final or valor_final == "0.00"):
            status = "FALHA CRÍTICA: IA não encontrou os dados"
        elif not linha_final:
            status = "AVISO (IA encontrou o valor, mas não a linha)"
        elif valor_final == "0.00":
            status = "AVISO (IA encontrou a linha, mas não o valor)"
        else:
            status = "SUCESSO (Extraído via IA)"

        return {"Arquivo": nome_arquivo, "Linha": linha_final, "Valor": valor_final, "Status": status, "Agencia": agencia, "Conta": conta}

    except Exception as e:
        return {"Arquivo": nome_arquivo, "Linha": "", "Valor": "0.00", "Status": f"Erro Crítico: {str(e)}", "Agencia": agencia, "Conta": conta}

# ======================================================
# 4) GERADOR DE EXCEL DUPLO (DADOS e RELATÓRIO)
# ======================================================
def gerar_workbook_padrao_dados():
    wb = Workbook()
    ws = wb.active
    ws.title = "Boletos"
    headers = ["Código de Barras", "Valor", "Data de Pagamento\n(dd/mm/aaaa)", "Identificação Interna\n(Opcional)", "Agência de Origem", "Conta de Origem"]
    fill_header = PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type="solid")
    font_header = Font(name='Calibri', size=11, bold=True)
    alignment_header = Alignment(horizontal="center", vertical="center", wrap_text=True)
    border_all = Border(top=Side(style='thin'), left=Side(style='thin'), right=Side(style='thin'), bottom=Side(style='thin'))
    ws.append(headers)
    for col_num, cell in enumerate(ws[1], 1):
        cell.fill = fill_header; cell.font = font_header; cell.alignment = alignment_header; cell.border = border_all
        ws.column_dimensions[ws.cell(row=1, column=col_num).column_letter].width = 20
    ws.column_dimensions['A'].width = 55
    ws.column_dimensions['D'].width = 30
    ws.freeze_panes = "A2"
    return wb

def gerar_excel_duplo(dados, pasta, modelo_path):
    data_str = datetime.now().strftime('%d%m%Y_%H%M')

    path_dados = os.path.join(pasta, f"planilha boletos BTG {data_str}.xlsx")
    path_relatorio = os.path.join(pasta, f"Relatorio_Status_{data_str}.xlsx")

    if modelo_path and os.path.exists(modelo_path):
        wb_dados = load_workbook(modelo_path)
    else:
        wb_dados = gerar_workbook_padrao_dados()

    ws_dados = wb_dados.active
    ws_dados.protection.sheet = False

    idx_d = 2
    while ws_dados.cell(row=idx_d, column=1).value: idx_d += 1

    wb_rel = Workbook()
    ws_rel = wb_rel.active
    ws_rel.title = "Relatório de Status"
    ws_rel.append(["Arquivo do Boleto", "Status do Processamento", "Linha Digitável Extraída", "Valor Extraído"])
    for col in ['A', 'B', 'C', 'D']:
        ws_rel.column_dimensions[col].width = 40
    idx_r = 2

    for d in dados:
        ws_dados.cell(row=idx_d, column=1, value=str(d["Linha"]))
        ws_dados.cell(row=idx_d, column=2, value=limpar_valor_brasileiro(str(d["Valor"])))
        ws_dados.cell(row=idx_d, column=3, value=date.today().strftime("%d/%m/%Y"))
        ws_dados.cell(row=idx_d, column=4, value=d["Arquivo"])
        ws_dados.cell(row=idx_d, column=5, value=d["Agencia"])
        ws_dados.cell(row=idx_d, column=6, value=d["Conta"])
        idx_d += 1

        ws_rel.cell(row=idx_r, column=1, value=d["Arquivo"])
        ws_rel.cell(row=idx_r, column=2, value=d["Status"])
        ws_rel.cell(row=idx_r, column=3, value=str(d["Linha"]))
        ws_rel.cell(row=idx_r, column=4, value=str(d["Valor"]))
        idx_r += 1

    wb_dados.save(path_dados)
    wb_dados.close()

    wb_rel.save(path_relatorio)
    wb_rel.close()

    return path_dados, path_relatorio

# ======================================================
# 5) PONTES DE COMUNICAÇÃO (EEL) E SISTEMA
# ======================================================
def obter_caminho_config():
    pasta_usuario = os.path.expanduser("~")
    return os.path.join(pasta_usuario, "config_berh.json")

@eel.expose
def salvar_config_modelo(caminho):
    try:
        arquivo_config = obter_caminho_config()
        with open(arquivo_config, "w", encoding="utf-8") as f:
            json.dump({"modelo_salvo": caminho}, f)
    except Exception as e:
        print(f"Erro ao salvar config: {e}")

@eel.expose
def carregar_config_modelo():
    try:
        arquivo_config = obter_caminho_config()
        if os.path.exists(arquivo_config):
            with open(arquivo_config, "r", encoding="utf-8") as f:
                dados = json.load(f)
                return dados.get("modelo_salvo", "")
    except Exception:
        pass
    return ""

@eel.expose
def selecionar_arquivo_zip():
    root = tk.Tk()
    root.withdraw()
    root.wm_attributes('-topmost', 1)
    caminhos = filedialog.askopenfilenames(parent=root, title="Selecione os arquivos ZIP", filetypes=[("ZIP", "*.zip")])
    root.destroy()
    return list(caminhos) if caminhos else []

@eel.expose
def selecionar_arquivo_modelo():
    root = tk.Tk()
    root.withdraw()
    root.wm_attributes('-topmost', 1)
    caminho = filedialog.askopenfilename(parent=root, title="Selecione a Planilha Modelo", filetypes=[("Excel", "*.xlsx")])
    root.destroy()
    return caminho if caminho else ""

@eel.expose
def abrir_arquivo_excel(caminho_arquivo):
    try:
        if os.path.exists(caminho_arquivo):
            if sys.platform == "win32":
                os.startfile(caminho_arquivo)
            else:
                # macOS usa 'open', Linux usa 'xdg-open'
                opener = "open" if sys.platform == "darwin" else "xdg-open"
                subprocess.call([opener, caminho_arquivo])
    except Exception as e:
        print(f"Erro ao abrir arquivo: {e}")

@eel.expose
def abrir_pasta_do_arquivo(caminho_arquivo):
    try:
        if os.path.exists(caminho_arquivo):
            if sys.platform == "win32":
                subprocess.Popen(rf'explorer /select,"{os.path.normpath(caminho_arquivo)}"')
            else:
                pasta = os.path.dirname(caminho_arquivo)
                opener = "open" if sys.platform == "darwin" else "xdg-open"
                subprocess.call([opener, pasta])
    except Exception as e:
        print(f"Erro ao abrir pasta: {e}")

@eel.expose
def minimizar_janela():
    try:
        if sys.platform == "win32":
            ctypes.windll.user32.ShowWindow(ctypes.windll.user32.GetForegroundWindow(), 6)
    except Exception as e:
        print(f"Erro ao minimizar: {e}")

@eel.expose
def iniciar_extracao(zip_paths, modelo_path):
    if not zip_paths or len(zip_paths) == 0:
        return {"status": "erro", "mensagem": "Erro: Nenhum arquivo ZIP selecionado."}

    res = []
    try:
        pdfs_para_processar = []

        for zp in zip_paths:
            if os.path.exists(zp):
                with zipfile.ZipFile(zp, 'r') as z:
                    for nome in z.namelist():
                        if nome.lower().endswith('.pdf'):
                            pdfs_para_processar.append((nome, z.read(nome)))

        total = len(pdfs_para_processar)
        if total == 0:
            return {"status": "erro", "mensagem": "Erro: Nenhum PDF encontrado nos ZIPs."}

        eel.atualizar_status_js(f"Preparando {total} faturas...")()

        # Ajuste a concorrência se notar 429 (cotas). Reduza max_workers conforme necessário.
        with concurrent.futures.ThreadPoolExecutor(max_workers=10) as exc:
            stop = DummyStop()
            futuros = {exc.submit(processar_pdf, nome, bytes_pdf, stop): nome for nome, bytes_pdf in pdfs_para_processar}

            processados = 0
            for f in concurrent.futures.as_completed(futuros):
                nome_arquivo = futuros[f]
                try:
                    r = f.result(timeout=60)
                    if r: res.append(r)
                except concurrent.futures.TimeoutError:
                    res.append({"Arquivo": nome_arquivo, "Linha": "", "Valor": "0.00", "Status": "ERRO: Thread travada (Timeout)", "Agencia": "0050", "Conta": "569539-7"})
                except Exception as e:
                    res.append({"Arquivo": nome_arquivo, "Linha": "", "Valor": "0.00", "Status": f"Erro interno: {str(e)}", "Agencia": "0050", "Conta": "569539-7"})

                processados += 1
                eel.atualizar_status_js(f"Processando: {processados} de {total} faturas...")()

        if res:
            pasta_destino = os.path.dirname(zip_paths[0])
            caminho_dados, caminho_relatorio = gerar_excel_duplo(res, pasta_destino, modelo_path)

            return {
                "status": "sucesso",
                "mensagem": "Extração concluída com sucesso!",
                "planilha_dados": caminho_dados,
                "planilha_relatorio": caminho_relatorio
            }

    except Exception as e:
        return {"status": "erro", "mensagem": f"Erro fatal: {str(e)}"}

if __name__ == "__main__":
    eel.init('web')
    eel.start('index.html', size=(850, 750), position=(300, 150), port=0)
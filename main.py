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
API_KEY = "AIzaSyC5R7I9TycXDKUhThkuvk79piZ7JkIT2xk"

# ======================================================
# 2) DEPENDÊNCIAS DO EXCEL, GENAI E OCR
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

try:
    import pdfplumber
except ImportError:
    print("Erro: Instale a biblioteca pdfplumber (pip install pdfplumber)")
    sys.exit(1)

# ======================================================
# 3) LÓGICA DE LIMPEZA E EXTRAÇÃO (OCR + IA + AUDITOR)
# ======================================================
def limpar_valor_brasileiro(valor_str):
    if not valor_str or valor_str == "0.00":
        return "0.00"
    limpo = re.sub(r'[^\d\.,]', '', str(valor_str))
    if '.' in limpo and ',' in limpo:
        limpo = limpo.replace('.', '')
    limpo = limpo.replace(',', '.')
    return limpo

def extrair_com_ocr(pdf_bytes):
    linha, valor = "", "0.00"
    try:
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            texto = ""
            # OTIMIZAÇÃO 1: Lê apenas a primeira página para economizar tempo
            for pagina in pdf.pages[:1]:
                texto += pagina.extract_text() + "\n"
                
            # CORREÇÃO: Lookarounds evitam que o Regex pegue um dígito adjacente aleatório
            match_linha = re.search(r'(?<!\d)(?:\d[\.\s]*){47}(?!\d)|(?<!\d)(?:\d[\.\s]*){48}(?!\d)', texto)
            if match_linha:
                linha = re.sub(r'\D', '', match_linha.group(0))
                
            match_valor = re.search(r'(?:Valor do Documento|R\$|Valor|Total)\s*[:|-]?\s*([\d\.]+(?:,\d{2}))', texto, re.IGNORECASE)
            if match_valor:
                valor = limpar_valor_brasileiro(match_valor.group(1))
    except Exception:
        pass
    return linha, valor

def processar_com_gemini_vision(api_key, pdf_bytes):
    try:
        client = genai.Client(api_key=api_key)
        # CORREÇÃO: Prompt mais diretivo sobre a quantidade exata de dígitos
        prompt = """
        Analise este boleto bancário. Extraia os dados abaixo:
        1. LINHA DIGITÁVEL: Código de barras numérico.
           - Boletos bancários padrão: EXATAMENTE 47 DÍGITOS.
           - Contas de consumo/impostos: EXATAMENTE 48 DÍGITOS.
           - Não invente nem adicione zeros extras!
        2. VALOR DO DOCUMENTO: Valor total.
        Responda EXATAMENTE neste formato:
        LINHA: [numeros]
        VALOR: [valor]
        """
        config = types.GenerateContentConfig(temperature=0.0)
        response = client.models.generate_content(
            model="gemini-2.0-flash",
            contents=[types.Part.from_bytes(data=pdf_bytes, mime_type="application/pdf"), prompt],
            config=config
        )
        
        linha, valor = "", "0.00"
        if response and response.text:
            texto_resp = response.text
            match_linha = re.search(r'LINHA:\s*([\d\.\s]+)', texto_resp)
            if match_linha:
                l = re.sub(r'\D', '', match_linha.group(1))
                if len(l) >= 44: linha = l
                
            match_valor = re.search(r'VALOR:\s*(.+)', texto_resp)
            if match_valor: valor = limpar_valor_brasileiro(match_valor.group(1).strip())
                
            if not linha:
                # CORREÇÃO: Limpa apenas pontos e espaços antes de buscar o bloco isolado
                texto_limpo = re.sub(r'[\.\-\s]', '', texto_resp)
                bruto = re.search(r'\b\d{47}\b|\b\d{48}\b', texto_limpo)
                if bruto: linha = bruto.group(0)
                
        return linha, valor
    except Exception as e:
        erro = str(e)
        if "429" in erro: return "", "Erro: Cota Excedida"
        return "", f"Erro IA: {type(e).__name__}"

def auditoria_final_gemini(api_key, pdf_bytes, linha_ocr, valor_ocr, linha_ia, valor_ia):
    try:
        time.sleep(0.5) 
        client = genai.Client(api_key=api_key)
        prompt = f"""
        Você é um auditor financeiro sênior. Houve uma divergência na leitura deste boleto.
        - O sistema 1 (OCR) leu: LINHA: {linha_ocr} | VALOR: {valor_ocr}
        - O sistema 2 (IA) leu: LINHA: {linha_ia} | VALOR: {valor_ia}

        Analise a imagem novamente com atenção máxima.
        Lembre-se: Boletos normais têm EXATAMENTE 47 dígitos. Guias de imposto têm EXATAMENTE 48 dígitos. Não adicione zeros.
        Extraia a informação DEFINITIVA e correta que está impressa no boleto.
        Responda EXATAMENTE neste formato:
        LINHA: [numeros definitivos]
        VALOR: [valor definitivo]
        """
        config = types.GenerateContentConfig(temperature=0.0)
        response = client.models.generate_content(
            model="gemini-2.0-flash",
            contents=[types.Part.from_bytes(data=pdf_bytes, mime_type="application/pdf"), prompt],
            config=config
        )
        
        linha, valor = "", "0.00"
        if response and response.text:
            texto_resp = response.text
            match_linha = re.search(r'LINHA:\s*([\d\.\s]+)', texto_resp)
            if match_linha:
                l = re.sub(r'\D', '', match_linha.group(1))
                if len(l) >= 44: linha = l
            match_valor = re.search(r'VALOR:\s*(.+)', texto_resp)
            if match_valor: valor = limpar_valor_brasileiro(match_valor.group(1).strip())
        return linha, valor
    except Exception:
        return "", "0.00"

class DummyStop:
    def is_set(self): return False

def processar_pdf(api_key, nome_arquivo, conteudo_bytes, stop_event):
    if stop_event.is_set(): return None
    agencia, conta = "0050", "569539-7" 
    
    try:
        with concurrent.futures.ThreadPoolExecutor(max_workers=2) as executor_interno:
            futuro_ocr = executor_interno.submit(extrair_com_ocr, conteudo_bytes)
            futuro_ia = executor_interno.submit(processar_com_gemini_vision, api_key, conteudo_bytes)
            
            linha_ocr, valor_ocr = futuro_ocr.result()
            linha_ia, valor_ia = futuro_ia.result()
        
        if "Erro" in valor_ia:
            status = valor_ia
            linha_final, valor_final = linha_ocr, valor_ocr
        elif not linha_ia and not linha_ocr:
            status = "FALHA CRÍTICA: OCR e IA não encontraram nada"
            linha_final, valor_final = "", "0.00"
        elif linha_ocr == linha_ia and valor_ocr == valor_ia:
            status = "SUCESSO (Validado de primeira)"
            linha_final, valor_final = linha_ia, valor_ia
        else:
            if linha_ocr and valor_ocr != "0.00":
                linha_def, valor_def = auditoria_final_gemini(api_key, conteudo_bytes, linha_ocr, valor_ocr, linha_ia, valor_ia)
                if linha_def == linha_ocr and valor_def == valor_ocr:
                    status = "SUCESSO (OCR venceu a disputa)"
                    linha_final, valor_final = linha_def, valor_def
                elif linha_def == linha_ia and valor_def == valor_ia:
                    status = "SUCESSO (IA Gerente venceu a disputa)"
                    linha_final, valor_final = linha_def, valor_def
                else:
                    status = f"REVISÃO MANUAL! (OCR: {valor_ocr} | IA: {valor_ia} | Auditor: {valor_def})"
                    linha_final, valor_final = linha_def, valor_def
            else:
                status = "SUCESSO (Salvo pela IA)"
                linha_final, valor_final = linha_ia, valor_ia
            
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
    # Formata a data e hora atual (exemplo: 05032026_1430)
    data_str = datetime.now().strftime('%d%m%Y_%H%M')
    
    # Atualizado: Novo nome da planilha principal com data e hora
    path_dados = os.path.join(pasta, f"planilha boletos BTG {data_str}.xlsx")
    
    # Mantém o nome do relatório de status também com data e hora para organização
    path_relatorio = os.path.join(pasta, f"Relatorio_Status_{data_str}.xlsx")

    if modelo_path and os.path.exists(modelo_path):
        wb_dados = load_workbook(modelo_path)
    else:
        wb_dados = gerar_workbook_padrao_dados()
        
    ws_dados = wb_dados.active
    
    # Força o desbloqueio da planilha para garantir que fique editável
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

    # Fecha os arquivos no sistema logo após salvar para liberar o uso
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
                subprocess.call(["open", caminho_arquivo])
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
                subprocess.call(["open", pasta])
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
        if total == 0: return {"status": "erro", "mensagem": "Erro: Nenhum PDF encontrado nos ZIPs."}

        eel.atualizar_status_js(f"Preparando {total} faturas...")()

        with concurrent.futures.ThreadPoolExecutor(max_workers=10) as exc:
            stop = DummyStop()
            futuros = {exc.submit(processar_pdf, API_KEY, nome, bytes_pdf, stop): nome for nome, bytes_pdf in pdfs_para_processar}
            
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
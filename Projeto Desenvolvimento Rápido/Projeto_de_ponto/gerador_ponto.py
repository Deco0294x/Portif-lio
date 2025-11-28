# gerador_ponto.py
# GERADOR DE FOLHA DE PONTO - VERSÃO ATUALIZADA (V.5.39.00 - Histórico Postos)
# Requisitos: Python 3.10+, bibliotecas: pandas, tkcalendar, reportlab
# pip install pandas tkcalendar reportlab

from __future__ import annotations
import os
import io
import json
import base64
import traceback
from datetime import datetime, timedelta, date
from typing import Dict, Any, List, Optional, Tuple
import unicodedata

# GUI
from tkinter import (
    Tk, Toplevel, Frame, Label, Entry, Button, StringVar, BooleanVar, Spinbox, Canvas, Listbox,
    ttk, filedialog, messagebox, LEFT, RIGHT, X, Y, BOTH, NSEW, RIDGE, SUNKEN, FLAT, WORD, Text, Scrollbar, END, DISABLED, NORMAL, EXTENDED
)
from tkcalendar import DateEntry

# PDF
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm
from reportlab.lib.utils import ImageReader
from reportlab.lib.colors import black # Adicionado para uso na grade/linhas

# ---------------------------
# CONFIGURAÇÕES E CONSTANTES
# ---------------------------
DATA_STORE = "escalas_store.json"
LOGO_B64_PATH = "logo.txt"
OUTPUT_FOLDER = "Pontos Gerados"

WEEKDAY_PT_SHORT = {0: "SEG", 1: "TER", 2: "QUA", 3: "QUI", 4: "SEX", 5: "SÁB", 6: "DOM"}

# Meses em português
MONTH_NAMES_PT = {
    1: "JANEIRO", 2: "FEVEREIRO", 3: "MARÇO", 4: "ABRIL",
    5: "MAIO", 6: "JUNHO", 7: "JULHO", 8: "AGOSTO",
    9: "SETEMBRO", 10: "OUTUBRO", 11: "NOVEMBRO", 12: "DEZEMBRO"
}

# Lógica das escalas
SCALE_TYPES = {
    "5X2": {"cycle": [1, 1, 1, 1, 1, 0, 0]}, # 5 Trabalha, 2 Folga
    "5X1": {"cycle": [1, 1, 1, 1, 1, 0]}, # 5 Trabalha, 1 Folga
    "6X1 (FIXO)": {"cycle": None}, # Lógica de folga fixa no Domingo
    "6X1 (INTERCALADA)": {"cycle": None}, # Lógica de folga a cada 7 dias, alternada
    "12X36": {"cycle": [0, 1]}, # 1 Folga, 1 Trabalha (inicia com Folga, alinhando com a 1ª Folga)
}

# DEFAULT_WEEKLY_TEMPLATE removido - não mais necessário após remoção de horários automáticos

# ---------------------------
# UTILITÁRIOS
# ---------------------------
def normalize_text(s: str) -> str:
    """Remove acentos e coloca em minúsculas para comparação flexível."""
    if s is None:
        s = ""
    # NFD decompõe acentos; removemos letras da categoria Mn (marcas)
    decomposed = unicodedata.normalize('NFD', str(s))
    no_accents = ''.join(c for c in decomposed if unicodedata.category(c) != 'Mn')
    return no_accents.casefold()
def now_str():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

def safe_mkdir(path: str):
    try:
        os.makedirs(path, exist_ok=True)
    except Exception:
        pass

def safe_path_name(name: str) -> str:
    """Converte um nome para um nome de pasta seguro, removendo caracteres especiais."""
    if not name:
        return "SEM_POSTO"
    # Remove caracteres especiais, mantém letras, números, espaços e alguns símbolos
    safe = "".join(c for c in name if (c.isalnum() or c in (' ', '_', '-')))
    # Substitui espaços por underscore e remove underscores duplicados
    safe = safe.strip().replace(" ", "_").replace("__", "_").upper()
    return safe or "SEM_POSTO"

def load_store() -> dict:
    if os.path.exists(DATA_STORE):
        try:
            with open(DATA_STORE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}

def save_store(store: dict):
    try:
        with open(DATA_STORE, "w", encoding="utf-8") as f:
            json.dump(store, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print("ERRO AO SALVAR STORE:", e)

def safe_remove_file(path: str) -> bool:
    if not os.path.exists(path):
        return True
    try:
        os.remove(path)
        return True
    except PermissionError:
        return False
    except Exception:
        try:
            os.remove(path)
            return True
        except Exception:
            return False

def parse_hhmm(s: Optional[str]) -> Optional[tuple]:
    if s is None:
        return None
    s = str(s).strip()
    if s == "":
        return None
    s = s.replace(",", ":").replace(".", ":")
    if ":" in s:
        parts = s.split(":")
        try:
            h = int(parts[0]); m = int(parts[1]) if len(parts) > 1 and parts[1] != "" else 0
            if 0 <= h <= 23 and 0 <= m <= 59:
                return (h, m)
            return None
        except Exception:
            return None
    if s.isdigit():
        try:
            if len(s) <= 2:
                h = int(s); m = 0
            elif len(s) == 3:
                h = int(s[0]); m = int(s[1:])
            else:
                h = int(s[:2]); m = int(s[2:])
            if 0 <= h <= 23 and 0 <= m <= 59:
                return (h, m)
            return None
        except Exception:
            return None
    return None

def normalize_to_hhmm(s: Optional[str]) -> str:
    parsed = parse_hhmm(s)
    if not parsed:
        return ""
    h, m = parsed
    return f"{h:02d}:{m:02d}"

def hhmm_to_minutes(hhmm: Optional[str]) -> int:
    if not hhmm:
        return 0
    try:
        parts = str(hhmm).split(":")
        h = int(parts[0]); m = int(parts[1]) if len(parts) > 1 else 0
        return h * 60 + m
    except Exception:
        return 0

def minutes_to_hhmm_signed(total_minutes: int) -> str:
    try:
        tm = int(total_minutes)
    except Exception:
        return ""
    if tm > 0:
        h = tm // 60; m = tm % 60
        return f"+{h:02d}:{m:02d}"
    elif tm < 0:
        tm_abs = abs(tm); h = tm_abs // 60; m = tm_abs % 60
        return f"-{h:02d}:{m:02d}"
    else:
        return "00:00"

def daterange(start_date: date, end_date: date):
    cur = start_date
    while cur <= end_date:
        yield cur
        cur += timedelta(days=1)

def load_logo_image(filial=None):
    """
    Carrega logo do store.
    Se filial for fornecida, busca logo específico da filial em logos_filiais.
    Caso contrário, usa logo_base64 global como fallback.
    Retorna ImageReader ou None.
    """
    try:
        store = load_store()
        b64 = None
        
        # Tenta carregar logo específico da filial
        if filial:
            logos_filiais = store.get("logos_filiais", {})
            b64 = logos_filiais.get(filial, "").strip()
        
        # Fallback: logo global
        if not b64:
            b64 = store.get("logo_base64", "").strip()
        
        # Fallback: arquivo logo.txt
        if not b64 and os.path.exists(LOGO_B64_PATH):
            with open(LOGO_B64_PATH, "r", encoding="utf-8") as f:
                b64 = f.read().strip()
        
        if not b64:
            return None
        
        data = base64.b64decode(b64)
        return ImageReader(io.BytesIO(data))
    
    except Exception as e:
        print("ERRO AO CARREGAR LOGO:", e)
        return None

# ---------------------------
# SCHEDULE (SIMULAÇÃO E REGRAS)
# ---------------------------
def generate_employee_schedule(employee: dict, config: dict, emp_events: dict):
    """
    Gera schedule considerando:
      - admissão do funcionário (não gera antes)
      - eventos manuais (folgas e feriados extras)
      - regras de escalas (5X1, 12X36, 6X1 Fixo/Intercalada)
      - feriados nacionais (todos folgam) e locais (apenas postos selecionados folgam)
      - EXCEÇÃO: Escala 12X36 trabalha normalmente nos feriados
    """
    start = config["start_date"]
    end = config["end_date"]
    holidays = config.get("holidays", {})
    holiday_type = config.get("holiday_type", {})
    holiday_postos = config.get("holiday_postos", {})
    scale_type = config.get("scale_type", "6X1 (FIXO)")
    cycle_info = SCALE_TYPES.get(scale_type, SCALE_TYPES["5X2"])
    cycle = cycle_info.get("cycle", SCALE_TYPES["5X2"]["cycle"])
    cycle_len = len(cycle) if cycle else None

    admissao = employee.get("admissao", None)
    emp_name = employee.get("nome", "")
    emp_posto = employee.get("posto", "")

    emp_events_for_emp = emp_events.get(emp_name, {}).copy()  # copiar para podermos alterar localmente

    # --- Lógica de Alinhamento de Ciclo (para 5X2, 12X36, 5X1) ---
    cycle_start_ref = start # Padrão: alinha com a data de início do período
    first_off_date = None
    first_off_str = config.get("first_off", None)
    if first_off_str:
        try:
            first_off_date = datetime.strptime(first_off_str, "%Y-%m-%d").date()
        except Exception:
            pass
            
    if first_off_date and cycle:
        # A 'first_off_date' deve alinhar com a primeira Folga (0) no ciclo.
        try:
            index_of_first_off = cycle.index(0) 
        except ValueError:
            index_of_first_off = 0
            
        if index_of_first_off >= 0:
            # A data que corresponde ao índice 0 do ciclo é:
            # Subtrai o índice da folga para encontrar o dia que marca o início do ciclo (índice 0)
            cycle_start_ref = first_off_date - timedelta(days=index_of_first_off)
    # --- Fim Lógica de Alinhamento de Ciclo ---

    # --- Pré-cálculo do 3º Domingo para 5X1 ---
    third_sundays_by_month: Dict[Tuple[int, int], date] = {}
    if scale_type == "5X1":
        sundays_by_month: Dict[Tuple[int, int], List[date]] = {}
        for d_range in daterange(start, end):
            if d_range.weekday() == 6: # Sunday
                sundays_by_month.setdefault((d_range.year, d_range.month), []).append(d_range)
        
        for (yr, mo), s_list in sundays_by_month.items():
            if len(s_list) >= 3:
                third_sundays_by_month[(yr, mo)] = s_list[2] # índice 2 é o 3º Domingo

    # --- PASSAGEM 1: Geração do Agendamento (Coleta de dados) ---
    final_schedule: Dict[str, Any] = {}

    for d in daterange(start, end):
        ds = d.strftime("%Y-%m-%d")
        
        # não gerar antes da admissão
        if admissao and d < admissao:
            final_schedule[ds] = {"type": "ANTES_ADMISSAO"}
            continue

        # eventos manuais
        ev = emp_events_for_emp.get(ds)
        if ev in ("FOLGA", "FERIADO"):
            final_schedule[ds] = {"type": ev}
            continue

        # feriado global - verifica se a data atual corresponde a algum feriado cadastrado
        # A partir da data de cadastro, o feriado se replica para todos os anos seguintes (mesmo dia/mês)
        is_hol = False
        matched_holiday_key = None
        for hol_date_str in holidays.keys():
            try:
                hol_date = datetime.strptime(hol_date_str, "%Y-%m-%d").date()
                # Verifica se dia e mês são iguais E se o ano atual é >= ao ano do cadastro
                if d.day == hol_date.day and d.month == hol_date.month and d.year >= hol_date.year:
                    is_hol = True
                    matched_holiday_key = hol_date_str
                    break
            except Exception:
                continue

        wd = d.weekday()

        # Verifica se é feriado e se o funcionário folga nele
        # EXCEÇÃO: Escala 12X36 trabalha normalmente nos feriados (ignora feriado)
        should_be_off_holiday = False
        if is_hol and matched_holiday_key and scale_type != "12X36":
            # Verifica o tipo de feriado usando a chave original do feriado cadastrado
            hol_type = holiday_type.get(matched_holiday_key, "NACIONAL")
            if hol_type == "NACIONAL":
                # Feriado nacional: todos folgam
                should_be_off_holiday = True
            elif hol_type == "LOCAL":
                # Feriado local: verifica se o posto do funcionário está na lista
                postos_com_direito = holiday_postos.get(matched_holiday_key, [])
                if emp_posto in postos_com_direito:
                    should_be_off_holiday = True
        
        if should_be_off_holiday:
            final_schedule[ds] = {"type": "FERIADO"}
            continue

        # --- Lógica das Escalas para determinar se é dia de trabalho ---
        # Assume dia de trabalho para dias úteis (segunda a sexta-feira)
        worked = True if wd < 5 else False  # Segunda a Sexta trabalha, Sábado e Domingo folga (padrão)

        if scale_type == "6X1 (FIXO)":
            # 6X1 Fixo: Folga todos os Domingos
            if wd == 6:
                worked = False
                
        elif scale_type == "6X1 (INTERCALADA)" and first_off_date:
            # 6X1 Intercalada: Folga alterna a cada 7 dias (ciclo de 14 dias), contando a partir da 1ª Folga.
            days_since_first_off = (d - first_off_date).days
            
            # Folgas caem quando days_since_first_off % 14 é 0 (1ª Folga) ou 7 (2ª Folga)
            if days_since_first_off % 14 == 0 or days_since_first_off % 14 == 7:
                worked = False
        
        elif cycle_len:
            # Ciclo padrão (5X2, 12X36, 5X1) usando alinhamento por cycle_start_ref
            idx = (d - cycle_start_ref).days % cycle_len
            worked = bool(cycle[idx])
        
        # --- FIM Lógica das Escalas ---

        if not worked:
            final_schedule[ds] = {"type": "FOLGA"}
            continue

        entry_type = "TRABALHADO"

        # Horários deixados em branco para preenchimento manual
        entry = {
            "type": entry_type,
            "entrada": "",
            "int_start": "",
            "int_end": "",
            "saida": ""
        }
        if entry_type == "FERIADO_TRABALHADO":
            entry["name"] = holidays.get(ds, "FERIADO")
            
        final_schedule[ds] = entry
    
    # --- PASSAGEM 2: Aplica Regra da Folga do 3º Domingo para 5X1 ---
    if scale_type == "5X1":
        
        # Identifica os meses que *já* possuem uma folga de domingo (de escala regular)
        months_with_sunday_folga = set()
        for ds, entry in final_schedule.items():
            if entry.get("type") == "FOLGA":
                dt_key = datetime.strptime(ds, "%Y-%m-%d")
                if dt_key.weekday() == 6: # Sunday
                    months_with_sunday_folga.add((dt_key.year, dt_key.month))
        
        for (yr, mo), third_sunday in third_sundays_by_month.items():
            
            if (yr, mo) not in months_with_sunday_folga:
                # Condição: Nenhuma folga de escala caiu em um Domingo neste mês
                
                ds_sunday = third_sunday.strftime("%Y-%m-%d")
                current_entry = final_schedule.get(ds_sunday)
                
                # Aplica a folga extra APENAS se o dia for TRABALHADO (ou Folga, mas a folga anterior não conta no set)
                is_overrideable = not current_entry or current_entry.get("type") in ("TRABALHADO", "FERIADO_TRABALHADO")
                
                if is_overrideable:
                    # Folga extra para compensar domingo
                    final_schedule[ds_sunday] = {"type": "FOLGA_DOMINGO_EXTRA", "obs": "FOLGA DOMINGO EXTRA"}
                
    # --- Saída Final: Retorna o agendamento completo ---
    for ds in sorted(final_schedule.keys()):
        yield ds, final_schedule[ds]

# ---------------------------
# PDF GENERATION (A4, MARGENS 10mm, HELVETICA 8pt, GRADE, SALDO TOTAL, RODAPÉ)
# ---------------------------
def generate_pdf_for_employee(
    nome: str,
    cpf: str,
    matricula: str,
    funcao: str,
    posto_global: str,
    filial: str,
    cnpj: str,
    endereco: str,
    cidade: str,
    schedule_map: Dict[str, Any],

    out_folder: str = OUTPUT_FOLDER,
    version_index: Optional[int] = None,
) -> Optional[str]:
    safe_mkdir(out_folder)
    logo = load_logo_image(filial)

    # agrupa por mês
    by_month: Dict[Tuple[int,int], Dict[str,dict]] = {}
    for ds, entry in schedule_map.items():
        try:
            dt = datetime.strptime(ds, "%Y-%m-%d").date()
        except:
            continue
        key = (dt.year, dt.month)
        by_month.setdefault(key, {})[ds] = entry

    # gera por mês
    saved_paths = []
    for (yr, mo), entries in sorted(by_month.items()):
        # se nenhum dia desse mês tem tipo TRABALHADO ou similar, pulamos
        has_day = False
        sorted_dates = sorted(entries.keys())
        for ds, e in entries.items():
            t = (e.get("type") or "").upper()
            if t not in ("ANTES_ADMISSAO", "FERIADO", "FOLGA"):
                has_day = True
                break
        if not has_day:
            # pulamos geração deste mês (sem dias úteis)
            continue

        # Estrutura de pastas desejada:
        # out_folder/<POSTO>/<ANO>/<FUNCIONARIO>/<Opção X se duplicado>/<arquivo.pdf>

        # 1) Pasta do posto (se vazio, usa o nome do funcionário)
        if posto_global and posto_global.strip():
            posto_folder = safe_path_name(posto_global)
        else:
            # Se posto vazio, usa o nome do funcionário como pasta
            posto_folder = safe_path_name(nome)
        posto_path = os.path.join(out_folder, posto_folder)
        safe_mkdir(posto_path)

        # 2) Pasta do ano
        year_folder = str(yr)
        year_path = os.path.join(posto_path, year_folder)
        safe_mkdir(year_path)

        # 3) Pasta do funcionário
        nome_base = (nome or "").rstrip(".").strip().upper()
        # remove caracteres proibidos em path (Windows): \ / : * ? " < > |
        nome_base = "".join(ch for ch in nome_base if ch not in "\\/:*?\"<>|")
        funcionario_folder = nome_base
        funcionario_path = os.path.join(year_path, funcionario_folder)
        safe_mkdir(funcionario_path)

        # 4) Se for duplicado, cria subpasta "Opção X"
        if version_index:
            opcao_folder = f"Opcao {version_index}"
            final_path = os.path.join(funcionario_path, opcao_folder)
            safe_mkdir(final_path)
        else:
            final_path = funcionario_path

        # Nome do arquivo PDF: MM.AAAA_NOME_FUNCIONARIO.pdf
        month_name = MONTH_NAMES_PT.get(mo, f"MES{mo:02d}")
        fname = f"{mo:02d}.{yr}_{nome_base}.pdf"
        path = os.path.join(final_path, fname)

        # tenta remover arquivo existente
        if os.path.exists(path):
            ok = safe_remove_file(path)
            if not ok:
                print(f"[{now_str()}] AVISO: NÃO FOI POSSÍVEL SOBRESCREVER {path}. Arquivo pode estar aberto.")
                continue

        c = canvas.Canvas(path, pagesize=A4)
        page_w, page_h = A4

        left = 10 * mm
        right = page_w - 10 * mm
        top_margin = page_h - 10 * mm
        bottom_margin = 10 * mm

        font_main = "Helvetica"
        font_bold = "Helvetica-Bold"
        font_size_table = 8

        row_height = 4.2 * mm  # Reduzido de 6.0mm para 4.2mm (-5px aprox.)
        cell_padding = row_height * 0.35  # Aumentado para 35% da altura para mais espaço vertical
        text_vertical_shift = -(row_height * 0.5)  # Desloca o conteúdo meia linha para baixo

        # CÓDIGO DE 13 DÍGITOS
        first_day_dt = datetime.strptime(sorted_dates[0], "%Y-%m-%d").date()
        mat_5_digits = str(matricula or '0').zfill(5)[:5]
        first_day_dmy = first_day_dt.strftime("%d%m%Y")
        codigo_13_digitos = f"{mat_5_digits}{first_day_dmy}"

        # cabeçalho
        info_x = left
        if logo:
            try:
                logo_target_h = 26.0 * mm * 0.65
                lw, lh = logo.getSize()
                scale = logo_target_h / lh
                logo_target_w = lw * scale
                logo_y = (top_margin - 2 * mm) - (logo_target_h / 2) - (6.3 * mm) - (3.5 * mm) - (3.5 * mm)  # 10px + 10px = 20px para baixo
                logo_x = left + (2.1 * mm) + (3.5 * mm)  # 6px + 10px = 16px para direita
                c.drawImage(logo, logo_x, logo_y, width=logo_target_w, height=logo_target_h, mask='auto')
                info_x = left + logo_target_w + 12 * mm
            except Exception:
                info_x = left

        # 1. EMPRESA/FILIAL (10pt) e CNPJ/ENDEREÇO/CIDADE (7pt)
        y_empresa = top_margin - 6 * mm
        c.setFont(font_bold, 10) # 10px para nome da empresa
        c.drawString(info_x, y_empresa, (filial or "EMPRESA").upper())

        y_cnpj = top_margin - 12 * mm
        cnpj_txt_prefix = "CNPJ: "
        c.setFont(font_main, 7) # 7px para CNPJ
        c.drawString(info_x, y_cnpj, cnpj_txt_prefix + (cnpj or '---').upper())
        
        # Endereço e Cidade (7pt) - AJUSTADO: Usa I e J separados por ", "
        y_end = y_cnpj - 4 * mm 
        end_prefix = "ENDEREÇO: "
        
        # Concatena Endereço e Cidade com vírgula + espaço, se existirem
        end_city_txt = ""
        if endereco and cidade:
            end_city_txt = f"{endereco}, {cidade}".upper()
        elif endereco:
            end_city_txt = endereco.upper()
        elif cidade:
            end_city_txt = cidade.upper()
        
        # Escreve ENDEREÇO: [ENDEREÇO, CIDADE]
        c.setFont(font_bold, 7)
        c.drawString(info_x, y_end, end_prefix.upper())
        end_prefix_w = c.stringWidth(end_prefix, font_bold, 7)

        c.setFont(font_main, 7)
        # O valor do endereço começa logo após o prefixo negrito
        c.drawString(info_x + end_prefix_w, y_end, end_city_txt)
        # FIM AJUSTE ENDEREÇO/CIDADE

        # 2. FUNCIONÁRIO/MATRÍCULA (Negrito no rótulo, normal no valor)
        y_nome = y_end - 4 * mm # Puxa para baixo 4mm
        
        # FUNCIONÁRIO
        c.setFont(font_bold, 9)
        nome_prefix = "FUNCIONÁRIO: "
        c.drawString(info_x, y_nome, nome_prefix.upper())
        
        nome_prefix_w = c.stringWidth(nome_prefix, font_bold, 9)
        c.setFont(font_main, 9)
        c.drawString(info_x + nome_prefix_w, y_nome, (nome or '').upper())
        
        # MATRÍCULA
        y_mat = y_nome
        mat_prefix = "MATRÍCULA: "
        nome_w = c.stringWidth((nome or ''), font_main, 9)
        mat_x = info_x + nome_prefix_w + nome_w + 8 * mm
        
        c.setFont(font_bold, 9)
        c.drawString(mat_x, y_mat, mat_prefix.upper())
        
        mat_prefix_w = c.stringWidth(mat_prefix, font_bold, 9)
        c.setFont(font_main, 9)
        c.drawString(mat_x + mat_prefix_w, y_mat, (matricula or '').upper())

        # 3. CPF (8pt)
        y_cpf = y_nome - 4 * mm
        c.setFont(font_main, 8)
        if cpf:
            c.drawString(info_x, y_cpf, f"CPF: {cpf}".upper())

        # 4. POSTO e FUNÇÃO (8pt - Negrito no rótulo, normal no valor)
        y_posto = y_cpf - 4 * mm
        posto_prefix = "POSTO: "
        
        c.setFont(font_bold, 8)
        c.drawString(info_x, y_posto, posto_prefix.upper())
        
        posto_prefix_w = c.stringWidth(posto_prefix, font_bold, 8)
        c.setFont(font_main, 8)
        c.drawString(info_x + posto_prefix_w, y_posto, (posto_global or '').upper())
        
        # FUNÇÃO ao lado do POSTO
        posto_w = c.stringWidth((posto_global or ''), font_main, 8)
        funcao_prefix = " - FUNÇÃO: "
        funcao_x = info_x + posto_prefix_w + posto_w
        
        c.setFont(font_bold, 8)
        c.drawString(funcao_x, y_posto, funcao_prefix.upper())
        
        funcao_prefix_w = c.stringWidth(funcao_prefix, font_bold, 8)
        c.setFont(font_main, 8)
        c.drawString(funcao_x + funcao_prefix_w, y_posto, (funcao or '').upper())
        
        # 5. Título da Tabela (PONTOS REALIZADOS, CÓDIGO, PERÍODO)
        y_title = y_posto - 5 * mm
        c.setLineWidth(0.5)
        c.line(left, y_title, right, y_title)
        
        y_line = y_title - 5 * mm
        
        # PONTOS REALIZADOS (CENTRO)
        c.setFont(font_bold, 9)
        pontos_txt = "PONTOS REALIZADOS"
        pontos_x = page_w / 2
        c.drawCentredString(pontos_x, y_line, pontos_txt)
        
        # CÓDIGO DE 13 DÍGITOS (ESQUERDA)
        c.setFont(font_main, 8) # 8px para Código
        c.drawString(left + 2 * mm, y_line, codigo_13_digitos)
        
        # PERÍODO (DIREITA)
        # Formata a data do período como DD/MM/YYYY
        start_date_str = first_day_dt.strftime("%d/%m/%Y")
        end_date_str = datetime.strptime(sorted_dates[-1], "%Y-%m-%d").date().strftime("%d/%m/%Y")
        periodo_txt = f"PERÍODO: {start_date_str} A {end_date_str}"
        
        c.setFont(font_main, 8) # 8px para PERÍODO
        periodo_w = c.stringWidth(periodo_txt, font_main, 8)
        c.drawString(right - periodo_w - 2 * mm, y_line, periodo_txt)

        # 6. Início da Tabela
        y_header = y_line - 5 * mm

        sorted_dates = sorted(entries.keys())
        # row_height definido acima (6.0 mm)

        # Tabela: DIA | ENTRADA | INT. SAÍDA | INT. RETORNO | SAÍDA | SALDO | OCORRÊNCIA
        headers = ["DIA", "ENTRADA", "INT. SAÍDA", "INT. RETORNO", "SAÍDA", "SALDO", "OCORRÊNCIA"]
        table_left = left
        table_right = right
        table_width = table_right - table_left

        # Colunas reajustadas
        col_widths = [
            25 * mm,  # DIA
            25 * mm, 25 * mm, 25 * mm, 25 * mm,  # ENTRADA/SAÍDA x4 (100 mm)
            25 * mm, # SALDO
            table_width - (25 * mm * 6) # OCORRÊNCIA (190 - 150 = 40mm)
        ]

        c.setFont(font_bold, font_size_table)
        x = table_left
        header_text_y = y_header - (row_height - cell_padding)  # Base do cabeçalho considerando padding
        for i, h in enumerate(headers):
            cw = col_widths[i]
            # Centraliza e aplica deslocamento meia linha para baixo
            c.drawCentredString(x + cw / 2, header_text_y + (cell_padding/2) + text_vertical_shift, h.upper())
            x += cw

        c.setFont(font_main, font_size_table)
        y = header_text_y - (row_height * 0.5)  # Posição inicial das linhas

        # Não há mais cálculo de saldo - todos os horários ficam em branco

        for ds in sorted_dates:
            entry = entries[ds]
            dt = datetime.strptime(ds, "%Y-%m-%d").date()
            day_label = f"{dt.strftime('%d/%m')} - {WEEKDAY_PT_SHORT[dt.weekday()]}"
            etype = (entry.get("type", "") or "").upper()

            obs_text = ""
            if etype in ("FOLGA", "FOLGA_DOMINGO_EXTRA", "FERIADO"):
                col_vals = ["-", "-", "-", "-"]
                obs_text = etype.replace("_DOMINGO_EXTRA", " (EXTRA)")
                saldo_display = ""
                saldo_minutes = 0
            elif etype in ("ANTES_ADMISSAO"):
                col_vals = ["-", "-", "-", "-"]
                obs_text = "PRÉ-ADMISSAO"
                saldo_display = ""
                saldo_minutes = 0
            else:
                # Deixa todos os horários em branco para preenchimento manual
                col_vals = ["", "", "", ""]
                saldo_display = ""
                saldo_minutes = 0
                obs_text = (entry.get("name") or entry.get("obs") or "").upper()

            # Calcula posição Y centralizada com o novo padding e deslocamento meia linha para baixo
            text_y = y - (row_height / 2) + cell_padding + text_vertical_shift

            x = table_left
            c.drawCentredString(x + col_widths[0] / 2, text_y, day_label.upper()); x += col_widths[0]

            # 4 Colunas de Ponto
            for i_col in range(4):
                txt = col_vals[i_col] if i_col < len(col_vals) else ""
                c.drawCentredString(x + col_widths[i_col + 1] / 2, text_y, (txt or "").upper()); x += col_widths[i_col + 1]

            # Saldo
            c.drawCentredString(x + col_widths[5] / 2, text_y, (saldo_display or "").upper()); x += col_widths[5]

            # Ocorrência
            c.drawString(x + 2 * mm, text_y, (obs_text or "").upper()); x += col_widths[6]

            y -= row_height

        # ADICIONA LINHA DE SALDO TOTAL
        # Fundo cinza claro para toda a linha
        c.setFillColorRGB(0.9, 0.9, 0.9)
        c.rect(table_left, y - row_height, table_right - table_left, row_height, fill=1, stroke=0)
        
        # Sobrescreve a coluna SALDO com branco
        x_saldo_inicio = table_left + col_widths[0] + col_widths[1] + col_widths[2] + col_widths[3] + col_widths[4]
        c.setFillColorRGB(1, 1, 1)  # Branco
        c.rect(x_saldo_inicio, y - row_height, col_widths[5], row_height, fill=1, stroke=0)
        
        c.setFillColorRGB(0, 0, 0)  # Volta para preto para o texto
        
        text_y_saldo_total = y - (row_height / 2) + cell_padding + text_vertical_shift
        x = table_left
        
        # Coluna DATA com texto "SALDO TOTAL:" (em negrito, centralizado)
        c.setFont(font_bold, font_size_table)
        c.drawCentredString(x + col_widths[0] / 2, text_y_saldo_total, "SALDO TOTAL:")
        c.setFont(font_main, font_size_table)  # Volta para fonte normal
        x += col_widths[0]
        
        # 4 Colunas de Ponto vazias
        for i_col in range(4):
            x += col_widths[i_col + 1]
        
        # Coluna SALDO vazia (branca, para preenchimento manual)
        x += col_widths[5]
        
        # Coluna OCORRÊNCIA vazia
        x += col_widths[6]
        
        y -= row_height

        # grade (sem linha de SALDO MÊS - removida)
        c.setLineWidth(0.3)
        top_y = header_text_y + (cell_padding * 1.5)  # Ajustado para novo padding
        bottom_y = y  # Borda inferior
        if bottom_y >= top_y:
            bottom_y = top_y - row_height

        # Desenha as linhas horizontais da grade
        c.line(table_left, top_y, table_right, top_y) # Linha topo
        for i in range(int(round((top_y - bottom_y) / float(row_height))) + 1):
            yy = top_y - i * row_height
            c.line(table_left, yy, table_right, yy)
        c.line(table_left, bottom_y, table_right, bottom_y) # Linha baixo (após Saldo)

        # Grade Vertical
        x = table_left
        for cw in col_widths:
            c.line(x, top_y, x, bottom_y)
            x += cw
        c.line(x, top_y, x, bottom_y) # Linha final direita

        # RODAPÉ SOLICITADO

        # 1. OBSERVAÇÕES
        y_obs = y - row_height - 4 * mm
        c.setFont(font_main, 6)
        obs_text = "OBSERVAÇÕES PARA USO EXCLUSIVO DO DEPARTAMENTO PESSOAL:"
        c.drawString(table_left + 2 * mm, y_obs, obs_text.upper())
        
        obs_line_y = y_obs - 1 * mm
        obs_line_height = 6 * mm  # Aumentado de 5mm para 6mm para melhor distribuição
        num_obs_lines = 3 # Reduzido de 4 para 3 linhas
        
        # Desenha as linhas das observações
        for i in range(num_obs_lines):
            c.line(table_left, obs_line_y - (i * obs_line_height), table_right, obs_line_y - (i * obs_line_height))
        
        # 2. DECLARAÇÃO - Descer 2 linhas
        y_decl = obs_line_y - (num_obs_lines * obs_line_height) - 2 * mm + (row_height / 2) - (2 * row_height)  # Desce 2 linhas adicionais
        decl = "DECLARO QUE O HORÁRIO ACIMA REGISTRADO, FOI O ÚNICO POR MIM REALIZADO NO PERÍODO."
        c.setFont(font_bold, 7)
        decl_w = c.stringWidth(decl, font_bold, 7)
        decl_x = (table_left + table_right) / 2 - decl_w / 2
        c.drawString(decl_x, y_decl, decl.upper())

        # 3. ASSINATURA E DATA (REMOVIDAS AS BARRAS "/")
        # Espaço de 2 linhas abaixo da declaração antes das informações
        sign_y = y_decl - (2 * row_height) - 6 * mm
        
        # DATA
        date_line_w = 40 * mm 
        date_x = table_left + 8 * mm
        c.line(date_x, sign_y, date_x + date_line_w, sign_y)
        
        # Texto DATA (CENTRALIZADO ABAIXO DA LINHA)
        date_txt = "DATA"
        c.setFont(font_main, 6)
        c.drawCentredString(date_x + date_line_w / 2, sign_y - 4 * mm, date_txt)

        # ASSINATURA
        sig_line_w = 80 * mm
        sig_x = date_x + date_line_w + 12 * mm 
        c.line(sig_x, sign_y, sig_x + sig_line_w, sign_y)

        # Texto ASSINATURA DO FUNCIONÁRIO (CENTRALIZADO ABAIXO DA LINHA)
        sig_txt = "ASSINATURA DO FUNCIONÁRIO"
        c.drawCentredString(sig_x + sig_line_w / 2, sign_y - 4 * mm, sig_txt)
        
        # 4. BORDA AO REDOR DE TODO O PDF
        # Desenha retângulo conectando as bordas da tabela
        c.setLineWidth(0.5)
        # Usa as mesmas coordenadas das margens e da tabela
        border_rect_x = left
        border_rect_y = bottom_margin
        border_rect_w = right - left
        border_rect_h = top_margin - bottom_margin
        c.rect(border_rect_x, border_rect_y, border_rect_w, border_rect_h, stroke=1, fill=0)

        try:
            c.showPage()
            c.save()
            saved_paths.append(path)
        except Exception as e:
            print(f"[{now_str()}] ERRO AO SALVAR PDF PARA {nome}: {e}\n{traceback.format_exc()}")

    return saved_paths

# ---------------------------
# APLICAÇÃO GUI (JANELA PRINCIPAL)
# ---------------------------
class PontoApp:

    def __init__(self, root: Tk):
        self.root = root
        self.root.title("GERADOR DE FOLHA DE PONTO - SALDO TOTAL:")
        self.store = load_store()

        # dados
        self.funcionarios: List[Dict[str, Any]] = self.store.get("funcionarios", [])
        self.global_holidays: Dict[str, str] = self.store.get("global_holidays", {})
        # Tipo de feriado: "NACIONAL" ou "LOCAL" - Dict[data_str, tipo]
        self.holiday_type: Dict[str, str] = self.store.get("holiday_type", {})
        # Postos que tem direito a cada feriado local - Dict[data_str, List[posto]]
        self.holiday_postos: Dict[str, List[str]] = self.store.get("holiday_postos", {})
        self.emp_personal_hols: Dict[str, List[str]] = self.store.get("emp_personal_hols", {})
        self.emp_scale_choice: Dict[str, str] = self.store.get("emp_scale_choice", {})
        self.emp_first_off: Dict[str, str] = self.store.get("emp_first_off", {})
        self.emp_faltas_atestados: Dict[str, dict] = self.store.get("emp_faltas_atestados", {})
        self.emp_trabalha_feriado: Dict[str, bool] = self.store.get("emp_trabalha_feriado", {})
        # Estado de ordenação por coluna da tabela principal (True = descendente)
        self._emp_sort_dir: Dict[str, bool] = {}
        # Estado de revisão por funcionário (True = revisado, mostra "OK" verde)
        self.emp_revisado: Dict[str, bool] = self.store.get("emp_revisado", {})
        # Sistema de cidades: Dict[cidade_nome, List[postos]]
        self.cidades: Dict[str, List[str]] = self.store.get("cidades", {})
        # Cidades vinculadas aos feriados locais: Dict[data_str, cidade_nome]
        self.holiday_cidades: Dict[str, str] = self.store.get("holiday_cidades", {})
        # Registro histórico de todos os postos já vistos (mantém mesmo após trocar planilha)
        self.all_postos_historico: set = set(self.store.get("all_postos_historico", []))

        self._build_top_frame()
        self._build_mid_frame()
        self._build_bottom_frame()
        self.update_employee_tree()

    def _is_cpf_format(self, text: str) -> bool:
        """Verifica se o texto parece ser um CPF (contém principalmente números, pontos e traços)."""
        if not text:
            return False
        # Remove espaços
        text = text.strip()
        # Conta quantos caracteres são dígitos, pontos ou traços
        valid_chars = sum(1 for c in text if c.isdigit() or c in '.-')
        # Se mais de 70% dos caracteres são números/pontos/traços, provavelmente é CPF
        return len(text) > 0 and (valid_chars / len(text)) > 0.7
    
    def _parse_date_str(self, s: str) -> Optional[date]:
        s = (s or "").strip()
        # Aceita tanto DD/MM/AAAA quanto 8 dígitos (DDMMAAAA)
        try:
            if "/" in s:
                return datetime.strptime(s, "%d/%m/%Y").date()
            digits = "".join(ch for ch in s if ch.isdigit())
            if len(digits) == 8:
                s_fmt = f"{digits[:2]}/{digits[2:4]}/{digits[4:]}"
                return datetime.strptime(s_fmt, "%d/%m/%Y").date()
        except Exception:
            pass
        return None

    def _auto_complete_date(self, var: StringVar):
        """Se o conteúdo tiver 8 dígitos (DDMMAAAA), formata para DD/MM/AAAA.
        Não valida o calendário; a validação completa acontece ao gerar/usar o período."""
        s = (var.get() or "").strip()
        digits = "".join(ch for ch in s if ch.isdigit())
        if len(digits) == 8:
            var.set(f"{digits[:2]}/{digits[2:4]}/{digits[4:]}")

    def _get_period_dates(self) -> Optional[Tuple[date, date]]:
        """Lê os campos de INÍCIO e FIM (DD/MM/AAAA) e retorna (start, end).
        Mostra erro amigável se inválido."""
        s = self._parse_date_str(self.start_date_var.get() if hasattr(self, 'start_date_var') else '')
        e = self._parse_date_str(self.end_date_var.get() if hasattr(self, 'end_date_var') else '')
        if not s or not e:
            messagebox.showerror("ERRO", "Informe datas válidas no formato DD/MM/AAAA para INÍCIO e FIM.")
            return None
        if e < s:
            messagebox.showerror("ERRO", "PERÍODO INVÁLIDO: FIM anterior ao INÍCIO.")
            return None
        return s, e

    def _date_validate_on_key(self, proposed: str) -> bool:
        """Validação on-the-fly: permite apenas dígitos e '/', e no máximo 10 caracteres."""
        if proposed is None:
            return True
        s = proposed.strip()
        if len(s) > 10:
            self.root.bell()
            return False
        for ch in s:
            if not (ch.isdigit() or ch == '/'):
                self.root.bell()
                return False
        # No máximo dois '/'
        if s.count('/') > 2:
            self.root.bell()
            return False
        return True

    def _date_mask_on_key(self, event, var: StringVar):
        """Insere automaticamente '/' após DD e DD/MM quando digitando no fim.
        Não mexe quando navegando no meio do texto."""
        try:
            w = event.widget
            s = (var.get() or "").strip()
            # Ignora teclas de navegação/edição
            if event.keysym in ("BackSpace", "Delete", "Left", "Right", "Home", "End", "Tab"):
                return
            # Só automatiza quando cursor está no fim
            if w.index("insert") != len(s):
                return
            # Adiciona barra após 2 ou 5 caracteres
            if len(s) == 2 and s.isdigit():
                var.set(s + "/")
                w.icursor("end")
            elif len(s) == 5 and (s[:2].isdigit() and s[2] == '/' and s[3:5].isdigit()):
                var.set(s + "/")
                w.icursor("end")
        except Exception:
            pass

    def _build_top_frame(self):
        # Frame principal do topo com design moderno
        top_container = Frame(self.root, bg="#2c3e50", relief="flat", bd=0)
        top_container.pack(fill=X, padx=0, pady=0)
        
        # Título do sistema
        title_frame = Frame(top_container, bg="#1a252f")
        title_frame.pack(fill=X)
        Label(title_frame, text="📋 GERADOR DE FOLHA DE PONTO", 
              font=("Segoe UI", 14, "bold"), bg="#1a252f", fg="#ecf0f1", 
              pady=12).pack()

        top = Frame(top_container, bg="#2c3e50")
        top.pack(fill=X, padx=15, pady=12)

        # Linha única: Período + Botões de ação
        # Coluna 0-1: Período
        periodo_frame = Frame(top, bg="#34495e", relief="flat", bd=0)
        periodo_frame.grid(row=0, column=0, sticky="nsew", padx=4, pady=4)
        
        # Borda interna para efeito card
        periodo_inner = Frame(periodo_frame, bg="#34495e", relief="solid", bd=1)
        periodo_inner.pack(fill=BOTH, expand=True, padx=2, pady=2)

        Label(periodo_inner, text="📅 INÍCIO:", font=("Segoe UI", 9, "bold"), 
              bg="#34495e", fg="#ecf0f1").grid(row=0, column=0, sticky="e", padx=(10, 5), pady=8)
        vcmd = (self.root.register(self._date_validate_on_key), "%P")
        self.start_date_var = StringVar(value=date.today().strftime("%d/%m/%Y"))
        self.start_date_entry = Entry(periodo_inner, textvariable=self.start_date_var, 
                                      width=13, font=("Segoe UI", 9), 
                                      relief="groove", bd=2, bg="white", 
                                      fg="#2c3e50", insertbackground="#3498db",
                                      validate="key", validatecommand=vcmd)
        self.start_date_entry.grid(row=0, column=1, padx=5, pady=8)
        self.start_date_entry.bind("<FocusOut>", lambda e: self._auto_complete_date(self.start_date_var))
        self.start_date_entry.bind("<Return>", lambda e: self._auto_complete_date(self.start_date_var))
        self.start_date_entry.bind("<KeyRelease>", lambda e: self._date_mask_on_key(e, self.start_date_var))

        Label(periodo_inner, text="FIM:", font=("Segoe UI", 9, "bold"), 
              bg="#34495e", fg="#ecf0f1").grid(row=0, column=2, sticky="e", padx=(10, 5), pady=8)
        self.end_date_var = StringVar(value=date.today().strftime("%d/%m/%Y"))
        self.end_date_entry = Entry(periodo_inner, textvariable=self.end_date_var, 
                                    width=13, font=("Segoe UI", 9), 
                                    relief="groove", bd=2, bg="white", 
                                    fg="#2c3e50", insertbackground="#3498db",
                                    validate="key", validatecommand=vcmd)
        self.end_date_entry.grid(row=0, column=3, padx=(5, 10), pady=8)
        self.end_date_entry.bind("<FocusOut>", lambda e: self._auto_complete_date(self.end_date_var))
        self.end_date_entry.bind("<Return>", lambda e: self._auto_complete_date(self.end_date_var))
        self.end_date_entry.bind("<KeyRelease>", lambda e: self._date_mask_on_key(e, self.end_date_var))

        # Botões com design moderno e cores vibrantes
        btn_style = {"font": ("Segoe UI", 9, "bold"), "fg": "white", 
                    "relief": "flat", "bd": 0, "cursor": "hand2", 
                    "padx": 14, "pady": 10}
        
        Button(top, text="🖼️ LOGO", command=self.select_logo_image, 
               bg="#9b59b6", **btn_style).grid(row=0, column=1, sticky="ew", padx=4, pady=4)
        
        Button(top, text="🗓️ FERIADOS", command=self.mark_holidays, 
               bg="#8e44ad", **btn_style).grid(row=0, column=2, sticky="ew", padx=4, pady=4)
        
        Button(top, text="🏙️ CIDADES", command=self.manage_cities, 
               bg="#3498db", **btn_style).grid(row=0, column=3, sticky="ew", padx=4, pady=4)
        
        Button(top, text="📂 CARREGAR", command=self.show_load_menu, 
               bg="#1abc9c", **btn_style).grid(row=0, column=4, sticky="ew", padx=4, pady=4)
        
        Button(top, text="🗑️ LIMPAR", command=self.clear_all_employees, 
               bg="#e67e22", **btn_style).grid(row=0, column=5, sticky="ew", padx=4, pady=4)
        
        Button(top, text="💾 SALVAR", command=self.save_config, 
               bg="#27ae60", **btn_style).grid(row=0, column=6, sticky="ew", padx=4, pady=4)
        
        # Botão destaque para gerar PDFs
        Button(top, text="📄 GERAR PDFs", command=self.generate_all_pdfs, 
               bg="#e74c3c", fg="white", font=("Segoe UI", 10, "bold"),
               relief="flat", bd=0, cursor="hand2", padx=18, pady=12
              ).grid(row=0, column=7, sticky="ew", padx=4, pady=4)

        # Barra de pesquisa moderna
        search_frame = Frame(top, bg="white", relief="groove", bd=2)
        search_frame.grid(row=1, column=0, columnspan=8, sticky="ew", pady=(8, 0))
        
        Label(search_frame, text="🔍", font=("Segoe UI", 13), bg="white"
             ).grid(row=0, column=0, padx=(12, 5), pady=10)
        
        self.search_var = StringVar()
        self.search_var.trace_add("write", lambda *_: self.update_employee_tree())
        search_entry = Entry(search_frame, textvariable=self.search_var, 
                            width=65, font=("Segoe UI", 10), 
                            relief="flat", bd=0, bg="white",
                            fg="#2c3e50", insertbackground="#3498db")
        search_entry.grid(row=0, column=1, sticky="ew", padx=5, pady=10)
        
        # Linha separadora sutil
        Frame(search_entry, height=2, bg="#bdc3c7").place(relx=0, rely=1, relwidth=1)
        
        Button(search_frame, text="✕", command=lambda: self.search_var.set(""), 
               bg="#95a5a6", fg="white", font=("Segoe UI", 9, "bold"), 
               relief="flat", bd=0, cursor="hand2", padx=12, pady=6,
               activebackground="#7f8c8d", activeforeground="white"
              ).grid(row=0, column=2, padx=(5, 12), pady=10)
        
        search_frame.columnconfigure(1, weight=1)
        for i in range(8):
            top.columnconfigure(i, weight=1)
        top.columnconfigure(0, weight=2)

    def _build_mid_frame(self):
        # Container da tabela com design moderno
        mid_container = Frame(self.root, bg="#ecf0f1", relief="flat", bd=0)
        mid_container.pack(fill=BOTH, expand=True, padx=10, pady=10)
        
        # Título da seção com gradiente
        title_frame = Frame(mid_container, bg="#2c3e50")
        title_frame.pack(fill=X)
        Label(title_frame, text="👥 LISTA DE FUNCIONÁRIOS", font=("Segoe UI", 13, "bold"),
              bg="#2c3e50", fg="#ecf0f1", pady=12).pack()
        
        # Frame de estatísticas moderno
        stats_frame = Frame(mid_container, bg="#ecf0f1")
        stats_frame.pack(fill=X, padx=5, pady=(8, 5))
        
        # Labels de estatísticas com design card
        self.label_revisados = Label(stats_frame, text="✅ REVISADOS: 0", 
                                     font=("Segoe UI", 10, "bold"), 
                                     bg="#27ae60", fg="white", 
                                     padx=18, pady=8, relief="flat", bd=0)
        self.label_revisados.pack(side="left", padx=8, pady=5)
        
        self.label_nao_revisados = Label(stats_frame, text="⚠ NÃO REVISADOS: 0", 
                                         font=("Segoe UI", 10, "bold"), 
                                         bg="#e74c3c", fg="white", 
                                         padx=18, pady=8, relief="flat", bd=0)
        self.label_nao_revisados.pack(side="left", padx=8, pady=5)
        
        # Frame da tabela
        mid = Frame(mid_container, bg="white")
        mid.pack(fill=BOTH, expand=True, padx=2, pady=2)

        # Estilo moderno para Treeview
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Treeview", 
                       background="#ffffff",
                       foreground="#2c3e50",
                       rowheight=45,
                       fieldbackground="#ffffff",
                       font=("Segoe UI", 9),
                       borderwidth=0)
        style.configure("Treeview.Heading",
                       background="#34495e",
                       foreground="white",
                       font=("Segoe UI", 9, "bold"),
                       relief="flat",
                       borderwidth=0,
                       padding=10)
        style.map("Treeview.Heading",
                 background=[("active", "#2980b9")])
        style.map("Treeview",
                 background=[("selected", "#3498db")],
                 foreground=[("selected", "white")])
        
        # Criação da Treeview com cores alternadas
        self.emp_tree = ttk.Treeview(mid, columns=("mat", "cpf", "scale", "posto", "revisado"), show="tree headings", height=16)
        self.emp_tree.tag_configure('oddrow', background='#f8f9fa')
        self.emp_tree.tag_configure('evenrow', background='#ffffff')
        
        self.emp_tree.heading("#0", text="FUNCIONÁRIO", command=lambda: self.sort_emp_tree("#0"))
        self.emp_tree.column("#0", width=400, anchor="w")
        self.emp_tree.heading("mat", text="MATRÍCULA", command=lambda: self.sort_emp_tree("mat")); self.emp_tree.column("mat", width=90, anchor="center")
        self.emp_tree.heading("cpf", text="CPF", command=lambda: self.sort_emp_tree("cpf")); self.emp_tree.column("cpf", width=140, anchor="center")
        self.emp_tree.heading("scale", text="ESCALA", command=lambda: self.sort_emp_tree("scale")); self.emp_tree.column("scale", width=140, anchor="center")
        self.emp_tree.heading("posto", text="POSTO", command=lambda: self.sort_emp_tree("posto")); self.emp_tree.column("posto", width=100, anchor="center")
        self.emp_tree.heading("revisado", text="REVISADO"); self.emp_tree.column("revisado", width=90, anchor="center")

        self.emp_tree.pack(side=LEFT, fill=BOTH, expand=True)
        self.emp_tree.bind("<Double-1>", self.on_tree_double)
        self.emp_tree.bind("<Button-1>", self.on_tree_click)

        sb = ttk.Scrollbar(mid, orient="vertical", command=self.emp_tree.yview)
        self.emp_tree.configure(yscroll=sb.set)
        sb.pack(side=LEFT, fill=Y)

    def _build_bottom_frame(self):
        # Rodapé moderno com dica
        footer_frame = Frame(self.root, bg="#34495e", relief="flat", bd=0)
        footer_frame.pack(fill=X, padx=0, pady=0)
        
        Label(footer_frame, text="💡 DICA: DÊ DUPLO-CLIQUE NO FUNCIONÁRIO PARA CONFIGURAR HORÁRIOS POR DIA.", 
              font=("Segoe UI", 9, "italic"), bg="#34495e", fg="#ecf0f1", pady=10).pack()

    # ---------------------------
    # FUNÇÕES DE AÇÃO
    # ---------------------------
    def clear_all_employees(self):
        """
        Limpa lista de funcionários atual (interface e memória em runtime),
        não remove do JSON salvo até o usuário clicar SALVAR (decisão).
        """
        if not self.funcionarios:
            messagebox.showinfo("LIMPAR", "NÃO HÁ FUNCIONÁRIOS PARA LIMPAR.")
            return
        if not messagebox.askyesno("CONFIRMAR", "DESEJA REMOVER TODOS OS FUNCIONÁRIOS DA TELA ATUAL? (ESSA AÇÃO NÃO EXCLUI OS DADOS SALVOS NO ARQUIVO JSON ATÉ VOCÊ SALVAR)"):
            return
        self.funcionarios = []
        self.update_employee_tree()
        messagebox.showinfo("LIMPO", "LISTA DE FUNCIONÁRIOS LIMPA. CARREGUE UMA NOVA PLANILHA PARA ADICIONAR.")

    def manage_cities(self):
        """Gerenciador de cidades - cadastra cidades e vincula postos a elas"""
        # IMPORTANTE: Reconstrói o histórico de postos a partir dos funcionários carregados
        # para garantir que sempre mostramos NOMES DOS POSTOS e não CPFs
        if hasattr(self, 'funcionarios') and self.funcionarios:
            self.all_postos_historico.clear()
            postos_atuais = set(
                emp.get("posto", "").strip()
                for emp in self.funcionarios
                if emp.get("posto", "").strip() and not self._is_cpf_format(emp.get("posto", "").strip())
            )
            self.all_postos_historico.update(postos_atuais)
            # Salva o histórico limpo
            self.store["all_postos_historico"] = sorted(list(self.all_postos_historico))
            save_store(self.store)
        
        top = Toplevel(self.root)
        top.title("GERENCIAR CIDADES E POSTOS")
        top.geometry("850x650")
        top.configure(bg="#ecf0f1")
        top.transient(self.root)
        top.grab_set()
        
        # Título moderno
        title_frame = Frame(top, bg="#2c3e50")
        title_frame.pack(fill=X)
        Label(title_frame, text="🏙️ GERENCIAR CIDADES E POSTOS", 
              font=("Segoe UI", 13, "bold"), bg="#2c3e50", fg="white", 
              pady=15).pack()
        
        # Frame principal com duas colunas (lista + botões)
        main_frame = Frame(top, bg="#ecf0f1")
        main_frame.pack(fill=BOTH, expand=True, padx=20, pady=15)
        
        # COLUNA ESQUERDA: Lista de cidades
        left_frame = Frame(main_frame, bg="#ecf0f1")
        left_frame.pack(side=LEFT, fill=BOTH, expand=True, padx=(0, 15))
        
        Label(left_frame, text="CIDADES CADASTRADAS", 
              font=("Segoe UI", 11, "bold"), bg="#ecf0f1", 
              fg="#2c3e50").pack(pady=8)
        
        # Listbox de cidades com scroll
        list_container = Frame(left_frame)
        list_container.pack(fill=BOTH, expand=True)
        
        scrollbar = ttk.Scrollbar(list_container, orient="vertical")
        cities_listbox = ttk.Treeview(list_container, columns=("cidade", "postos"), show="headings", 
                                      yscrollcommand=scrollbar.set, height=15)
        scrollbar.config(command=cities_listbox.yview)
        
        cities_listbox.heading("cidade", text="CIDADE")
        cities_listbox.heading("postos", text="Nº POSTOS")
        cities_listbox.column("cidade", width=300)
        cities_listbox.column("postos", width=100)
        
        cities_listbox.pack(side=LEFT, fill=BOTH, expand=True)
        scrollbar.pack(side=RIGHT, fill=Y)
        
        def refresh_cities():
            cities_listbox.delete(*cities_listbox.get_children())
            for cidade, postos in sorted(self.cidades.items()):
                cities_listbox.insert("", "end", values=(cidade.upper(), len(postos)))
        
        refresh_cities()
        
        # Evento de duplo clique para editar postos
        def on_double_click(event):
            edit_city_postos()
        
        cities_listbox.bind("<Double-Button-1>", on_double_click)
        
        # COLUNA DIREITA: Botões de ação
        right_frame = Frame(main_frame, bg="#ecf0f1")
        right_frame.pack(side=RIGHT, fill=Y, padx=(15, 0))
        
        Label(right_frame, text="AÇÕES", 
              font=("Segoe UI", 11, "bold"), bg="#ecf0f1", 
              fg="#2c3e50").pack(pady=(0, 15))
        
        def add_city():
            dlg = Toplevel(top)
            dlg.title("ADICIONAR CIDADE")
            dlg.geometry("550x320")
            dlg.configure(bg="#ecf0f1")
            dlg.transient(top)
            dlg.grab_set()
            dlg.resizable(False, False)
            
            # Centralizar
            dlg.update_idletasks()
            x = (dlg.winfo_screenwidth() // 2) - 275
            y = (dlg.winfo_screenheight() // 2) - 160
            dlg.geometry(f"550x320+{x}+{y}")
            
            # Título
            title_add = Frame(dlg, bg="#27ae60")
            title_add.pack(fill=X)
            Label(title_add, text="🏙️ CADASTRAR NOVA CIDADE", 
                  font=("Segoe UI", 13, "bold"), bg="#27ae60", fg="white", 
                  pady=15).pack()
            
            # Frame principal
            content = Frame(dlg, bg="white", relief="groove", bd=2)
            content.pack(fill=BOTH, expand=True, padx=20, pady=20)
            
            Label(content, text="NOME DA CIDADE:", bg="white", 
                  font=("Segoe UI", 11, "bold"), fg="#2c3e50").pack(pady=(20, 5))
            
            city_var = StringVar()
            entry = Entry(content, textvariable=city_var, width=40, 
                         font=("Segoe UI", 11), relief="solid", bd=1,
                         justify="center")
            entry.pack(pady=5, ipady=8)
            entry.focus_set()
            
            def confirm_add():
                city_name = city_var.get().strip().upper()
                if not city_name:
                    messagebox.showwarning("AVISO", "Digite o nome da cidade.")
                    return
                if city_name in self.cidades:
                    messagebox.showwarning("AVISO", f"A cidade '{city_name}' já está cadastrada.")
                    return
                self.cidades[city_name] = []
                refresh_cities()
                messagebox.showinfo("✓ Sucesso", f"Cidade '{city_name}' cadastrada com sucesso!")
                dlg.destroy()
            
            # Enter para confirmar
            entry.bind("<Return>", lambda e: confirm_add())
            
            # Separador visual
            separator = Frame(content, bg="#d5dbdb", height=2)
            separator.pack(fill=X, padx=30, pady=(30, 25))
            
            # Botões modernos com sombra e espaçamento
            btn_container = Frame(content, bg="white")
            btn_container.pack(pady=(0, 25))
            
            # Botão Cadastrar (verde, destaque)
            btn_cadastrar = Button(btn_container, text="✓  CADASTRAR", command=confirm_add,
                                   bg="#27ae60", fg="white", font=("Segoe UI", 11, "bold"),
                                   relief="solid", bd=1, cursor="hand2", 
                                   padx=40, pady=14, width=14,
                                   activebackground="#229954", activeforeground="white",
                                   highlightthickness=0)
            btn_cadastrar.pack(side=LEFT, padx=10)
            
            # Botão Cancelar (cinza)
            btn_cancelar = Button(btn_container, text="✖  CANCELAR", command=dlg.destroy,
                                  bg="#bdc3c7", fg="#2c3e50", font=("Segoe UI", 11, "bold"),
                                  relief="solid", bd=1, cursor="hand2", 
                                  padx=40, pady=14, width=14,
                                  activebackground="#95a5a6", activeforeground="white",
                                  highlightthickness=0)
            btn_cancelar.pack(side=LEFT, padx=10)
        
        def remove_city():
            selection = cities_listbox.selection()
            if not selection:
                messagebox.showwarning("AVISO", "Selecione uma cidade para remover.")
                return
            
            cidade = cities_listbox.item(selection[0])["values"][0]
            if messagebox.askyesno("CONFIRMAR", f"Remover a cidade '{cidade}'?\n\nIsso também removerá a associação com feriados locais."):
                # Remove cidade
                if cidade in self.cidades:
                    del self.cidades[cidade]
                # Remove associações com feriados
                feriados_to_remove = [d for d, c in self.holiday_cidades.items() if c == cidade]
                for d in feriados_to_remove:
                    del self.holiday_cidades[d]
                refresh_cities()
        
        def edit_city_postos():
            selection = cities_listbox.selection()
            if not selection:
                messagebox.showwarning("AVISO", "Selecione uma cidade para editar postos.")
                return
            
            cidade = cities_listbox.item(selection[0])["values"][0]
            
            # Diálogo para editar postos da cidade - LAYOUT HORIZONTAL
            dlg = Toplevel(top)
            dlg.title(f"POSTOS DA CIDADE: {cidade}")
            dlg.geometry("1100x650")
            dlg.configure(bg="#ecf0f1")
            dlg.transient(top)
            dlg.grab_set()
            
            # Centralizar janela
            dlg.update_idletasks()
            x = (dlg.winfo_screenwidth() // 2) - (550)
            y = (dlg.winfo_screenheight() // 2) - (325)
            dlg.geometry(f"1100x650+{x}+{y}")
            
            # Título
            title_dlg = Frame(dlg, bg="#3498db")
            title_dlg.pack(fill=X)
            Label(title_dlg, text=f"🏙️ GERENCIAR POSTOS - {cidade}", 
                  font=("Segoe UI", 14, "bold"), bg="#3498db", fg="white", 
                  pady=18).pack()
            
            # Pega TODOS os postos do histórico (não apenas da planilha atual)
            postos_unicos = sorted(list(self.all_postos_historico))
            if not postos_unicos:
                messagebox.showwarning("AVISO", "Não há postos cadastrados. Carregue a planilha primeiro.")
                dlg.destroy()
                return
            
            # Carrega postos já associados à cidade atual
            postos_cidade_atual = self.cidades.get(cidade, [])
            
            # Identifica postos já usados em OUTRAS cidades
            postos_usados_outras_cidades = set()
            for outra_cidade, postos_lista in self.cidades.items():
                if outra_cidade != cidade:
                    postos_usados_outras_cidades.update(postos_lista)
            
            # Postos disponíveis = postos não usados em outras cidades
            postos_disponiveis = [p for p in postos_unicos if p not in postos_usados_outras_cidades]
            
            # Separa em duas listas: postos livres e postos da cidade
            postos_livres = sorted([p for p in postos_disponiveis if p not in postos_cidade_atual])
            postos_na_cidade = sorted(postos_cidade_atual)
            
            # Frame principal com 3 colunas
            main_container = Frame(dlg, bg="#ecf0f1")
            main_container.pack(fill=BOTH, expand=True, padx=25, pady=20)
            
            # COLUNA ESQUERDA: Postos Disponíveis
            left_column = Frame(main_container, bg="#ecf0f1")
            left_column.pack(side=LEFT, fill=BOTH, expand=True, padx=(0, 15))
            
            Label(left_column, text="📋 POSTOS DISPONÍVEIS", 
                  font=("Segoe UI", 11, "bold"), bg="#ecf0f1", fg="#2c3e50").pack(pady=8)
            
            left_frame = Frame(left_column, bg="white", relief="groove", bd=2)
            left_frame.pack(fill=BOTH, expand=True)
            
            left_scroll = ttk.Scrollbar(left_frame, orient="vertical")
            left_listbox = Listbox(left_frame, yscrollcommand=left_scroll.set, 
                                   selectmode=EXTENDED, font=("Segoe UI", 10),
                                   bg="white", fg="#2c3e50", selectbackground="#3498db",
                                   activestyle="none", relief="flat", bd=0, highlightthickness=1,
                                   highlightcolor="#3498db")
            left_scroll.config(command=left_listbox.yview)
            left_listbox.pack(side=LEFT, fill=BOTH, expand=True, padx=2, pady=2)
            left_scroll.pack(side=RIGHT, fill=Y)
            
            for posto in postos_livres:
                left_listbox.insert(END, posto.upper())
            
            Label(left_column, text=f"Total: {len(postos_livres)}", 
                  font=("Segoe UI", 9, "bold"), bg="#ecf0f1", fg="#7f8c8d").pack(pady=5)
            
            # COLUNA CENTRAL: Botões de Transferência
            center_column = Frame(main_container, bg="#ecf0f1")
            center_column.pack(side=LEFT, padx=15)
            
            # Espaçador para centralizar verticalmente
            Frame(center_column, bg="#ecf0f1").pack(expand=True)
            
            def mover_para_direita():
                """Move postos selecionados da esquerda para direita"""
                selecionados = left_listbox.curselection()
                if not selecionados:
                    return
                
                # Pega os postos selecionados
                postos_mover = [left_listbox.get(idx) for idx in reversed(selecionados)]
                
                # Remove da esquerda
                for idx in reversed(selecionados):
                    left_listbox.delete(idx)
                
                # Adiciona na direita
                for posto in postos_mover:
                    right_listbox.insert(END, posto)
                
                # Atualiza contadores
                left_column.winfo_children()[-1].config(text=f"Total: {left_listbox.size()}")
                right_column.winfo_children()[-1].config(text=f"Total: {right_listbox.size()}")
            
            def mover_para_esquerda():
                """Move postos selecionados da direita para esquerda"""
                selecionados = right_listbox.curselection()
                if not selecionados:
                    return
                
                # Pega os postos selecionados
                postos_mover = [right_listbox.get(idx) for idx in reversed(selecionados)]
                
                # Remove da direita
                for idx in reversed(selecionados):
                    right_listbox.delete(idx)
                
                # Adiciona na esquerda
                for posto in postos_mover:
                    left_listbox.insert(END, posto)
                
                # Atualiza contadores
                left_column.winfo_children()[-1].config(text=f"Total: {left_listbox.size()}")
                right_column.winfo_children()[-1].config(text=f"Total: {right_listbox.size()}")
            
            btn_style = {"font": ("Segoe UI", 14, "bold"), "width": 6,
                        "bg": "#3498db", "fg": "white", "relief": "flat",
                        "bd": 0, "cursor": "hand2", "pady": 12}
            
            Button(center_column, text="►", command=mover_para_direita, **btn_style).pack(pady=8)
            Button(center_column, text="◄", command=mover_para_esquerda, **btn_style).pack(pady=8)
            
            Frame(center_column, bg="#ecf0f1").pack(expand=True)
            
            # COLUNA DIREITA: Postos da Cidade
            right_column = Frame(main_container, bg="#ecf0f1")
            right_column.pack(side=LEFT, fill=BOTH, expand=True, padx=(15, 0))
            
            Label(right_column, text=f"🏙️ POSTOS EM {cidade.upper()}", 
                  font=("Segoe UI", 11, "bold"), bg="#ecf0f1", fg="#27ae60").pack(pady=8)
            
            right_frame = Frame(right_column, bg="white", relief="groove", bd=2)
            right_frame.pack(fill=BOTH, expand=True)
            
            right_scroll = ttk.Scrollbar(right_frame, orient="vertical")
            right_listbox = Listbox(right_frame, yscrollcommand=right_scroll.set,
                                    selectmode=EXTENDED, font=("Segoe UI", 10),
                                    bg="white", fg="#27ae60", selectbackground="#27ae60",
                                    activestyle="none", relief="flat", bd=0, highlightthickness=1,
                                    highlightcolor="#27ae60")
            right_scroll.config(command=right_listbox.yview)
            right_listbox.pack(side=LEFT, fill=BOTH, expand=True, padx=2, pady=2)
            right_scroll.pack(side=RIGHT, fill=Y)
            
            for posto in postos_na_cidade:
                right_listbox.insert(END, posto.upper())
            
            Label(right_column, text=f"Total: {len(postos_na_cidade)}", 
                  font=("Segoe UI", 9, "bold"), bg="#ecf0f1", fg="#7f8c8d").pack(pady=5)
            
            # Botões de ação
            btn_frame_dlg = Frame(dlg, bg="#ecf0f1")
            btn_frame_dlg.pack(pady=20)
            
            def salvar_alteracoes():
                """Salva os postos da cidade"""
                # Pega todos os postos da listbox direita
                novos_postos = [right_listbox.get(idx) for idx in range(right_listbox.size())]
                # Converte de volta para lowercase (como está armazenado)
                novos_postos = [p.lower() for p in novos_postos]
                
                self.cidades[cidade] = novos_postos
                
                # Atualiza holiday_postos para feriados locais desta cidade
                for data_feriado, cidade_feriado in self.holiday_cidades.items():
                    if cidade_feriado == cidade:
                        self.holiday_postos[data_feriado] = novos_postos
                
                refresh_cities()
                messagebox.showinfo("✓ Salvo", f"Postos da cidade {cidade} atualizados com sucesso!")
                dlg.destroy()
            
            Button(btn_frame_dlg, text="💾 SALVAR", command=salvar_alteracoes,
                   bg="#27ae60", fg="white", font=("Segoe UI", 10, "bold"),
                   relief="flat", bd=0, cursor="hand2", padx=30, pady=10).pack(side=LEFT, padx=5)
            
            Button(btn_frame_dlg, text="✖ CANCELAR", command=dlg.destroy,
                   bg="#95a5a6", fg="white", font=("Segoe UI", 10, "bold"),
                   relief="flat", bd=0, cursor="hand2", padx=30, pady=10).pack(side=LEFT, padx=5)
        
        def show_postos_sem_cidade():
            """Mostra lista de postos que não estão vinculados a nenhuma cidade"""
            # Pega TODOS os postos do histórico (não apenas da planilha atual)
            postos_unicos = sorted(list(self.all_postos_historico))
            if not postos_unicos:
                messagebox.showwarning("AVISO", "Não há postos cadastrados. Carregue a planilha primeiro.")
                return
            
            # Identifica postos já vinculados a alguma cidade
            postos_vinculados = set()
            for cidade, postos_lista in self.cidades.items():
                postos_vinculados.update(postos_lista)
            
            # Postos sem cidade = postos únicos - postos vinculados
            postos_sem_cidade = [p for p in postos_unicos if p not in postos_vinculados]
            
            # Diálogo para mostrar postos sem cidade
            dlg = Toplevel(top)
            dlg.title("POSTOS SEM CIDADE CADASTRADA")
            dlg.geometry("600x500")
            dlg.transient(top)
            dlg.grab_set()
            
            if not postos_sem_cidade:
                Label(dlg, text="✅ TODOS OS POSTOS ESTÃO VINCULADOS A UMA CIDADE!", 
                      font=("Helvetica", 12, "bold"), fg="green").pack(pady=30)
            else:
                Label(dlg, text=f"POSTOS SEM CIDADE: {len(postos_sem_cidade)}", 
                      font=("Helvetica", 12, "bold")).pack(pady=10)
                
                # Frame com lista de postos e scrollbar
                list_container = Frame(dlg)
                list_container.pack(fill=BOTH, expand=True, padx=20, pady=10)
                
                scrollbar = ttk.Scrollbar(list_container, orient="vertical")
                postos_tree = ttk.Treeview(list_container, columns=("posto",), show="headings", 
                                          yscrollcommand=scrollbar.set)
                scrollbar.config(command=postos_tree.yview)
                
                postos_tree.heading("posto", text="POSTO")
                postos_tree.column("posto", width=500)
                
                postos_tree.pack(side=LEFT, fill=BOTH, expand=True)
                scrollbar.pack(side=RIGHT, fill=Y)
                
                # Popula lista
                for posto in postos_sem_cidade:
                    postos_tree.insert("", "end", values=(posto.upper(),))
            
            Button(dlg, text="FECHAR", command=dlg.destroy).pack(pady=10)
        
        # Botões modernos em coluna vertical
        btn_style = {"fg": "white", "font": ("Segoe UI", 10, "bold"), 
                    "relief": "flat", "bd": 0, "cursor": "hand2", 
                    "padx": 20, "pady": 12, "width": 18}
        
        Button(right_frame, text="➕ ADICIONAR CIDADE", command=add_city, 
               bg="#27ae60", **btn_style).pack(fill=X, pady=(0, 10))
        Button(right_frame, text="✏️ EDITAR POSTOS", command=edit_city_postos, 
               bg="#3498db", **btn_style).pack(fill=X, pady=(0, 10))
        Button(right_frame, text="🗑️ REMOVER CIDADE", command=remove_city, 
               bg="#e74c3c", **btn_style).pack(fill=X, pady=(0, 10))
        Button(right_frame, text="📋 POSTOS LIVRES", command=show_postos_sem_cidade, 
               bg="#95a5a6", **btn_style).pack(fill=X, pady=(0, 10))
        
        # Botão fechar
        Button(top, text="✓ FECHAR", command=top.destroy,
               bg="#34495e", fg="white", font=("Segoe UI", 10, "bold"),
               relief="flat", bd=0, cursor="hand2", padx=25, pady=10).pack(pady=15)

    def mark_holidays(self):
        ans = messagebox.askyesno("CARREGAR FERIADOS", "DESEJA CARREGAR O ÚLTIMO CONJUNTO DE FERIADOS SALVOS?")
        if ans:
            # já está em self.global_holidays via load_store
            pass
        top = Toplevel(self.root)
        top.title("GERENCIAR FERIADOS")
        top.geometry("750x600")
        top.configure(bg="#ecf0f1")
        top.transient(self.root)
        top.grab_set()
        
        # Título moderno
        title_frame = Frame(top, bg="#2c3e50")
        title_frame.pack(fill=X)
        Label(title_frame, text="🗓️ GERENCIAR FERIADOS", 
              font=("Segoe UI", 13, "bold"), bg="#2c3e50", fg="white", 
              pady=15).pack()
        
        # Card para formulário
        frm = Frame(top, bg="white", relief="groove", bd=2)
        frm.pack(padx=15, pady=15, fill=X)
        
        frm_inner = Frame(frm, bg="white")
        frm_inner.pack(padx=15, pady=15, fill=X)
        
        Label(frm_inner, text="📅 DATA:", bg="white", fg="#2c3e50",
              font=("Segoe UI", 10, "bold")).grid(row=0, column=0, sticky="w", padx=8, pady=8)
        try:
            date_entry = DateEntry(
                frm_inner,
                date_pattern="dd/mm/yyyy",
                locale="pt_BR",
                showweeknumbers=False,
                firstweekday="sunday",
                selectmode="day",
                state='normal',
                cursor='hand2',
                takefocus=True,
                selectbackground='gray85',
                selectforeground='black',
                normalbackground='white',
                normalforeground='black',
                background='white',
                foreground='black',
                borderwidth=2,
                relief="groove",
                font=("Segoe UI", 9)
            )
        except TypeError:
            date_entry = DateEntry(frm_inner, date_pattern="dd/mm/yyyy", state='normal')
        date_entry.grid(row=0, column=1, padx=8, pady=8, sticky="w")
        
        Label(frm_inner, text="📝 NOME:", bg="white", fg="#2c3e50",
              font=("Segoe UI", 10, "bold")).grid(row=1, column=0, sticky="w", padx=8, pady=8)
        name_var = StringVar()
        Entry(frm_inner, textvariable=name_var, width=35, font=("Segoe UI", 9),
              relief="groove", bd=2).grid(row=1, column=1, padx=8, pady=8, sticky="w")
        
        Label(frm_inner, text="🏷️ TIPO:", bg="white", fg="#2c3e50",
              font=("Segoe UI", 10, "bold")).grid(row=2, column=0, sticky="w", padx=8, pady=8)
        tipo_var = StringVar(value="NACIONAL")
        tipo_frame = Frame(frm_inner, bg="white")
        tipo_frame.grid(row=2, column=1, sticky="w", padx=8, pady=8)
        ttk.Radiobutton(tipo_frame, text="🌍 NACIONAL (todos folgam)", 
                       variable=tipo_var, value="NACIONAL").pack(side=LEFT, padx=6)
        ttk.Radiobutton(tipo_frame, text="🏙️ LOCAL (selecionar cidade)", 
                       variable=tipo_var, value="LOCAL").pack(side=LEFT, padx=6)
        
        def add():
            d = date_entry.get_date().strftime("%Y-%m-%d")
            name = name_var.get().strip().upper() or "FERIADO"
            tipo = tipo_var.get()
            
            if tipo == "LOCAL":
                # Abre janela para selecionar cidade
                if not self.cidades:
                    messagebox.showwarning("AVISO", "Não há cidades cadastradas. Cadastre as cidades primeiro no menu 'GERENCIAR CIDADES'.")
                    return
                
                cidade_dlg = Toplevel(top)
                cidade_dlg.title(f"Selecionar Cidade - {name}")
                cidade_dlg.geometry("400x300")
                cidade_dlg.transient(top)
                cidade_dlg.grab_set()
                
                Label(cidade_dlg, text="Selecione a CIDADE que tem este feriado:", 
                      font=("Helvetica", 10, "bold")).pack(pady=10)
                
                # Listbox com cidades
                list_frame = Frame(cidade_dlg)
                list_frame.pack(padx=12, pady=8, fill=BOTH, expand=True)
                
                scrollbar = ttk.Scrollbar(list_frame, orient="vertical")
                cidade_listbox = Listbox(list_frame, yscrollcommand=scrollbar.set, height=10)
                scrollbar.config(command=cidade_listbox.yview)
                
                cidade_listbox.pack(side=LEFT, fill=BOTH, expand=True)
                scrollbar.pack(side=RIGHT, fill=Y)
                
                # Popula listbox
                cidades_list = sorted(self.cidades.keys())
                for cidade in cidades_list:
                    cidade_listbox.insert("end", cidade)
                
                # Se feriado já existe, seleciona cidade atual
                cidade_atual = self.holiday_cidades.get(d)
                if cidade_atual and cidade_atual in cidades_list:
                    idx = cidades_list.index(cidade_atual)
                    cidade_listbox.selection_set(idx)
                    cidade_listbox.see(idx)
                
                def confirmar_cidade():
                    selection = cidade_listbox.curselection()
                    if not selection:
                        messagebox.showwarning("AVISO", "Selecione uma cidade.")
                        return
                    
                    cidade_selecionada = cidades_list[selection[0]]
                    postos_da_cidade = self.cidades.get(cidade_selecionada, [])
                    
                    if not postos_da_cidade:
                        if not messagebox.askyesno("CONFIRMAR", 
                            f"A cidade '{cidade_selecionada}' não tem postos vinculados.\n\n"
                            "Este feriado LOCAL não se aplicará a ninguém. Continuar?"):
                            return
                    
                    self.global_holidays[d] = name
                    self.holiday_type[d] = "LOCAL"
                    self.holiday_cidades[d] = cidade_selecionada
                    # Atualiza holiday_postos com os postos da cidade
                    self.holiday_postos[d] = postos_da_cidade
                    cidade_dlg.destroy()
                    refresh()
                
                def cancelar_cidade():
                    cidade_dlg.destroy()
                
                btn_frame_dlg = Frame(cidade_dlg)
                btn_frame_dlg.pack(pady=10)
                Button(btn_frame_dlg, text="CONFIRMAR", command=confirmar_cidade).pack(side=LEFT, padx=6)
                Button(btn_frame_dlg, text="CANCELAR", command=cancelar_cidade).pack(side=LEFT, padx=6)
            else:
                # Feriado NACIONAL
                self.global_holidays[d] = name
                self.holiday_type[d] = "NACIONAL"
                if d in self.holiday_postos:
                    del self.holiday_postos[d]
                if d in self.holiday_cidades:
                    del self.holiday_cidades[d]
                refresh()
        
        Button(frm_inner, text="➕ ADICIONAR FERIADO", command=add,
               bg="#27ae60", fg="white", font=("Segoe UI", 10, "bold"),
               relief="flat", bd=0, cursor="hand2", padx=25, pady=10
              ).grid(row=3, column=0, columnspan=2, pady=15)
        
        tree = ttk.Treeview(top, columns=("data", "nome", "tipo", "cidade"), show="headings", height=12)
        tree.heading("data", text="DATA")
        tree.heading("nome", text="NOME")
        tree.heading("tipo", text="TIPO")
        tree.heading("cidade", text="CIDADE (LOCAL)")
        tree.column("data", width=100, anchor="center")
        tree.column("nome", width=250, anchor="w")
        tree.column("tipo", width=100, anchor="center")
        tree.column("cidade", width=220, anchor="w")
        tree.pack(padx=8, pady=8, fill=BOTH, expand=True)
        
        def refresh():
            for i in tree.get_children():
                tree.delete(i)
            for k, v in sorted(self.global_holidays.items()):
                try:
                    display = datetime.strptime(k, "%Y-%m-%d").strftime("%d/%m/%Y")
                except:
                    display = k
                tipo = self.holiday_type.get(k, "NACIONAL")
                cidade = self.holiday_cidades.get(k, "-") if tipo == "LOCAL" else "-"
                tree.insert("", "end", values=(display, v, tipo, cidade))
        
        def edit_selected(event=None):
            """Abre janela para editar feriado selecionado"""
            sel = tree.selection()
            if not sel:
                return
            
            val = tree.item(sel[0], "values")
            date_part = val[0]
            nome_feriado = val[1]
            tipo_feriado = val[2]
            
            try:
                dt = datetime.strptime(date_part, "%d/%m/%Y").date()
                key = dt.strftime("%Y-%m-%d")
            except Exception:
                key = date_part
                dt = None
            
            if key not in self.global_holidays:
                return
            
            # Janela de edição
            edit_dlg = Toplevel(top)
            edit_dlg.title(f"EDITAR FERIADO: {nome_feriado}")
            edit_dlg.geometry("500x400")
            edit_dlg.transient(top)
            edit_dlg.grab_set()
            
            Label(edit_dlg, text="EDITAR FERIADO", font=("Helvetica", 14, "bold")).pack(pady=10)
            
            # Frame de edição
            edit_frame = Frame(edit_dlg)
            edit_frame.pack(pady=10, padx=20, fill=BOTH, expand=True)
            
            # Data (editável)
            Label(edit_frame, text="DATA:", font=("Helvetica", 10, "bold")).grid(row=0, column=0, sticky="w", pady=5)
            data_entry = DateEntry(edit_frame, width=37, background='darkblue', foreground='white', 
                                  borderwidth=2, date_pattern='dd/mm/yyyy', locale='pt_BR',
                                  state='normal', cursor='hand2', takefocus=True,
                                  showweeknumbers=False, firstweekday='sunday', selectmode='day')
            if dt:
                data_entry.set_date(dt)
            data_entry.grid(row=0, column=1, sticky="w", pady=5)
            
            # Nome
            Label(edit_frame, text="NOME:", font=("Helvetica", 10, "bold")).grid(row=1, column=0, sticky="w", pady=5)
            nome_var = StringVar(value=nome_feriado)
            Entry(edit_frame, textvariable=nome_var, width=40, font=("Helvetica", 10)).grid(row=1, column=1, sticky="w", pady=5)
            
            # Tipo
            Label(edit_frame, text="TIPO:", font=("Helvetica", 10, "bold")).grid(row=2, column=0, sticky="w", pady=5)
            tipo_var = StringVar(value=tipo_feriado)
            tipo_frame_edit = Frame(edit_frame)
            tipo_frame_edit.grid(row=2, column=1, sticky="w", pady=5)
            ttk.Radiobutton(tipo_frame_edit, text="NACIONAL", variable=tipo_var, value="NACIONAL").pack(side=LEFT, padx=5)
            ttk.Radiobutton(tipo_frame_edit, text="LOCAL", variable=tipo_var, value="LOCAL").pack(side=LEFT, padx=5)
            
            # Cidade (para LOCAL)
            Label(edit_frame, text="CIDADE:", font=("Helvetica", 10, "bold")).grid(row=3, column=0, sticky="w", pady=5)
            cidade_var = StringVar(value=self.holiday_cidades.get(key, ""))
            cidade_combo = ttk.Combobox(edit_frame, textvariable=cidade_var, values=sorted(self.cidades.keys()), 
                                       state="readonly" if tipo_feriado == "LOCAL" else "disabled", width=37)
            cidade_combo.grid(row=3, column=1, sticky="w", pady=5)
            
            def update_cidade_state(*args):
                if tipo_var.get() == "LOCAL":
                    cidade_combo.config(state="readonly")
                else:
                    cidade_combo.config(state="disabled")
                    cidade_var.set("")
            
            tipo_var.trace("w", update_cidade_state)
            
            # Botões
            btn_frame = Frame(edit_dlg)
            btn_frame.pack(pady=20)
            
            def save_changes():
                nova_data = data_entry.get_date()
                nova_key = nova_data.strftime("%Y-%m-%d")
                novo_nome = nome_var.get().strip().upper() or "FERIADO"
                novo_tipo = tipo_var.get()
                
                if novo_tipo == "LOCAL":
                    nova_cidade = cidade_var.get()
                    if not nova_cidade:
                        messagebox.showwarning("AVISO", "Selecione uma cidade para feriado LOCAL.")
                        return
                    if nova_cidade not in self.cidades:
                        messagebox.showwarning("AVISO", "Cidade não encontrada. Cadastre a cidade primeiro.")
                        return
                
                # Se a data mudou, remove a entrada antiga
                if nova_key != key:
                    if key in self.global_holidays:
                        del self.global_holidays[key]
                    if key in self.holiday_type:
                        del self.holiday_type[key]
                    if key in self.holiday_postos:
                        del self.holiday_postos[key]
                    if key in self.holiday_cidades:
                        del self.holiday_cidades[key]
                
                # Salva com a nova chave (ou mesma se não mudou)
                if novo_tipo == "LOCAL":
                    self.global_holidays[nova_key] = novo_nome
                    self.holiday_type[nova_key] = "LOCAL"
                    self.holiday_cidades[nova_key] = nova_cidade
                    self.holiday_postos[nova_key] = self.cidades.get(nova_cidade, [])
                else:
                    # Feriado nacional
                    self.global_holidays[nova_key] = novo_nome
                    self.holiday_type[nova_key] = "NACIONAL"
                    if nova_key in self.holiday_cidades:
                        del self.holiday_cidades[nova_key]
                    if nova_key in self.holiday_postos:
                        del self.holiday_postos[nova_key]
                
                refresh()
                edit_dlg.destroy()
                messagebox.showinfo("SUCESSO", "Feriado atualizado com sucesso!")
            
            def delete_holiday():
                if messagebox.askyesno("CONFIRMAR EXCLUSÃO", 
                    f"Tem certeza que deseja EXCLUIR o feriado '{nome_feriado}'?\n\nEsta ação não pode ser desfeita."):
                    if key in self.global_holidays:
                        del self.global_holidays[key]
                    if key in self.holiday_type:
                        del self.holiday_type[key]
                    if key in self.holiday_postos:
                        del self.holiday_postos[key]
                    if key in self.holiday_cidades:
                        del self.holiday_cidades[key]
                    refresh()
                    edit_dlg.destroy()
                    messagebox.showinfo("SUCESSO", "Feriado excluído com sucesso!")
            
            Button(btn_frame, text="SALVAR ALTERAÇÕES", command=save_changes, 
                   bg="green", fg="white", font=("Helvetica", 10, "bold"), width=18).pack(side=LEFT, padx=5)
            Button(btn_frame, text="EXCLUIR FERIADO", command=delete_holiday, 
                   bg="red", fg="white", font=("Helvetica", 10, "bold"), width=18).pack(side=LEFT, padx=5)
            Button(btn_frame, text="CANCELAR", command=edit_dlg.destroy, 
                   font=("Helvetica", 10), width=18).pack(side=LEFT, padx=5)
        
        tree.bind("<Double-1>", edit_selected)
        refresh()
        
        btnf = Frame(top, bg="#ecf0f1")
        btnf.pack(pady=15)
        
        btn_style = {"fg": "white", "font": ("Segoe UI", 10, "bold"), 
                    "relief": "flat", "bd": 0, "cursor": "hand2", 
                    "padx": 20, "pady": 10}
        
        Button(btnf, text="💾 SALVAR", command=lambda: (self._save_holidays_and_close(top)),
               bg="#27ae60", **btn_style).pack(side=LEFT, padx=8)
        Button(btnf, text="✕ FECHAR", command=lambda: (top.grab_release(), top.destroy()),
               bg="#95a5a6", **btn_style).pack(side=LEFT, padx=8)
        
        self.root.wait_window(top)

    def _save_holidays_and_close(self, top_win):
        self.store["global_holidays"] = self.global_holidays
        self.store["holiday_type"] = self.holiday_type
        self.store["holiday_postos"] = self.holiday_postos
        save_store(self.store)
        top_win.grab_release(); top_win.destroy()
        messagebox.showinfo("SALVO", "FERIADOS SALVOS.")

    def load_spreadsheet(self):
        path = filedialog.askopenfilename(filetypes=[("Planilhas", "*.xlsx *.xls *.csv"), ("Todos", "*.*")])
        if not path:
            return
        ext = os.path.splitext(path)[1].lower()
        try:
            import pandas as pd
            if ext in (".xls", ".xlsx", ".xlsm"):
                df = pd.read_excel(path, dtype=str).fillna("")
            else:
                try:
                    df = pd.read_csv(path, dtype=str, encoding="utf-8", sep=None, engine="python").fillna("")
                except Exception:
                    df = pd.read_csv(path, dtype=str, encoding="latin-1", sep=None, engine="python").fillna("")
        except Exception as e:
            messagebox.showerror("ERRO AO ABRIR ARQUIVO", str(e))
            return

        cols = list(df.columns)

        def find(names: List[str]) -> Optional[str]:
            for n in names:
                for c in cols:
                    # Normaliza para comparação (ignora espaços, pontuações e acentos)
                    # Mantido o lower/strip para retrocompatibilidade
                    normalized_c = (
                        str(c).strip().lower()
                        .replace(' ', '').replace('.', '').replace(',', '')
                        .replace('ç', 'c').replace('ã', 'a').replace('á', 'a')
                        .replace('é', 'e').replace('ê', 'e').replace('õ', 'o')
                        .replace('ó', 'o').replace('ú', 'u')
                    )
                    normalized_n = (
                        n.strip().lower()
                        .replace(' ', '').replace('.', '').replace(',', '')
                        .replace('ç', 'c').replace('ã', 'a').replace('á', 'a')
                        .replace('é', 'e').replace('ê', 'e').replace('õ', 'o')
                        .replace('ó', 'o').replace('ú', 'u')
                    )
                    if normalized_c == normalized_n:
                        return c
            return None

        # Nomes de colunas ajustados para lidar com a formatação do usuário
        col_nome = find(["NOME", "name"])
        col_mat = find(["MATRICULA", "matricula"])
        col_adm = find(["ADMISSAO", "ADMISSÃO"])
        col_cpf = find(["CPF"])
        col_filial = find(["FILIAL", "Empresa"])
        col_cnpj = find(["CNPJ"])
        col_func = find(["FUNCAO", "FUNÇÃO", "cargo"])
        col_posto = find(["POSTO", "posto"])
        # Busca por "ENDEREÇO" e "CIDADE"
        col_end = find(["ENDERECO", "ENDEREÇO"]) 
        col_cidade = find(["CIDADE", "city"])
        # Busca por "TIPO DE JORNADA" e "PRIMEIRO DIA DE FOLGA" (colunas K e L)
        col_tipo_jornada = find(["JORNADA (5X1 / 5X2 / 6X1 FIXO / 6X1 INTERCALADA / 12X36)", "TIPO DE JORNADA", "JORNADA", "ESCALA"])
        col_primeira_folga = find(["PRIMEIRO DIA DE FOLGA", "PRIMEIRODIADEFOLGA", "1 FOLGA", "PRIMEIRA FOLGA"])

        if not col_nome:
            messagebox.showerror("ERRO", "Coluna 'NOME' não encontrada.")
            return

        new_list = []
        for _, row in df.iterrows():
            nome = str(row.get(col_nome, "")).strip().upper()
            if not nome:
                continue
            matric = str(row.get(col_mat, "")) if col_mat else ""
            adm_val = row.get(col_adm, "") if col_adm else ""
            adm_date = None
            if str(adm_val).strip():
                for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"):
                    try:
                        adm_date = datetime.strptime(str(adm_val).strip(), fmt).date()
                        break
                    except Exception:
                        pass
            func = str(row.get(col_func, "")).strip().upper() if col_func else ""
            filial = str(row.get(col_filial, "")).strip().upper() if col_filial else ""
            cnpj = str(row.get(col_cnpj, "")).strip() if col_cnpj else ""
            cpf = str(row.get(col_cpf, "")).strip() if col_cpf else ""
            posto = str(row.get(col_posto, "")).strip().upper() if col_posto else ""
            endereco = str(row.get(col_end, "")).strip().upper() if col_end else ""
            cidade = str(row.get(col_cidade, "")).strip().upper() if col_cidade else ""
            
            # Lê TIPO DE JORNADA e PRIMEIRO DIA DE FOLGA da planilha (colunas K e L)
            tipo_jornada = str(row.get(col_tipo_jornada, "")).strip().upper() if col_tipo_jornada else ""
            primeira_folga_str = str(row.get(col_primeira_folga, "")).strip() if col_primeira_folga else ""
            
            # Normaliza tipo de jornada para corresponder às chaves de SCALE_TYPES
            # Aceita variações como "6x1 fica", "6x1 fixo", "6x1 intercalada", etc.
            if tipo_jornada:
                tipo_normalizado = None
                tj_lower = tipo_jornada.lower().replace("(", "").replace(")", "").strip()
                
                if "5x2" in tj_lower or "5 x 2" in tj_lower:
                    tipo_normalizado = "5X2"
                elif "5x1" in tj_lower or "5 x 1" in tj_lower:
                    tipo_normalizado = "5X1"
                elif "12x36" in tj_lower or "12 x 36" in tj_lower:
                    tipo_normalizado = "12X36"
                elif "6x1" in tj_lower or "6 x 1" in tj_lower:
                    if "intercalad" in tj_lower:
                        tipo_normalizado = "6X1 (INTERCALADA)"
                    else:
                        tipo_normalizado = "6X1 (FIXO)"
                
                if tipo_normalizado:
                    self.emp_scale_choice[nome] = tipo_normalizado
            
            # Se primeira folga veio da planilha, tenta converter para data
            if primeira_folga_str:
                primeira_folga_date = None
                for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y"):
                    try:
                        primeira_folga_date = datetime.strptime(primeira_folga_str, fmt).date()
                        break
                    except Exception:
                        pass
                if primeira_folga_date:
                    self.emp_first_off[nome] = primeira_folga_date.strftime("%Y-%m-%d")

            new_list.append({
                "nome": nome,
                "matricula": matric,
                "admissao": adm_date,
                "funcao": func,
                "filial": filial,
                "cnpj": cnpj,
                "endereco": endereco, 
                "cidade": cidade, 
                "cpf": cpf,
                "posto": posto
            })

        # Adiciona novos postos ao histórico (preserva os antigos)
        # FILTRO: Ignora valores que parecem CPF (contém apenas números, pontos e traços)
        novos_postos = set(
            emp.get("posto", "").strip()
            for emp in new_list
            if emp.get("posto", "").strip() and not self._is_cpf_format(emp.get("posto", "").strip())
        )
        self.all_postos_historico.update(novos_postos)
        
        self.funcionarios = new_list
        self.store["funcionarios"] = self.funcionarios
        self.store["all_postos_historico"] = sorted(list(self.all_postos_historico))
        save_store(self.store)
        self.update_employee_tree()
        messagebox.showinfo("OK", f"{len(new_list)} FUNCIONÁRIOS CARREGADOS.")

    def select_logo_image(self):
        """Abre janela para gerenciar logos por filial."""
        # Obter lista de filiais únicas dos funcionários
        filiais = sorted(set(emp.get("filial", "").strip().upper() for emp in self.funcionarios if emp.get("filial", "").strip()))
        
        if not filiais:
            messagebox.showwarning("AVISO", "NENHUMA FILIAL ENCONTRADA.\nCarregue funcionários primeiro para definir filiais.")
            return
        
        # Janela principal
        logo_win = Toplevel(self.root)
        logo_win.title("GERENCIAR LOGOS POR FILIAL")
        logo_win.geometry("750x680")
        logo_win.configure(bg="#ecf0f1")
        logo_win.transient(self.root)
        logo_win.grab_set()
        logo_win.resizable(True, True)  # Permite redimensionar
        
        # Centralizar
        logo_win.update_idletasks()
        x = (logo_win.winfo_screenwidth() // 2) - (750 // 2)
        y = (logo_win.winfo_screenheight() // 2) - (680 // 2)
        logo_win.geometry(f"750x680+{x}+{y}")
        
        # Título
        title_frame = Frame(logo_win, bg="#2c3e50")
        title_frame.pack(fill=X)
        Label(title_frame, text="🏢 GERENCIAR LOGOS POR FILIAL", 
              font=("Segoe UI", 14, "bold"), bg="#2c3e50", fg="white", 
              pady=18).pack()
        
        # Instruções
        info_frame = Frame(logo_win, bg="#3498db", relief="flat", bd=0)
        info_frame.pack(fill=X, padx=15, pady=(10, 10))
        Label(info_frame, text="💡 Selecione uma FILIAL e clique em ANEXAR para adicionar um logo específico.", 
              font=("Segoe UI", 9), bg="#3498db", fg="white", 
              pady=10, wraplength=700, justify=LEFT).pack(padx=10)
        
        # Container principal (lista + preview)
        main_container = Frame(logo_win, bg="#ecf0f1")
        main_container.pack(fill=BOTH, expand=True, padx=15, pady=10)
        
        # Frame da lista (lado esquerdo)
        list_frame = Frame(main_container, bg="#ecf0f1")
        list_frame.pack(side=LEFT, fill=BOTH, expand=True, padx=(0, 10))
        
        # Estilo da Treeview
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Logos.Treeview",
                       background="white",
                       foreground="black",
                       rowheight=35,
                       fieldbackground="white",
                       font=("Segoe UI", 10))
        style.map("Logos.Treeview", background=[("selected", "#3498db")])
        style.configure("Logos.Treeview.Heading",
                       background="#34495e",
                       foreground="white",
                       font=("Segoe UI", 11, "bold"),
                       relief="flat")
        
        # Treeview
        tree = ttk.Treeview(list_frame, columns=("filial", "status"), 
                           show="headings", style="Logos.Treeview")
        tree.heading("filial", text="FILIAL")
        tree.heading("status", text="STATUS DO LOGO")
        tree.column("filial", width=300, anchor="w")
        tree.column("status", width=150, anchor="center")
        
        # Scrollbar
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=scrollbar.set)
        tree.pack(side=LEFT, fill=BOTH, expand=True)
        scrollbar.pack(side=RIGHT, fill=Y)
        
        # Frame de preview (lado direito)
        preview_frame = Frame(main_container, bg="white", relief="groove", bd=2, width=250)
        preview_frame.pack(side=RIGHT, fill=Y, padx=(10, 0))
        preview_frame.pack_propagate(False)
        
        # Título do preview
        Label(preview_frame, text="PREVIEW DO LOGO", font=("Segoe UI", 10, "bold"),
              bg="#34495e", fg="white", pady=10).pack(fill=X)
        
        # Label para mostrar a imagem
        img_label = Label(preview_frame, bg="white", text="Nenhum logo\nselecionado",
                         font=("Segoe UI", 9), fg="#7f8c8d")
        img_label.pack(pady=20, padx=10)
        
        # Botão de remover (inicialmente oculto)
        remove_btn = Button(preview_frame, text="🗑️ REMOVER LOGO",
                           bg="#e74c3c", fg="white", font=("Segoe UI", 10, "bold"),
                           relief="raised", bd=2, cursor="hand2", padx=15, pady=8)
        
        # Preencher lista
        def atualizar_lista():
            tree.delete(*tree.get_children())
            logos_filiais = self.store.get("logos_filiais", {})
            for filial in filiais:
                status = "✅ LOGO" if filial in logos_filiais else "❌ SEM LOGO"
                tree.insert("", END, values=(filial, status))
        
        atualizar_lista()
        
        # Botão fechar (canto inferior direito)
        close_btn = Button(logo_win, text="✕ FECHAR", command=logo_win.destroy,
                          bg="#95a5a6", fg="white", font=("Segoe UI", 10, "bold"),
                          relief="raised", bd=2, cursor="hand2", padx=20, pady=8)
        close_btn.pack(side="bottom", anchor="e", padx=15, pady=15)
        
        # Função para mostrar preview do logo
        def mostrar_preview(event=None):
            sel = tree.selection()
            if not sel:
                img_label.config(image="", text="Nenhum logo\nselecionado")
                remove_btn.pack_forget()
                return
            
            filial = tree.item(sel[0], "values")[0]
            logos_filiais = self.store.get("logos_filiais", {})
            
            if filial not in logos_filiais:
                img_label.config(image="", text="❌ SEM LOGO\n\nClique 2x para\nadicionar")
                remove_btn.pack_forget()
                return
            
            try:
                # Carrega e redimensiona a imagem
                from PIL import Image, ImageTk
                b64_data = logos_filiais[filial]
                img_data = base64.b64decode(b64_data)
                img = Image.open(io.BytesIO(img_data))
                
                # Redimensiona mantendo proporção
                img.thumbnail((220, 300), Image.Resampling.LANCZOS)
                photo = ImageTk.PhotoImage(img)
                
                img_label.config(image=photo, text="")
                img_label.image = photo  # Mantém referência
                
                # Mostra botão remover
                remove_btn.pack(pady=10, padx=10)
            except Exception as e:
                img_label.config(image="", text=f"Erro ao carregar\nimagem")
                remove_btn.pack_forget()
        
        # Bind para mostrar preview ao selecionar
        tree.bind("<<TreeviewSelect>>", mostrar_preview)
        
        # Função para anexar logo com duplo clique
        def anexar_logo(event=None):
            sel = tree.selection()
            if not sel:
                return
            
            filial = tree.item(sel[0], "values")[0]
            
            path = filedialog.askopenfilename(
                title=f"SELECIONAR LOGO - {filial}",
                filetypes=[
                    ("Imagens", "*.png *.jpg *.jpeg *.bmp *.gif"),
                    ("Todos", "*.*")
                ]
            )
            if not path:
                return
            
            try:
                file_size = os.path.getsize(path)
                if file_size > 5 * 1024 * 1024:
                    messagebox.showerror("ERRO", "IMAGEM MUITO GRANDE. MÁXIMO: 5MB")
                    return
                
                with open(path, "rb") as f:
                    img_data = f.read()
                b64_str = base64.b64encode(img_data).decode("utf-8")
                
                if "logos_filiais" not in self.store:
                    self.store["logos_filiais"] = {}
                self.store["logos_filiais"][filial] = b64_str
                save_store(self.store)
                
                atualizar_lista()
                mostrar_preview()
                messagebox.showinfo("SUCESSO", f"LOGO SALVO PARA:\n{filial}")
            
            except Exception as e:
                messagebox.showerror("ERRO", f"ERRO AO CARREGAR LOGO:\n{str(e)}")
        
        # Bind duplo clique para anexar
        tree.bind("<Double-1>", anexar_logo)
        
        # Função para remover logo
        def remover_logo():
            sel = tree.selection()
            if not sel:
                return
            
            filial = tree.item(sel[0], "values")[0]
            
            if messagebox.askyesno("CONFIRMAR", f"REMOVER LOGO DA FILIAL:\n{filial}?"):
                if "logos_filiais" in self.store and filial in self.store["logos_filiais"]:
                    del self.store["logos_filiais"][filial]
                    save_store(self.store)
                    atualizar_lista()
                    mostrar_preview()
                    messagebox.showinfo("SUCESSO", f"LOGO REMOVIDO DA FILIAL:\n{filial}")
        
        remove_btn.config(command=remover_logo)

    def show_load_menu(self):
        """Abre menu com opções: Carregar Arquivo ou Baixar Modelo Excel."""
        menu_win = Toplevel(self.root)
        menu_win.title("CARREGAR FUNCIONÁRIOS")
        menu_win.geometry("480x280")
        menu_win.resizable(False, False)
        menu_win.configure(bg="#ecf0f1")
        menu_win.transient(self.root)
        menu_win.grab_set()
        
        # Título moderno
        title_frame = Frame(menu_win, bg="#2c3e50")
        title_frame.pack(fill=X)
        Label(title_frame, text="📂 CARREGAR FUNCIONÁRIOS", 
              font=("Segoe UI", 13, "bold"), bg="#2c3e50", fg="white", 
              pady=18).pack()
        
        # Container de botões
        btn_container = Frame(menu_win, bg="#ecf0f1")
        btn_container.pack(fill=BOTH, expand=True, padx=25, pady=20)
        
        # Botão Carregar Arquivo
        Button(btn_container, text="📂 CARREGAR ARQUIVO", 
               command=lambda: [menu_win.destroy(), self.load_spreadsheet()],
               bg="#1abc9c", fg="white", font=("Segoe UI", 11, "bold"),
               relief="flat", bd=0, cursor="hand2", padx=25, pady=15
              ).pack(pady=10, fill=X)
        
        # Botão Baixar Modelo
        Button(btn_container, text="📥 BAIXAR MODELO EXCEL", 
               command=lambda: [menu_win.destroy(), self.download_excel_template()],
               bg="#3498db", fg="white", font=("Segoe UI", 11, "bold"),
               relief="flat", bd=0, cursor="hand2", padx=25, pady=15
              ).pack(pady=10, fill=X)
        
        # Botão Cancelar
        Button(btn_container, text="❌ CANCELAR", command=menu_win.destroy,
               bg="#95a5a6", fg="white", font=("Segoe UI", 10, "bold"),
               relief="flat", bd=0, cursor="hand2", padx=25, pady=12
              ).pack(pady=10, fill=X)

    def download_excel_template(self):
        """Cria e salva arquivo Excel modelo com as colunas necessárias e exemplos."""
        try:
            import pandas as pd
            from datetime import datetime
            
            # Dados de exemplo - ordem das colunas: A-L
            from collections import OrderedDict
            exemplos = [
                OrderedDict([
                    ("NOME", "JOÃO DA SILVA"),
                    ("CPF", "123.456.789-00"),
                    ("MATRICULA", "2024001"),
                    ("FUNÇÃO", "VIGILANTE"),
                    ("POSTO", "SHOPPING CENTER"),
                    ("ADMISSÃO", "01/01/2024"),
                    ("FILIAL", "MATRIZ RECIFE"),
                    ("CNPJ", "12.345.678/0001-90"),
                    ("ENDEREÇO", "RUA DA AURORA, 123 - BOA VISTA"),
                    ("CIDADE", "RECIFE"),
                    ("JORNADA (5X1 / 5X2 / 6X1 FIXO / 6X1 INTERCALADA / 12X36)", "6X1 FIXO"),
                    ("PRIMEIRO DIA DE FOLGA", "06/01/2025")
                ]),
                OrderedDict([
                    ("NOME", "MARIA SANTOS"),
                    ("CPF", "987.654.321-00"),
                    ("MATRICULA", "2024002"),
                    ("FUNÇÃO", "SUPERVISOR"),
                    ("POSTO", "HOSPITAL MUNICIPAL"),
                    ("ADMISSÃO", "15/02/2024"),
                    ("FILIAL", "FILIAL OLINDA"),
                    ("CNPJ", "98.765.432/0001-10"),
                    ("ENDEREÇO", "AV. GETÚLIO VARGAS, 456 - CENTRO"),
                    ("CIDADE", "OLINDA"),
                    ("JORNADA (5X1 / 5X2 / 6X1 FIXO / 6X1 INTERCALADA / 12X36)", "5X2"),
                    ("PRIMEIRO DIA DE FOLGA", "18/02/2025")
                ]),
                OrderedDict([
                    ("NOME", "CARLOS PEREIRA"),
                    ("CPF", "456.789.123-00"),
                    ("MATRICULA", "2024003"),
                    ("FUNÇÃO", "PORTEIRO"),
                    ("POSTO", "CONDOMINIO RESIDENCIAL"),
                    ("ADMISSÃO", "10/03/2024"),
                    ("FILIAL", "MATRIZ RECIFE"),
                    ("CNPJ", "12.345.678/0001-90"),
                    ("ENDEREÇO", "RUA DA AURORA, 123 - BOA VISTA"),
                    ("CIDADE", "RECIFE"),
                    ("JORNADA (5X1 / 5X2 / 6X1 FIXO / 6X1 INTERCALADA / 12X36)", "12X36"),
                    ("PRIMEIRO DIA DE FOLGA", "11/03/2025")
                ])
            ]
            
            # Cria DataFrame
            df = pd.DataFrame(exemplos)
            
            # Solicita onde salvar
            path = filedialog.asksaveasfilename(
                title="SALVAR MODELO EXCEL",
                defaultextension=".xlsx",
                filetypes=[("Excel", "*.xlsx"), ("Todos", "*.*")],
                initialfile=f"MODELO_FUNCIONARIOS.xlsx"
            )
            
            if not path:
                return
            
            # Salva arquivo com openpyxl para adicionar validação
            from openpyxl import load_workbook
            from openpyxl.worksheet.datavalidation import DataValidation
            
            df.to_excel(path, index=False, sheet_name="FUNCIONÁRIOS")
            
            # Adiciona dropdown na coluna JORNADA (coluna K)
            wb = load_workbook(path)
            ws = wb.active
            
            # Define validação de lista para coluna K (JORNADA)
            jornada_validation = DataValidation(
                type="list",
                formula1='"5X1,5X2,6X1 (FIXO),6X1 (INTERCALADA),12X36"',
                allow_blank=False,
                showDropDown=True
            )
            jornada_validation.error = 'Valor inválido'
            jornada_validation.errorTitle = 'Entrada Inválida'
            jornada_validation.prompt = 'Selecione o tipo de jornada'
            jornada_validation.promptTitle = 'TIPO DE JORNADA'
            
            # Aplica validação em toda a coluna K (linhas 2 a 1000)
            ws.add_data_validation(jornada_validation)
            jornada_validation.add(f'K2:K1000')
            
            wb.save(path)
            
            messagebox.showinfo(
                "MODELO CRIADO",
                f"MODELO EXCEL SALVO COM SUCESSO!\n\n"
                f"LOCAL: {os.path.basename(path)}\n\n"
                f"O arquivo contém 3 exemplos.\n"
                f"APAGUE-OS e insira seus dados reais.\n\n"
                f"ORDEM DAS COLUNAS (A-L):\n"
                f"A - NOME (obrigatório)\n"
                f"B - CPF\n"
                f"C - MATRICULA\n"
                f"D - FUNÇÃO\n"
                f"E - POSTO\n"
                f"F - ADMISSÃO (DD/MM/AAAA)\n"
                f"G - FILIAL\n"
                f"H - CNPJ\n"
                f"I - ENDEREÇO\n"
                f"J - CIDADE\n"
                f"K - JORNADA (dropdown: 5X1/5X2/6X1 FIXO/6X1 INTERCALADA/12X36)\n"
                f"L - PRIMEIRO DIA DE FOLGA (DD/MM/AAAA)"
            )
        
        except Exception as e:
            messagebox.showerror("ERRO", f"ERRO AO CRIAR MODELO:\n{str(e)}")

    def update_employee_tree(self):
        for i in self.emp_tree.get_children():
            self.emp_tree.delete(i)
        # Filtro de busca (por nome e matrícula)
        query = normalize_text(self.search_var.get()) if hasattr(self, 'search_var') else ""

        row_count = 0
        for emp in self.funcionarios:
            nome = emp.get("nome", "")
            funcao = emp.get("funcao", "")
            # Monta o texto com nome e função em linhas separadas
            nome_completo = f"{nome}\n{funcao}" if funcao else nome
            mat = emp.get("matricula", "")
            cpf = emp.get("cpf", "")
            escala = self.emp_scale_choice.get(nome, "6X1 (FIXO)")
            posto = emp.get("posto", "") or ""
            revisado = "✓" if self.emp_revisado.get(nome, False) else ""
            if not query:
                pass_filter = True
            else:
                nome_norm = normalize_text(nome)
                mat_norm = normalize_text(mat)
                pass_filter = (query in nome_norm) or (query in mat_norm)
            if pass_filter:
                # Cores alternadas nas linhas
                tag = 'evenrow' if row_count % 2 == 0 else 'oddrow'
                iid = self.emp_tree.insert("", "end", values=(mat, cpf, escala, posto, revisado), 
                                          text=nome_completo, tags=(tag,))
                # Marca fundo verde na coluna revisado se estiver OK
                if revisado == "✓":
                    self.emp_tree.tag_configure(f"rev_{iid}", background="#27ae60", foreground="white")
                    self.emp_tree.item(iid, tags=(f"rev_{iid}",))
                row_count += 1
        
        # Atualiza contadores de revisão
        self.update_revision_stats()
    
    def update_revision_stats(self):
        """Atualiza os contadores de funcionários revisados e não revisados"""
        total = len(self.funcionarios)
        revisados = sum(1 for nome in self.funcionarios if self.emp_revisado.get(nome.get("nome", ""), False))
        nao_revisados = total - revisados
        
        if hasattr(self, 'label_revisados'):
            self.label_revisados.config(text=f"✅ REVISADOS: {revisados}")
        if hasattr(self, 'label_nao_revisados'):
            self.label_nao_revisados.config(text=f"⚠ NÃO REVISADOS: {nao_revisados}")

    def sort_emp_tree(self, col: str):
        """Ordena a tabela principal por qualquer coluna ao clicar no cabeçalho.
        col pode ser "#0" (nome) ou uma das colunas definidas em columns().
        Alterna entre ascendente/descendente a cada clique.
        """
        # Direção atual (False = asc, True = desc). Primeiro clique = asc.
        reverse = self._emp_sort_dir.get(col, False)

        # Captura itens visíveis e valor de ordenação para cada um
        items = list(self.emp_tree.get_children(""))

        def coerce(val: str) -> Any:
            # Normaliza valor de comparação
            if col == "mat":
                try:
                    return int(str(val).strip())
                except Exception:
                    return 0
            return str(val).casefold()

        def get_val(item_id: str):
            if col == "#0":
                v = self.emp_tree.item(item_id, "text")
                # Extrai apenas o nome (primeira linha) para ordenação
                v = v.split("\n")[0] if "\n" in v else v
            else:
                v = self.emp_tree.set(item_id, col)
            return coerce(v)

        items_sorted = sorted(items, key=get_val, reverse=reverse)

        # Reordena visualmente os itens
        for idx, iid in enumerate(items_sorted):
            self.emp_tree.move(iid, "", idx)

        # Alterna direção para o próximo clique
        self._emp_sort_dir[col] = not reverse

    def on_tree_click(self, event=None):
        """Detecta clique na coluna REVISADO e alterna o estado de revisão."""
        region = self.emp_tree.identify("region", event.x, event.y)
        if region != "cell":
            return
        col = self.emp_tree.identify_column(event.x)
        # col retorna "#6" para a 6ª coluna (revisado)
        if col != "#6":
            return
        iid = self.emp_tree.identify_row(event.y)
        if not iid:
            return
        nome_completo = self.emp_tree.item(iid, "text")
        # Extrai apenas o nome (primeira linha, antes do \n)
        nome = nome_completo.split("\n")[0] if "\n" in nome_completo else nome_completo
        # Alterna estado
        current = self.emp_revisado.get(nome, False)
        self.emp_revisado[nome] = not current
        # Atualiza visualmente
        self.update_employee_tree()

    # ---------------------------
    # POPUP DE EDIÇÃO DE FUNCIONÁRIO (MODAL)
    # ---------------------------
    def on_tree_double(self, event=None):
        sel = self.emp_tree.selection()
        if not sel:
            return
        item = sel[0]
        nome_completo = self.emp_tree.item(item, "text")
        # Extrai apenas o nome (primeira linha, antes do \n)
        nome = nome_completo.split("\n")[0] if "\n" in nome_completo else nome_completo
        emp = next((e for e in self.funcionarios if e.get("nome") == nome), None)
        if not emp:
            messagebox.showerror("ERRO", "FUNCIONÁRIO NÃO ENCONTRADO")
            return

        top = Toplevel(self.root)
        top.title(f"CONFIGURAR FUNCIONÁRIO")
        top.configure(bg="#ecf0f1")
        top.transient(self.root)
        top.grab_set()
        top.columnconfigure(0, weight=1)
        top.rowconfigure(0, weight=1)
        
        # Título da janela
        title_frame = Frame(top, bg="#2c3e50")
        title_frame.pack(fill=X)
        Label(title_frame, text="⚙️ CONFIGURAR FUNCIONÁRIO", 
              font=("Segoe UI", 13, "bold"), bg="#2c3e50", fg="white", 
              pady=15).pack()

        frm = Frame(top, bg="#ecf0f1")
        frm.pack(fill=BOTH, expand=True, padx=15, pady=15)
        for i in range(6): 
            frm.columnconfigure(i, weight=1) 
        frm.columnconfigure(0, weight=0)

        # Cabeçalho com informações do funcionário (card moderno)
        header_frame = Frame(frm, bg="#34495e", relief="flat", bd=0)
        header_frame.grid(row=0, column=0, columnspan=6, sticky="ew", pady=(0, 15))
        
        # Info container
        info_container = Frame(header_frame, bg="#34495e")
        info_container.pack(fill=X, padx=15, pady=12)
        
        # Nome do funcionário
        Label(info_container, text="👤 ", font=("Segoe UI", 12), 
              fg="#ecf0f1", bg="#34495e").pack(side=LEFT)
        Label(info_container, text=nome, font=("Segoe UI", 12, "bold"), 
              fg="white", bg="#34495e").pack(side=LEFT, padx=(0, 20))
        
        # Função
        if emp.get('funcao'):
            Label(info_container, text="💼 ", font=("Segoe UI", 10), 
                  fg="#ecf0f1", bg="#34495e").pack(side=LEFT)
            Label(info_container, text=emp.get('funcao',''), font=("Segoe UI", 10), 
                  fg="#bdc3c7", bg="#34495e").pack(side=LEFT, padx=(0, 20))
        Label(header_frame, text=f"CPF: {emp.get('cpf','')}", font=("Helvetica", 10), 
              fg="white", bg="#2c3e50").pack(side=RIGHT, padx=10)

        # Frame de configurações da escala
        config_frame = Frame(frm, bg="#ffffff", relief="groove", bd=2)
        config_frame.grid(row=1, column=0, columnspan=6, sticky="ew", pady=(0, 15), padx=2)
        config_frame.columnconfigure(1, weight=1)
        config_frame.columnconfigure(3, weight=1)
        
        Label(config_frame, text="ESCALA:", font=("Helvetica", 10, "bold"), bg="#ffffff").grid(row=0, column=0, sticky="w", padx=10, pady=10)
        escala_var = StringVar(value=self.emp_scale_choice.get(nome, "6X1 (FIXO)"))
        escala_cb = ttk.Combobox(config_frame, textvariable=escala_var, values=list(SCALE_TYPES.keys()), state="readonly", width=20, font=("Helvetica", 10))
        escala_cb.grid(row=0, column=1, sticky="w", padx=5, pady=10)

        Label(config_frame, text="DATA 1ª FOLGA:", font=("Helvetica", 10, "bold"), bg="#ffffff").grid(row=0, column=2, sticky="w", padx=(30, 10), pady=10)
        try:
            first_entry = DateEntry(
                config_frame,
                date_pattern="dd/mm/yyyy",
                locale="pt_BR",
                width=16,
                font=("Helvetica", 10),
                showweeknumbers=False,
                firstweekday="sunday",
                selectmode="day",
                year=2025,
                month=1,
                state='normal',
                cursor='hand2',
                takefocus=True,
                selectbackground='#3498db',
                selectforeground='white',
                normalbackground='white',
                normalforeground='black',
                background='#3498db',
                foreground='white',
                borderwidth=2,
                relief="solid"
            )
        except TypeError:
            first_entry = DateEntry(config_frame, date_pattern="dd/mm/yyyy", width=16, font=("Helvetica", 10), state='normal')
        if self.emp_first_off.get(nome):
            try:
                first_entry.set_date(datetime.strptime(self.emp_first_off[nome], "%Y-%m-%d").date())
            except Exception:
                pass
        first_entry.grid(row=0, column=3, sticky="w", padx=5, pady=10)

        # Checkboxes
        checkbox_frame = Frame(frm, bg="#ffffff", relief="groove", bd=2)
        checkbox_frame.grid(row=2, column=0, columnspan=6, sticky="ew", pady=(0, 15), padx=2)
        
        trava_var = BooleanVar(value=self.emp_trabalha_feriado.get(nome, False))
        chk = ttk.Checkbutton(checkbox_frame, text="TRABALHA EM FERIADOS (PADRÃO: NÃO)", variable=trava_var)
        chk.grid(row=0, column=0, sticky="w", padx=15, pady=10)

        # Checkbox de período noturno removido - não mais necessário
        
        # Botões de ação com estilo
        btn_frame = Frame(frm, bg="#f0f0f0")
        btn_frame.grid(row=3, column=0, columnspan=6, pady=15, sticky="ew")
        btn_frame.columnconfigure(0, weight=1)

        Button(
            btn_frame, text="⚙ GERENCIADOR DE OCORRÊNCIAS", 
            command=lambda n=nome: self.popup_manage_absences(n),
            bg="#3498db", fg="white", font=("Helvetica", 10, "bold"),
            relief="raised", bd=2, cursor="hand2", padx=10, pady=8
        ).grid(row=0, column=0, sticky="ew", padx=5, pady=4)

        current_row = 4
        
        # Botões de controle inferior com estilo
        sc_frame = Frame(frm, bg="#f0f0f0")
        sc_frame.grid(row=current_row, column=0, columnspan=6, sticky="ew", pady=(10, 5))

        def salvar():
            self.emp_scale_choice[nome] = escala_var.get()
            try:
                self.emp_first_off[nome] = first_entry.get_date().strftime("%Y-%m-%d")
            except Exception:
                self.emp_first_off[nome] = ""
            self.emp_trabalha_feriado[nome] = trava_var.get()
            
            # Horários e período noturno removidos - não são mais configuráveis
            self.update_employee_tree()
            self.store["emp_scale_choice"] = self.emp_scale_choice
            self.store["emp_first_off"] = self.emp_first_off
            self.store["emp_trabalha_feriado"] = self.emp_trabalha_feriado
            self.store["emp_faltas_atestados"] = self.emp_faltas_atestados # Garante que está no save
            save_store(self.store)
            messagebox.showinfo("SALVO", f"CONFIGURAÇÕES DE {nome} SALVAS.")
            top.grab_release()
            top.destroy()

        Button(
            sc_frame, text="✓ SALVAR", command=salvar, 
            bg="#2ecc71", fg="white", font=("Helvetica", 10, "bold"),
            relief="raised", bd=2, cursor="hand2", padx=20, pady=6
        ).pack(side=RIGHT, padx=5)
        
        Button(
            sc_frame, text="✕ CANCELAR", 
            command=lambda: (top.grab_release(), top.destroy()),
            bg="#e74c3c", fg="white", font=("Helvetica", 10, "bold"),
            relief="raised", bd=2, cursor="hand2", padx=20, pady=6
        ).pack(side=RIGHT, padx=5)

        self.root.wait_window(top)

    def popup_manage_absences(self, nome: str):
        period = self._get_period_dates()
        if not period:
            return
        start, end = period
        if end < start:
            messagebox.showerror("ERRO", "PERÍODO INVÁLIDO")
            return

        top = Toplevel(self.root)
        top.title(f"GERENCIADOR DE OCORRÊNCIAS - {nome}")
        top.geometry("650x600")
        top.transient(self.root)
        top.grab_set()
        top.configure(bg="#ecf0f1")
        
        # Centralizar na tela
        top.update_idletasks()
        x = (top.winfo_screenwidth() // 2) - (650 // 2)
        y = (top.winfo_screenheight() // 2) - (600 // 2)
        top.geometry(f"650x600+{x}+{y}")

        # Cabeçalho
        header_frame = Frame(top, bg="#34495e", relief=RIDGE, bd=2)
        header_frame.pack(fill=X, padx=10, pady=10)
        
        Label(header_frame, text=f"PERÍODO: {start.strftime('%d/%m/%Y')} - {end.strftime('%d/%m/%Y')}", 
              font=("Helvetica", 12, "bold"), bg="#34495e", fg="white", pady=10).pack()

        # Container para a Treeview + Scrollbar
        list_frame = Frame(top, bg="#ecf0f1")
        list_frame.pack(fill=BOTH, expand=True, padx=10, pady=5)
        
        # Estilo para a Treeview
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Custom.Treeview", 
                       background="white",
                       foreground="black",
                       rowheight=25,
                       fieldbackground="white",
                       font=("Helvetica", 9))
        style.map("Custom.Treeview", background=[("selected", "#3498db")])
        style.configure("Custom.Treeview.Heading",
                       background="#2c3e50",
                       foreground="white",
                       font=("Helvetica", 10, "bold"),
                       relief=FLAT)

        tree = ttk.Treeview(list_frame, columns=("data", "dia", "situacao"), show="headings", 
                           height=15, style="Custom.Treeview")
        tree.heading("data", text="DATA")
        tree.heading("dia", text="DIA")
        tree.heading("situacao", text="SITUAÇÃO (CLIQUE)")
        tree.column("data", width=150, anchor="center")
        tree.column("dia", width=100, anchor="center")
        tree.column("situacao", width=280, anchor="center")
        tree.pack(side=LEFT, fill=BOTH, expand=True)

        # Scrollbar vertical
        sb = ttk.Scrollbar(list_frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=sb.set)
        sb.pack(side=LEFT, fill=Y)

        emp_events = self.emp_faltas_atestados.get(nome, {})
        # Mapa de data (dd/mm/YYYY) -> item id da Treeview, para aplicar períodos rapidamente
        date_to_iid: Dict[str, str] = {}

        for d in daterange(start, end):
            ds = d.strftime("%Y-%m-%d")
            display = d.strftime("%d/%m/%Y")
            day_name = WEEKDAY_PT_SHORT[d.weekday()]
            status = emp_events.get(ds, "")
            iid = tree.insert("", "end", values=(display, day_name, status.upper() if isinstance(status, str) else status))
            date_to_iid[display] = iid

        def cycle_status(e):
            iid = tree.identify_row(e.y)
            if not iid:
                return
            v = list(tree.item(iid, "values"))
            cur = v[2]
            seq = ["", "FOLGA", "FERIADO"]
            try:
                idx = seq.index(cur)
                v[2] = seq[(idx + 1) % len(seq)]
            except Exception:
                v[2] = "FOLGA"
            tree.item(iid, values=v)

        tree.bind("<Button-1>", cycle_status)

        def open_apply_period_dialog():
            # Janela simples para escolher tipo, data inicial e quantidade de dias
            dlg = Toplevel(top)
            dlg.title("Aplicar Ocorrência por Período")
            dlg.geometry("450x280")
            dlg.transient(top)
            dlg.grab_set()
            dlg.configure(bg="#ecf0f1")
            dlg.resizable(False, False)
            
            # Centralizar
            dlg.update_idletasks()
            x = (dlg.winfo_screenwidth() // 2) - (450 // 2)
            y = (dlg.winfo_screenheight() // 2) - (280 // 2)
            dlg.geometry(f"450x280+{x}+{y}")
            
            # Cabeçalho
            header = Frame(dlg, bg="#34495e", relief=RIDGE, bd=2)
            header.grid(row=0, column=0, columnspan=2, sticky="ew", padx=10, pady=10)
            Label(header, text="Configurar Período de Ocorrência", 
                  font=("Helvetica", 11, "bold"), bg="#34495e", fg="white", pady=8).pack()

            # Frame principal
            main_frame = Frame(dlg, bg="#ecf0f1")
            main_frame.grid(row=1, column=0, columnspan=2, padx=20, pady=10, sticky="ew")

            Label(main_frame, text="Tipo de ocorrência:", bg="#ecf0f1", 
                  font=("Helvetica", 10)).grid(row=0, column=0, sticky="w", padx=10, pady=10)
            tipos = ["FOLGA", "FERIADO"]
            tipo_var = StringVar(value=tipos[0])
            tipo_cb = ttk.Combobox(main_frame, values=tipos, textvariable=tipo_var, 
                                  state="readonly", width=18, font=("Helvetica", 10))
            tipo_cb.grid(row=0, column=1, padx=10, pady=10, sticky="w")

            Label(main_frame, text="Data de início:", bg="#ecf0f1", 
                  font=("Helvetica", 10)).grid(row=1, column=0, sticky="w", padx=10, pady=10)
            try:
                de = DateEntry(main_frame, date_pattern="dd/mm/yyyy", locale="pt_BR", width=18,
                              state='normal', cursor='hand2', takefocus=True,
                              showweeknumbers=False, firstweekday='sunday', selectmode='day',
                              font=("Helvetica", 10))
            except TypeError:
                de = DateEntry(main_frame, date_pattern="dd/mm/yyyy", width=18, 
                              state='normal', font=("Helvetica", 10))
            # valor inicial = início do período atual
            try:
                de.set_date(start)
            except Exception:
                pass
            de.grid(row=1, column=1, padx=10, pady=10, sticky="w")

            Label(main_frame, text="Quantidade de dias:", bg="#ecf0f1", 
                  font=("Helvetica", 10)).grid(row=2, column=0, sticky="w", padx=10, pady=10)
            days_var = StringVar(value="1")
            days_sb = Spinbox(main_frame, from_=1, to=365, textvariable=days_var, 
                             width=10, font=("Helvetica", 10))
            days_sb.grid(row=2, column=1, padx=10, pady=10, sticky="w")

            def apply_and_close():
                try:
                    first_day = de.get_date()
                except Exception:
                    messagebox.showerror("ERRO", "Data inicial inválida.")
                    return
                try:
                    qnt = int(days_var.get())
                except Exception:
                    messagebox.showerror("ERRO", "Quantidade de dias inválida.")
                    return
                sel_tipo = (tipo_var.get() or "").upper()
                if sel_tipo not in tipos:
                    messagebox.showerror("ERRO", "Selecione um tipo de ocorrência.")
                    return

                # Aplica na árvore apenas dentro do período da tela
                for i in range(max(1, qnt)):
                    cur_day = first_day + timedelta(days=i)
                    if cur_day < start or cur_day > end:
                        continue
                    disp = cur_day.strftime("%d/%m/%Y")
                    iid2 = date_to_iid.get(disp)
                    if not iid2:
                        continue
                    vals = list(tree.item(iid2, "values"))
                    if len(vals) >= 3:
                        vals[2] = sel_tipo
                        tree.item(iid2, values=vals)

                dlg.grab_release()
                dlg.destroy()

            frm_btn = Frame(dlg, bg="#ecf0f1")
            frm_btn.grid(row=2, column=0, columnspan=2, pady=20)
            
            Button(frm_btn, text="✓ APLICAR", command=apply_and_close,
                   bg="#27ae60", fg="white", font=("Helvetica", 10, "bold"),
                   relief="raised", bd=3, cursor="hand2",
                   padx=30, pady=10, width=12).pack(side=LEFT, padx=5)
            
            Button(frm_btn, text="✗ FECHAR", command=lambda: (dlg.grab_release(), dlg.destroy()),
                   bg="#95a5a6", fg="white", font=("Helvetica", 10, "bold"),
                   relief="raised", bd=3, cursor="hand2",
                   padx=30, pady=10, width=12).pack(side=LEFT, padx=5)
            
            dlg.wait_window(dlg)

        def save():
            new_events = {}
            for iid in tree.get_children():
                v = tree.item(iid, "values")
                # v[0] = "dd/mm/YYYY"
                if v[2] in ("FOLGA", "FERIADO"):
                    try:
                        dt = datetime.strptime(v[0], "%d/%m/%Y").date()
                        key = dt.strftime("%Y-%m-%d")
                    except Exception:
                        key = v[0]
                    new_events[key] = v[2]

            self.emp_faltas_atestados[nome] = new_events
            self.update_employee_tree()
            messagebox.showinfo("SALVO", "OCORRÊNCIAS SALVAS.")
            top.grab_release()
            top.destroy()

        # Frame de instruções
        info_frame = Frame(top, bg="#ecf0f1")
        info_frame.pack(fill=X, padx=10, pady=5)
        
        info_label = Label(info_frame, 
                          text="💡 Clique na coluna SITUAÇÃO para alternar entre os tipos de ocorrência",
                          font=("Helvetica", 9, "italic"),
                          bg="#ecf0f1", fg="#7f8c8d")
        info_label.pack()

        # Frame dos botões com design melhorado
        btn_frame = Frame(top, bg="#ecf0f1")
        btn_frame.pack(pady=15)
        
        Button(btn_frame, text="📅 APLICAR POR PERÍODO", command=open_apply_period_dialog,
               bg="#3498db", fg="white", font=("Helvetica", 10, "bold"),
               relief="raised", bd=3, cursor="hand2",
               padx=20, pady=10, width=20).pack(side=LEFT, padx=5)
        
        Button(btn_frame, text="✓ SALVAR", command=save,
               bg="#27ae60", fg="white", font=("Helvetica", 10, "bold"),
               relief="raised", bd=3, cursor="hand2",
               padx=20, pady=10, width=15).pack(side=LEFT, padx=5)
        
        Button(btn_frame, text="✗ CANCELAR", command=lambda: (top.grab_release(), top.destroy()),
               bg="#e74c3c", fg="white", font=("Helvetica", 10, "bold"),
               relief="raised", bd=3, cursor="hand2",
               padx=20, pady=10, width=15).pack(side=LEFT, padx=5)

        self.root.wait_window(top)

    def save_config(self):
        self.store["global_holidays"] = self.global_holidays
        self.store["holiday_type"] = self.holiday_type
        self.store["holiday_postos"] = self.holiday_postos
        self.store["holiday_cidades"] = self.holiday_cidades
        self.store["cidades"] = self.cidades
        self.store["emp_personal_hols"] = self.emp_personal_hols
        self.store["emp_scale_choice"] = self.emp_scale_choice
        self.store["emp_first_off"] = self.emp_first_off
        self.store["emp_faltas_atestados"] = self.emp_faltas_atestados
        self.store["emp_trabalha_feriado"] = self.emp_trabalha_feriado
        self.store["emp_revisado"] = self.emp_revisado
        self.store["funcionarios"] = self.funcionarios
        self.store["all_postos_historico"] = sorted(list(self.all_postos_historico))
        save_store(self.store)
        messagebox.showinfo("SALVO", "CONFIGURAÇÕES SALVAS LOCALMENTE.")

    def show_pdf_generation_dialog(self):
        """Mostra diálogo para escolher modo de geração de PDFs"""
        if not self.funcionarios:
            messagebox.showwarning("AVISO", "CARREGUE PLANILHA PRIMEIRO")
            return None
        
        dlg = Toplevel(self.root)
        dlg.title("OPÇÕES DE GERAÇÃO DE PDFs")
        dlg.geometry("650x650")
        dlg.transient(self.root)
        dlg.grab_set()
        dlg.configure(bg="#ecf0f1")
        dlg.resizable(False, False)  # Impede redimensionamento
        
        result = {"mode": None, "selection": []}
        
        # Frame principal
        main_frame = Frame(dlg, bg="#ffffff", relief="raised", bd=2)
        main_frame.pack(fill=BOTH, expand=True, padx=15, pady=15)
        
        # Título com fundo colorido
        title_frame = Frame(main_frame, bg="#34495e")
        title_frame.pack(fill=X)
        Label(title_frame, text="📄 ESCOLHA O MODO DE GERAÇÃO DE PDFs", 
              font=("Helvetica", 13, "bold"), bg="#34495e", fg="white", pady=12).pack()
        
        # Frame de opções
        options_frame = Frame(main_frame, bg="#ffffff")
        options_frame.pack(fill=X, padx=20, pady=15)
        
        # Variável de seleção
        mode_var = StringVar(value="todos")
        
        # Estilo personalizado para os radio buttons
        style = ttk.Style()
        style.configure("Custom.TRadiobutton", font=("Helvetica", 10))
        
        # Opção 1: Todos
        rb1_frame = Frame(options_frame, bg="#e8f8f5", relief="groove", bd=2)
        rb1_frame.pack(fill=X, pady=5)
        rb1 = ttk.Radiobutton(rb1_frame, text="📄 GERAR TODOS OS FUNCIONÁRIOS", 
                              variable=mode_var, value="todos", style="Custom.TRadiobutton")
        rb1.pack(anchor="w", padx=10, pady=10)
        
        # Opção 2: Revisados (OK)
        rb2_frame = Frame(options_frame, bg="#d5f4e6", relief="groove", bd=2)
        rb2_frame.pack(fill=X, pady=5)
        rb2 = ttk.Radiobutton(rb2_frame, text="✅ GERAR APENAS REVISADOS (MARCADOS COM OK)", 
                              variable=mode_var, value="revisados", style="Custom.TRadiobutton")
        rb2.pack(anchor="w", padx=10, pady=10)
        
        # Opção 3: Por Posto
        rb3_frame = Frame(options_frame, bg="#fff3cd", relief="groove", bd=2)
        rb3_frame.pack(fill=X, pady=5)
        rb3 = ttk.Radiobutton(rb3_frame, text="🏙️ ESCOLHER POSTOS ESPECÍFICOS", 
                              variable=mode_var, value="postos", style="Custom.TRadiobutton")
        rb3.pack(anchor="w", padx=10, pady=10)
        
        # Opção 4: Por Funcionário
        rb4_frame = Frame(options_frame, bg="#cfe2ff", relief="groove", bd=2)
        rb4_frame.pack(fill=X, pady=5)
        rb4 = ttk.Radiobutton(rb4_frame, text="👤 ESCOLHER FUNCIONÁRIOS ESPECÍFICOS", 
                              variable=mode_var, value="funcionarios", style="Custom.TRadiobutton")
        rb4.pack(anchor="w", padx=10, pady=10)
        
        # Frame para listas de seleção (altura fixa para não cobrir botões)
        selection_frame = Frame(main_frame, bg="#f0f0f0", relief="groove", bd=2, height=200)
        selection_frame.pack(fill=X, pady=(15, 0))
        selection_frame.pack_propagate(False)  # Mantém altura fixa
        
        Label(selection_frame, text="SELECIONE OS ITENS:", 
              font=("Helvetica", 10, "bold"), bg="#f0f0f0").pack(pady=5)
        
        # Listbox com scrollbar (altura limitada)
        list_container = Frame(selection_frame, bg="#f0f0f0")
        list_container.pack(fill=BOTH, expand=True, padx=10, pady=(0, 5))
        
        scrollbar = ttk.Scrollbar(list_container)
        scrollbar.pack(side=RIGHT, fill=Y)
        
        listbox = Listbox(list_container, selectmode="multiple", 
                         yscrollcommand=scrollbar.set, font=("Helvetica", 9))
        listbox.pack(side=LEFT, fill=BOTH, expand=True)
        scrollbar.config(command=listbox.yview)
        
        # Botões de seleção rápida
        btn_frame = Frame(selection_frame, bg="#f0f0f0")
        btn_frame.pack(pady=5)
        
        def select_all():
            listbox.select_set(0, "end")
        
        def deselect_all():
            listbox.selection_clear(0, "end")
        
        Button(btn_frame, text="✓ MARCAR TODOS", command=select_all,
               bg="#27ae60", fg="white", font=("Helvetica", 8, "bold"),
               padx=8, pady=3).pack(side=LEFT, padx=3)
        
        Button(btn_frame, text="✗ DESMARCAR TODOS", command=deselect_all,
               bg="#e74c3c", fg="white", font=("Helvetica", 8, "bold"),
               padx=8, pady=3).pack(side=LEFT, padx=3)
        
        # Função para atualizar a lista baseado na seleção
        def update_list(*args):
            listbox.delete(0, "end")
            mode = mode_var.get()
            
            if mode == "postos":
                # Lista postos únicos
                postos = sorted(set(emp.get("posto", "SEM POSTO") for emp in self.funcionarios))
                for posto in postos:
                    listbox.insert("end", posto)
                selection_frame.pack(fill=X, pady=(15, 0))
            elif mode == "funcionarios":
                # Lista funcionários
                for emp in sorted(self.funcionarios, key=lambda x: x.get("nome", "")):
                    nome = emp.get("nome", "")
                    funcao = emp.get("funcao", "")
                    display_text = f"{nome} - {funcao}" if funcao else nome
                    listbox.insert("end", display_text)
                selection_frame.pack(fill=X, pady=(15, 0))
            else:
                # Esconde a lista para "todos" e "revisados"
                selection_frame.pack_forget()
        
        mode_var.trace("w", update_list)
        update_list()  # Atualiza inicialmente
        
        # Botões de ação com estilo melhorado
        action_frame = Frame(dlg, bg="#ecf0f1", relief="raised", bd=2)
        action_frame.pack(side="bottom", fill=X, padx=0, pady=0)
        
        button_container = Frame(action_frame, bg="#ecf0f1")
        button_container.pack(pady=15)
        
        def on_confirm():
            mode = mode_var.get()
            result["mode"] = mode
            
            if mode == "postos":
                # Pega postos selecionados
                selected_indices = listbox.curselection()
                result["selection"] = [listbox.get(i) for i in selected_indices]
                if not result["selection"]:
                    messagebox.showwarning("AVISO", "SELECIONE PELO MENOS UM POSTO", parent=dlg)
                    return
            elif mode == "funcionarios":
                # Pega funcionários selecionados (extrai nome antes do hífen)
                selected_indices = listbox.curselection()
                result["selection"] = []
                for i in selected_indices:
                    text = listbox.get(i)
                    # Formato: "NOME - FUNÇÃO" - extrai o nome
                    if " - " in text:
                        nome = text.split(" - ")[0]
                        result["selection"].append(nome)
                    else:
                        result["selection"].append(text)
                if not result["selection"]:
                    messagebox.showwarning("AVISO", "SELECIONE PELO MENOS UM FUNCIONÁRIO", parent=dlg)
                    return
            
            dlg.destroy()
        
        def on_cancel():
            result["mode"] = None
            dlg.destroy()
        
        Button(button_container, text="✓ GERAR PDFs", command=on_confirm,
               bg="#27ae60", fg="white", font=("Helvetica", 11, "bold"),
               relief="raised", bd=3, cursor="hand2",
               padx=30, pady=10).pack(side=LEFT, padx=10)
        
        Button(button_container, text="✗ CANCELAR", command=on_cancel,
               bg="#c0392b", fg="white", font=("Helvetica", 11, "bold"),
               relief="raised", bd=3, cursor="hand2",
               padx=30, pady=10).pack(side=LEFT, padx=10)
        
        dlg.wait_window()
        return result if result["mode"] else None

    def _show_professional_report(self, funcionarios_to_process, generated_files, nao_gerados_motivo):
        """Mostra relatório simples de geração de PDFs."""
        print(f"[DEBUG] Total de arquivos gerados: {len(generated_files)}")
        print(f"[DEBUG] Arquivos: {generated_files[:3] if generated_files else 'Nenhum'}")
        
        # Identifica quem teve PDF gerado - simplificado
        nomes_gerados = set()
        if generated_files:
            # Extrai nomes únicos dos funcionários que tiveram PDFs gerados
            for f in generated_files:
                try:
                    # Formato do arquivo: 11.2025_NOME COMPLETO.pdf
                    basename = os.path.basename(f).replace('.pdf', '')
                    
                    # Remove a parte da data (MM.YYYY_)
                    if '_' in basename:
                        nome_from_file = basename.split('_', 1)[1]  # Pega tudo depois do primeiro _
                    else:
                        nome_from_file = basename
                    
                    # Compara com os funcionários processados (case insensitive)
                    nome_from_file_upper = nome_from_file.upper().strip()
                    
                    for emp in funcionarios_to_process:
                        nome_emp = emp.get("nome", "").strip().upper()
                        
                        # Match exato ou contém
                        if nome_emp == nome_from_file_upper or nome_emp in nome_from_file_upper:
                            nomes_gerados.add(emp.get("nome", "").strip())
                            print(f"[DEBUG] Match encontrado: {emp.get('nome', '')} -> {basename}")
                            break
                except Exception as e:
                    print(f"[DEBUG] Erro ao processar arquivo {f}: {e}")
                    continue
        
        total_gerados = len(nomes_gerados)
        print(f"[DEBUG] Total de funcionários com PDFs: {total_gerados}")
        
        # Calcula quantos não foram gerados
        total_processados = len(funcionarios_to_process)
        nao_gerados = total_processados - total_gerados
        
        # Janela principal (altura ajustada se houver não gerados)
        altura_janela = 520 if nao_gerados > 0 else 350
        
        dlg = Toplevel(self.root)
        dlg.title("GERAÇÃO DE PDFs - CONCLUÍDA")
        dlg.configure(bg="#ecf0f1")
        dlg.transient(self.root)
        dlg.grab_set()
        dlg.geometry(f"500x{altura_janela}")
        dlg.resizable(False, False)
        
        # Centralizar na tela
        dlg.update_idletasks()
        x = (dlg.winfo_screenwidth() // 2) - (250)
        y = (dlg.winfo_screenheight() // 2) - (altura_janela // 2)
        dlg.geometry(f"500x{altura_janela}+{x}+{y}")
        
        # Header
        header_frame = Frame(dlg, bg="#27ae60" if total_gerados > 0 else "#e74c3c")
        header_frame.pack(fill=X)
        
        if total_gerados > 0:
            Label(header_frame, text="✓", font=("Segoe UI", 48), 
                  bg="#27ae60", fg="white").pack(pady=(30, 10))
            Label(header_frame, text="PROCESSO CONCLUÍDO!", 
                  font=("Segoe UI", 16, "bold"), bg="#27ae60", fg="white").pack(pady=(0, 30))
        else:
            Label(header_frame, text="⚠", font=("Segoe UI", 48), 
                  bg="#e74c3c", fg="white").pack(pady=(30, 10))
            Label(header_frame, text="NENHUM PDF GERADO", 
                  font=("Segoe UI", 16, "bold"), bg="#e74c3c", fg="white").pack(pady=(0, 30))
        
        # Conteúdo
        content_frame = Frame(dlg, bg="white")
        content_frame.pack(fill=BOTH, expand=True)
        
        # Mensagem principal
        if total_gerados > 0:
            msg_frame = Frame(content_frame, bg="white")
            msg_frame.pack(expand=True, pady=15)
            
            Label(msg_frame, text=str(total_gerados), 
                  font=("Segoe UI", 42, "bold"), bg="white", fg="#27ae60").pack()
            Label(msg_frame, text=f"de {total_processados} funcionário{'s' if total_processados > 1 else ''}", 
                  font=("Segoe UI", 11), bg="white", fg="#7f8c8d").pack(pady=(3, 0))
            
            if nao_gerados > 0:
                Label(msg_frame, text=f"({nao_gerados} não gerado{'s' if nao_gerados > 1 else ''})", 
                      font=("Segoe UI", 9), bg="white", fg="#e67e22").pack(pady=(3, 0))
        else:
            msg_frame = Frame(content_frame, bg="white")
            msg_frame.pack(expand=True, pady=20)
            
            Label(msg_frame, text="Nenhum PDF foi gerado", 
                  font=("Segoe UI", 13, "bold"), bg="white", fg="#7f8c8d").pack(pady=5)
            Label(msg_frame, text="Possíveis causas:", 
                  font=("Segoe UI", 10, "bold"), bg="white", fg="#95a5a6").pack(pady=(15, 5))
            
            causes_frame = Frame(msg_frame, bg="white")
            causes_frame.pack()
            Label(causes_frame, text="• Período selecionado sem dias úteis", 
                  font=("Segoe UI", 9), bg="white", fg="#95a5a6").pack(anchor="w", padx=20)
            Label(causes_frame, text="• Funcionários sem horários configurados", 
                  font=("Segoe UI", 9), bg="white", fg="#95a5a6").pack(anchor="w", padx=20)
            Label(causes_frame, text="• Data de admissão posterior ao período", 
                  font=("Segoe UI", 9), bg="white", fg="#95a5a6").pack(anchor="w", padx=20)
        
        # Botões
        btn_frame = Frame(dlg, bg="#ecf0f1")
        btn_frame.pack(fill=X, padx=30, pady=25)
        
        if total_gerados > 0:
            # Função para mostrar lista de funcionários
            def show_employee_list():
                list_dlg = Toplevel(dlg)
                list_dlg.title("FUNCIONÁRIOS GERADOS")
                list_dlg.configure(bg="#ecf0f1")
                list_dlg.transient(dlg)
                list_dlg.grab_set()
                list_dlg.geometry("600x500")
                
                # Centralizar
                list_dlg.update_idletasks()
                lx = (list_dlg.winfo_screenwidth() // 2) - (300)
                ly = (list_dlg.winfo_screenheight() // 2) - (250)
                list_dlg.geometry(f"600x500+{lx}+{ly}")
                
                # Header
                list_header = Frame(list_dlg, bg="#2c3e50")
                list_header.pack(fill=X)
                Label(list_header, text="👥 FUNCIONÁRIOS COM PDFs GERADOS", 
                      font=("Segoe UI", 13, "bold"), bg="#2c3e50", fg="white", 
                      pady=20).pack()
                
                # Lista com scroll
                list_container = Frame(list_dlg, bg="white")
                list_container.pack(fill=BOTH, expand=True, padx=15, pady=15)
                
                canvas = Canvas(list_container, bg="white", highlightthickness=0)
                scrollbar = ttk.Scrollbar(list_container, orient="vertical", command=canvas.yview)
                scrollable_frame = Frame(canvas, bg="white")
                
                scrollable_frame.bind(
                    "<Configure>",
                    lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
                )
                
                canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
                canvas.configure(yscrollcommand=scrollbar.set)
                
                # Preencher lista
                for idx, nome in enumerate(sorted(nomes_gerados), 1):
                    item_frame = Frame(scrollable_frame, bg="#f8f9fa" if idx % 2 == 0 else "white")
                    item_frame.pack(fill=X, pady=1)
                    
                    Label(item_frame, text=f"{idx}.", font=("Segoe UI", 10), 
                          bg=item_frame["bg"], fg="#95a5a6", width=4).pack(side=LEFT, padx=(15, 5), pady=15)
                    Label(item_frame, text=nome, font=("Segoe UI", 10), 
                          bg=item_frame["bg"], fg="#2c3e50").pack(side=LEFT, anchor="w", pady=15)
                
                canvas.pack(side=LEFT, fill=BOTH, expand=True)
                scrollbar.pack(side=RIGHT, fill=Y)
                
                # Botão OK
                Button(list_dlg, text="OK", command=list_dlg.destroy,
                       bg="#27ae60", fg="white", font=("Segoe UI", 11, "bold"),
                       relief="flat", bd=0, cursor="hand2", padx=50, pady=12).pack(pady=(0, 20))
                
                # Scroll com mouse
                def _on_mousewheel(event):
                    canvas.yview_scroll(int(-1*(event.delta/120)), "units")
                canvas.bind_all("<MouseWheel>", _on_mousewheel)
                
                def on_close():
                    canvas.unbind_all("<MouseWheel>")
                    list_dlg.destroy()
                
                list_dlg.protocol("WM_DELETE_WINDOW", on_close)
            
            Button(btn_frame, text="📋 LISTAR FUNCIONÁRIOS GERADOS", command=show_employee_list,
                   bg="#3498db", fg="white", font=("Segoe UI", 10, "bold"),
                   relief="flat", bd=0, cursor="hand2", padx=20, pady=12).pack(fill=X, pady=(0, 10))
            
            # Botão para mostrar não gerados (se houver)
            if nao_gerados > 0:
                def show_not_generated():
                    # Identifica quem não teve PDF gerado
                    nomes_nao_gerados = [emp.get("nome", "") for emp in funcionarios_to_process 
                                        if emp.get("nome", "").strip() not in nomes_gerados]
                    
                    list_dlg = Toplevel(dlg)
                    list_dlg.title("FUNCIONÁRIOS NÃO GERADOS")
                    list_dlg.configure(bg="#ecf0f1")
                    list_dlg.transient(dlg)
                    list_dlg.grab_set()
                    list_dlg.geometry("700x500")
                    
                    # Centralizar
                    list_dlg.update_idletasks()
                    lx = (list_dlg.winfo_screenwidth() // 2) - (350)
                    ly = (list_dlg.winfo_screenheight() // 2) - (250)
                    list_dlg.geometry(f"700x500+{lx}+{ly}")
                    
                    # Header
                    list_header = Frame(list_dlg, bg="#e74c3c")
                    list_header.pack(fill=X)
                    Label(list_header, text=f"⚠ {nao_gerados} FUNCIONÁRIOS NÃO GERARAM PDFs", 
                          font=("Segoe UI", 13, "bold"), bg="#e74c3c", fg="white", 
                          pady=20).pack()
                    
                    # Lista com scroll
                    list_container = Frame(list_dlg, bg="white")
                    list_container.pack(fill=BOTH, expand=True, padx=15, pady=15)
                    
                    canvas = Canvas(list_container, bg="white", highlightthickness=0)
                    scrollbar = ttk.Scrollbar(list_container, orient="vertical", command=canvas.yview)
                    scrollable_frame = Frame(canvas, bg="white")
                    
                    scrollable_frame.bind(
                        "<Configure>",
                        lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
                    )
                    
                    canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
                    canvas.configure(yscrollcommand=scrollbar.set)
                    
                    # Preencher lista com motivos
                    for idx, nome in enumerate(sorted(nomes_nao_gerados), 1):
                        item_frame = Frame(scrollable_frame, bg="#f8f9fa" if idx % 2 == 0 else "white", 
                                          relief="groove", bd=1)
                        item_frame.pack(fill=X, pady=2, padx=5)
                        
                        Label(item_frame, text=f"{idx}.", font=("Segoe UI", 9), 
                              bg=item_frame["bg"], fg="#95a5a6", width=4).pack(side=LEFT, padx=(10, 5), pady=10)
                        
                        name_motivo_frame = Frame(item_frame, bg=item_frame["bg"])
                        name_motivo_frame.pack(side=LEFT, fill=X, expand=True, pady=10)
                        
                        Label(name_motivo_frame, text=nome, font=("Segoe UI", 10, "bold"), 
                              bg=item_frame["bg"], fg="#2c3e50").pack(anchor="w")
                        
                        motivo = nao_gerados_motivo.get(nome, "Motivo não especificado")
                        Label(name_motivo_frame, text=f"Motivo: {motivo}", font=("Segoe UI", 8), 
                              bg=item_frame["bg"], fg="#7f8c8d", wraplength=550, justify=LEFT).pack(anchor="w")
                    
                    canvas.pack(side=LEFT, fill=BOTH, expand=True)
                    scrollbar.pack(side=RIGHT, fill=Y)
                    
                    # Botão OK
                    Button(list_dlg, text="OK", command=list_dlg.destroy,
                           bg="#95a5a6", fg="white", font=("Segoe UI", 11, "bold"),
                           relief="flat", bd=0, cursor="hand2", padx=50, pady=12).pack(pady=(0, 20))
                    
                    # Scroll com mouse
                    def _on_mousewheel(event):
                        canvas.yview_scroll(int(-1*(event.delta/120)), "units")
                    canvas.bind_all("<MouseWheel>", _on_mousewheel)
                    
                    def on_close():
                        canvas.unbind_all("<MouseWheel>")
                        list_dlg.destroy()
                    
                    list_dlg.protocol("WM_DELETE_WINDOW", on_close)
                
                Button(btn_frame, text=f"⚠ VER {nao_gerados} NÃO GERADOS", command=show_not_generated,
                       bg="#e67e22", fg="white", font=("Segoe UI", 10, "bold"),
                       relief="flat", bd=0, cursor="hand2", padx=20, pady=12).pack(fill=X, pady=(0, 10))
        
        Button(btn_frame, text="OK", command=dlg.destroy,
               bg="#27ae60" if total_gerados > 0 else "#95a5a6", fg="white", font=("Segoe UI", 11, "bold"),
               relief="flat", bd=0, cursor="hand2", padx=50, pady=12).pack(fill=X)
    
    def _show_scrollable_info(self, title: str, message: str):
        """Mostra uma caixa de diálogo com informações roláveis.
        
        Args:
            title: Título da janela
            message: Conteúdo da mensagem (pode ser longo)
        """
        dlg = Toplevel(self.root)
        dlg.title(title)
        dlg.configure(bg="#ecf0f1")
        dlg.transient(self.root)
        dlg.grab_set()
        
        # Configurar tamanho maior para melhor visualização
        dlg.geometry("750x550")
        
        # Centralizar na tela
        dlg.update_idletasks()
        x = (dlg.winfo_screenwidth() // 2) - (750 // 2)
        y = (dlg.winfo_screenheight() // 2) - (550 // 2)
        dlg.geometry(f"750x550+{x}+{y}")
        
        # Título moderno
        title_frame = Frame(dlg, bg="#2c3e50")
        title_frame.pack(fill=X)
        Label(title_frame, text=title, 
              font=("Segoe UI", 13, "bold"), bg="#2c3e50", fg="white", 
              pady=15).pack()
        
        # Frame principal moderno
        main_frame = Frame(dlg, bg="white", relief="flat", bd=0)
        main_frame.pack(fill=BOTH, expand=True, padx=15, pady=15)
        
        # Área de texto com scrollbar
        text_frame = Frame(main_frame, bg="white")
        text_frame.pack(fill=BOTH, expand=True)
        
        scrollbar = ttk.Scrollbar(text_frame)
        scrollbar.pack(side=RIGHT, fill=Y)
        
        text_widget = Text(text_frame, wrap=WORD, yscrollcommand=scrollbar.set,
                          font=("Consolas", 10), bg="#f8f9fa", fg="#2c3e50",
                          relief="flat", bd=0, padx=20, pady=20,
                          selectbackground="#3498db", selectforeground="white")
        text_widget.pack(side=LEFT, fill=BOTH, expand=True)
        scrollbar.config(command=text_widget.yview)
        
        # Inserir conteúdo - IMPORTANTE: inserir ANTES de desabilitar
        if message:
            text_widget.insert("1.0", message)
        else:
            text_widget.insert("1.0", "Nenhuma informação disponível.")
        
        text_widget.config(state=DISABLED)
        
        # Botão OK moderno
        btn_frame = Frame(dlg, bg="#ecf0f1")
        btn_frame.pack(pady=15)
        
        Button(btn_frame, text="✓ OK", command=dlg.destroy,
               bg="#27ae60", fg="white", font=("Segoe UI", 11, "bold"),
               relief="flat", bd=0, cursor="hand2",
               padx=50, pady=12).pack()
        
        dlg.wait_window()

    def generate_all_pdfs(self):
        # Verificação inicial
        if not self.funcionarios:
            messagebox.showerror("ERRO", "NENHUM FUNCIONÁRIO CARREGADO!\n\nCarregue uma planilha antes de gerar PDFs.")
            return
        
        # Mostra diálogo de seleção
        options = self.show_pdf_generation_dialog()
        if not options:
            return  # Usuário cancelou
        
        mode = options["mode"]
        selection = options["selection"]
        
        # Filtra funcionários baseado no modo selecionado
        funcionarios_to_process = []
        
        if mode == "todos":
            funcionarios_to_process = self.funcionarios
        elif mode == "revisados":
            # Filtra apenas funcionários marcados como revisados (OK)
            funcionarios_to_process = [emp for emp in self.funcionarios 
                                      if self.emp_revisado.get(emp.get("nome", ""), False)]
            if not funcionarios_to_process:
                messagebox.showwarning("AVISO", "NENHUM FUNCIONÁRIO MARCADO COMO OK (REVISADO)")
                return
        elif mode == "postos":
            # Filtra por postos selecionados
            funcionarios_to_process = [emp for emp in self.funcionarios 
                                      if emp.get("posto", "SEM POSTO") in selection]
            if not funcionarios_to_process:
                messagebox.showwarning("AVISO", "NENHUM FUNCIONÁRIO ENCONTRADO NOS POSTOS SELECIONADOS")
                return
        elif mode == "funcionarios":
            # Filtra por nomes selecionados
            funcionarios_to_process = [emp for emp in self.funcionarios 
                                      if emp.get("nome", "") in selection]
            if not funcionarios_to_process:
                messagebox.showwarning("AVISO", "NENHUM FUNCIONÁRIO SELECIONADO FOI ENCONTRADO")
                return
        
        period = self._get_period_dates()
        if not period:
            return
        start, end = period

        # accumuladores de relatório para feedback
        adm_issues = []   # (nome, admissao) -> quando nada foi gerado por estar todo o período antes da admissão
        adm_started = []  # (nome, admissao) -> quando gerou parcialmente a partir do mês de admissão
        generated_files = []

        # Mapeia duplicidades por posto considerando nome com e sem ponto final
        # Chave do grupo: (posto, nome_sem_ponto). Valor: lista de nomes originais desse posto
        dup_index_map: Dict[Tuple[str, str, str], int] = {}
        from collections import defaultdict
        grupos = defaultdict(list)
        for emp in funcionarios_to_process:
            posto_emp = emp.get("posto", "SEM POSTO") or "SEM POSTO"
            nome_emp = emp.get("nome", "") or ""
            nome_base = nome_emp.rstrip(".").strip()
            grupos[(posto_emp, nome_base)].append(nome_emp)
        # Atribui índices (1,2,...) quando houver mais de um no mesmo grupo
        for (posto_emp, nome_base), nomes_lista in grupos.items():
            if len(nomes_lista) > 1:
                # Ordena para deixar o sem ponto primeiro
                nomes_ordenados = sorted(nomes_lista, key=lambda n: (n.endswith("."), n))
                for idx, nome_original in enumerate(nomes_ordenados, start=1):
                    dup_index_map[(posto_emp, nome_base, nome_original)] = idx
        nao_gerados_motivo = {}  # Dict[nome, motivo] - rastreia motivo de cada não gerado

        # Primeira passada: preparar agendas e coletar relatórios
        prepared = []  # cada item: {nome, emp, conf, filtered_schedule, version_index}

        for emp in funcionarios_to_process:
            nome = emp.get("nome", "")
            posto_do_emp = emp.get("posto", "SEM POSTO") or "SEM POSTO"
            nome_base = nome.rstrip(".").strip()
            version_index = dup_index_map.get((posto_do_emp, nome_base, nome))
            admissao = emp.get("admissao", None)
            
            # Passando a escala e 1ª folga para a função de agendamento
            scale_type = self.emp_scale_choice.get(nome, "6X1 (FIXO)")
            first_off_str = self.emp_first_off.get(nome)
            
            conf = {
                "start_date": start,
                "end_date": end,
                "holidays": self.global_holidays,
                "holiday_type": self.holiday_type,
                "holiday_postos": self.holiday_postos,
                "scale_type": scale_type,
                "first_off": first_off_str
            }

            # build schedule and detect admission month constraints
            sch = {}
            
            # A função generate_employee_schedule agora usa os novos parâmetros de escala
            for ds, entry in generate_employee_schedule(emp, conf, self.emp_faltas_atestados):
                # record schedule
                sch[ds] = entry

            # If admission after start: we need to skip months before admissao.
            # Build list of months that would be generated; skip months entirely before admission
            # and report them.
            months = sorted({(datetime.strptime(ds, "%Y-%m-%d").year, datetime.strptime(ds, "%Y-%m-%d").month) for ds in sch.keys()})
            # filter months based on admission
            months_to_generate = []
            months_skipped_by_adm = []
            for (yr, mo) in months:
                # if admission exists and the whole month is before admission, skip it
                if admissao:
                    adm_month_year = (admissao.year, admissao.month)
                    if (yr, mo) < adm_month_year:
                        months_skipped_by_adm.append((yr, mo))
                        continue
                months_to_generate.append((yr, mo))

            # generate a filtered schedule that only includes days of months_to_generate
            filtered_schedule = {}
            months_set = set(months_to_generate)
            for ds, e in sch.items():
                try:
                    dt = datetime.strptime(ds, "%Y-%m-%d").date()
                except Exception:
                    continue
                key = (dt.year, dt.month)
                if key in months_set:
                    filtered_schedule[ds] = e

            # if filtered_schedule empty (e.g., admissão após o fim do período ou sem dias no months_to_generate)
            if not filtered_schedule:
                # report admission issue if any months were skipped by admission
                if months_skipped_by_adm:
                    adm_issues.append((nome, emp.get("admissao")))
                    nao_gerados_motivo[nome] = f"Admitido após o período ({emp.get('admissao').strftime('%d/%m/%Y') if emp.get('admissao') else 'data não informada'})"
                else:
                    nao_gerados_motivo[nome] = "Sem dias no período selecionado (agenda vazia)"
                # nothing to generate for this employee
                continue

            # Se houve meses anteriores pulados por causa da admissão, registra mensagem de "gerado a partir de"
            if months_skipped_by_adm and admissao:
                try:
                    adm_started.append((nome, admissao))
                except Exception:
                    pass

            # ensure posto from spreadsheet is used if present; put into entries if empty
            for ds, e in filtered_schedule.items():
                if not e.get("posto"):
                    e["posto"] = emp.get("posto", "") or ""

            prepared.append({
                "nome": nome,
                "emp": emp,
                "conf": conf,
                "filtered_schedule": filtered_schedule,
                "version_index": version_index,
            })

        # Segunda passada: gerar PDFs
        print(f"[DEBUG] Iniciando geração de PDFs para {len(prepared)} funcionários")
        for item in prepared:
            nome = item["nome"]
            emp = item["emp"]
            filtered_schedule = item["filtered_schedule"]
            
            print(f"[DEBUG] Gerando PDF para {nome}, dias no schedule: {len(filtered_schedule)}")

            try:
                saved = generate_pdf_for_employee(
                    nome=emp.get("nome", ""),
                    cpf=emp.get("cpf", ""),
                    matricula=emp.get("matricula", ""),
                    funcao=emp.get("funcao", ""),
                    posto_global=emp.get("posto", ""),
                    filial=emp.get("filial", ""),
                    cnpj=emp.get("cnpj", ""),
                    endereco=emp.get("endereco", ""),
                    cidade=emp.get("cidade", ""), 
                    schedule_map=filtered_schedule,

                    out_folder=OUTPUT_FOLDER,
                    version_index=item.get("version_index")
                )
                print(f"[DEBUG] {nome}: saved={len(saved) if saved else 0} arquivos")
                if saved:
                    generated_files.extend(saved)
                else:
                    # Se não salvou nenhum arquivo
                    if nome not in nao_gerados_motivo:
                        nao_gerados_motivo[nome] = "Nenhum mês com dias úteis para gerar PDF"
            except Exception as e:
                print(f"[{now_str()}] ERRO AO GERAR PDF PARA {nome}: {e}\n{traceback.format_exc()}")
                nao_gerados_motivo[nome] = f"Erro durante a geração: {str(e)[:50]}"

        # Criar relatório visual profissional
        self._show_professional_report(funcionarios_to_process, generated_files, nao_gerados_motivo)
        
def main():
    safe_mkdir(OUTPUT_FOLDER)
    root = Tk()
    root.title("GERADOR DE FOLHA DE PONTO - V.5.39")
    root.geometry("1280x800")
    root.configure(bg="#ecf0f1")
    
    # Ícone da janela (opcional, se disponível)
    try:
        root.iconbitmap("icon.ico")
    except:
        pass

    # Tratamento para restaurar janela ao clicar quando minimizada
    def restore_window(event=None):
        try:
            if root.state() == 'iconic':  # Se está minimizada
                root.deiconify()
                root.state('normal')
                root.lift()
                root.focus_force()
        except Exception:
            pass
    
    def check_window_state(event=None):
        try:
            if root.state() == 'iconic':
                root.after(50, restore_window)
        except Exception:
            pass
    
    # Bind múltiplos eventos para garantir restauração
    root.bind('<Map>', restore_window)
    root.bind('<Visibility>', check_window_state)

    app = PontoApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
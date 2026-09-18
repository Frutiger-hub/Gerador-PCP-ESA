# -*- coding: utf-8 -*-
"""
GERADOR DE RELATORIO PCP - ESA | V9.8.7
Aplicacao simples para Windows: Tkinter + PyMuPDF + ReportLab.

Como usar:
A aplicação pode ser instalada pelo instalador do pacote, sem configuração manual de Python, pip ou bibliotecas.

A aplicacao permite:
- importar um ou varios relatorios PDF;
- identificar automaticamente fase, referencia e data de emissao;
- extrair OS, cliente, marca, potencia, previsao, colaborador e entrada;
- ignorar descricoes como BOMBA ABS, BOMBA FLYGT, BOMBAS KSB e MOTOR//BOMBA C/
  como se fossem colaboradores;
- gerar um PDF gerencial por relatorio;
- gerar analise executiva, prioridades, aging, entregas, colaboradores, clientes, marcas, OS em multiplas fases, alertas de dados e base completa;
- consultar as O.S. carregadas por perguntas e filtros rápidos e gerar PDF da consulta.
"""

import os
import re
import sys
import subprocess
import json
import ctypes
from pathlib import Path
from datetime import datetime, date, timedelta
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

try:
    import pymupdf
except ImportError:
    pymupdf = None

try:
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.platypus import (
        SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
    )
    from reportlab.lib.units import mm
except ImportError:
    colors = None


APP_TITLE = "Gerador de Relatorios PCP - ESA"
WINDOW_APP_ID = "ESA.GeradorPCP"
TODAY = date.today()


def set_windows_app_user_model_id():
    """Define um AppUserModelID próprio para o Windows identificar o app.

    Isso evita que a barra de tarefas trate a janela como um simples processo
    Python e permite que o ícone do aplicativo seja usado como ícone da janela
    e da entrada correspondente na barra de tarefas.
    """
    if sys.platform != "win32":
        return
    try:
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(WINDOW_APP_ID)
    except Exception:
        pass

# Cores de status para destacar registros inteiros nas tabelas do PDF.
# A paleta é propositalmente mais contrastante para facilitar a leitura impressa e na tela.
STATUS_ROW_COLORS = {
    "Atrasada": "#FFD1D6",
    "Vence hoje": "#FFE39A",
    "Próximos 7 dias": "#FFF0B8",
    "Prazo futuro": "#FFFFFF",
    "Sem previsão": "#E7EAEE",
}

def _status_fill(status):
    return colors.HexColor(STATUS_ROW_COLORS.get(status, "#FFFFFF"))


# Palavras que indicam que o campo nao e um nome de pessoa.
# A heuristica e conservadora para evitar classificar o tipo de bomba
# como colaborador.
NON_PERSON_WORDS = {
    "BOMBA", "BOMBAS", "MOTOR", "MOTORES", "MOTOR//BOMBA",
    "CA", "ABS", "FLYGT", "KSB", "WEG", "SEW", "EBARA", "SULZER",
    "SCHNEIDER", "GRUNDFOS", "IMBIL", "SIEMENS", "VOGES",
    "MOTOBOMBA", "EQUIPAMENTO", "OFICINA", "MANUTENCAO",
}

def clean_text(text):
    text = text.replace("\xa0", " ")
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def parse_date(value):
    try:
        return datetime.strptime(value, "%d/%m/%Y").date()
    except Exception:
        return None


def is_person_name(segment):
    """
    Decide se o texto parece nome de colaborador.
    Os PDFs de origem normalmente trazem nomes em caixa alta e abreviados.
    """
    if not segment:
        return False

    s = clean_text(segment).upper()

    if any(word in s.split() for word in NON_PERSON_WORDS):
        return False
    if "BOMBA" in s or "MOTOR" in s or "MOTOBOMBA" in s:
        return False

    # remove ruido tipico
    s = re.sub(r"[^A-ZÀ-Ú ]", " ", s)
    s = clean_text(s)

    words = s.split()
    if len(words) < 2:
        return False
    if len(words) > 5:
        return False

    # Pessoa normalmente aparece como nome + sobrenome, mesmo truncado.
    return all(len(w) >= 2 for w in words)


def extract_collaborator(segment):
    """Extrai somente o nome de pessoa do trecho antes da data de inicio."""
    segment = clean_text(segment)
    if not segment:
        return ""

    # O relatorio costuma trazer o colaborador seguido do tipo/equipamento:
    # "CLECIO RAMOS BOMBA FLYGT". Corta a partir da primeira palavra operacional.
    words = segment.split()
    cut = len(words)
    operational_markers = {
        "BOMBA", "BOMBAS", "MOTOR", "MOTOBOMBA", "CA", "ABS", "FLYGT",
        "KSB", "WEG", "SEW", "EBARA", "SULZER", "SCHNEIDER", "GRUNDFOS",
        "IMBIL", "SIEMENS", "VOGES"
    }
    for i, w in enumerate(words):
        token = re.sub(r"[^A-ZÀ-Ú/]+", "", w.upper())
        if token in operational_markers or "BOMBA" in token or "MOTOR" in token:
            cut = i
            break

    candidate = clean_text(" ".join(words[:cut]))
    return candidate if is_person_name(candidate) else ""


def _parse_phase_section(section_text, phase_num, phase_name, emission, referencia):
    """Parseia uma seção de fase do relatório ESA."""
    lines = [clean_text(x) for x in section_text.splitlines() if clean_text(x)]
    date_re = re.compile(r"\d{2}/\d{2}/\d{4}")
    row_start = re.compile(
        r"^(?P<os>\d{5,8})\s+(?P<os2>\d{5,8})\s+(?P<data_os>\d{2}/\d{2}/\d{4})\s+(?P<rest>.+)$"
    )
    known_brands = {
        "WEG", "FLYGT", "ABS", "EBARA", "KSB", "SULZER", "SCHNEIDER",
        "GRUNDFOS", "IMBIL", "SEW", "SIEMENS", "VOGES", "ASTEN", "EBERLE",
        "REALINCE", "REALIANCE", "MERCO SUL", "MERCOSUL", "GLASS", "ECOFLUX",
        "SPV"
    }
    total_fase = None
    mtotal = re.search(r"Total desta Fase\s*-+>\s*(\d+)", section_text, re.I)
    if mtotal:
        total_fase = int(mtotal.group(1))

    rows = []
    for line in lines:
        m = row_start.match(line)
        if not m or m.group("os") != m.group("os2"):
            continue

        data_os = parse_date(m.group("data_os"))
        restline = clean_text(m.group("rest"))
        dates = list(date_re.finditer(restline))
        if not dates:
            continue

        prev = parse_date(dates[0].group(0)) if len(dates) >= 2 else None
        initial_text = clean_text(restline[:dates[0].start()])

        if len(dates) >= 2:
            data_inicio = parse_date(dates[1].group(0))
            after_start = restline[dates[1].end():]
            hm = re.search(r"\b(\d{2}:\d{2})\b", after_start)
            hora_inicio = hm.group(1) if hm else ""
            collab_segment = clean_text(restline[dates[0].end():dates[1].start()])
        else:
            data_inicio = data_os
            hm = re.search(r"\b(\d{2}:\d{2})\b", restline[dates[0].end():])
            hora_inicio = hm.group(1) if hm else ""
            collab_segment = ""

        tokens = initial_text.split()
        client_code = ""
        client = ""
        marca = ""
        potencia = ""
        modelo_tag = ""

        power_idx = None
        power_len = 1
        for i in range(len(tokens)):
            if (
                re.fullmatch(r"\d+(?:[.,]\d+)?", tokens[i])
                and i + 1 < len(tokens)
                and tokens[i + 1].upper() in {"CV", "KW"}
            ):
                power_idx = i
                power_len = 2
                break
            if re.fullmatch(r"\d+(?:[.,]\d+)?(?:CV|KW)", tokens[i].upper()):
                power_idx = i
                break

        before = tokens
        if power_idx is not None:
            potencia = " ".join(tokens[power_idx:power_idx + power_len])
            before = tokens[:power_idx]

        if before and re.fullmatch(r"\d{2,6}", before[0]):
            client_code = before[0]
            before = before[1:]

        brand_idx = next((i for i, t in enumerate(before) if t.upper() in known_brands), None)
        if brand_idx is not None:
            marca = before[brand_idx]
            client = " ".join(before[:brand_idx]).strip()
            modelo_tag = " ".join(before[brand_idx + 1:]).strip()
        else:
            client = " ".join(before).strip()

        collaborator = extract_collaborator(collab_segment)

        row = {
            "OS": m.group("os"),
            "Data OS": data_os,
            "Cliente": client,
            "Cliente Código": client_code,
            "Marca": marca,
            "Modelo/TAG": modelo_tag,
            "Potência": potencia,
            "Prev. Entrega": prev,
            "Colaborador": collaborator,
            "Data Início": data_inicio,
            "Hora Início": hora_inicio,
            "Fase Num": str(phase_num),
            "Fase Nome": phase_name,
        }
        # Uma OS pode aparecer em várias fases. A deduplicação ocorre apenas dentro da fase.
        if not any(r["OS"] == row["OS"] for r in rows):
            rows.append(row)

    return {
        "fase_num": str(phase_num),
        "fase_nome": phase_name,
        "total_fase": total_fase if total_fase is not None else len(rows),
        "rows": rows,
        "emissao": emission,
        "referencia": referencia,
    }


def extract_report(pdf_path):
    """Extrai uma ou várias fases presentes no mesmo PDF ESA.

    O parser identifica cada bloco iniciado por ``Fase:`` e mantém a fase
    associada a cada OS. A mesma OS pode aparecer novamente em outra fase,
    sem ser descartada, pois ela representa uma etapa diferente do processo.
    """
    if pymupdf is None:
        raise RuntimeError("PyMuPDF nao esta instalado. Rode: pip install pymupdf")

    doc = pymupdf.open(pdf_path)
    full_text = "\n".join(page.get_text("text") for page in doc)
    lines = full_text.splitlines()

    emission = None
    referencia = ""
    m = re.search(r"Emissao:\s*(\d{2}/\d{2}/\d{4})\s*-\s*(\d{2}:\d{2})", full_text, re.I)
    if m:
        emission = datetime.strptime(f"{m.group(1)} {m.group(2)}", "%d/%m/%Y %H:%M")
    m = re.search(r"Referencia:\s*([A-Za-z]{3}/\d{4})", full_text, re.I)
    if m:
        referencia = m.group(1)

    phase_matches = list(re.finditer(r"(?mi)^Fase:\s*(\d+)\s+(.+?)\s*(?=\s+Periodo de:|$)", full_text))
    phases = []
    for idx, pm in enumerate(phase_matches):
        section_start = pm.start()
        section_end = phase_matches[idx + 1].start() if idx + 1 < len(phase_matches) else len(full_text)
        section = full_text[section_start:section_end]
        phase_name = clean_text(pm.group(2))
        phases.append(_parse_phase_section(section, pm.group(1), phase_name, emission, referencia))

    phases = [ph for ph in phases if ph["rows"] or ph["total_fase"]]
    if not phases:
        raise ValueError(
            f"Nao foi possivel localizar nenhuma fase/OS no formato esperado em:\n{pdf_path}\n\n"
            "Confira se o PDF possui texto selecionavel e segue o layout da ESA."
        )

    rows = []
    for ph in phases:
        rows.extend(ph["rows"])

    report = {
        "file": str(pdf_path),
        "emissao": emission,
        "referencia": referencia,
        "fase_num": phases[0]["fase_num"] if len(phases) == 1 else "",
        "fase_nome": phases[0]["fase_nome"] if len(phases) == 1 else "MULTIFASE",
        "total_fase": phases[0]["total_fase"] if len(phases) == 1 else sum(ph["total_fase"] for ph in phases),
        "rows": rows,
        "phases": phases,
        "multi_fase": len(phases) > 1,
    }
    return report

DELAY_REASONS = [
    "Aguardando material",
    "Aguardando peça",
    "Aguardando aprovação",
    "Aguardando cliente",
    "Falta de mão de obra",
    "Sobrecarga da equipe",
    "Máquina/equipamento parado",
    "Manutenção",
    "Retrabalho",
    "Problema técnico",
    "Falta de informação",
    "Prioridade alterada",
    "Dependência de outra fase",
    "O.S. não iniciada",
    "Outros",
]

def delay_key(row):
    return f"{row.get('Fase Num','-')}::{row.get('OS','-')}"

def safe_pct(n, total):
    return f"{(n/total):.1%}" if total else "0.0%"

def classify_status(prev, ref_date):
    if not prev:
        return "Sem previsão"
    if prev < ref_date:
        return "Atrasada"
    if prev == ref_date:
        return "Vence hoje"
    days = (prev - ref_date).days
    if days <= 7:
        return "Próximos 7 dias"
    return "Prazo futuro"


def enrich(report):
    ref_date = report["emissao"].date() if report["emissao"] else TODAY
    for row in report["rows"]:
        row["Status"] = classify_status(row["Prev. Entrega"], ref_date)
        if row.get("Data Início"):
            row["Dias na Fase"] = max(0, (ref_date - row["Data Início"]).days)
        else:
            row["Dias na Fase"] = None
        row.setdefault("Motivo Atraso", "")
        row.setdefault("Observação Atraso", "")
        row.setdefault("Responsável Atraso", "")
        row.setdefault("Ação Atraso", "")
        row.setdefault("Prazo Ação", "")
        score, label = _criticality(row, ref_date)
        row["Criticidade Score"] = score
        row["Criticidade"] = label
    for ph in report.get("phases", []):
        ph["ref_date"] = ref_date
        ph["rows"] = [r for r in report["rows"] if r.get("Fase Num") == ph["fase_num"]]
        ph["status_counts"] = {
            s: sum(1 for r in ph["rows"] if r.get("Status") == s)
            for s in ["Atrasada", "Vence hoje", "Próximos 7 dias", "Prazo futuro", "Sem previsão"]
        }
    report["ref_date"] = ref_date
    report["status_counts"] = {
        s: sum(1 for r in report["rows"] if r.get("Status") == s)
        for s in ["Atrasada", "Vence hoje", "Próximos 7 dias", "Prazo futuro", "Sem previsão"]
    }
    return report

def ptxt(s):
    if s is None:
        return ""
    return str(s)


def fmt_date(d):
    return d.strftime("%d/%m/%Y") if d else "Sem previsão"


def generate_pdf_single(report, output_dir):
    if colors is None:
        raise RuntimeError("ReportLab nao esta instalado. Rode: pip install reportlab")

    os.makedirs(output_dir, exist_ok=True)

    phase_num = report["fase_num"] or "fase"
    phase_name = report["fase_nome"] or "relatorio"
    safe_name = re.sub(r'[^A-Za-z0-9_-]+', '_', f"Fase_{phase_num}_{phase_name}").strip("_")
    output_path = os.path.join(output_dir, safe_name + ".pdf")

    rows = report["rows"]
    ref_date = report["ref_date"]

    total = len(rows)
    status_order = ["Atrasada", "Vence hoje", "Próximos 7 dias", "Prazo futuro", "Sem previsão"]
    status_counts = {s: sum(1 for r in rows if r["Status"] == s) for s in status_order}

    # Tempo na fase: data de referência do relatório menos a data de início registrada.
    for r in rows:
        if r.get("Data Início"):
            r["Dias na Fase"] = max(0, (ref_date - r["Data Início"]).days)
        else:
            r["Dias na Fase"] = None

    time_values = [r["Dias na Fase"] for r in rows if r["Dias na Fase"] is not None]
    avg_days = sum(time_values) / len(time_values) if time_values else 0
    max_days = max(time_values) if time_values else 0

    clients = {}
    brands = {}
    collabs = {}
    for r in rows:
        clients[r["Cliente"] or "Não informado"] = clients.get(r["Cliente"] or "Não informado", 0) + 1
        brands[r["Marca"] or "Não informado"] = brands.get(r["Marca"] or "Não informado", 0) + 1
        if r["Colaborador"]:
            collabs[r["Colaborador"]] = collabs.get(r["Colaborador"], 0) + 1

    doc = SimpleDocTemplate(
        output_path,
        pagesize=A4,
        rightMargin=10*mm,
        leftMargin=10*mm,
        topMargin=10*mm,
        bottomMargin=10*mm
    )

    styles = getSampleStyleSheet()
    title = ParagraphStyle(
        "title", parent=styles["Title"], fontSize=17, leading=20,
        textColor=colors.HexColor("#1F4E78"), spaceAfter=5
    )
    h = ParagraphStyle(
        "h", parent=styles["Heading2"], fontSize=10.5,
        textColor=colors.HexColor("#1F4E78"), spaceBefore=6, spaceAfter=4
    )
    body = ParagraphStyle(
        "body", parent=styles["BodyText"], fontSize=8.1, leading=10.5
    )
    small = ParagraphStyle(
        "small", parent=body, fontSize=6.5, leading=8
    )

    story = [
        Paragraph(
            f"RELATÓRIO PCP | FASE {phase_num} - {phase_name}",
            title
        ),
        Paragraph(
            f"Referência: {report['referencia'] or '-'} | "
            f"Emissão: {report['emissao'].strftime('%d/%m/%Y %H:%M') if report['emissao'] else '-'}",
            body
        ),
        Spacer(1, 5)
    ]

    kpi = Table(
        [
            ["TOTAL OS", "ATRASADAS", "VENCEM HOJE", "PRÓX. 7 DIAS", "PRAZO FUTURO"],
            [total, status_counts["Atrasada"], status_counts["Vence hoje"],
             status_counts["Próximos 7 dias"], status_counts["Prazo futuro"]]
        ],
        colWidths=[34*mm]*5, rowHeights=[7*mm, 9*mm]
    )
    kpi.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#D9EAF7")),
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
        ("FONTNAME", (0,1), (-1,1), "Helvetica-Bold"),
        ("FONTSIZE", (0,0), (-1,0), 6.5),
        ("FONTSIZE", (0,1), (-1,1), 13),
        ("ALIGN", (0,0), (-1,-1), "CENTER"),
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
        ("GRID", (0,0), (-1,-1), .3, colors.HexColor("#D0D0D0")),
        ("BACKGROUND", (1,1), (1,1), colors.HexColor("#FCE4D6")),
    ]))
    story.append(kpi)
    tempo_kpi = Table(
        [
            ["TEMPO MÉDIO NA FASE", "MAIOR TEMPO NA FASE", "OS COM INÍCIO REGISTRADO"],
            [f"{avg_days:.1f} dias", f"{max_days} dias", str(len(time_values))]
        ],
        colWidths=[56*mm, 56*mm, 56*mm], rowHeights=[6.5*mm, 9*mm]
    )
    tempo_kpi.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(-1,0),colors.HexColor("#EAF1F7")),
        ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),
        ("FONTNAME",(0,1),(-1,1),"Helvetica-Bold"),
        ("FONTSIZE",(0,0),(-1,0),6.3),
        ("FONTSIZE",(0,1),(-1,1),11),
        ("ALIGN",(0,0),(-1,-1),"CENTER"),
        ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
        ("GRID",(0,0),(-1,-1),.3,colors.HexColor("#CDD8E2")),
    ]))
    story.append(Spacer(1, 2))
    story.append(tempo_kpi)
    story.append(Spacer(1, 4))

    story.append(Paragraph("1. Diagnóstico executivo", h))
    atrasadas = status_counts["Atrasada"]
    proximas = status_counts["Próximos 7 dias"] + status_counts["Vence hoje"]
    if total:
        story.append(Paragraph(
            f"A carteira possui {total} OS. Em relação à data de referência "
            f"{ref_date.strftime('%d/%m/%Y')}, {atrasadas} estão atrasadas "
            f"({atrasadas/total:.1%}) e {proximas} exigem acompanhamento de curto prazo. "
            f"O tempo médio das OS com início registrado na fase é de {avg_days:.1f} dias, "
            f"com máximo de {max_days} dias. Há {status_counts['Sem previsão']} OS sem previsão "
            "quando aplicável. A análise prioriza recuperação de atraso e controle das próximas entregas.",
            body
        ))

    story.append(Paragraph("2. Distribuição da carteira por prazo", h))
    st = [["Status", "Quantidade", "Participação"]]
    for s in status_order:
        n = status_counts[s]
        st.append([s, n, f"{n/total:.1%}" if total else "0.0%"])

    t = Table(st, colWidths=[58*mm, 30*mm, 35*mm], repeatRows=1)
    t.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(-1,0),colors.HexColor("#1F4E78")),
        ("TEXTCOLOR",(0,0),(-1,0),colors.white),
        ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),
        ("GRID",(0,0),(-1,-1),.35,colors.HexColor("#BFC9D4")),
        ("ALIGN",(1,1),(-1,-1),"CENTER"),
        ("FONTSIZE",(0,0),(-1,-1),8),
    ]))
    for i, st_name in enumerate(status_order, 1):
        t.setStyle(TableStyle([("BACKGROUND", (0,i), (-1,i), _status_fill(st_name))]))
    story.append(t)

    # Prioridade
    story.append(Paragraph("3. Prioridade de atuação do PCP", h))
    priority_rows = [
        ["Prioridade", "Qtd.", "Ação recomendada"],
        ["1 - Recuperação", status_counts["Atrasada"],
         "Tratar OS atrasadas, identificar causa e definir nova data real."],
        ["2 - Curto prazo", status_counts["Vence hoje"] + status_counts["Próximos 7 dias"],
         "Acompanhar diariamente, validando capacidade, material e sequência."],
        ["3 - Programação", status_counts["Prazo futuro"],
         "Manter na programação e acompanhar avanço conforme plano."],
        ["4 - Dados", status_counts["Sem previsão"],
         "Definir previsão de entrega para melhorar o controle do PCP."]
    ]
    pt = Table(priority_rows, colWidths=[40*mm,18*mm,105*mm], repeatRows=1)
    pt.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(-1,0),colors.HexColor("#1F4E78")),
        ("TEXTCOLOR",(0,0),(-1,0),colors.white),
        ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),
        ("GRID",(0,0),(-1,-1),.3,colors.HexColor("#D0D0D0")),
        ("FONTSIZE",(0,0),(-1,-1),7.3),
        ("VALIGN",(0,0),(-1,-1),"TOP")
    ]))
    story.append(pt)

    story.append(Paragraph("4. Colaboradores registrados na entrada", title))
    if collabs:
        crows = [["Colaborador", "OS", "Quantidade"]]
        for name, count in sorted(collabs.items(), key=lambda x: (-x[1], x[0])):
            oslist = ", ".join(str(r["OS"]) for r in rows if r["Colaborador"] == name)
            crows.append([name, oslist, count])
    else:
        crows = [["Colaborador", "OS", "Quantidade"], ["Nenhum nome identificado", "-", 0]]

    ct = Table(crows, colWidths=[55*mm,75*mm,25*mm], repeatRows=1)
    ct.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(-1,0),colors.HexColor("#1F4E78")),
        ("TEXTCOLOR",(0,0),(-1,0),colors.white),
        ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),
        ("GRID",(0,0),(-1,-1),.3,colors.HexColor("#D0D0D0")),
        ("ALIGN",(2,1),(2,-1),"CENTER"),
        ("FONTSIZE",(0,0),(-1,-1),7.2),
        ("VALIGN",(0,0),(-1,-1),"TOP"),
    ]))
    story.append(ct)

    story.append(Paragraph("5. Tempo na fase por OS", h))
    time_rows = [["OS", "Data de início", "Tempo na fase", "Status"]]
    time_sorted = sorted(
        rows, key=lambda x: (x["Dias na Fase"] is None, -(x["Dias na Fase"] or 0))
    )
    for r in time_sorted:
        days_txt = f"{r['Dias na Fase']} dias" if r["Dias na Fase"] is not None else "Sem informação"
        start_txt = r["Data Início"].strftime("%d/%m/%Y") if r["Data Início"] else "Não informado"
        time_rows.append([r["OS"], start_txt, days_txt, r["Status"]])
    tt = Table(time_rows, colWidths=[24*mm,35*mm,38*mm,53*mm], repeatRows=1)
    tt.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(-1,0),colors.HexColor("#173A5E")),
        ("TEXTCOLOR",(0,0),(-1,0),colors.white),
        ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),
        ("GRID",(0,0),(-1,-1),.3,colors.HexColor("#D2D8DE")),
        ("ALIGN",(0,1),(2,-1),"CENTER"),
        ("FONTSIZE",(0,0),(-1,-1),6.6),
    ]))
    for i, r in enumerate(time_sorted, 1):
        tt.setStyle(TableStyle([("BACKGROUND", (0,i), (-1,i), _status_fill(r.get("Status"))) ]))
    story.append(tt)

    story.append(Paragraph("6. OS atrasadas", h))
    late = [r for r in rows if r["Status"] == "Atrasada"]
    late_rows = [["OS", "Cliente", "Marca", "Pot.", "Prev. Entrega", "Colaborador", "Motivo"]]
    for r in late:
        late_rows.append([
            r["OS"], _cell_para(r["Cliente"] or "N/I"), _cell_para(r["Marca"] or "N/I"),
            _cell_para(r["Potência"] or "N/I"), fmt_date(r["Prev. Entrega"]),
            _cell_para(r["Colaborador"] or "Não informado"), _cell_para(r.get("Motivo Atraso") or "Não informado")
        ])

    if len(late_rows) == 1:
        late_rows.append(["-", "Nenhuma OS atrasada", "-", "-", "-", "-", "-"])

    lt = Table(late_rows, colWidths=[14*mm,42*mm,18*mm,15*mm,25*mm,33*mm,28*mm], repeatRows=1)
    lt.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(-1,0),colors.HexColor("#1F4E78")),
        ("TEXTCOLOR",(0,0),(-1,0),colors.white),
        ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),
        ("GRID",(0,0),(-1,-1),.25,colors.HexColor("#D0D0D0")),
        ("FONTSIZE",(0,0),(-1,-1),6.4),
        ("BACKGROUND",(0,1),(-1,-1),colors.HexColor("#FCE4D6")),
        ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
    ]))
    story.append(lt)

    story.append(Paragraph("7. Motivos e ações dos atrasos", h))
    if late:
        reason_counts = {}
        for r in late:
            reason = r.get("Motivo Atraso") or "Não informado"
            reason_counts[reason] = reason_counts.get(reason, 0) + 1
        reason_rows = [["Motivo", "OS", "%", "Responsável / ação"]]
        for reason, count in sorted(reason_counts.items(), key=lambda x: (-x[1], x[0])):
            acts = [
                f"{r.get('Responsável Atraso')} | {r.get('Ação Atraso')}"
                for r in late if (r.get("Motivo Atraso") or "Não informado") == reason
                and (r.get("Responsável Atraso") or r.get("Ação Atraso"))
            ]
            reason_rows.append([_cell_para(reason), count, safe_pct(count, len(late)), _cell_para(" ; ".join(acts[:3]) if acts else "-")])
        rt = Table(reason_rows, colWidths=[47*mm,15*mm,18*mm,65*mm], repeatRows=1)
        rt.setStyle(TableStyle([
            ("BACKGROUND",(0,0),(-1,0),colors.HexColor("#1F4E78")),
            ("TEXTCOLOR",(0,0),(-1,0),colors.white),
            ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),
            ("GRID",(0,0),(-1,-1),.25,colors.HexColor("#D0D0D0")),
            ("FONTSIZE",(0,0),(-1,-1),6.1),
            ("ALIGN",(1,1),(2,-1),"CENTER"),
            ("VALIGN",(0,0),(-1,-1),"TOP"),
            ("ROWBACKGROUNDS",(0,1),(-1,-1),[colors.white, colors.HexColor("#F7F9FB")]),
        ]))
        story.append(rt)
    else:
        story.append(Paragraph("Não há O.S. atrasadas para classificar.", body))

    story.append(Paragraph("8. Concentração por cliente", h))
    client_rows = [["Cliente", "OS"]]
    for client, count in sorted(clients.items(), key=lambda x: (-x[1], x[0])):
        client_rows.append([client, count])
    cct = Table(client_rows, colWidths=[105*mm,25*mm], repeatRows=1)
    cct.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(-1,0),colors.HexColor("#1F4E78")),
        ("TEXTCOLOR",(0,0),(-1,0),colors.white),
        ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),
        ("GRID",(0,0),(-1,-1),.3,colors.HexColor("#D0D0D0")),
        ("ALIGN",(1,1),(1,-1),"CENTER"),
        ("FONTSIZE",(0,0),(-1,-1),7.2)
    ]))
    story.append(cct)

    story.append(Paragraph("9. Concentração por marca", h))
    brand_rows = [["Marca", "OS", "%"]]
    for brand, count in sorted(brands.items(), key=lambda x: (-x[1], x[0])):
        brand_rows.append([brand, count, f"{count/total:.1%}" if total else "0.0%"])
    bt = Table(brand_rows, colWidths=[60*mm,25*mm,25*mm], repeatRows=1)
    bt.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(-1,0),colors.HexColor("#1F4E78")),
        ("TEXTCOLOR",(0,0),(-1,0),colors.white),
        ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),
        ("GRID",(0,0),(-1,-1),.3,colors.HexColor("#D0D0D0")),
        ("ALIGN",(1,1),(-1,-1),"CENTER"),
        ("FONTSIZE",(0,0),(-1,-1),7.2)
    ]))
    story.append(bt)

    story.append(Paragraph(f"10. Base completa - {total} OS", title))
    base = [["OS","Cliente","Marca","Pot.","Prev. Entrega","Colaborador","Entrada","Dias na fase"]]
    for r in rows:
        entrada = "-"
        if r["Data Início"]:
            entrada = r["Data Início"].strftime("%d/%m/%Y")
            if r["Hora Início"]:
                entrada += " " + r["Hora Início"]
        dias_fase = f"{r['Dias na Fase']} dias" if r["Dias na Fase"] is not None else "-"
        base.append([
            r["OS"],
            r["Cliente"] or "N/I",
            r["Marca"] or "N/I",
            r["Potência"] or "N/I",
            fmt_date(r["Prev. Entrega"]),
            r["Colaborador"] or "Não informado",
            entrada,
            dias_fase
        ])

    btbl = Table(base, colWidths=[14*mm,39*mm,17*mm,15*mm,24*mm,34*mm,25*mm,21*mm], repeatRows=1)
    btbl.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(-1,0),colors.HexColor("#1F4E78")),
        ("TEXTCOLOR",(0,0),(-1,0),colors.white),
        ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),
        ("GRID",(0,0),(-1,-1),.25,colors.HexColor("#D0D0D0")),
        ("FONTSIZE",(0,0),(-1,-1),5.9),
        ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
    ]))
    for i, r in enumerate(rows, 1):
        fill = "#FCE4D6" if r["Status"] == "Atrasada" else "#FFF2CC" if r["Status"] == "Vence hoje" else "#FFFFFF"
        btbl.setStyle(TableStyle([("BACKGROUND",(0,i),(-1,i),colors.HexColor(fill))]))
    story.append(btbl)

    story.append(Spacer(1, 5))
    story.append(Paragraph(
        "Observação: somente nomes identificados como pessoas são mostrados no campo Colaborador. "
        "Descrições operacionais ou tipos de bomba não são tratados como nomes de colaboradores.",
        small
    ))

    doc.build(story)
    return output_path


def generate_pdf_multiphase(report, output_dir):
    """Gera um único PDF consolidando todas as fases de um mesmo arquivo."""
    if colors is None:
        raise RuntimeError("ReportLab nao esta instalado. Rode: pip install reportlab")
    os.makedirs(output_dir, exist_ok=True)

    source_name = Path(report["file"]).stem
    safe_name = re.sub(r"[^A-Za-z0-9_-]+", "_", f"Relatorio_PCP_{source_name}_Multifases").strip("_")
    output_path = os.path.join(output_dir, safe_name + ".pdf")
    rows = report["rows"]
    ref_date = report["ref_date"]
    phases = report.get("phases", [])
    counts = report["status_counts"]
    total = len(rows)

    time_values = [r["Dias na Fase"] for r in rows if r.get("Dias na Fase") is not None]
    avg_days = sum(time_values) / len(time_values) if time_values else 0
    max_days = max(time_values) if time_values else 0

    collabs = {}
    for r in rows:
        if r.get("Colaborador"):
            collabs[r["Colaborador"]] = collabs.get(r["Colaborador"], 0) + 1

    doc = SimpleDocTemplate(
        output_path, pagesize=A4,
        rightMargin=10*mm, leftMargin=10*mm, topMargin=10*mm, bottomMargin=10*mm
    )
    styles = getSampleStyleSheet()
    title = ParagraphStyle("title_multi", parent=styles["Title"], fontSize=17, leading=20,
                           textColor=colors.HexColor("#1F4E78"), spaceAfter=5)
    h = ParagraphStyle("h_multi", parent=styles["Heading2"], fontSize=10,
                       textColor=colors.HexColor("#1F4E78"), spaceBefore=7, spaceAfter=4)
    body = ParagraphStyle("body_multi", parent=styles["BodyText"], fontSize=8.1, leading=10.5)
    small = ParagraphStyle("small_multi", parent=body, fontSize=6.5, leading=8)
    phase_cell = ParagraphStyle("phase_cell_multi", parent=body, fontSize=5.0, leading=5.9,
                                textColor=colors.HexColor("#263442"), wordWrap="CJK", spaceAfter=0, spaceBefore=0)

    story = [
        Paragraph("RELATÓRIO PCP | CONSOLIDADO MULTIFASE", title),
        Paragraph(
            f"Arquivo: {source_name} | Referência: {report['referencia'] or '-'} | "
            f"Emissão: {report['emissao'].strftime('%d/%m/%Y %H:%M') if report['emissao'] else '-'} | "
            f"Fases identificadas: {len(phases)}",
            body
        ), Spacer(1, 5)
    ]

    kpi = Table([
        ["REGISTROS", "FASES", "ATRASADAS", "VENCEM HOJE", "PRÓX. 7 DIAS"],
        [total, len(phases), counts["Atrasada"], counts["Vence hoje"], counts["Próximos 7 dias"]]
    ], colWidths=[34*mm, 28*mm, 34*mm, 36*mm, 42*mm], rowHeights=[7*mm, 9*mm])
    kpi.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#D9EAF7")),
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
        ("FONTNAME", (0,1), (-1,1), "Helvetica-Bold"),
        ("FONTSIZE", (0,0), (-1,0), 6.4), ("FONTSIZE", (0,1), (-1,1), 12),
        ("ALIGN", (0,0), (-1,-1), "CENTER"), ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
        ("GRID", (0,0), (-1,-1), .3, colors.HexColor("#D0D0D0")),
        ("BACKGROUND", (2,1), (2,1), colors.HexColor("#FCE4D6")),
    ]))
    story.append(kpi)

    tempo = Table([
        ["TEMPO MÉDIO NA FASE", "MAIOR TEMPO NA FASE", "OS/REGISTROS COM INÍCIO"],
        [f"{avg_days:.1f} dias", f"{max_days} dias", str(len(time_values))]
    ], colWidths=[56*mm, 56*mm, 56*mm], rowHeights=[6.5*mm, 9*mm])
    tempo.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#EAF1F7")),
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
        ("FONTNAME", (0,1), (-1,1), "Helvetica-Bold"),
        ("FONTSIZE", (0,0), (-1,0), 6.2), ("FONTSIZE", (0,1), (-1,1), 11),
        ("ALIGN", (0,0), (-1,-1), "CENTER"), ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
        ("GRID", (0,0), (-1,-1), .3, colors.HexColor("#CDD8E2")),
    ]))
    story += [Spacer(1, 2), tempo, Spacer(1, 4)]

    story.append(Paragraph("1. Resumo por fase", h))
    phase_rows = [["Fase", "Descrição", "Registros", "Atrasadas", "Hoje", "Próx. 7d"]]
    for ph in phases:
        c = ph.get("status_counts", {})
        phase_rows.append([
            ph["fase_num"], ph["fase_nome"], len(ph["rows"]), c.get("Atrasada", 0),
            c.get("Vence hoje", 0), c.get("Próximos 7 dias", 0)
        ])
    pt = Table(phase_rows, colWidths=[15*mm, 76*mm, 20*mm, 22*mm, 18*mm, 23*mm], repeatRows=1)
    pt.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#1F4E78")),
        ("TEXTCOLOR", (0,0), (-1,0), colors.white),
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
        ("GRID", (0,0), (-1,-1), .3, colors.HexColor("#D0D0D0")),
        ("ALIGN", (0,1), (-1,-1), "CENTER"),
        ("FONTSIZE", (0,0), (-1,-1), 6.9),
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
    ]))
    story.append(pt)

    story.append(Paragraph("2. Diagnóstico executivo", h))
    if total:
        urgent = counts["Atrasada"] + counts["Vence hoje"] + counts["Próximos 7 dias"]
        story.append(Paragraph(
            f"O documento reúne {len(phases)} fases e {total} registros de OS. Em {ref_date.strftime('%d/%m/%Y')}, "
            f"{counts['Atrasada']} registros estão atrasados ({counts['Atrasada']/total:.1%}) e "
            f"{urgent} exigem acompanhamento imediato ou de curto prazo. "
            f"A leitura multifase mantém a mesma OS em fases diferentes quando ela aparece no documento, "
            f"permitindo enxergar o fluxo do processo sem apagar etapas repetidas.", body
        ))

    story.append(Paragraph("3. Colaboradores registrados", h))
    if collabs:
        crows = [["Colaborador", "Registros"]]
        for name, count in sorted(collabs.items(), key=lambda x: (-x[1], x[0])):
            crows.append([name, count])
    else:
        crows = [["Colaborador", "Registros"], ["Nenhum nome identificado", 0]]
    ct = Table(crows, colWidths=[115*mm, 25*mm], repeatRows=1)
    ct.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#1F4E78")),
        ("TEXTCOLOR", (0,0), (-1,0), colors.white), ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
        ("GRID", (0,0), (-1,-1), .3, colors.HexColor("#D0D0D0")),
        ("ALIGN", (1,1), (1,-1), "CENTER"), ("FONTSIZE", (0,0), (-1,-1), 7.2),
    ]))
    story.append(ct)

    story.append(Paragraph(f"4. Base completa | {total} registros", title))
    base = [["Fase", "OS", "Cliente", "Marca", "Pot.", "Prev.", "Colaborador", "Entrada", "Dias"]]
    ordered = sorted(rows, key=lambda r: (int(r.get("Fase Num") or 0), int(r["OS"])))
    for r in ordered:
        entrada = r["Data Início"].strftime("%d/%m/%Y") if r.get("Data Início") else "-"
        if entrada != "-" and r.get("Hora Início"):
            entrada += " " + r["Hora Início"]
        dias = str(r["Dias na Fase"]) if r.get("Dias na Fase") is not None else "-"
        base.append([
            r.get("Fase Num", "-"), r["OS"], r["Cliente"] or "N/I", r["Marca"] or "N/I",
            r["Potência"] or "N/I", fmt_date(r["Prev. Entrega"]), r["Colaborador"] or "Não informado",
            entrada, dias
        ])
    bt = Table(base, colWidths=[10*mm, 13*mm, 34*mm, 15*mm, 13*mm, 21*mm, 31*mm, 25*mm, 15*mm], repeatRows=1)
    bt.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#1F4E78")),
        ("TEXTCOLOR", (0,0), (-1,0), colors.white), ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
        ("GRID", (0,0), (-1,-1), .25, colors.HexColor("#D0D0D0")),
        ("FONTSIZE", (0,0), (-1,-1), 5.7), ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
        ("ROWBACKGROUNDS", (0,1), (-1,-1), [colors.white, colors.HexColor("#F6F8FA")]),
    ]))
    for i, r in enumerate(ordered, 1):
        if r.get("Status") == "Atrasada":
            bt.setStyle(TableStyle([("BACKGROUND", (0,i), (-1,i), colors.HexColor("#FCE4E4"))]))
        elif r.get("Status") == "Vence hoje":
            bt.setStyle(TableStyle([("BACKGROUND", (0,i), (-1,i), colors.HexColor("#FFF2CC"))]))
    story.append(bt)

    story.append(Spacer(1, 5))
    story.append(Paragraph(
        "Observação: o mesmo número de OS pode aparecer em mais de uma fase e será mantido em cada ocorrência, "
        "pois cada linha representa uma etapa do processo. Descrições operacionais como BOMBA ABS, BOMBA FLYGT, "
        "BOMBAS KSB e MOTOR//BOMBA C/ não são tratadas como nomes de colaboradores.", small
    ))
    doc.build(story)
    return output_path


def _safe_para(value):
    """Escapa texto para uso seguro em Paragraph do ReportLab."""
    import html
    return html.escape(ptxt(value))


def _cell_para(value, style=None):
    """Cria uma célula de tabela que quebra linha sem depender de estilos locais."""
    from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
    if style is None:
        base = getSampleStyleSheet()["BodyText"]
        st = ParagraphStyle(
            "CellWrap", parent=base, fontName="Helvetica", fontSize=6.2,
            leading=7.2, wordWrap="CJK", spaceAfter=0, spaceBefore=0
        )
    else:
        st = style
    return Paragraph(_safe_para(value), st)

def _phase_name_cell(row, style=None):
    """Mostra o nome da etapa em célula com quebra automática."""
    return _cell_para(row.get("Fase Nome") or "Não informado", style)


def _pct(n, total):
    return f"{(n / total):.1%}" if total else "0.0%"


def _fmt_dt(row):
    if not row.get("Data Início"):
        return "Não informado"
    txt = row["Data Início"].strftime("%d/%m/%Y")
    if row.get("Hora Início"):
        txt += " " + row["Hora Início"]
    return txt


def _days_late(row, ref_date):
    prev = row.get("Prev. Entrega")
    if not prev:
        return None
    return max(0, (ref_date - prev).days)


def _criticality(row, ref_date):
    """Calcula uma criticidade operacional simples e explicável para o PCP."""
    late = _days_late(row, ref_date) or 0
    prev = row.get("Prev. Entrega")
    days_phase = row.get("Dias na Fase") or 0
    score = 0
    if late > 0:
        score += min(65, 45 + min(late, 20))
    elif prev == ref_date:
        score += 42
    elif prev:
        days_to_due = (prev - ref_date).days
        score += 34 if days_to_due <= 3 else 24 if days_to_due <= 7 else 8
    else:
        score += 28
    if days_phase >= 30:
        score += 22
    elif days_phase >= 15:
        score += 14
    elif days_phase >= 8:
        score += 7
    elif row.get("Data Início") is None:
        score += 4
    score = min(100, score)
    if score >= 70:
        label = "Crítica"
    elif score >= 50:
        label = "Alta"
    elif score >= 30:
        label = "Média"
    else:
        label = "Baixa"
    return score, label


def generate_pdf_enhanced(report, output_dir):
    """Gera um relatorio PCP mais completo, mantendo analise por fase e consolidado."""
    if colors is None:
        raise RuntimeError("ReportLab nao esta instalado. Rode: pip install reportlab")

    os.makedirs(output_dir, exist_ok=True)
    source_name = Path(report["file"]).stem
    phases = report.get("phases", [])
    rows = report.get("rows", [])
    ref_date = report.get("ref_date") or (report.get("emissao").date() if report.get("emissao") else TODAY)

    safe_name = re.sub(
        r"[^A-Za-z0-9_-]+", "_",
        f"Relatorio_PCP_{source_name}_Analise_Completa"
    ).strip("_")
    output_path = os.path.join(output_dir, safe_name + ".pdf")

    status_order = ["Atrasada", "Vence hoje", "Próximos 7 dias", "Prazo futuro", "Sem previsão"]
    status_counts = {s: sum(1 for r in rows if r.get("Status") == s) for s in status_order}
    total = len(rows)

    # --- indicadores derivados ---
    time_values = [r.get("Dias na Fase") for r in rows if r.get("Dias na Fase") is not None]
    avg_days = sum(time_values) / len(time_values) if time_values else 0
    max_days = max(time_values) if time_values else 0

    clients = {}
    brands = {}
    collabs = {}
    phases_by_os = {}
    for r in rows:
        client = r.get("Cliente") or "Não informado"
        brand = r.get("Marca") or "Não informado"
        clients[client] = clients.get(client, 0) + 1
        brands[brand] = brands.get(brand, 0) + 1
        if r.get("Colaborador"):
            collabs[r["Colaborador"]] = collabs.get(r["Colaborador"], 0) + 1
        phases_by_os.setdefault(r["OS"], []).append(r.get("Fase Num", "-"))

    os_multi_phase = {os_: fases for os_, fases in phases_by_os.items() if len(fases) > 1}
    with_start = len(time_values)
    unplanned = sum(1 for r in rows if not r.get("Prev. Entrega"))
    due_or_urgent = status_counts["Atrasada"] + status_counts["Vence hoje"] + status_counts["Próximos 7 dias"]
    until_3 = sum(1 for r in rows if r.get("Status") == "Próximos 7 dias" and r.get("Prev. Entrega") and 0 < (r["Prev. Entrega"] - ref_date).days <= 3)
    days_4_7 = sum(1 for r in rows if r.get("Status") == "Próximos 7 dias" and r.get("Prev. Entrega") and 4 <= (r["Prev. Entrega"] - ref_date).days <= 7)
    critical_counts = {lab: sum(1 for r in rows if r.get("Criticidade") == lab) for lab in ("Crítica", "Alta", "Média", "Baixa")}

    # Prioridade: atraso primeiro, depois prazo, depois tempo na fase.
    priority_rows = sorted(
        rows,
        key=lambda r: (
            0 if r.get("Status") == "Atrasada" else 1 if r.get("Status") == "Vence hoje" else 2 if r.get("Status") == "Próximos 7 dias" else 3,
            -(_days_late(r, ref_date) or 0),
            -(r.get("Dias na Fase") or 0)
        )
    )
    overdue = [r for r in rows if r.get("Status") == "Atrasada"]
    aging_sorted = sorted(rows, key=lambda r: -(r.get("Dias na Fase") or -1))

    # --- documento ---
    phase_header_parts = [
        f"{ph.get('fase_num', '-')} - {ph.get('fase_nome', '-')}"
        for ph in phases
    ]
    phase_header_text = " | ".join(phase_header_parts) if phase_header_parts else "-"

    def _wrap_header(text, max_chars=105):
        # Quebra o cabeçalho das fases em poucas linhas para caber em A4.
        parts = text.split(" | ")
        lines = []
        current = ""
        for part in parts:
            candidate = part if not current else current + " | " + part
            if current and len(candidate) > max_chars:
                lines.append(current)
                current = part
            else:
                current = candidate
        if current:
            lines.append(current)
        return lines[:2] or ["-"]

    phase_header_lines = _wrap_header(phase_header_text)

    def draw_page(canvas, doc):
        canvas.saveState()
        w, h = A4

        # Cabeçalho repetido em todas as páginas, deixando claro qual(is) fase(s)
        # estão sendo analisadas.
        canvas.setFillColor(colors.HexColor("#6B7785"))
        canvas.setFont("Helvetica-Bold", 5.8)
        canvas.drawString(10*mm, h-5.5*mm, "FASE(S):")
        canvas.setFont("Helvetica", 5.8)
        y = h-5.5*mm
        x = 22*mm
        for i, line in enumerate(phase_header_lines):
            if i == 0:
                canvas.drawString(x, y, line[:150])
            else:
                canvas.drawString(22*mm, h-9*mm, line[:150])
        canvas.setStrokeColor(colors.HexColor("#D9E2EC"))
        canvas.line(10*mm, h-11*mm, w-10*mm, h-11*mm)

        canvas.line(10*mm, 9*mm, w-10*mm, 9*mm)
        canvas.setFillColor(colors.HexColor("#6B7785"))
        canvas.setFont("Helvetica", 6.5)
        canvas.drawString(10*mm, 5.5*mm, f"Fonte: {source_name} | Referência {report.get('referencia') or '-'}")
        canvas.drawRightString(w-10*mm, 5.5*mm, f"Página {doc.page}")
        canvas.restoreState()

    doc = SimpleDocTemplate(
        output_path,
        pagesize=A4,
        rightMargin=10*mm,
        leftMargin=10*mm,
        topMargin=10*mm,
        bottomMargin=14*mm,
        title=f"Relatório PCP - {source_name}",
        author="Gerador de Relatórios PCP - ESA",
    )

    styles = getSampleStyleSheet()
    title = ParagraphStyle(
        "title_enh", parent=styles["Title"], fontSize=18, leading=21,
        textColor=colors.HexColor("#173A5E"), spaceAfter=4
    )
    subtitle = ParagraphStyle(
        "sub_enh", parent=styles["BodyText"], fontSize=8.2, leading=10.2,
        textColor=colors.HexColor("#5E6B78"), spaceAfter=6
    )
    h1 = ParagraphStyle(
        "h1_enh", parent=styles["Heading2"], fontSize=11, leading=13,
        textColor=colors.HexColor("#173A5E"), spaceBefore=8, spaceAfter=4
    )
    h2 = ParagraphStyle(
        "h2_enh", parent=h1, fontSize=9.5, leading=11.5, spaceBefore=6
    )
    body = ParagraphStyle(
        "body_enh", parent=styles["BodyText"], fontSize=7.8, leading=10.2,
        textColor=colors.HexColor("#263442")
    )
    small = ParagraphStyle(
        "small_enh", parent=body, fontSize=6.1, leading=7.6,
        textColor=colors.HexColor("#5E6B78")
    )
    tiny = ParagraphStyle(
        "tiny_enh", parent=small, fontSize=5.7, leading=6.8
    )
    phase_cell = ParagraphStyle(
        "phase_cell_enh", parent=body, fontSize=5.0, leading=5.9,
        textColor=colors.HexColor("#263442"), wordWrap="CJK", spaceAfter=0, spaceBefore=0
    )
    obs_cell = ParagraphStyle(
        "obs_cell_enh", parent=body, fontSize=5.2, leading=7.0,
        textColor=colors.HexColor("#263442"), wordWrap="CJK", spaceAfter=0, spaceBefore=0
    )

    story = []
    if report.get("multi_fase"):
        phase_nums = [str(ph.get("fase_num", "-")).strip() for ph in phases if ph.get("fase_num")]
        mode_title = "ANÁLISE PCP | CONSOLIDADO | FASES " + (", ".join(phase_nums) if phase_nums else "MULTIFASE")
    else:
        fase_num = report.get("fase_num", "-")
        fase_nome = report.get("fase_nome") or (phases[0].get("fase_nome") if phases else "-")
        mode_title = f"ANÁLISE PCP | FASE {fase_num} - {_safe_para(fase_nome)}"
    story.append(Paragraph(mode_title, title))

    # Identificação explícita da fase no cabeçalho do relatório.
    if report.get("multi_fase"):
        phase_label = "<b>FASES ANALISADAS:</b> " + " | ".join(
            f"{ph.get('fase_num', '-')} - {_safe_para(ph.get('fase_nome', '-'))}" for ph in phases
        )
    else:
        phase_label = (
            f"<b>FASE {report.get('fase_num', '-')}:</b> "
            f"{_safe_para(report.get('fase_nome') or (phases[0].get('fase_nome') if phases else '-'))}"
        )
    story.append(Paragraph(phase_label, subtitle))
    story.append(Paragraph(
        f"Arquivo: {_safe_para(source_name)} | Fases identificadas: {len(phases)} | "
        f"Registros: {total} | Emissão: {report.get('emissao').strftime('%d/%m/%Y %H:%M') if report.get('emissao') else '-'} | "
        f"Data-base da análise: {ref_date.strftime('%d/%m/%Y')}", subtitle
    ))

    # 1. Visao executiva
    story.append(Paragraph("1. Visão executiva", h1))
    kpi_data = [
        ["REGISTROS", "FASES", "ATRASADAS", "VENCEM HOJE", "ATÉ 3 DIAS", "4-7 DIAS", "SEM PREVISÃO"],
        [total, len(phases), status_counts["Atrasada"], status_counts["Vence hoje"], until_3, days_4_7, status_counts["Sem previsão"]]
    ]
    k = Table(kpi_data, colWidths=[27*mm, 19*mm, 25*mm, 27*mm, 25*mm, 25*mm, 28*mm], rowHeights=[6.2*mm, 8.5*mm])
    k.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#EAF1F7")),
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
        ("FONTNAME", (0,1), (-1,1), "Helvetica-Bold"),
        ("FONTSIZE", (0,0), (-1,0), 5.7),
        ("FONTSIZE", (0,1), (-1,1), 11.5),
        ("ALIGN", (0,0), (-1,-1), "CENTER"),
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
        ("GRID", (0,0), (-1,-1), .3, colors.HexColor("#CBD5DF")),
        ("BACKGROUND", (2,1), (2,1), colors.HexColor("#FCE4E4")),
        ("BACKGROUND", (3,1), (5,1), colors.HexColor("#FFF2CC")),
    ]))
    story.append(k)
    story.append(Spacer(1, 3))

    k2 = Table([
        ["TEMPO MÉDIO NA FASE", "MAIOR TEMPO", "COM INÍCIO", "OS EM MAIS DE UMA FASE", "URGENTES"],
        [f"{avg_days:.1f} dias", f"{max_days} dias", str(with_start), str(len(os_multi_phase)), str(due_or_urgent)]
    ], colWidths=[35*mm, 28*mm, 27*mm, 40*mm, 31*mm], rowHeights=[6.2*mm, 8.5*mm])
    k2.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#F4F7FA")),
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
        ("FONTNAME", (0,1), (-1,1), "Helvetica-Bold"),
        ("FONTSIZE", (0,0), (-1,0), 5.7),
        ("FONTSIZE", (0,1), (-1,1), 10.5),
        ("ALIGN", (0,0), (-1,-1), "CENTER"),
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
        ("GRID", (0,0), (-1,-1), .3, colors.HexColor("#D3DCE5")),
    ]))
    story.append(k2)

    story.append(Paragraph("Leitura do PCP", h2))
    if total:
        diagnosis = (
            f"A carteira analisada possui {total} registros distribuídos em {len(phases)} fase(s). "
            f"Há {status_counts['Atrasada']} atrasados ({_pct(status_counts['Atrasada'], total)}), "
            f"{status_counts['Vence hoje']} com vencimento hoje e {status_counts['Próximos 7 dias']} previstos para os próximos 7 dias. "
            f"Somando esses grupos, {due_or_urgent} registros exigem ação de curto prazo. "
            f"{with_start} registros possuem início de fase e o tempo médio calculado a partir desses registros é de {avg_days:.1f} dias. "
        )
        if len(os_multi_phase):
            diagnosis += f"Há {len(os_multi_phase)} OS repetidas em duas ou mais fases, o que permite acompanhar o avanço delas pelo fluxo. "
        if unplanned:
            diagnosis += f"Também foram encontrados {unplanned} registros sem previsão de entrega; eles merecem validação de prazo."
        else:
            diagnosis += "Todos os registros possuem previsão de entrega no relatório de origem."
        story.append(Paragraph(diagnosis, body))

    story.append(Paragraph("2. Criticidade e atenção imediata", h1))
    crit_table = Table([
        ["Crítica", "Alta", "Média", "Baixa"],
        [critical_counts["Crítica"], critical_counts["Alta"], critical_counts["Média"], critical_counts["Baixa"]],
    ], colWidths=[39*mm, 39*mm, 39*mm, 39*mm], rowHeights=[6.2*mm, 8.5*mm])
    crit_table.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#F4F7FA")),
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
        ("FONTNAME", (0,1), (-1,1), "Helvetica-Bold"),
        ("ALIGN", (0,0), (-1,-1), "CENTER"),
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
        ("GRID", (0,0), (-1,-1), .3, colors.HexColor("#D3DCE5")),
        ("BACKGROUND", (0,1), (0,1), colors.HexColor("#FCE4E4")),
        ("BACKGROUND", (1,1), (1,1), colors.HexColor("#FFF0D9")),
    ]))
    story.append(crit_table)
    immediate = sorted(rows, key=lambda r: (-int(r.get("Criticidade Score", 0)), -(_days_late(r, ref_date) or 0), -(r.get("Dias na Fase") or 0)))[:8]
    immediate_rows = [["#", "Fase", "Etapa", "OS", "Cliente", "Criticidade", "Atraso", "Dias fase"]]
    for i, r in enumerate(immediate, 1):
        late = _days_late(r, ref_date)
        immediate_rows.append([i, r.get("Fase Num", "-"), _phase_name_cell(r, phase_cell), r.get("OS", "-"), _cell_para(r.get("Cliente") or "N/I"), r.get("Criticidade", "-"), str(late) if late is not None else "-", str(r.get("Dias na Fase")) if r.get("Dias na Fase") is not None else "-"])
    it = Table(immediate_rows, colWidths=[7*mm, 9*mm, 35*mm, 12*mm, 43*mm, 26*mm, 24*mm, 25*mm], repeatRows=1)
    it.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#173A5E")),
        ("TEXTCOLOR", (0,0), (-1,0), colors.white),
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
        ("GRID", (0,0), (-1,-1), .22, colors.HexColor("#D2D9E0")),
        ("FONTSIZE", (0,0), (-1,-1), 6.0),
        ("ALIGN", (0,1), (1,-1), "CENTER"),
        ("ALIGN", (3,1), (3,-1), "CENTER"),
        ("ALIGN", (5,1), (-1,-1), "CENTER"),
    ]))
    for i, r in enumerate(immediate, 1):
        fill = "#FCE4E4" if r.get("Criticidade") == "Crítica" else "#FFF0D9" if r.get("Criticidade") == "Alta" else "#FFF8E7" if r.get("Criticidade") == "Média" else "#FFFFFF"
        it.setStyle(TableStyle([("BACKGROUND", (0,i), (-1,i), colors.HexColor(fill))]))
    story.append(it)

    # 3. Analise por fase
    story.append(Paragraph("3. Análise individual por fase", h1))
    phase_rows = [["Fase", "Descrição", "OS", "Atras.", "Críticas", "Hoje", "Próx. 7d", "Futuro", "Média dias"]]
    for ph in phases:
        pr = ph.get("rows", [])
        c = ph.get("status_counts", {})
        vals = [r.get("Dias na Fase") for r in pr if r.get("Dias na Fase") is not None]
        avg = sum(vals)/len(vals) if vals else 0
        phase_rows.append([
            ph.get("fase_num", "-"), _safe_para(ph.get("fase_nome", "-")), len(pr),
            c.get("Atrasada", 0), sum(1 for r in pr if r.get("Criticidade") == "Crítica"), c.get("Vence hoje", 0), c.get("Próximos 7 dias", 0),
            c.get("Prazo futuro", 0), f"{avg:.1f}"
        ])
    pt = Table(phase_rows, colWidths=[12*mm, 63*mm, 12*mm, 13*mm, 15*mm, 13*mm, 17*mm, 15*mm, 20*mm], repeatRows=1)
    pt.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#173A5E")),
        ("TEXTCOLOR", (0,0), (-1,0), colors.white),
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
        ("GRID", (0,0), (-1,-1), .25, colors.HexColor("#D1D9E1")),
        ("FONTSIZE", (0,0), (-1,-1), 6.2),
        ("ALIGN", (2,1), (-1,-1), "CENTER"),
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
        ("ROWBACKGROUNDS", (0,1), (-1,-1), [colors.white, colors.HexColor("#F7F9FB")]),
    ]))
    story.append(pt)

    # 3. Prioridades
    # As larguras das tabelas abaixo são mantidas dentro da área útil da página para evitar overflow.
    story.append(Paragraph("4. Fila de prioridades do PCP", h1))
    story.append(Paragraph(
        "Ordenação sugerida para a atuação: primeiro atrasos, depois vencimentos imediatos e, por fim, maior tempo acumulado na fase. "
        "A lista abaixo serve como fila operacional e não substitui decisões específicas de produção, material ou qualidade.", body
    ))
    prio = [["#", "Etapa", "OS", "Cliente", "Prev.", "Status", "Criticidade", "Dias atraso", "Dias na fase", "Colaborador", "OBS"]]
    for i, r in enumerate(priority_rows[:20], 1):
        late = _days_late(r, ref_date)
        prio.append([
            i, _phase_name_cell(r, phase_cell), r.get("OS", "-"), _cell_para(r.get("Cliente") or "N/I"),
            fmt_date(r.get("Prev. Entrega")), r.get("Status", "-"), r.get("Criticidade", "-"),
            str(late) if late is not None else "-", str(r.get("Dias na Fase")) if r.get("Dias na Fase") is not None else "-",
            _cell_para(r.get("Colaborador") or "Não informado"), Paragraph("<br/>&nbsp;", obs_cell)
        ])
    prt = Table(prio, colWidths=[6*mm, 24*mm, 11*mm, 30*mm, 18*mm, 19*mm, 18*mm, 16*mm, 16*mm, 18*mm, 14*mm], repeatRows=1)
    prt.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#173A5E")),
        ("TEXTCOLOR", (0,0), (-1,0), colors.white),
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
        ("GRID", (0,0), (-1,-1), .22, colors.HexColor("#D2D9E0")),
        ("FONTSIZE", (0,0), (-1,-1), 5.9),
        ("ALIGN", (0,1), (1,-1), "CENTER"),
        ("ALIGN", (3,1), (3,-1), "CENTER"),
        ("ALIGN", (5,1), (-1,-1), "CENTER"),
    ]))
    for i, r in enumerate(priority_rows[:20], 1):
        prt.setStyle(TableStyle([
            ("BACKGROUND", (0,i), (-1,i), _status_fill(r.get("Status"))),
            ("BACKGROUND", (-1,i), (-1,i), colors.white),
        ]))
    story.append(prt)

    # 4. Aging
    story.append(Paragraph("5. Aging: tempo das OS na fase", h1))
    aging_buckets = [("0 a 3 dias", lambda d: d is not None and d <= 3),
                     ("4 a 7 dias", lambda d: d is not None and 4 <= d <= 7),
                     ("8 a 14 dias", lambda d: d is not None and 8 <= d <= 14),
                     ("15+ dias", lambda d: d is not None and d >= 15),
                     ("Sem início", lambda d: d is None)]
    aging_table = [["Faixa", "Registros", "%", "Leitura PCP"]]
    for label, fn in aging_buckets:
        n = sum(1 for r in rows if fn(r.get("Dias na Fase")))
        leitura = "Normal" if label == "0 a 3 dias" else "Atenção" if label == "4 a 7 dias" else "Crítico" if label in ("8 a 14 dias", "15+ dias") else "Validar informação"
        aging_table.append([label, n, _pct(n, total), leitura])
    at = Table(aging_table, colWidths=[34*mm, 25*mm, 25*mm, 77*mm], repeatRows=1)
    at.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#173A5E")),
        ("TEXTCOLOR", (0,0), (-1,0), colors.white),
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
        ("GRID", (0,0), (-1,-1), .25, colors.HexColor("#D0D7DF")),
        ("ALIGN", (1,1), (2,-1), "CENTER"),
        ("FONTSIZE", (0,0), (-1,-1), 6.6),
        ("ROWBACKGROUNDS", (0,1), (-1,-1), [colors.white, colors.HexColor("#F7F9FB")]),
    ]))
    story.append(at)
    top_aging = aging_sorted[:15]
    story.append(Paragraph("15 maiores tempos na fase", h2))
    aging_rows = [["Etapa", "OS", "Cliente", "Início", "Dias", "Status", "Colaborador", "OBS"]]
    for r in top_aging:
        aging_rows.append([
            _phase_name_cell(r, phase_cell), r.get("OS", "-"), _cell_para(r.get("Cliente") or "N/I"),
            _fmt_dt(r), str(r.get("Dias na Fase")) if r.get("Dias na Fase") is not None else "-", r.get("Status", "-"),
            _cell_para(r.get("Colaborador") or "Não informado"), Paragraph("<br/>&nbsp;", obs_cell)
        ])
    agt = Table(aging_rows, colWidths=[28*mm, 11*mm, 39*mm, 24*mm, 14*mm, 24*mm, 20*mm, 24*mm], repeatRows=1)
    agt.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#2A577D")),
        ("TEXTCOLOR", (0,0), (-1,0), colors.white),
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
        ("GRID", (0,0), (-1,-1), .2, colors.HexColor("#D3D9E0")),
        ("FONTSIZE", (0,0), (-1,-1), 6.0),
        ("ALIGN", (1,1), (1,-1), "CENTER"),
        ("ALIGN", (3,1), (5,-1), "CENTER"),
    ]))
    story.append(agt)

    # 5. Entregas
    story.append(Paragraph("6. Calendário de vencimentos", h1))
    due_buckets = [
        ("Atrasadas", [r for r in rows if r.get("Status") == "Atrasada"]),
        ("Hoje", [r for r in rows if r.get("Status") == "Vence hoje"]),
        ("Próximos 7 dias", [r for r in rows if r.get("Status") == "Próximos 7 dias"]),
        ("Mais de 7 dias", [r for r in rows if r.get("Status") == "Prazo futuro"]),
        ("Sem previsão", [r for r in rows if r.get("Status") == "Sem previsão"]),
    ]
    due = [["Janela", "Qtd.", "%", "Primeira observação"]]
    for label, subset in due_buckets:
        first = "-"
        if subset:
            ordered = sorted(subset, key=lambda r: r.get("Prev. Entrega") or date.max)
            first = f"OS {ordered[0]['OS']} | {fmt_date(ordered[0].get('Prev. Entrega'))}"
        due.append([label, len(subset), _pct(len(subset), total), first])
    dt = Table(due, colWidths=[42*mm, 20*mm, 24*mm, 74*mm], repeatRows=1)
    dt.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#173A5E")),
        ("TEXTCOLOR", (0,0), (-1,0), colors.white),
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
        ("GRID", (0,0), (-1,-1), .25, colors.HexColor("#D0D7DF")),
        ("ALIGN", (1,1), (2,-1), "CENTER"),
        ("FONTSIZE", (0,0), (-1,-1), 6.5),
    ]))
    story.append(dt)

    # 6. Colaboradores
    story.append(Paragraph("7. Distribuição por colaborador", h1))
    if collabs:
        collab_rows = [["Colaborador", "Registros", "%", "Atrasadas"]]
        for name, count in sorted(collabs.items(), key=lambda x: (-x[1], x[0])):
            late_count = sum(1 for r in rows if r.get("Colaborador") == name and r.get("Status") == "Atrasada")
            collab_rows.append([_safe_para(name), count, _pct(count, total), late_count])
        missing = sum(1 for r in rows if not r.get("Colaborador"))
        if missing:
            collab_rows.append(["Sem colaborador identificado", missing, _pct(missing, total), sum(1 for r in rows if not r.get("Colaborador") and r.get("Status") == "Atrasada")])
    else:
        collab_rows = [["Colaborador", "Registros", "%", "Atrasadas"], ["Nenhum nome identificado", 0, "0.0%", 0]]
    ct = Table(collab_rows, colWidths=[75*mm, 27*mm, 25*mm, 27*mm], repeatRows=1)
    ct.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#173A5E")),
        ("TEXTCOLOR", (0,0), (-1,0), colors.white),
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
        ("GRID", (0,0), (-1,-1), .25, colors.HexColor("#D0D7DF")),
        ("ALIGN", (1,1), (-1,-1), "CENTER"),
        ("FONTSIZE", (0,0), (-1,-1), 6.5),
    ]))
    story.append(ct)

    # 7. Clientes e marcas
    story.append(Paragraph("8. Perfil da carteira: clientes e marcas", h1))
    left = [["Top clientes", "OS"]]
    for name, count in sorted(clients.items(), key=lambda x: (-x[1], x[0]))[:12]:
        left.append([_safe_para(name), count])
    right = [["Marcas", "OS", "%"]]
    for name, count in sorted(brands.items(), key=lambda x: (-x[1], x[0])):
        right.append([_safe_para(name), count, _pct(count, total)])
    left_t = Table(left, colWidths=[58*mm, 16*mm], repeatRows=1)
    left_t.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#2A577D")),
        ("TEXTCOLOR", (0,0), (-1,0), colors.white), ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
        ("GRID", (0,0), (-1,-1), .2, colors.HexColor("#D3D9E0")), ("ALIGN", (1,1), (1,-1), "CENTER"),
        ("FONTSIZE", (0,0), (-1,-1), 6.0),
    ]))
    right_t = Table(right, colWidths=[42*mm, 18*mm, 20*mm], repeatRows=1)
    right_t.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#2A577D")),
        ("TEXTCOLOR", (0,0), (-1,0), colors.white), ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
        ("GRID", (0,0), (-1,-1), .2, colors.HexColor("#D3D9E0")), ("ALIGN", (1,1), (-1,-1), "CENTER"),
        ("FONTSIZE", (0,0), (-1,-1), 6.0),
    ]))
    paired = Table([[left_t, right_t]], colWidths=[78*mm, 80*mm])
    paired.setStyle(TableStyle([("VALIGN", (0,0), (-1,-1), "TOP")]))
    story.append(paired)

    # 8. OS em multiplas fases
    story.append(Paragraph("9. OS presentes em mais de uma fase", h1))
    if os_multi_phase:
        multi_rows = [["OS", "Fases encontradas", "Qtd. fases", "Leitura"]]
        for os_, fase_list in sorted(os_multi_phase.items(), key=lambda x: (-len(x[1]), x[0])):
            uniq = []
            for f in fase_list:
                if f not in uniq:
                    uniq.append(f)
            multi_rows.append([
                os_, ", ".join(uniq), len(uniq),
                "Acompanhar movimentação entre etapas" if len(uniq) > 1 else "-"
            ])
        mt = Table(multi_rows, colWidths=[24*mm, 45*mm, 25*mm, 64*mm], repeatRows=1)
        mt.setStyle(TableStyle([
            ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#173A5E")),
            ("TEXTCOLOR", (0,0), (-1,0), colors.white), ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
            ("GRID", (0,0), (-1,-1), .25, colors.HexColor("#D0D7DF")),
            ("ALIGN", (0,1), (2,-1), "CENTER"), ("FONTSIZE", (0,0), (-1,-1), 6.4),
        ]))
        story.append(mt)
    else:
        story.append(Paragraph("Nenhuma OS repetida em mais de uma fase foi encontrada neste documento.", body))

    # 9. Detalhamento de cada fase
    story.append(Paragraph("10. Detalhamento operacional por fase", h1))
    for ph in phases:
        ph_rows = ph.get("rows", [])
        if not ph_rows:
            continue
        ph_status = ph.get("status_counts", {})
        ph_times = [r.get("Dias na Fase") for r in ph_rows if r.get("Dias na Fase") is not None]
        ph_avg = sum(ph_times)/len(ph_times) if ph_times else 0
        ph_max = max(ph_times) if ph_times else 0
        story.append(Paragraph(
            f"Fase { _safe_para(ph.get('fase_num')) } - {_safe_para(ph.get('fase_nome'))}", h2
        ))
        story.append(Paragraph(
            f"Registros: {len(ph_rows)} | Atrasadas: {ph_status.get('Atrasada',0)} | "
            f"Hoje: {ph_status.get('Vence hoje',0)} | Próx. 7d: {ph_status.get('Próximos 7 dias',0)} | "
            f"Tempo médio: {ph_avg:.1f} dias | Maior tempo: {ph_max} dias.", body
        ))
        # Top 8 criticas por fase
        ph_sorted = sorted(ph_rows, key=lambda r: (0 if r.get("Status") == "Atrasada" else 1, -(_days_late(r, ref_date) or 0), -(r.get("Dias na Fase") or 0)))[:8]
        pht = [["OS", "Cliente", "Marca", "Prev.", "Status", "Criticidade", "Dias atraso", "Dias fase", "Colaborador", "OBS"]]
        for r in ph_sorted:
            pht.append([
                r.get("OS", "-"), _safe_para(r.get("Cliente") or "N/I"), _safe_para(r.get("Marca") or "N/I"),
                fmt_date(r.get("Prev. Entrega")), r.get("Status", "-"), r.get("Criticidade", "-"),
                str(_days_late(r, ref_date)) if _days_late(r, ref_date) is not None else "-",
                str(r.get("Dias na Fase")) if r.get("Dias na Fase") is not None else "-",
                _safe_para(r.get("Colaborador") or "Não informado"), Paragraph("<br/>&nbsp;", obs_cell)
            ])
        phtbl = Table(pht, colWidths=[11*mm, 31*mm, 13*mm, 17*mm, 20*mm, 20*mm, 15*mm, 15*mm, 19*mm, 19*mm], repeatRows=1)
        phtbl.setStyle(TableStyle([
            ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#345F7D")),
            ("TEXTCOLOR", (0,0), (-1,0), colors.white), ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
            ("GRID", (0,0), (-1,-1), .2, colors.HexColor("#D4DBE2")),
            ("FONTSIZE", (0,0), (-1,-1), 5.6), ("ALIGN", (0,1), (0,-1), "CENTER"),
            ("ALIGN", (3,1), (6,-1), "CENTER"),
        ]))
        for i, r in enumerate(ph_sorted, 1):
            if r.get("Status") == "Atrasada":
                phtbl.setStyle(TableStyle([("BACKGROUND", (0,i), (-1,i), colors.HexColor("#FCE4E4"))]))
        story.append(phtbl)

    # 10. Alertas de qualidade dos dados
    story.append(Paragraph("11. Alertas e qualidade dos dados", h1))
    quality = []
    no_collab = sum(1 for r in rows if not r.get("Colaborador"))
    no_start = sum(1 for r in rows if not r.get("Data Início"))
    no_power = sum(1 for r in rows if not r.get("Potência"))
    no_brand = sum(1 for r in rows if not r.get("Marca"))
    if unplanned: quality.append(f"{unplanned} registros sem previsão de entrega.")
    if no_collab: quality.append(f"{no_collab} registros sem colaborador identificado como pessoa.")
    if no_start: quality.append(f"{no_start} registros sem data de início de fase utilizável.")
    if no_power: quality.append(f"{no_power} registros sem potência informada.")
    if no_brand: quality.append(f"{no_brand} registros sem marca identificada.")
    if not quality:
        quality.append("Nenhum alerta básico de qualidade foi detectado nos campos analisados.")
    story.append(Paragraph(" ".join(quality), body))
    story.append(Paragraph(
        "Observação sobre colaboradores: descrições operacionais como BOMBA ABS, BOMBA FLYGT, BOMBAS KSB e MOTOR//BOMBA C/ não são tratadas como nomes de colaboradores.", small
    ))

    # 11. Base completa
    story.append(Paragraph(f"12. Base completa | {total} registros", h1))
    full = [["Etapa", "OS", "Cliente", "Marca", "Pot.", "Prev.", "Status", "Colab.", "Entrada", "Dias", "OBS"]]
    ordered = sorted(rows, key=lambda r: (int(r.get("Fase Num") or 0), int(r.get("OS") or 0)))
    for r in ordered:
        full.append([
            _phase_name_cell(r, phase_cell), r.get("OS", "-"), _cell_para(r.get("Cliente") or "N/I"),
            _cell_para(r.get("Marca") or "N/I"), _cell_para(r.get("Potência") or "N/I"),
            fmt_date(r.get("Prev. Entrega")), r.get("Status", "-"),
            _cell_para(r.get("Colaborador") or "Não informado"), _fmt_dt(r),
            str(r.get("Dias na Fase")) if r.get("Dias na Fase") is not None else "-",
            Paragraph("<br/>&nbsp;", obs_cell)
        ])
    ft = Table(full, colWidths=[22*mm, 10*mm, 23*mm, 10*mm, 9*mm, 17*mm, 19*mm, 19*mm, 22*mm, 10*mm, 19*mm], repeatRows=1)
    ft.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#173A5E")),
        ("TEXTCOLOR", (0,0), (-1,0), colors.white), ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
        ("GRID", (0,0), (-1,-1), .2, colors.HexColor("#D2D9E0")), ("FONTSIZE", (0,0), (-1,-1), 5.1),
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
        ("ALIGN", (1,1), (1,-1), "CENTER"), ("ALIGN", (4,1), (6,-1), "CENTER"), ("ALIGN", (9,1), (9,-1), "CENTER"),
    ]))
    for i, r in enumerate(ordered, 1):
        ft.setStyle(TableStyle([
            ("BACKGROUND", (0,i), (-1,i), _status_fill(r.get("Status"))),
            ("BACKGROUND", (-1,i), (-1,i), colors.white),
        ]))
    story.append(ft)

    # 12. Recomendacoes
    story.append(Paragraph("13. Recomendações de atuação do PCP", h1))
    recs = []
    if status_counts["Atrasada"]:
        recs.append(f"Atacar primeiro as {status_counts['Atrasada']} OS atrasadas, registrando causa do atraso e nova previsão realista.")
    if status_counts["Vence hoje"]:
        recs.append(f"Confirmar hoje a execução das {status_counts['Vence hoje']} OS com vencimento imediato.")
    if status_counts["Próximos 7 dias"]:
        recs.append(f"Reservar capacidade para as {status_counts['Próximos 7 dias']} OS dos próximos 7 dias antes que entrem em atraso.")
    long_aging = sum(1 for r in rows if (r.get("Dias na Fase") or 0) >= 8)
    if long_aging:
        recs.append(f"Revisar as {long_aging} OS com 8 dias ou mais na fase para identificar gargalo, espera ou falta de material/informação.")
    if unplanned:
        recs.append(f"Validar prazo das {unplanned} OS sem previsão para evitar carteira sem horizonte de entrega.")
    if len(os_multi_phase):
        recs.append(f"Acompanhar as {len(os_multi_phase)} OS presentes em mais de uma fase para evitar que o avanço entre etapas fique sem sequência.")
    recs.append("Usar a base completa deste relatório como ponto de partida para a programação diária e para a reunião de produção.")
    rec_table = [[str(i), t] for i, t in enumerate(recs, 1)]
    rt = Table([["#", "Ação"]] + rec_table, colWidths=[10*mm, 150*mm], repeatRows=1)
    rt.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#173A5E")),
        ("TEXTCOLOR", (0,0), (-1,0), colors.white), ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
        ("GRID", (0,0), (-1,-1), .2, colors.HexColor("#D0D7DF")),
        ("FONTSIZE", (0,0), (-1,-1), 6.6), ("VALIGN", (0,0), (-1,-1), "TOP"),
    ]))
    story.append(rt)

    doc.build(story, onFirstPage=draw_page, onLaterPages=draw_page)
    return output_path



def generate_pdf_summary(report, output_dir):
    """Gera a versão resumida do relatório: visão executiva, OS atrasadas e base completa."""
    if colors is None:
        raise RuntimeError("ReportLab nao esta instalado. Rode: pip install reportlab")

    os.makedirs(output_dir, exist_ok=True)
    source_name = Path(report["file"]).stem
    phases = report.get("phases", [])
    rows = report.get("rows", [])
    ref_date = report.get("ref_date") or (report.get("emissao").date() if report.get("emissao") else TODAY)

    safe_name = re.sub(
        r"[^A-Za-z0-9_-]+", "_",
        f"Relatorio_PCP_{source_name}_Resumo"
    ).strip("_")
    output_path = os.path.join(output_dir, safe_name + ".pdf")

    status_order = ["Atrasada", "Vence hoje", "Próximos 7 dias", "Prazo futuro", "Sem previsão"]
    status_counts = {s: sum(1 for r in rows if r.get("Status") == s) for s in status_order}
    total = len(rows)
    overdue = [r for r in rows if r.get("Status") == "Atrasada"]
    overdue = sorted(overdue, key=lambda r: (-(r.get("Dias na Fase") or 0), r.get("Prev. Entrega") or date.min))
    critical_overdue = sum(1 for r in overdue if r.get("Criticidade") == "Crítica")

    time_values = [r.get("Dias na Fase") for r in rows if r.get("Dias na Fase") is not None]
    avg_days = sum(time_values) / len(time_values) if time_values else 0
    max_days = max(time_values) if time_values else 0
    phase_header_parts = [f"{ph.get('fase_num', '-')} - {ph.get('fase_nome', '-')}" for ph in phases]
    phase_header_text = " | ".join(phase_header_parts) if phase_header_parts else "-"

    def _wrap_header(text, max_chars=105):
        parts = text.split(" | ")
        lines, current = [], ""
        for part in parts:
            candidate = part if not current else current + " | " + part
            if current and len(candidate) > max_chars:
                lines.append(current)
                current = part
            else:
                current = candidate
        if current:
            lines.append(current)
        return lines[:2] or ["-"]

    phase_header_lines = _wrap_header(phase_header_text)

    def draw_page(canvas, doc):
        canvas.saveState()
        w, h = A4
        canvas.setFillColor(colors.HexColor("#6B7785"))
        canvas.setFont("Helvetica-Bold", 5.8)
        canvas.drawString(10*mm, h-5.5*mm, "FASE(S):")
        canvas.setFont("Helvetica", 5.8)
        for i, line in enumerate(phase_header_lines):
            canvas.drawString(22*mm, h-(5.5 + 3.5*i)*mm, line[:150])
        line_y = h-(8.5 if len(phase_header_lines) > 1 else 11)*mm
        canvas.setStrokeColor(colors.HexColor("#D9E2EC"))
        canvas.line(10*mm, line_y-2.5*mm, w-10*mm, line_y-2.5*mm)
        canvas.line(10*mm, 9*mm, w-10*mm, 9*mm)
        canvas.setFillColor(colors.HexColor("#6B7785"))
        canvas.setFont("Helvetica", 6.5)
        canvas.drawString(10*mm, 5.5*mm, f"Fonte: {source_name} | Referência {report.get('referencia') or '-'}")
        canvas.drawRightString(w-10*mm, 5.5*mm, f"Página {doc.page}")
        canvas.restoreState()

    doc = SimpleDocTemplate(
        output_path, pagesize=A4,
        rightMargin=10*mm, leftMargin=10*mm, topMargin=15*mm, bottomMargin=14*mm,
        title=f"Relatório PCP Resumido - {source_name}",
        author="Gerador de Relatórios PCP - ESA",
    )

    styles = getSampleStyleSheet()
    title = ParagraphStyle("title_sum", parent=styles["Title"], fontSize=18, leading=21,
                           textColor=colors.HexColor("#173A5E"), spaceAfter=4)
    subtitle = ParagraphStyle("sub_sum", parent=styles["BodyText"], fontSize=8.2, leading=10.2,
                              textColor=colors.HexColor("#5E6B78"), spaceAfter=6)
    h1 = ParagraphStyle("h1_sum", parent=styles["Heading2"], fontSize=11, leading=13,
                        textColor=colors.HexColor("#173A5E"), spaceBefore=8, spaceAfter=4)
    h2 = ParagraphStyle("h2_sum", parent=h1, fontSize=9.3, leading=11.2, spaceBefore=6)
    body = ParagraphStyle("body_sum", parent=styles["BodyText"], fontSize=7.8, leading=10.2,
                          textColor=colors.HexColor("#263442"))
    small = ParagraphStyle("small_sum", parent=body, fontSize=6.1, leading=7.6,
                           textColor=colors.HexColor("#5E6B78"))
    phase_cell = ParagraphStyle("phase_cell_sum", parent=body, fontSize=5.0, leading=5.9,
                                textColor=colors.HexColor("#263442"), wordWrap="CJK", spaceAfter=0, spaceBefore=0)
    obs_cell = ParagraphStyle("obs_cell_sum", parent=body, fontSize=5.2, leading=7.0,
                              textColor=colors.HexColor("#263442"), wordWrap="CJK", spaceAfter=0, spaceBefore=0)

    story = []
    mode_title = "RELATÓRIO PCP | RESUMIDO MULTIFASE" if report.get("multi_fase") else f"RELATÓRIO PCP | RESUMIDO | FASE {report.get('fase_num', '')}"
    story.append(Paragraph(mode_title, title))
    if report.get("multi_fase"):
        phase_label = "<b>FASES ANALISADAS:</b> " + " | ".join(
            f"{ph.get('fase_num', '-')} - {_safe_para(ph.get('fase_nome', '-'))}" for ph in phases
        )
    else:
        phase_label = f"<b>FASE {report.get('fase_num', '-')}:</b> {_safe_para(report.get('fase_nome') or (phases[0].get('fase_nome') if phases else '-'))}"
    story.append(Paragraph(phase_label, subtitle))
    story.append(Paragraph(
        f"Arquivo: {_safe_para(source_name)} | Fases: {len(phases)} | Registros: {total} | "
        f"Emissão: {report.get('emissao').strftime('%d/%m/%Y %H:%M') if report.get('emissao') else '-'} | "
        f"Data-base: {ref_date.strftime('%d/%m/%Y')}", subtitle
    ))

    # 1. Visão executiva
    story.append(Paragraph("1. Visão executiva", h1))
    kpi = Table([
        ["REGISTROS", "FASES", "ATRASADAS", "VENCEM HOJE", "PRÓX. 7 DIAS", "SEM PREVISÃO"],
        [total, len(phases), status_counts["Atrasada"], status_counts["Vence hoje"], status_counts["Próximos 7 dias"], status_counts["Sem previsão"]]
    ], colWidths=[28*mm, 22*mm, 29*mm, 29*mm, 32*mm, 30*mm], rowHeights=[6.2*mm, 8.5*mm])
    kpi.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#EAF1F7")),
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
        ("FONTNAME", (0,1), (-1,1), "Helvetica-Bold"),
        ("FONTSIZE", (0,0), (-1,0), 5.7),
        ("FONTSIZE", (0,1), (-1,1), 11.5),
        ("ALIGN", (0,0), (-1,-1), "CENTER"), ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
        ("GRID", (0,0), (-1,-1), .3, colors.HexColor("#CBD5DF")),
        ("BACKGROUND", (2,1), (2,1), colors.HexColor("#FCE4E4")),
        ("BACKGROUND", (3,1), (4,1), colors.HexColor("#FFF2CC")),
    ]))
    story.append(kpi)
    story.append(Spacer(1, 3))
    k2 = Table([
        ["TEMPO MÉDIO NA FASE", "MAIOR TEMPO", "COM INÍCIO", "URGENTES / CURTO PRAZO", "ATRASOS CRÍTICOS"],
        [f"{avg_days:.1f} dias", f"{max_days} dias", str(len(time_values)), str(status_counts["Atrasada"] + status_counts["Vence hoje"] + status_counts["Próximos 7 dias"]), str(critical_overdue)]
    ], colWidths=[33*mm, 29*mm, 24*mm, 42*mm, 37*mm], rowHeights=[6.2*mm, 8.5*mm])
    k2.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#F4F7FA")),
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
        ("FONTNAME", (0,1), (-1,1), "Helvetica-Bold"),
        ("FONTSIZE", (0,0), (-1,0), 5.7), ("FONTSIZE", (0,1), (-1,1), 10.5),
        ("ALIGN", (0,0), (-1,-1), "CENTER"), ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
        ("GRID", (0,0), (-1,-1), .3, colors.HexColor("#D3DCE5")),
    ]))
    story.append(k2)
    if total:
        urgent = status_counts["Atrasada"] + status_counts["Vence hoje"] + status_counts["Próximos 7 dias"]
        diagnosis = (
            f"A carteira possui {total} registros em {len(phases)} fase(s). "
            f"Há {status_counts['Atrasada']} atrasados ({_pct(status_counts['Atrasada'], total)}), "
            f"{status_counts['Vence hoje']} vencendo hoje e {status_counts['Próximos 7 dias']} nos próximos 7 dias. "
            f"Assim, {urgent} registros exigem acompanhamento imediato ou de curto prazo. "
            f"O tempo médio das OS com início de fase é de {avg_days:.1f} dias."
        )
        story.append(Spacer(1, 3))
        story.append(Paragraph(diagnosis, body))

    # 2. OS atrasadas
    # OBS permanece com fundo branco mesmo quando a linha recebe a cor do status.
    story.append(Paragraph(f"2. O.S atrasadas | {len(overdue)} registros", h1))
    late = [["Etapa", "OS", "Cliente", "Marca", "Pot.", "Prev. Entrega", "Atraso", "Dias fase", "Criticidade", "Colaborador", "Motivo", "OBS"]]
    for r in overdue:
        late_days = _days_late(r, ref_date)
        late.append([
            _phase_name_cell(r, phase_cell), r.get("OS", "-"), _cell_para(r.get("Cliente") or "N/I"),
            _cell_para(r.get("Marca") or "N/I"), _cell_para(r.get("Potência") or "N/I"),
            fmt_date(r.get("Prev. Entrega")), str(late_days) if late_days is not None else "-",
            str(r.get("Dias na Fase")) if r.get("Dias na Fase") is not None else "-", r.get("Criticidade", "-"),
            _cell_para(r.get("Colaborador") or "Não informado"), _cell_para(r.get("Motivo Atraso") or "Não informado"),
            Paragraph("<br/>&nbsp;", obs_cell)
        ])
    if not overdue:
        late.append(["-", "-", "Nenhuma O.S atrasada", "-", "-", "-", "-", "-", "-", "-", "-", "-"])
    lt = Table(late, colWidths=[20*mm, 9*mm, 20*mm, 9*mm, 8*mm, 16*mm, 9*mm, 10*mm, 14*mm, 20*mm, 20*mm, 21*mm], repeatRows=1)
    lt.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#173A5E")),
        ("TEXTCOLOR", (0,0), (-1,0), colors.white), ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
        ("GRID", (0,0), (-1,-1), .35, colors.HexColor("#C7B9BE")),
        ("FONTSIZE", (0,0), (-1,-1), 5.6), ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
        ("ALIGN", (1,1), (1,-1), "CENTER"),
        ("ALIGN", (4,1), (7,-1), "CENTER"),
    ]))
    for i in range(1, len(late)):
        lt.setStyle(TableStyle([
            ("BACKGROUND", (0,i), (-1,i), _status_fill("Atrasada")),
            ("BACKGROUND", (-1,i), (-1,i), colors.white),
        ]))
    story.append(lt)
    story.append(Spacer(1, 3))
    story.append(Paragraph("Causas dos atrasos", h2))
    if overdue:
        reason_counts = {}
        for r in overdue:
            reason = r.get("Motivo Atraso") or "Não informado"
            reason_counts[reason] = reason_counts.get(reason, 0) + 1
        reason_table = [["Motivo", "O.S.", "%"]] + [
            [_cell_para(reason), count, safe_pct(count, len(overdue))]
            for reason, count in sorted(reason_counts.items(), key=lambda x: (-x[1], x[0]))
        ]
    else:
        reason_table = [["Motivo", "O.S.", "%"], ["Não há atrasos", 0, "0.0%"]]
    rt = Table(reason_table, colWidths=[112*mm, 22*mm, 22*mm], repeatRows=1)
    rt.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(-1,0),colors.HexColor("#173A5E")),
        ("TEXTCOLOR",(0,0),(-1,0),colors.white), ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),
        ("GRID",(0,0),(-1,-1),.2,colors.HexColor("#D2D9E0")),
        ("FONTSIZE",(0,0),(-1,-1),6.2), ("ALIGN",(1,1),(-1,-1),"CENTER"),
    ]))
    story.append(rt)

    # 3. Base completa
    story.append(Paragraph(f"3. Base completa | {total} registros", h1))
    full = [["Etapa", "OS", "Cliente", "Marca", "Pot.", "Prev.", "Status", "Colab.", "Entrada", "Dias", "OBS"]]
    ordered = sorted(rows, key=lambda r: (int(r.get("Fase Num") or 0), int(r.get("OS") or 0)))
    for r in ordered:
        full.append([
            _phase_name_cell(r, phase_cell), r.get("OS", "-"), _cell_para(r.get("Cliente") or "N/I"),
            _cell_para(r.get("Marca") or "N/I"), _cell_para(r.get("Potência") or "N/I"),
            fmt_date(r.get("Prev. Entrega")), r.get("Status", "-"),
            _cell_para(r.get("Colaborador") or "Não informado"), _fmt_dt(r),
            str(r.get("Dias na Fase")) if r.get("Dias na Fase") is not None else "-",
            Paragraph("<br/>&nbsp;", obs_cell)
        ])
    ft = Table(full, colWidths=[22*mm, 10*mm, 22*mm, 10*mm, 9*mm, 17*mm, 21*mm, 19*mm, 23*mm, 10*mm, 18*mm], repeatRows=1)
    ft.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#173A5E")),
        ("TEXTCOLOR", (0,0), (-1,0), colors.white), ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
        ("GRID", (0,0), (-1,-1), .3, colors.HexColor("#BEC8D2")), ("FONTSIZE", (0,0), (-1,-1), 5.1),
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
        ("ALIGN", (1,1), (1,-1), "CENTER"), ("ALIGN", (4,1), (6,-1), "CENTER"), ("ALIGN", (9,1), (9,-1), "CENTER"),
    ]))
    for i, r in enumerate(ordered, 1):
        ft.setStyle(TableStyle([
            ("BACKGROUND", (0,i), (-1,i), _status_fill(r.get("Status"))),
            ("BACKGROUND", (-1,i), (-1,i), colors.white),
        ]))
    story.append(ft)
    story.append(Spacer(1, 5))
    story.append(Paragraph(
        "Este modelo é intencionalmente resumido: concentra a visão executiva, a fila de O.S atrasadas e a base completa, "
        "mantendo o nome das fases no cabeçalho e em cada registro da base; a coluna OBS fica reservada para anotações.", small
    ))

    doc.build(story, onFirstPage=draw_page, onLaterPages=draw_page)
    return output_path

def generate_pdf(report, output_dir, mode="completo"):
    if mode == "resumido":
        return generate_pdf_summary(report, output_dir)
    return generate_pdf_enhanced(report, output_dir)


QUERY_STOPWORDS = {
    "MOSTRE", "MOSTRAR", "QUAIS", "QUAL", "AS", "OS", "O", "DE", "DA", "DO", "DAS", "DOS",
    "QUE", "ESTAO", "ESTÃO", "NA", "NO", "NAS", "NOS", "EM", "PARA", "POR", "COM", "E",
    "ME", "TEM", "TÊM", "TENHAM", "SOMENTE", "APENAS", "TODAS", "TODOS", "A", "UM", "UMA",
    "MAIS", "MENOS", "FASE", "FASES", "DIAS", "DIA", "HOJE", "AMANHA", "AMANHÃ", "ATE", "ATÉ",
    "PRÓXIMOS", "PROXIMOS", "CRITICAS", "CRÍTICAS", "ATRASADAS", "ATRASADOS", "VENCEM", "VENCE",
}


def _normalize_query(text):
    text = (text or "").lower().strip()
    repl = str.maketrans({
        "á":"a","à":"a","ã":"a","â":"a","ä":"a",
        "é":"e","ê":"e","ë":"e",
        "í":"i","ï":"i",
        "ó":"o","ô":"o","õ":"o","ö":"o",
        "ú":"u","ü":"u",
        "ç":"c",
    })
    return text.translate(repl)


def _query_days_from_text(text):
    m = re.search(r"(\d+)\s*dias?", _normalize_query(text))
    return int(m.group(1)) if m else None


def _query_os_numbers(text):
    return re.findall(r"\b\d{5,8}\b", text or "")


def _query_match_word(text, *words):
    n = _normalize_query(text)
    return any(w in n for w in words)


def filter_rows_by_query(reports, query):
    """Interpreta perguntas simples de PCP e retorna (rows, descricao)."""
    q = _normalize_query(query)
    if not q:
        rows = []
        for rep in reports:
            rows.extend(rep.get("rows", []))
        return rows, "Todas as O.S. carregadas"

    all_rows = []
    ref_dates = []
    for rep in reports:
        ref_dates.append(rep.get("ref_date") or TODAY)
        all_rows.extend(rep.get("rows", []))
    ref_date = max(ref_dates) if ref_dates else TODAY

    rows = list(all_rows)
    description = []

    # OS específica
    os_numbers = _query_os_numbers(query)
    if os_numbers:
        wanted = set(os_numbers)
        rows = [r for r in rows if str(r.get("OS", "")) in wanted]
        description.append("O.S. " + ", ".join(os_numbers))

    # Datas / situação
    atraso_sort = None
    if _query_match_word(q, "atrasadas por tempo na fase", "atrasados por tempo na fase", "atrasadas pelo tempo na fase", "atrasados pelo tempo na fase"):
        rows = [r for r in rows if r.get("Status") == "Atrasada"]
        description.append("atrasadas • ordenadas por dias na fase")
        atraso_sort = "fase"
    elif _query_match_word(q, "atrasadas por data de entrega", "atrasados por data de entrega", "atrasadas pela data de entrega", "atrasados pela data de entrega"):
        rows = [r for r in rows if r.get("Status") == "Atrasada"]
        description.append("atrasadas • ordenadas por data de entrega")
        atraso_sort = "entrega"
    elif _query_match_word(q, "entraram hoje", "iniciaram hoje", "comecaram hoje", "de hoje"):
        rows = [r for r in rows if r.get("Data Início") == ref_date]
        description.append("com início hoje")
    elif _query_match_word(q, "data da os hoje", "abertas hoje", "os abertas hoje"):
        rows = [r for r in rows if r.get("Data OS") == ref_date]
        description.append("com Data OS de hoje")
    elif _query_match_word(q, "vencem hoje", "vence hoje"):
        rows = [r for r in rows if r.get("Prev. Entrega") == ref_date]
        description.append("vencimento hoje")
    elif _query_match_word(q, "amanha", "amanhã"):
        target = ref_date + timedelta(days=1)
        rows = [r for r in rows if r.get("Prev. Entrega") == target]
        description.append("vencimento amanhã")
    elif _query_match_word(q, "proximos 3", "próximos 3"):
        end = ref_date + timedelta(days=3)
        rows = [r for r in rows if r.get("Prev. Entrega") and ref_date < r["Prev. Entrega"] <= end]
        description.append("vencimento nos próximos 3 dias")
    elif _query_match_word(q, "proximos 7", "próximos 7", "proxima semana", "próxima semana"):
        end = ref_date + timedelta(days=7)
        rows = [r for r in rows if r.get("Prev. Entrega") and ref_date < r["Prev. Entrega"] <= end]
        description.append("vencimento nos próximos 7 dias")
    elif "atrasad" in q:
        rows = [r for r in rows if r.get("Status") == "Atrasada"]
        description.append("atrasadas")
    elif "crit" in q:
        rows = [r for r in rows if r.get("Criticidade") == "Crítica"]
        description.append("críticas")
    elif "sem previs" in q:
        rows = [r for r in rows if not r.get("Prev. Entrega") or r.get("Status") == "Sem previsão"]
        description.append("sem previsão")

    # Fase
    mphase = re.search(r"fase\s*(?:n[º°]?\s*)?(\d+)", q)
    if mphase:
        phase = mphase.group(1)
        rows = [r for r in rows if str(r.get("Fase Num", "")) == phase]
        description.append(f"fase {phase}")

    # Dias na fase
    threshold = _query_days_from_text(q)
    if threshold is not None and ("fase" in q or "tempo" in q or "parad" in q or "perman" in q):
        if _query_match_word(q, "mais de", "acima de", "superior a", "a mais de"):
            rows = [r for r in rows if r.get("Dias na Fase") is not None and r["Dias na Fase"] > threshold]
            description.append(f"mais de {threshold} dias na fase")
        elif _query_match_word(q, "menos de", "abaixo de"):
            rows = [r for r in rows if r.get("Dias na Fase") is not None and r["Dias na Fase"] < threshold]
            description.append(f"menos de {threshold} dias na fase")

    # Colaborador explicitamente citado
    mc = re.search(r"colaborador(?:\s*[:=-]?\s*)(.+?)(?:\s+na\s+fase|\s+da\s+fase|$)", q)
    if mc:
        name = mc.group(1).strip(" .,-")
        if name:
            rows = [r for r in rows if _normalize_query(r.get("Colaborador", "")) and name in _normalize_query(r.get("Colaborador", ""))]
            description.append(f"colaborador {name.title()}")

    # Cliente / marca explicitamente citados
    for label, field in (("cliente", "Cliente"), ("marca", "Marca")):
        mm = re.search(rf"{label}\s+(?:=|igual a|chamado|da|do)?\s*([^,;]+?)(?:\s+na\s+fase|\s+atrasad|$)", q)
        if mm:
            value = mm.group(1).strip(" .:-")
            value = re.sub(r"\b(fase\s*\d+|hoje|amanha|amanhã)\b", "", value).strip()
            if value:
                nv = _normalize_query(value)
                rows = [r for r in rows if nv in _normalize_query(r.get(field, ""))]
                description.append(f"{label} {value}")

    # Estado no prazo
    if "no prazo" in q:
        rows = [r for r in rows if r.get("Status") in {"Prazo futuro", "Próximos 7 dias", "Vence hoje"}]
        description.append("no prazo")

    # Ordenação: cada consulta de atraso usa o critério solicitado.
    if atraso_sort == "fase":
        rows.sort(key=lambda r: (-(r.get("Dias na Fase") or -1), r.get("Prev. Entrega") or date.max))
    elif atraso_sort == "entrega":
        rows.sort(key=lambda r: (r.get("Prev. Entrega") or date.min, -(r.get("Dias na Fase") or -1)))
    else:
        rows.sort(key=lambda r: (-float(r.get("Criticidade Score") or 0), -(r.get("Dias na Fase") or 0), r.get("Prev. Entrega") or date.max))
    desc = " • ".join(description) if description else query.strip()
    return rows, desc


def filter_rows_by_selected_os(reports, os_text):
    """Filtra uma lista de O.S. digitadas pelo usuário e ordena da mais atrasada à mais recente."""
    numbers = re.findall(r"\b\d{5,8}\b", os_text or "")
    # preserva a ordem em que o usuário digitou, mas usa conjunto para filtrar rapidamente
    wanted = set(numbers)
    if not wanted:
        return [], "Nenhuma O.S. válida informada"

    all_rows = []
    ref_dates = []
    for rep in reports:
        ref_dates.append(rep.get("ref_date") or TODAY)
        all_rows.extend(rep.get("rows", []))
    ref_date = max(ref_dates) if ref_dates else TODAY

    rows = [r for r in all_rows if str(r.get("OS", "")) in wanted]

    # Mais atrasada primeiro. Para registros sem atraso, as mais recentes
    # (menor diferença negativa para a referência) ficam depois das atrasadas.
    def atraso_valor(r):
        prev = r.get("Prev. Entrega")
        if prev:
            return (ref_date - prev).days
        return -10**9

    rows.sort(key=lambda r: (-(atraso_valor(r)), -(r.get("Dias na Fase") or 0), str(r.get("OS", ""))))
    desc = f"O.S. selecionadas • {len(rows)} registro(s) • ordenadas da mais atrasada à mais recente"
    return rows, desc


def generate_query_pdf(rows, query, output_dir, ref_date):
    if colors is None:
        raise RuntimeError("ReportLab nao esta instalado.")
    os.makedirs(output_dir, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    safe_q = re.sub(r"[^A-Za-z0-9_-]+", "_", _normalize_query(query)).strip("_")[:55] or "consulta_pcp"
    path = os.path.join(output_dir, f"Consulta_PCP_{safe_q}_{stamp}.pdf")

    # A consulta pode conter registros de varias fases. Em paisagem temos espaco
    # para mostrar o numero e o nome da fase sem estourar a largura da pagina.
    doc = SimpleDocTemplate(
        path, pagesize=landscape(A4),
        leftMargin=10*mm, rightMargin=10*mm, topMargin=12*mm, bottomMargin=12*mm
    )
    styles = getSampleStyleSheet()
    title = ParagraphStyle("qt", parent=styles["Title"], fontName="Helvetica-Bold", fontSize=18, leading=22, textColor=colors.HexColor("#1E2450"), spaceAfter=4)
    sub = ParagraphStyle("qs", parent=styles["Normal"], fontSize=9, leading=12, textColor=colors.HexColor("#667085"), spaceAfter=9)
    cell = ParagraphStyle("qc", parent=styles["Normal"], fontSize=7.1, leading=8.4, wordWrap="CJK")
    phase_cell = ParagraphStyle("qphase", parent=cell, fontSize=6.7, leading=8.0)
    obs_cell = ParagraphStyle("qobs", parent=cell, fontSize=6.0, leading=7.2)
    story = [Paragraph("CONSULTA PCP", title), Paragraph(f"Consulta: {query} • Referência: {fmt_date(ref_date)} • {len(rows)} registro(s)", sub)]

    data = [["Etapa", "O.S.", "Cliente", "Prev. Entrega", "Atraso", "Dias Fase", "Status", "Criticidade", "Colaborador", "OBS"]]
    for r in rows:
        prev = r.get("Prev. Entrega")
        atraso = (ref_date - prev).days if prev and prev < ref_date else 0
        etapa = r.get("Fase Nome") or "Não informado"
        data.append([
            Paragraph(ptxt(etapa), phase_cell),
            Paragraph(ptxt(r.get("OS")), cell),
            Paragraph(ptxt(r.get("Cliente")) or "N/I", cell),
            Paragraph(fmt_date(prev), cell),
            Paragraph(str(atraso if atraso > 0 else 0), cell),
            Paragraph(str(r.get("Dias na Fase") if r.get("Dias na Fase") is not None else "-"), cell),
            Paragraph(ptxt(r.get("Status")), cell),
            Paragraph(ptxt(r.get("Criticidade")), cell),
            Paragraph(ptxt(r.get("Colaborador") or "Não informado"), cell),
            Paragraph("<br/>&nbsp;", obs_cell)
        ])

    tbl = Table(
        data, repeatRows=1,
        colWidths=[48*mm, 14*mm, 44*mm, 23*mm, 14*mm, 17*mm, 24*mm, 20*mm, 26*mm, 34*mm]
    )
    ts = [
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#1E2450")),
        ("TEXTCOLOR", (0,0), (-1,0), colors.white),
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
        ("FONTSIZE", (0,0), (-1,-1), 7),
        ("GRID", (0,0), (-1,-1), .35, colors.HexColor("#C9CFD8")),
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
        ("ROWBACKGROUNDS", (0,1), (-1,-1), [colors.white, colors.HexColor("#F7F8FA")]),
        ("LEFTPADDING", (0,0), (-1,-1), 4), ("RIGHTPADDING", (0,0), (-1,-1), 4),
        ("TOPPADDING", (0,0), (-1,-1), 4), ("BOTTOMPADDING", (0,0), (-1,-1), 4),
    ]
    for i, r in enumerate(rows, 1):
        ts.append(("BACKGROUND", (0,i), (-1,i), _status_fill(r.get("Status"))))
        ts.append(("BACKGROUND", (-1,i), (-1,i), colors.white))
    tbl.setStyle(TableStyle(ts))
    story.append(tbl)
    story.append(Spacer(1, 8))
    story.append(Paragraph("Atraso = dias após a previsão de entrega. Dias na Fase = permanência desde o início na fase atual. Etapa = nome completo da fase. A coluna OBS foi deixada em branco para anotações.", sub))
    doc.build(story)
    return path


class App(tk.Tk):
    """Interface inspirada diretamente na referência visual enviada pelo usuário."""

    def __init__(self):
        set_windows_app_user_model_id()
        super().__init__()

        self.title("Gerador de Relatórios PCP - ESA")
        # Ícone profissional do aplicativo
        try:
            icon_path = Path(__file__).resolve().parent / "icones" / "ESA_PCP.ico"
            png_icon = Path(__file__).resolve().parent / "icones" / "ESA_PCP_128.png"
            if icon_path.exists():
                self.iconbitmap(default=str(icon_path))
            if png_icon.exists():
                self._app_icon_img = tk.PhotoImage(file=str(png_icon))
                self.iconphoto(True, self._app_icon_img)
        except Exception:
            pass
        self.geometry("1480x900")
        self.minsize(1180, 720)
        self.configure(bg="#08111F")

        self.files = []
        self.output_dir = tk.StringVar(
            value=str(Path.home() / "Desktop" / "Relatorios_PCP")
        )
        self.status = tk.StringVar(
            value="Tudo pronto para começar."
        )
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", lambda *_: self.refresh_tree())
        self.report_mode = tk.StringVar(value="completo")
        self.auto_open = tk.BooleanVar(value=True)
        self.generated_history = []
        self.history_file = Path(os.getenv("LOCALAPPDATA", str(Path.home()))) / "Gerador_PCP_ESA" / "historico.json"
        self.delay_file = Path(os.getenv("LOCALAPPDATA", str(Path.home()))) / "Gerador_PCP_ESA" / "atrasos.json"
        self.delay_reasons = {}
        self._load_history()
        self._load_delay_reasons()

        self.palette = {
            "bg": "#08111F",
            "sidebar": "#0C1728",
            "sidebar_border": "#1B2940",
            "surface": "#101D30",
            "surface2": "#15253A",
            "surface3": "#1C2D45",
            "text": "#F4F7FB",
            "muted": "#8EA0B8",
            "purple": "#6D4AFF",
            "purple2": "#8F69FF",
            "blue": "#42A5F5",
            "green": "#22C55E",
            "green_dark": "#0D3B2A",
            "red": "#FF4D67",
            "yellow": "#F2B84B",
            "border": "#243550",
        }

        self._setup_styles()
        self._build()

    def _setup_styles(self):
        c = self.palette
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass

        style.configure(".", font=("Segoe UI", 10))
        style.configure("TFrame", background=c["bg"])
        style.configure("TButton",
                        background=c["surface2"],
                        foreground=c["text"],
                        padding=(12, 9),
                        borderwidth=0,
                        relief="flat",
                        focusthickness=0)
        style.map("TButton",
                  background=[("active", c["surface3"]), ("pressed", c["purple"])],
                  foreground=[("active", c["text"]), ("pressed", c["text"])])

        style.configure("Primary.TButton",
                        background=c["purple"],
                        foreground="#FFFFFF",
                        font=("Segoe UI", 10, "bold"),
                        padding=(15, 10),
                        borderwidth=0)
        style.map("Primary.TButton",
                  background=[("active", c["purple2"]), ("pressed", "#5734DB")])

        style.configure("Search.TEntry",
                        fieldbackground=c["surface2"],
                        foreground=c["text"],
                        insertcolor=c["text"],
                        borderwidth=0,
                        padding=(10, 8))

        style.configure(
            "Treeview",
            background=c["surface"],
            fieldbackground=c["surface"],
            foreground=c["text"],
            bordercolor=c["border"],
            rowheight=46,
            font=("Segoe UI", 10),
        )
        style.configure(
            "Treeview.Heading",
            background=c["surface3"],
            foreground="#B9C7D9",
            font=("Segoe UI", 9, "bold"),
            padding=(8, 10),
            relief="flat",
        )
        style.map(
            "Treeview",
            background=[("selected", "#2B1D6C")],
            foreground=[("selected", "#FFFFFF")],
        )
        style.configure(
            "Vertical.TScrollbar",
            background=c["surface3"],
            troughcolor=c["surface"],
            bordercolor=c["surface"],
            arrowcolor=c["muted"],
        )

    def _rounded_rect(self, canvas, x1, y1, x2, y2, radius=18, fill="#1F2937",
                      outline=None, width=1, tags=None):
        """Desenha um card arredondado no Canvas."""
        if outline is None:
            outline = fill
        r = min(radius, (x2-x1)/2, (y2-y1)/2)
        points = [
            x1+r, y1,
            x2-r, y1,
            x2, y1,
            x2, y1+r,
            x2, y2-r,
            x2, y2,
            x2-r, y2,
            x1+r, y2,
            x1, y2,
            x1, y2-r,
            x1, y1+r,
            x1, y1
        ]
        return canvas.create_polygon(
            points, smooth=True, splinesteps=20,
            fill=fill, outline=outline, width=width, tags=tags
        )

    def _build(self):
        c = self.palette

        # Custom title bar matching the reference.
        titlebar = tk.Frame(self, bg="#17144C", height=48)
        titlebar.pack(fill="x", side="top")
        titlebar.pack_propagate(False)

        tk.Label(titlebar, text="⚡", bg="#17144C", fg="#8F69FF",
                 font=("Segoe UI Symbol", 14, "bold")).pack(side="left", padx=(18, 7))
        tk.Label(titlebar, text="Gerador de Relatórios PCP - ESA",
                 bg="#17144C", fg="#F4F7FB",
                 font=("Segoe UI", 10, "bold")).pack(side="left")

        controls = tk.Frame(titlebar, bg="#17144C")
        controls.pack(side="right", padx=8)
        tk.Button(controls, text="—", command=self.iconify, bg="#17144C", fg="#B8C3D6",
                  activebackground="#211D62", activeforeground="#FFFFFF", relief="flat", borderwidth=0,
                  font=("Segoe UI", 11, "bold"), width=3).pack(side="left")
        tk.Button(controls, text="□", command=self._toggle_maximize, bg="#17144C", fg="#B8C3D6",
                  activebackground="#211D62", activeforeground="#FFFFFF", relief="flat", borderwidth=0,
                  font=("Segoe UI", 10), width=3).pack(side="left")
        tk.Button(controls, text="×", command=self.destroy, bg="#17144C", fg="#B8C3D6",
                  activebackground="#8B1E32", activeforeground="#FFFFFF", relief="flat", borderwidth=0,
                  font=("Segoe UI", 13, "bold"), width=3).pack(side="left")

        shell = tk.Frame(self, bg=c["bg"])
        shell.pack(fill="both", expand=True)

        # Sidebar modernizada: navegação estável, sem reconstruir a página a cada clique.
        sidebar = tk.Frame(shell, width=250, bg=c["sidebar"],
                           highlightbackground=c["sidebar_border"], highlightthickness=1)
        sidebar.pack(side="left", fill="y")
        sidebar.pack_propagate(False)

        brand = tk.Frame(sidebar, bg=c["sidebar"], height=126)
        brand.pack(fill="x", padx=20, pady=(18, 4))
        brand.pack_propagate(False)
        try:
            brand_icon = tk.PhotoImage(file=str(Path(__file__).resolve().parent / "icones" / "ESA_PCP_128.png"))
            brand_icon = brand_icon.subsample(2, 2)
            self._brand_icon = brand_icon
            tk.Label(brand, image=brand_icon, bg=c["sidebar"], bd=0).pack(side="left", padx=(0, 12), pady=6)
        except Exception:
            tk.Label(brand, text="ESA", bg=c["sidebar"], fg="#FFFFFF", font=("Segoe UI", 24, "bold")).pack(side="left", padx=(0, 12))
        brand_text = tk.Frame(brand, bg=c["sidebar"])
        brand_text.pack(side="left", fill="both", expand=True, pady=18)
        tk.Label(brand_text, text="PCP", bg=c["sidebar"], fg="#FFFFFF", font=("Segoe UI", 16, "bold")).pack(anchor="w")
        tk.Label(brand_text, text="Planejamento e\nControle da Produção", bg=c["sidebar"], fg="#91A0B6",
                 font=("Segoe UI", 8), justify="left").pack(anchor="w", pady=(3, 0))

        tk.Label(sidebar, text="NAVEGAÇÃO", bg=c["sidebar"], fg="#66758D",
                 font=("Segoe UI", 8, "bold")).pack(anchor="w", padx=24, pady=(8, 8))

        self.sidebar_buttons = {}
        self.sidebar_icons = {}
        self._nav_images = {}
        self._page_emoji_images = {}
        emoji_dir = Path(__file__).resolve().parent / "icones" / "menu_emoji"
        menu = [
            ("Início", "inicio"),
            ("Gerar Relatórios", "gerar"),
            ("Consultas PCP", "consultas"),
            ("Histórico", "historico"),
            ("Gestão de Atrasos", "atrasos"),
            ("Configurações", "config"),
            ("Sobre", "sobre"),
        ]
        for label, page in menu:
            row = tk.Frame(sidebar, bg=c["sidebar"], height=48, cursor="hand2")
            row.pack(fill="x", padx=14, pady=2)
            row.pack_propagate(False)
            accent = tk.Frame(row, bg=c["sidebar"], width=3)
            accent.pack(side="left", fill="y")
            icon_canvas = tk.Canvas(row, width=30, height=30, bg=c["sidebar"], highlightthickness=0, bd=0, cursor="hand2")
            icon_canvas.pack(side="left", padx=(10, 10))
            self._draw_nav_icon(icon_canvas, page, "#AAB7CA")
            text_label = tk.Label(row, text=label, bg=c["sidebar"], fg="#BFC9D8",
                                  font=("Segoe UI", 10), cursor="hand2", anchor="w")
            text_label.pack(side="left", fill="x", expand=True)
            for widget in (row, accent, icon_canvas, text_label):
                widget.bind("<Button-1>", lambda e, pg=page: self.show_page(pg))
            self.sidebar_buttons[page] = row
            self.sidebar_icons[page] = icon_canvas
            row._accent = accent

        footer = tk.Frame(sidebar, bg=c["sidebar"])
        footer.pack(side="bottom", fill="x", padx=24, pady=20)
        tk.Frame(footer, bg="#7C5CFF", width=28, height=3).pack(anchor="w", pady=(0, 11))
        tk.Label(footer, text='"Organização hoje,\n maior produtividade amanhã."',
                 bg=c["sidebar"], fg="#91A0B6", justify="left", font=("Segoe UI", 8, "italic")).pack(anchor="w")
        tk.Label(footer, text="ESA Eletrotécnica", bg=c["sidebar"], fg="#DCE3EE", font=("Segoe UI", 9, "bold")).pack(anchor="w", pady=(18, 0))
        tk.Label(footer, text="PCP  •  V9.8.7", bg=c["sidebar"], fg="#6F7D92", font=("Segoe UI", 8)).pack(anchor="w", pady=(2, 0))

        self.content_root = tk.Frame(shell, bg=c["bg"])
        self.content_root.pack(side="left", fill="both", expand=True, padx=25, pady=28)
        self.page_host = tk.Frame(self.content_root, bg=c["bg"])
        self.page_host.pack(fill="both", expand=True)
        self.content = self.page_host
        self.pages = {}
        self.current_page = None
        self.show_page("inicio")

    def _build_generator_page(self):
        c = self.palette
        # Hero / welcome area
        hero = tk.Canvas(self.content, height=124, bg=c["bg"], highlightthickness=0)
        hero.pack(fill="x", pady=(0, 24))
        self._rounded_rect(hero, 0, 0, 1100, 124, radius=18,
                           fill="#111C2E", outline="#20324F", width=1)

        # Decorative "machine gears" motif
        for cx, cy, rr, outline in [
            (780, 40, 80, "#16243A"),
            (900, 72, 68, "#1A2942"),
            (1040, 42, 92, "#14233A"),
            (1110, 110, 110, "#101F34")
        ]:
            hero.create_oval(cx-rr, cy-rr, cx+rr, cy+rr,
                             outline=outline, width=12)
            hero.create_oval(cx-rr+15, cy-rr+15, cx+rr-15, cy+rr-15,
                             outline="#1A2B46", width=2)
        hero.create_text(28, 36, text="Bem-vindo!",
                         anchor="w", fill="#F7F9FC",
                         font=("Segoe UI", 24, "bold"))
        hero.create_text(
            28, 79,
            text="Gere e acompanhe os relatórios do PCP de forma simples e rápida.",
            anchor="w", fill="#9FB6D5",
            font=("Segoe UI", 11)
        )
        hero.create_text(
            960, 48, text='"Informação certa,\ndecisões melhores."',
            anchor="center", fill="#E6EAF2", justify="center",
            font=("Segoe UI", 10)
        )

        # Action cards
        cards = tk.Frame(self.content, bg=c["bg"])
        cards.pack(fill="x", pady=(0, 24))

        self._action_card(cards, 0, 0, "gerar", "Adicionar PDFs",
                          "Selecione um ou mais arquivos", c["purple"],
                          self.add_files, "left")
        self._action_card(cards, 0, 1, "inicio", "Limpar lista",
                          "Remove todos os arquivos", c["blue"],
                          self.clear_files, "left")
        self._action_card(cards, 0, 2, "config", "Pasta de saída",
                          "Defina onde salvar o relatório", c["green"],
                          self.choose_output, "left")
        self._action_card(cards, 0, 3, "gerar", "GERAR RELATÓRIOS",
                          "Processa e cria o relatório em PDF", c["purple2"],
                          self.generate_all, "left", primary=True)

        # Output card
        out = tk.Frame(cards, bg=c["surface"], highlightbackground=c["border"], highlightthickness=1)
        out_card = tk.Frame(self.content, bg=c["surface"],
                            highlightbackground=c["border"], highlightthickness=1)
        out_card.pack(fill="x", pady=(0, 24))
        self._rounded_frame(out_card, c["surface"])

        top_out = tk.Frame(out_card, bg=c["surface"])
        top_out.pack(fill="x", padx=18, pady=(13, 5))
        out_icon = self._get_emoji_image("config", "small")
        if out_icon is not None:
            tk.Label(top_out, image=out_icon, bg=c["surface"], bd=0).pack(side="left", padx=(0, 10))
        else:
            tk.Label(top_out, text="⚙", bg=c["surface"], fg="#7D5BFF", font=("Segoe UI Emoji", 13)).pack(side="left", padx=(0, 10))
        tk.Label(top_out, text="Pasta de saída", bg=c["surface"], fg="#F3F6FB",
                 font=("Segoe UI", 10, "bold")).pack(side="left")

        pathline = tk.Frame(out_card, bg=c["surface"])
        pathline.pack(fill="x", padx=18, pady=(0, 15))
        tk.Entry(pathline, textvariable=self.output_dir,
                 bg=c["surface2"], fg="#DCE4F1", insertbackground="#DCE4F1",
                 relief="flat", highlightthickness=1,
                 highlightbackground="#30425F", highlightcolor=c["purple"],
                 font=("Segoe UI", 10)).pack(side="left", fill="x", expand=True, ipady=7)
        tk.Button(pathline, text="▢  Escolher...", command=self.choose_output,
                  bg=c["surface2"], fg="#E8EDF5", activebackground=c["surface3"],
                  activeforeground="#FFFFFF", relief="flat", borderwidth=0,
                  font=("Segoe UI", 10), padx=12, pady=7).pack(side="left", padx=(10, 0))

        # Report mode selector
        mode_card = tk.Frame(self.content, bg=c["surface"],
                             highlightbackground=c["border"], highlightthickness=1)
        mode_card.pack(fill="x", pady=(0, 18))
        mode_icon = self._get_emoji_image("gerar", "small")
        if mode_icon is not None:
            tk.Label(mode_card, image=mode_icon, bg=c["surface"], bd=0).pack(side="left", padx=(18, 10), pady=12)
        else:
            tk.Label(mode_card, text="📄", bg=c["surface"], fg="#7D5BFF", font=("Segoe UI Emoji", 13)).pack(side="left", padx=(18, 10), pady=12)
        info = tk.Frame(mode_card, bg=c["surface"])
        info.pack(side="left", fill="both", expand=True, pady=10)
        tk.Label(info, text="Tipo de relatório", bg=c["surface"], fg="#F3F6FB",
                 font=("Segoe UI", 10, "bold")).pack(anchor="w")
        tk.Label(info, text="Escolha entre a análise completa ou um resumo mais direto para gestão. O cabeçalho identifica a(s) fase(s).",
                 bg=c["surface"], fg="#92A2B7", font=("Segoe UI", 8)).pack(anchor="w", pady=(2,0))

        mode_buttons = tk.Frame(mode_card, bg=c["surface"])
        mode_buttons.pack(side="right", padx=16, pady=10)

        self.mode_complete_btn = tk.Button(
            mode_buttons, text="●  Relatório completo", command=lambda: self.set_report_mode("completo"),
            bg=c["purple"], fg="#FFFFFF", activebackground=c["purple2"], activeforeground="#FFFFFF",
            relief="flat", borderwidth=0, font=("Segoe UI", 9, "bold"), padx=13, pady=8
        )
        self.mode_complete_btn.pack(side="left", padx=(0, 8))
        self.mode_summary_btn = tk.Button(
            mode_buttons, text="○  Relatório resumido", command=lambda: self.set_report_mode("resumido"),
            bg=c["surface2"], fg="#C6D0DF", activebackground=c["surface3"], activeforeground="#FFFFFF",
            relief="flat", borderwidth=0, font=("Segoe UI", 9, "bold"), padx=13, pady=8
        )
        self.mode_summary_btn.pack(side="left")

        # Queue header
        queue_header = tk.Frame(self.content, bg=c["bg"])
        queue_header.pack(fill="x", pady=(0, 9))
        queue_icon = self._get_emoji_image("gerar", "small")
        if queue_icon is not None:
            tk.Label(queue_header, image=queue_icon, bg=c["bg"], bd=0).pack(side="left", padx=(5, 9))
        else:
            tk.Label(queue_header, text="📄", bg=c["bg"], fg="#8F69FF", font=("Segoe UI Emoji", 13)).pack(side="left", padx=(5, 9))
        tk.Label(queue_header, text="Arquivos na fila", bg=c["bg"], fg="#F5F7FB",
                 font=("Segoe UI", 13, "bold")).pack(side="left")

        self.queue_badge = tk.Label(queue_header, text="0 arquivos",
                                    bg="#2D1B6F", fg="#BFAEFF",
                                    font=("Segoe UI", 9, "bold"),
                                    padx=9, pady=4)
        self.queue_badge.pack(side="left", padx=10)

        searchbox = tk.Frame(queue_header, bg=c["surface2"],
                             highlightbackground=c["border"], highlightthickness=1)
        searchbox.pack(side="right")
        tk.Label(searchbox, text="⌕", bg=c["surface2"], fg="#93A2B8",
                 font=("Segoe UI Symbol", 15)).pack(side="left", padx=(9, 3))
        tk.Entry(searchbox, textvariable=self.search_var, width=24,
                 bg=c["surface2"], fg="#CBD5E1", insertbackground="#FFFFFF",
                 relief="flat", font=("Segoe UI", 9),
                 highlightthickness=0).pack(side="left", ipady=7)

        # Queue panel
        panel = tk.Frame(self.content, bg=c["surface"],
                         highlightbackground=c["border"], highlightthickness=1)
        panel.pack(fill="both", expand=True)
        self._rounded_frame(panel, c["surface"])

        cols = ("num", "arquivo", "fase", "os", "atrasadas", "colaboradores", "status", "acoes")
        self.tree = ttk.Treeview(panel, columns=cols, show="headings", selectmode="browse")

        widths = {
            "num": 46, "arquivo": 290, "fase": 340, "os": 75,
            "atrasadas": 90, "colaboradores": 120, "status": 120, "acoes": 100
        }
        headings = {
            "num": "#", "arquivo": "Arquivo", "fase": "Fase", "os": "OS",
            "atrasadas": "Atrasadas", "colaboradores": "Colaboradores",
            "status": "Status", "acoes": "Ações"
        }
        for col in cols:
            self.tree.heading(col, text=headings[col])
            self.tree.column(col, width=widths[col],
                             anchor="center" if col != "arquivo" else "w")

        scroll = ttk.Scrollbar(panel, orient="vertical",
                               command=self.tree.yview,
                               style="Vertical.TScrollbar")
        self.tree.configure(yscrollcommand=scroll.set)
        self.tree.pack(side="left", fill="both", expand=True, padx=(10, 0), pady=10)
        scroll.pack(side="right", fill="y", padx=(0, 8), pady=10)

        # Footer status
        footerbar = tk.Frame(self.content, bg="#06382B",
                             highlightbackground="#138A60", highlightthickness=1)
        footerbar.pack(fill="x", pady=(16, 0))
        self._rounded_frame(footerbar, "#06382B")

        tk.Label(footerbar, text="✓", bg="#06382B", fg="#4ADE80",
                 font=("Segoe UI", 17, "bold")).pack(side="left", padx=(18, 11), pady=10)
        self.footer_title = tk.Label(
            footerbar, text="Tudo pronto para gerar seu relatório!",
            bg="#06382B", fg="#E7FFF2", font=("Segoe UI", 10, "bold")
        )
        self.footer_title.pack(side="left")
        tk.Label(
            footerbar,
            text="Revise os arquivos e clique em Gerar Relatórios.",
            bg="#06382B", fg="#B1CDBE", font=("Segoe UI", 9)
        ).pack(side="right", padx=18)

        self._update_queue_count()

    def _clear_content(self):
        """Mantido por compatibilidade. As páginas agora são cacheadas para evitar flicker."""
        if not hasattr(self, "pages"):
            return
        for frame in self.pages.values():
            frame.place_forget()

    def _draw_nav_icon(self, canvas, page, color):
        """Exibe emojis coloridos pré-renderizados para uma aparência consistente no Windows."""
        canvas.delete("all")
        emoji_dir = Path(__file__).resolve().parent / "icones" / "menu_emoji"
        icon_path = emoji_dir / f"{page}.png"
        if page not in self._nav_images and icon_path.exists():
            try:
                self._nav_images[page] = tk.PhotoImage(file=str(icon_path))
            except Exception:
                self._nav_images[page] = None
        img = self._nav_images.get(page)
        if img is not None:
            canvas.create_image(15, 15, image=img)
        else:
            fallback = {
                "inicio": "🏠", "gerar": "📄", "consultas": "🔎", "historico": "🕒",
                "atrasos": "⏱️", "config": "⚙️", "sobre": "ℹ️"
            }.get(page, "•")
            canvas.create_text(15, 15, text=fallback, fill=color, font=("Segoe UI Emoji", 14))

    def _style_sidebar(self, active_page):
        c = self.palette
        for page, row in self.sidebar_buttons.items():
            active = page == active_page
            row_bg = "#2B1E63" if active else c["sidebar"]
            row.configure(bg=row_bg)
            row._accent.configure(bg=c["purple"] if active else row_bg)
            icon = self.sidebar_icons.get(page)
            if icon is not None:
                icon.configure(bg=row_bg)
            labels = [w for w in row.winfo_children() if isinstance(w, tk.Label)]
            if labels:
                labels[-1].configure(bg=row_bg, fg="#FFFFFF" if active else "#BFC9D8",
                                     font=("Segoe UI", 10, "bold" if active else "normal"))

    def show_page(self, page):
        if not hasattr(self, "page_host"):
            return
        self.current_page = page
        self._style_sidebar(page)
        frame = self.pages.get(page)
        if frame is None:
            frame = tk.Frame(self.page_host, bg=self.palette["bg"])
            self.pages[page] = frame
            frame.place(relx=0, rely=0, relwidth=1, relheight=1)
            previous = self.content
            self.content = frame
            try:
                if page == "inicio":
                    self._build_home_page()
                elif page == "gerar":
                    self._build_generator_page()
                elif page == "consultas":
                    self._build_query_page()
                elif page == "historico":
                    self._build_history_page()
                elif page == "atrasos":
                    self._build_delay_page()
                elif page == "config":
                    self._build_config_page()
                elif page == "sobre":
                    self._build_about_page()
            finally:
                self.content = previous
        frame.lift()
        self.content = self.page_host

    def _page_title(self, title, subtitle, icon="▦"):
        c = self.palette
        head = tk.Frame(self.content, bg=c["bg"])
        head.pack(fill="x", pady=(0, 18))

        # Usa os mesmos emojis coloridos do menu lateral, em tamanho menor,
        # para manter a identidade visual dentro de cada tela/aba.
        page_map = {
            "Início": "inicio",
            "Gerar Relatórios": "gerar",
            "Consultas PCP": "consultas",
            "Histórico": "historico",
            "Gestão de Atrasos": "atrasos",
            "Configurações": "config",
            "Sobre": "sobre",
        }
        page_key = page_map.get(title)
        title_img = None
        if page_key:
            small_dir = Path(__file__).resolve().parent / "icones" / "menu_emoji" / "small"
            icon_path = small_dir / f"{page_key}.png"
            if page_key not in self._page_emoji_images and icon_path.exists():
                try:
                    self._page_emoji_images[page_key] = tk.PhotoImage(file=str(icon_path))
                except Exception:
                    self._page_emoji_images[page_key] = None
            title_img = self._page_emoji_images.get(page_key)

        if title_img is not None:
            tk.Label(head, image=title_img, bg=c["bg"], width=20, height=20).pack(side="left", padx=(2, 12), pady=(2, 0))
        else:
            tk.Label(head, text=icon, bg=c["bg"], fg=c["purple"], font=("Segoe UI Symbol", 18)).pack(side="left", padx=(2, 12))

        text = tk.Frame(head, bg=c["bg"])
        text.pack(side="left")
        tk.Label(text, text=title, bg=c["bg"], fg=c["text"], font=("Segoe UI", 22, "bold")).pack(anchor="w")
        tk.Label(text, text=subtitle, bg=c["bg"], fg=c["muted"], font=("Segoe UI", 10)).pack(anchor="w", pady=(3, 0))

    def _get_emoji_image(self, name, size="small"):
        """Carrega um emoji colorido em PNG para uso em qualquer tela."""
        cache = getattr(self, "_content_emoji_images", None)
        if cache is None:
            cache = {}
            self._content_emoji_images = cache
        key = f"{size}:{name}"
        if key in cache:
            return cache[key]
        base = Path(__file__).resolve().parent / "icones" / "menu_emoji"
        folder = base / size if size in ("small", "small24") else base
        path = folder / f"{name}.png"
        try:
            img = tk.PhotoImage(file=str(path)) if path.exists() else None
        except Exception:
            img = None
        cache[key] = img
        return img

    def _emoji_label(self, parent, name, size="small", bg=None, padx=(0, 8), pady=0):
        """Cria um Label com emoji colorido, com fallback silencioso."""
        c = self.palette
        bg = c["bg"] if bg is None else bg
        img = self._get_emoji_image(name, size)
        if img is not None:
            lbl = tk.Label(parent, image=img, bg=bg, bd=0, highlightthickness=0)
        else:
            fallback = {
                "inicio": "🏠", "gerar": "📄", "historico": "🕒",
                "atrasos": "⏱", "config": "⚙", "sobre": "ℹ"
            }.get(name, "•")
            lbl = tk.Label(parent, text=fallback, bg=bg, fg=c["purple"],
                           font=("Segoe UI Emoji", 13))
        lbl.pack(side="left", padx=padx, pady=pady)
        return lbl

    def _card(self, parent, title, value, accent=None, subtitle=""):
        c = self.palette
        accent = accent or c["purple"]
        frame = tk.Frame(parent, bg=c["surface"], highlightbackground=c["border"], highlightthickness=1)
        frame.pack(side="left", fill="both", expand=True, padx=(0, 12))
        tk.Frame(frame, bg=accent, width=4).pack(side="left", fill="y")
        body = tk.Frame(frame, bg=c["surface"])
        body.pack(fill="both", expand=True, padx=15, pady=14)
        tk.Label(body, text=title.upper(), bg=c["surface"], fg=c["muted"], font=("Segoe UI", 8, "bold")).pack(anchor="w")
        tk.Label(body, text=value, bg=c["surface"], fg=c["text"], font=("Segoe UI", 20, "bold")).pack(anchor="w", pady=(3, 0))
        if subtitle:
            tk.Label(body, text=subtitle, bg=c["surface"], fg=c["muted"], font=("Segoe UI", 8)).pack(anchor="w", pady=(3, 0))
        return frame

    def _build_home_page(self):
        c = self.palette
        self._page_title("Início", "Visão rápida do ambiente de trabalho do PCP.", "⌂")
        cards = tk.Frame(self.content, bg=c["bg"])
        cards.pack(fill="x", pady=(0, 22))
        self._card(cards, "Arquivos na fila", str(len(self.files)), c["purple"], "PDFs aguardando processamento")
        self._card(cards, "Relatórios gerados", str(len(self.generated_history)), c["green"], "Registros no histórico")
        self._card(cards, "Modo atual", "Completo" if self.report_mode.get() == "completo" else "Resumido", c["blue"], "Modelo selecionado")
        self._card(cards, "Saída", "Definida" if self.output_dir.get() else "Não definida", c["yellow"], "Pasta onde os PDFs serão salvos")

        hero = tk.Frame(self.content, bg=c["surface"], highlightbackground=c["border"], highlightthickness=1)
        hero.pack(fill="x", pady=(0, 18))
        tk.Label(hero, text="Central de PCP", bg=c["surface"], fg=c["text"], font=("Segoe UI", 16, "bold")).pack(anchor="w", padx=20, pady=(18, 4))
        tk.Label(hero, text="Importe relatórios ESA, escolha o tipo de análise e gere o PDF sem precisar abrir o relatório manualmente.",
                 bg=c["surface"], fg=c["muted"], font=("Segoe UI", 10), wraplength=900, justify="left").pack(anchor="w", padx=20)
        actions = tk.Frame(hero, bg=c["surface"])
        actions.pack(fill="x", padx=20, pady=18)
        tk.Button(actions, text="📄  Gerar Relatórios", command=lambda: self.show_page("gerar"), bg=c["purple"], fg="#FFFFFF",
                  activebackground=c["purple2"], activeforeground="#FFFFFF", relief="flat", borderwidth=0,
                  font=("Segoe UI", 10, "bold"), padx=16, pady=9).pack(side="left")
        tk.Button(actions, text="🔎  Consultas PCP", command=lambda: self.show_page("consultas"), bg=c["surface2"], fg=c["text"],
                  activebackground=c["surface3"], activeforeground="#FFFFFF", relief="flat", borderwidth=0,
                  font=("Segoe UI", 10), padx=16, pady=9).pack(side="left", padx=(10, 0))
        tk.Button(actions, text="🕒  Ver Histórico", command=self.refresh_history_page, bg=c["surface2"], fg=c["text"],
                  activebackground=c["surface3"], activeforeground="#FFFFFF", relief="flat", borderwidth=0,
                  font=("Segoe UI", 10), padx=16, pady=9).pack(side="left", padx=(10, 0))

        info = tk.Frame(self.content, bg=c["bg"])
        info.pack(fill="both", expand=True)
        left = tk.Frame(info, bg=c["surface"], highlightbackground=c["border"], highlightthickness=1)
        left.pack(side="left", fill="both", expand=True, padx=(0, 10))
        tk.Label(left, text="Fluxo recomendado", bg=c["surface"], fg=c["text"], font=("Segoe UI", 13, "bold")).pack(anchor="w", padx=18, pady=(16, 8))
        for n, txt in enumerate(["Adicione um ou mais PDFs ESA.", "Escolha completo ou resumido.", "Confira a fila e gere os relatórios.", "Consulte o histórico para reabrir arquivos anteriores."], 1):
            tk.Label(left, text=f"{n:02d}   {txt}", bg=c["surface"], fg="#D4DCE8", font=("Segoe UI", 10)).pack(anchor="w", padx=18, pady=6)
        right = tk.Frame(info, bg=c["surface"], highlightbackground=c["border"], highlightthickness=1)
        right.pack(side="left", fill="both", expand=True, padx=(10, 0))
        tk.Label(right, text="Último relatório", bg=c["surface"], fg=c["text"], font=("Segoe UI", 13, "bold")).pack(anchor="w", padx=18, pady=(16, 8))
        if self.generated_history:
            item = self.generated_history[-1]
            tk.Label(right, text=item.get("file", "-"), bg=c["surface"], fg="#E7ECF4", font=("Segoe UI", 10, "bold"), wraplength=450, justify="left").pack(anchor="w", padx=18)
            tk.Label(right, text=f"{item.get('date','-')}  |  {item.get('mode','-')}  |  {item.get('phases','-')}", bg=c["surface"], fg=c["muted"], font=("Segoe UI", 9)).pack(anchor="w", padx=18, pady=(5, 0))
        else:
            tk.Label(right, text="Nenhum relatório gerado ainda.", bg=c["surface"], fg=c["muted"], font=("Segoe UI", 10)).pack(anchor="w", padx=18)

    def _build_query_page(self):
        c = self.palette
        self._page_title("Consultas PCP", "Faça perguntas rápidas sobre as O.S. carregadas e gere um PDF da consulta.", "🔎")

        box = tk.Frame(self.content, bg=c["surface"], highlightbackground=c["border"], highlightthickness=1)
        box.pack(fill="x", pady=(0, 12))
        tk.Label(box, text="Pergunta", bg=c["surface"], fg=c["text"], font=("Segoe UI", 11, "bold")).pack(anchor="w", padx=16, pady=(14, 5))
        line = tk.Frame(box, bg=c["surface"])
        line.pack(fill="x", padx=16, pady=(0, 12))
        self.query_var = tk.StringVar()
        entry = tk.Entry(line, textvariable=self.query_var, bg=c["surface2"], fg=c["text"], insertbackground=c["text"], relief="flat", font=("Segoe UI", 10))
        entry.pack(side="left", fill="x", expand=True, ipady=9)
        entry.bind("<Return>", lambda e: self.run_query())
        tk.Button(line, text="🔎 Pesquisar", command=self.run_query, bg=c["purple"], fg="white", activebackground=c["purple2"], relief="flat", borderwidth=0, font=("Segoe UI", 9, "bold"), padx=15, pady=8).pack(side="left", padx=(8,0))
        tk.Button(line, text="Limpar", command=self.clear_query, bg=c["surface2"], fg=c["text"], relief="flat", borderwidth=0, font=("Segoe UI", 9), padx=12, pady=8).pack(side="left", padx=(8,0))

        shortcuts = tk.Frame(self.content, bg=c["bg"])
        shortcuts.pack(fill="x", pady=(0, 10))
        quick = [
            ("📅 Hoje", "quais os entraram hoje?"),
            ("🔴 Atrasadas", "quais os estao atrasadas?"),
            ("⚠️ Críticas", "quais os criticas?"),
            ("📆 Vencem hoje", "quais os vencem hoje?"),
            ("⏱ Mais de 15 dias", "quais os com mais de 15 dias na fase?"),
            ("📋 Todas", "todas as os"),
        ]
        for label, q in quick:
            tk.Button(shortcuts, text=label, command=lambda qq=q: self._set_query_and_run(qq), bg=c["surface2"], fg=c["text"], activebackground=c["surface3"], relief="flat", borderwidth=0, font=("Segoe UI", 9), padx=11, pady=7).pack(side="left", padx=(0,7))

        overdue_box = tk.Frame(self.content, bg=c["surface"], highlightbackground=c["border"], highlightthickness=1)
        overdue_box.pack(fill="x", pady=(0, 12))
        tk.Label(overdue_box, text="🔴 Filtrar O.S. atrasadas", bg=c["surface"], fg=c["text"], font=("Segoe UI", 11, "bold")).pack(anchor="w", padx=16, pady=(12, 2))
        tk.Label(overdue_box, text="Escolha como deseja organizar as O.S. atrasadas:", bg=c["surface"], fg=c["muted"], font=("Segoe UI", 9)).pack(anchor="w", padx=16, pady=(0, 8))
        self.overdue_filter_var = tk.StringVar(value="fase")
        radio_line = tk.Frame(overdue_box, bg=c["surface"])
        radio_line.pack(fill="x", padx=16, pady=(0, 12))
        tk.Radiobutton(radio_line, text="⏱  Por Dias na Fase", variable=self.overdue_filter_var, value="fase", bg=c["surface"], fg=c["text"], selectcolor=c["surface2"], activebackground=c["surface"], activeforeground=c["text"], font=("Segoe UI", 9, "bold")).pack(side="left", padx=(0,20))
        tk.Radiobutton(radio_line, text="📆  Por Data de Entrega", variable=self.overdue_filter_var, value="entrega", bg=c["surface"], fg=c["text"], selectcolor=c["surface2"], activebackground=c["surface"], activeforeground=c["text"], font=("Segoe UI", 9, "bold")).pack(side="left", padx=(0,20))
        tk.Button(radio_line, text="🔎 Consultar Atrasadas", command=self.run_overdue_filter, bg=c["red"], fg="white", activebackground="#A72F42", relief="flat", borderwidth=0, font=("Segoe UI", 9, "bold"), padx=13, pady=7).pack(side="left")

        selected_box = tk.Frame(self.content, bg=c["surface"], highlightbackground=c["border"], highlightthickness=1)
        selected_box.pack(fill="x", pady=(0, 12))
        tk.Label(selected_box, text="📌 Relatório por O.S. selecionadas", bg=c["surface"], fg=c["text"], font=("Segoe UI", 11, "bold")).pack(anchor="w", padx=16, pady=(12, 2))
        tk.Label(selected_box, text="Digite uma ou várias O.S. (separe por vírgula, espaço ou quebra de linha). O resultado será ordenado da mais atrasada para a mais recente.", bg=c["surface"], fg=c["muted"], font=("Segoe UI", 9), wraplength=1080, justify="left").pack(anchor="w", padx=16, pady=(0, 8))
        selected_line = tk.Frame(selected_box, bg=c["surface"])
        selected_line.pack(fill="x", padx=16, pady=(0, 12))
        self.selected_os_var = tk.StringVar()
        selected_entry = tk.Entry(selected_line, textvariable=self.selected_os_var, bg=c["surface2"], fg=c["text"], insertbackground=c["text"], relief="flat", font=("Segoe UI", 10))
        selected_entry.pack(side="left", fill="x", expand=True, ipady=9)
        selected_entry.bind("<Return>", lambda e: self.run_selected_os_filter())
        tk.Button(selected_line, text="🔎 Filtrar O.S.", command=self.run_selected_os_filter, bg=c["blue"], fg="white", activebackground="#1976D2", relief="flat", borderwidth=0, font=("Segoe UI", 9, "bold"), padx=15, pady=8).pack(side="left", padx=(8,0))
        tk.Button(selected_line, text="📄 Gerar PDF", command=self.export_selected_os_pdf, bg=c["purple"], fg="white", activebackground=c["purple2"], relief="flat", borderwidth=0, font=("Segoe UI", 9, "bold"), padx=15, pady=8, state="disabled").pack(side="left", padx=(8,0))
        self.selected_os_pdf_btn = selected_line.winfo_children()[-1]
        tk.Button(selected_line, text="Limpar", command=self.clear_selected_os, bg=c["surface2"], fg=c["text"], relief="flat", borderwidth=0, font=("Segoe UI", 9), padx=12, pady=8).pack(side="left", padx=(8,0))

        hint = tk.Frame(self.content, bg="#10233B", highlightbackground="#2B4D76", highlightthickness=1)
        hint.pack(fill="x", pady=(0, 12))
        tk.Label(hint, text="💡 Exemplos: “OS da fase 33 atrasadas”, “atrasadas por tempo na fase”, “atrasadas por data de entrega”, “OS que vencem amanhã”, “OS com mais de 20 dias na fase”, “OS do colaborador Carlos”.", bg="#10233B", fg="#CFE2F7", font=("Segoe UI", 9), wraplength=1000, justify="left").pack(anchor="w", padx=14, pady=10)

        result = tk.Frame(self.content, bg=c["surface"], highlightbackground=c["border"], highlightthickness=1)
        result.pack(fill="both", expand=True)
        header = tk.Frame(result, bg=c["surface"])
        header.pack(fill="x", padx=12, pady=(10, 0))
        self.query_result_label = tk.Label(header, text="Nenhuma consulta executada.", bg=c["surface"], fg=c["text"], font=("Segoe UI", 10, "bold"))
        self.query_result_label.pack(side="left")
        self.query_export_btn = tk.Button(header, text="📄 Gerar PDF da consulta", command=self.export_query_pdf, bg=c["purple"], fg="white", activebackground=c["purple2"], relief="flat", borderwidth=0, font=("Segoe UI", 9, "bold"), padx=12, pady=7, state="disabled")
        self.query_export_btn.pack(side="right")
        cols=("fase","os","cliente","prev","atraso","dias","status","crit")
        tree=ttk.Treeview(result, columns=cols, show="headings", selectmode="browse")
        self.query_tree=tree
        heads={"fase":"Etapa","os":"O.S.","cliente":"Cliente","prev":"Prev. Entrega","atraso":"Atraso","dias":"Dias Fase","status":"Status","crit":"Criticidade"}
        widths={"fase":80,"os":90,"cliente":360,"prev":100,"atraso":80,"dias":90,"status":130,"crit":100}
        for col in cols:
            tree.heading(col, text=heads[col]); tree.column(col, width=widths[col], anchor="w")
        scroll=ttk.Scrollbar(result, orient="vertical", command=tree.yview, style="Vertical.TScrollbar"); tree.configure(yscrollcommand=scroll.set)
        tree.pack(side="left", fill="both", expand=True, padx=(10,0), pady=10); scroll.pack(side="right", fill="y", padx=(0,8), pady=10)
        self.query_rows=[]
        self.query_text=""
        self.query_ref_date=TODAY
        self.selected_os_rows=[]
        self.selected_os_text=""
        self.selected_os_ref_date=TODAY

    def run_selected_os_filter(self):
        text = self.selected_os_var.get().strip()
        reports = self._read_queued_reports()
        if not reports:
            messagebox.showwarning("Nenhum dado carregado", "Adicione pelo menos um PDF na aba Gerar Relatórios antes de filtrar O.S.")
            return
        if not re.findall(r"\b\d{5,8}\b", text):
            messagebox.showwarning("O.S. não informada", "Digite uma ou mais O.S., por exemplo: 289172, 289174, 290260")
            return
        rows, desc = filter_rows_by_selected_os(reports, text)
        refs = [rep.get("ref_date") for rep in reports if rep.get("ref_date")]
        ref_date = max(refs) if refs else TODAY
        self.selected_os_rows = rows
        self.selected_os_text = text
        self.selected_os_ref_date = ref_date

        for item in self.query_tree.get_children():
            self.query_tree.delete(item)
        for r in rows:
            prev = r.get("Prev. Entrega")
            atraso = (ref_date - prev).days if prev and prev < ref_date else 0
            self.query_tree.insert("", "end", values=(
                r.get("Fase Nome") or "Não informado",
                r.get("OS", "-"),
                r.get("Cliente") or "N/I",
                fmt_date(prev),
                f"{atraso} d",
                f"{r.get('Dias na Fase')} d" if r.get("Dias na Fase") is not None else "-",
                r.get("Status", "-"),
                r.get("Criticidade", "-")
            ))
        self.query_result_label.config(text=f"{len(rows)} registro(s) encontrados • {desc}")
        self.query_export_btn.config(state="normal" if rows else "disabled")
        self.selected_os_pdf_btn.config(state="normal" if rows else "disabled")
        self.status.set(f"Filtro de O.S.: {len(rows)} registro(s) encontrados.")

    def clear_selected_os(self):
        self.selected_os_var.set("")
        self.selected_os_rows = []
        self.selected_os_text = ""
        self.selected_os_pdf_btn.config(state="disabled")

    def export_selected_os_pdf(self):
        if not self.selected_os_rows:
            return
        try:
            os.makedirs(self.output_dir.get(), exist_ok=True)
            query = f"O.S. selecionadas: {self.selected_os_text} (mais atrasada para mais recente)"
            path = generate_query_pdf(self.selected_os_rows, query, self.output_dir.get(), self.selected_os_ref_date)
            if self.auto_open.get():
                self._open_path(path)
            self.status.set("PDF das O.S. selecionadas gerado com sucesso.")
            self.footer_title.config(text="PDF das O.S. selecionadas gerado.")
        except Exception as e:
            messagebox.showerror("Erro ao gerar consulta por O.S.", str(e))

    def _set_query_and_run(self, q):
        self.query_var.set(q)
        self.run_query()

    def run_overdue_filter(self):
        mode = getattr(self, "overdue_filter_var", None)
        mode = mode.get() if mode is not None else "fase"
        if mode == "entrega":
            self._set_query_and_run("quais os atrasadas por data de entrega?")
        else:
            self._set_query_and_run("quais os atrasadas por tempo na fase?")

    def clear_query(self):
        self.query_var.set("")
        self.query_rows=[]
        self.query_result_label.config(text="Nenhuma consulta executada.")
        self.query_export_btn.config(state="disabled")
        for item in self.query_tree.get_children():
            self.query_tree.delete(item)

    def run_query(self):
        query = self.query_var.get().strip()
        reports = self._read_queued_reports()
        if not reports:
            messagebox.showwarning("Nenhum dado carregado", "Adicione pelo menos um PDF na aba Gerar Relatórios para fazer consultas.")
            return
        rows, desc = filter_rows_by_query(reports, query)
        self.query_rows = rows
        self.query_text = query or "Todas as O.S."
        refs = [rep.get("ref_date") for rep in reports if rep.get("ref_date")]
        self.query_ref_date = max(refs) if refs else TODAY
        for item in self.query_tree.get_children():
            self.query_tree.delete(item)
        for r in rows:
            prev=r.get("Prev. Entrega")
            atraso=(self.query_ref_date-prev).days if prev and prev < self.query_ref_date else 0
            self.query_tree.insert("", "end", values=(r.get("Fase Num","-"),r.get("OS","-"),r.get("Cliente") or "N/I",fmt_date(prev),f"{atraso} d",f"{r.get('Dias na Fase')} d" if r.get("Dias na Fase") is not None else "-",r.get("Status","-"),r.get("Criticidade","-")))
        self.query_result_label.config(text=f"{len(rows)} registro(s) encontrados • {desc}")
        self.query_export_btn.config(state="normal" if rows else "disabled")
        self.status.set(f"Consulta: {len(rows)} registro(s) encontrados.")

    def export_query_pdf(self):
        if not self.query_rows:
            return
        try:
            os.makedirs(self.output_dir.get(), exist_ok=True)
            path = generate_query_pdf(self.query_rows, self.query_text, self.output_dir.get(), self.query_ref_date)
            if self.auto_open.get():
                self._open_path(path)
            self.status.set("PDF da consulta gerado com sucesso.")
            self.footer_title.config(text="PDF da consulta gerado.")
        except Exception as e:
            messagebox.showerror("Erro ao gerar consulta", str(e))

    def _build_history_page(self):
        c = self.palette
        self._page_title("Histórico", "Acompanhe e reabra os relatórios já gerados.", "◷")
        toolbar = tk.Frame(self.content, bg=c["bg"])
        toolbar.pack(fill="x", pady=(0, 10))
        tk.Button(toolbar, text="🔄  Atualizar", command=lambda: self.show_page("historico"), bg=c["surface2"], fg=c["text"], relief="flat", borderwidth=0, font=("Segoe UI", 9), padx=12, pady=7).pack(side="left")
        tk.Button(toolbar, text="🗑️  Limpar histórico", command=self.clear_history, bg=c["surface2"], fg="#FFB8C4", relief="flat", borderwidth=0, font=("Segoe UI", 9), padx=12, pady=7).pack(side="right")
        panel = tk.Frame(self.content, bg=c["surface"], highlightbackground=c["border"], highlightthickness=1)
        panel.pack(fill="both", expand=True)
        cols=("data","arquivo","modo","fases","total","atrasadas","criticas","media","pasta")
        tree=ttk.Treeview(panel, columns=cols, show="headings", selectmode="browse")
        self.history_tree=tree
        headings={"data":"Data/Hora","arquivo":"Relatório","modo":"Modo","fases":"Fases","total":"OS","atrasadas":"Atrasadas","criticas":"Críticas","media":"Média fase","pasta":"Pasta de saída"}
        widths={"data":135,"arquivo":290,"modo":85,"fases":310,"total":60,"atrasadas":75,"criticas":70,"media":80,"pasta":360}
        for col in cols:
            tree.heading(col,text=headings[col]); tree.column(col,width=widths[col],anchor="w")
        scroll=ttk.Scrollbar(panel,orient="vertical",command=tree.yview,style="Vertical.TScrollbar"); tree.configure(yscrollcommand=scroll.set)
        tree.pack(side="left",fill="both",expand=True,padx=(10,0),pady=10); scroll.pack(side="right",fill="y",padx=(0,8),pady=10)
        for idx, item in enumerate(reversed(self.generated_history)):
            tree.insert("","end",iid=str(idx),values=(item.get("date","-"),item.get("file","-"),item.get("mode","-"),item.get("phases","-"),item.get("total","-"),item.get("overdue","-"),item.get("critical","-"),item.get("avg_phase","-"),item.get("path","-")))
        actions=tk.Frame(self.content,bg=c["bg"]); actions.pack(fill="x",pady=(12,0))
        tk.Button(actions,text="📂 Abrir selecionado",command=self.open_history_selected,bg=c["purple"],fg="#FFFFFF",activebackground=c["purple2"],relief="flat",borderwidth=0,font=("Segoe UI",9,"bold"),padx=13,pady=8).pack(side="left")
        tk.Button(actions,text="📁 Abrir pasta",command=self.open_history_folder,bg=c["surface2"],fg=c["text"],relief="flat",borderwidth=0,font=("Segoe UI",9),padx=13,pady=8).pack(side="left",padx=(8,0))

    def _build_delay_page(self):
        c = self.palette
        self._page_title("Gestão de Atrasos", "Classifique as O.S. atrasadas e registre a ação necessária.", "⏱")
        if not self.files:
            note = tk.Frame(self.content, bg=c["surface"], highlightbackground=c["border"], highlightthickness=1)
            note.pack(fill="x", pady=(0, 14))
            tk.Label(note, text="Nenhum PDF está na fila.", bg=c["surface"], fg=c["text"], font=("Segoe UI", 13, "bold")).pack(anchor="w", padx=18, pady=(18,5))
            tk.Label(note, text="Abra Gerar Relatórios e adicione um ou mais PDFs para carregar as O.S. atrasadas.", bg=c["surface"], fg=c["muted"], font=("Segoe UI", 10)).pack(anchor="w", padx=18, pady=(0,18))
            return

        reports = self._read_queued_reports()
        overdue = []
        for report in reports:
            for r in report.get("rows", []):
                if r.get("Status") == "Atrasada":
                    overdue.append(r)
        overdue.sort(key=lambda r: (-(_days_late(r, r.get("_ref_date")) or 0), int(r.get("OS") or 0)))
        self.delay_rows = overdue

        summary = tk.Frame(self.content, bg=c["bg"]); summary.pack(fill="x", pady=(0, 14))
        self._card(summary, "O.S. atrasadas", str(len(overdue)), c["red"], "Identificadas nos PDFs da fila")
        classified = sum(1 for r in overdue if r.get("Motivo Atraso"))
        self._card(summary, "Classificadas", str(classified), c["green"], "Motivo já registrado")
        self._card(summary, "Sem motivo", str(len(overdue)-classified), c["yellow"], "Ainda precisam de análise")

        panel = tk.Frame(self.content, bg=c["surface"], highlightbackground=c["border"], highlightthickness=1)
        panel.pack(fill="both", expand=True)
        left = tk.Frame(panel, bg=c["surface"]); left.pack(side="left", fill="both", expand=True, padx=(10,0), pady=10)
        cols=("fase","os","cliente","atraso","criticidade","motivo")
        tree=ttk.Treeview(left, columns=cols, show="headings", selectmode="browse", height=16)
        self.delay_tree=tree
        heads={"fase":"Fase","os":"OS","cliente":"Cliente","atraso":"Atraso","criticidade":"Criticidade","motivo":"Motivo"}
        widths={"fase":70,"os":70,"cliente":220,"atraso":70,"criticidade":95,"motivo":180}
        for col in cols:
            tree.heading(col,text=heads[col]); tree.column(col,width=widths[col],anchor="center" if col in ("fase","os","atraso") else "w")
        sb=ttk.Scrollbar(left,orient="vertical",command=tree.yview,style="Vertical.TScrollbar"); tree.configure(yscrollcommand=sb.set)
        tree.pack(side="left",fill="both",expand=True); sb.pack(side="right",fill="y")
        for i,r in enumerate(overdue):
            ref = r.get("_ref_date") or date.today()
            late=_days_late(r, ref)
            tree.insert("","end",iid=str(i),values=(r.get("Fase Num","-"),r.get("OS","-"),r.get("Cliente") or "N/I", f"{late or 0} d", r.get("Criticidade","-"), r.get("Motivo Atraso") or "Não informado"))
        tree.bind("<<TreeviewSelect>>", self._on_delay_select)

        right=tk.Frame(panel,bg=c["surface"],width=420); right.pack(side="right",fill="y",padx=12,pady=10); right.pack_propagate(False)
        self.delay_detail=right
        tk.Label(right,text="Detalhes do atraso",bg=c["surface"],fg=c["text"],font=("Segoe UI",13,"bold")).pack(anchor="w",pady=(4,2))
        self.delay_info=tk.Label(right,text="Selecione uma O.S.",bg=c["surface"],fg=c["muted"],font=("Segoe UI",9),justify="left",wraplength=390); self.delay_info.pack(anchor="w",pady=(0,10))
        tk.Label(right,text="Motivo do atraso",bg=c["surface"],fg=c["muted"],font=("Segoe UI",9,"bold")).pack(anchor="w")
        self.delay_reason_widget=tk.Text(right,bg=c["surface2"],fg=c["text"],insertbackground=c["text"],relief="flat",height=3,wrap="word",undo=True)
        self.delay_reason_widget.pack(fill="x",pady=(3,9))
        fields=[("Observação","observacao",3),("Responsável","responsavel",1),("Ação necessária","acao",2),("Prazo da ação","prazo",1)]
        self.delay_vars={}
        for label,key,height in fields:
            tk.Label(right,text=label,bg=c["surface"],fg=c["muted"],font=("Segoe UI",9,"bold")).pack(anchor="w")
            var=tk.StringVar(); self.delay_vars[key]=var
            ent=tk.Text(right,bg=c["surface2"],fg=c["text"],insertbackground=c["text"],relief="flat",height=height,wrap="word")
            ent.pack(fill="x",pady=(3,8))
            setattr(self, f"delay_{key}_widget", ent)
        tk.Button(right,text="💾 Salvar classificação",command=self._save_selected_delay,bg=c["purple"],fg="#FFFFFF",activebackground=c["purple2"],relief="flat",borderwidth=0,font=("Segoe UI",10,"bold"),padx=14,pady=9).pack(fill="x",pady=(4,6))
        tk.Label(right,text="Os dados ficam salvos localmente e serão reutilizados nos próximos relatórios.",bg=c["surface"],fg=c["muted"],font=("Segoe UI",8),wraplength=390,justify="left").pack(anchor="w")

    def _on_delay_select(self, event=None):
        if not hasattr(self, "delay_tree"): return
        sel=self.delay_tree.selection()
        if not sel: return
        row=self.delay_rows[int(sel[0])]
        self.delay_current_index=int(sel[0])
        ref=row.get("_ref_date") or date.today()
        late=_days_late(row, ref)
        self.delay_info.config(text=f"O.S.: {row.get('OS','-')}\nFase: {row.get('Fase Num','-')} - {row.get('Fase Nome','-')}\nCliente: {row.get('Cliente') or 'N/I'}\nAtraso: {late or 0} dias | Previsão: {fmt_date(row.get('Prev. Entrega'))}\nCriticidade: {row.get('Criticidade','-')}")
        self.delay_reason_widget.delete("1.0","end")
        self.delay_reason_widget.insert("1.0", row.get("Motivo Atraso") or "")
        for key in self.delay_vars:
            w=getattr(self, f"delay_{key}_widget")
            w.delete("1.0","end"); w.insert("1.0", row.get({"observacao":"Observação Atraso","responsavel":"Responsável Atraso","acao":"Ação Atraso","prazo":"Prazo Ação"}[key]) or "")

    def _save_selected_delay(self):
        if not hasattr(self, "delay_current_index"): return
        row=self.delay_rows[self.delay_current_index]
        data={
            "motivo": self.delay_reason_widget.get("1.0","end").strip(),
            "observacao": self.delay_observacao_widget.get("1.0","end").strip(),
            "responsavel": self.delay_responsavel_widget.get("1.0","end").strip(),
            "acao": self.delay_acao_widget.get("1.0","end").strip(),
            "prazo": self.delay_prazo_widget.get("1.0","end").strip(),
        }
        self.delay_reasons[delay_key(row)] = data
        self._save_delay_reasons()
        row["Motivo Atraso"]=data["motivo"]; row["Observação Atraso"]=data["observacao"]; row["Responsável Atraso"]=data["responsavel"]; row["Ação Atraso"]=data["acao"]; row["Prazo Ação"]=data["prazo"]
        self.delay_tree.item(str(self.delay_current_index), values=(row.get("Fase Num","-"),row.get("OS","-"),row.get("Cliente") or "N/I",f"{_days_late(row,row.get('_ref_date') or date.today()) or 0} d",row.get("Criticidade","-"),row.get("Motivo Atraso") or "Não informado"))
        self.status.set(f"Atraso da O.S. {row.get('OS')} salvo.")

    def _build_config_page(self):
        c=self.palette
        self._page_title("Configurações", "Defina o comportamento padrão do gerador.", "⚙")
        card=tk.Frame(self.content,bg=c["surface"],highlightbackground=c["border"],highlightthickness=1); card.pack(fill="x",pady=(0,14))
        tk.Label(card,text="Pasta de saída padrão",bg=c["surface"],fg=c["text"],font=("Segoe UI",11,"bold")).pack(anchor="w",padx=18,pady=(16,4))
        tk.Label(card,text="Os PDFs gerados serão salvos nesta pasta.",bg=c["surface"],fg=c["muted"],font=("Segoe UI",9)).pack(anchor="w",padx=18)
        line=tk.Frame(card,bg=c["surface"]); line.pack(fill="x",padx=18,pady=14)
        tk.Entry(line,textvariable=self.output_dir,bg=c["surface2"],fg=c["text"],insertbackground=c["text"],relief="flat",font=("Segoe UI",9)).pack(side="left",fill="x",expand=True,ipady=7)
        tk.Button(line,text="📁 Escolher...",command=self.choose_output,bg=c["surface2"],fg=c["text"],relief="flat",borderwidth=0,font=("Segoe UI",9),padx=12,pady=7).pack(side="left",padx=(8,0))

        opts=tk.Frame(self.content,bg=c["surface"],highlightbackground=c["border"],highlightthickness=1); opts.pack(fill="x",pady=(0,14))
        tk.Label(opts,text="Preferências",bg=c["surface"],fg=c["text"],font=("Segoe UI",11,"bold")).pack(anchor="w",padx=18,pady=(16,10))
        tk.Checkbutton(opts,text="Abrir o PDF automaticamente após gerar",variable=self.auto_open,bg=c["surface"],activebackground=c["surface"],fg="#D8E0EC",selectcolor=c["surface2"],activeforeground=c["text"],font=("Segoe UI",10),anchor="w").pack(anchor="w",padx=18,pady=5)
        modes=tk.Frame(opts,bg=c["surface"]); modes.pack(fill="x",padx=18,pady=(8,16))
        tk.Label(modes,text="Modo padrão:",bg=c["surface"],fg=c["muted"],font=("Segoe UI",9)).pack(side="left")
        ttk.Radiobutton(modes,text="Completo",variable=self.report_mode,value="completo").pack(side="left",padx=(12,0))
        ttk.Radiobutton(modes,text="Resumido",variable=self.report_mode,value="resumido").pack(side="left",padx=(12,0))
        note=tk.Frame(self.content,bg="#06382B",highlightbackground="#138A60",highlightthickness=1); note.pack(fill="x",pady=(2,0))
        tk.Label(note,text="✓",bg="#06382B",fg="#4ADE80",font=("Segoe UI",14,"bold")).pack(side="left",padx=(16,10),pady=10)
        tk.Label(note,text="As preferências ficam ativas enquanto o programa estiver aberto.",bg="#06382B",fg="#E7FFF2",font=("Segoe UI",9)).pack(side="left")

    def _build_about_page(self):
        c=self.palette
        self._page_title("Sobre", "Informações da ferramenta e do processamento.", "ⓘ")
        card=tk.Frame(self.content,bg=c["surface"],highlightbackground=c["border"],highlightthickness=1); card.pack(fill="x")
        tk.Label(card,text="Gerador de Relatórios PCP - ESA",bg=c["surface"],fg=c["text"],font=("Segoe UI",18,"bold")).pack(anchor="w",padx=20,pady=(20,5))
        tk.Label(card,text="Versão 9.8.0 | Multifases + consultas PCP + gestão de atrasos + criticidade + histórico",bg=c["surface"],fg=c["purple2"],font=("Segoe UI",10,"bold")).pack(anchor="w",padx=20)
        for text in [
            "• Lê um ou vários PDFs ESA e identifica automaticamente as fases.",
            "• Mantém cada O.S. vinculada à sua fase de origem.",
            "• Gera relatório completo ou resumido.",
            "• Mantém histórico dos relatórios gerados.",
            "• Permite definir a pasta de saída e abertura automática dos PDFs.",
            "• Desenvolvido para apoiar o acompanhamento operacional do PCP.",
        ]:
            tk.Label(card,text=text,bg=c["surface"],fg="#D2DCE9",font=("Segoe UI",10),anchor="w").pack(anchor="w",padx=24,pady=5)
        tk.Label(card,text="Tecnologias: Python + Tkinter + PyMuPDF + ReportLab",bg=c["surface"],fg=c["muted"],font=("Segoe UI",9)).pack(anchor="w",padx=20,pady=(16,20))

    def _load_delay_reasons(self):
        try:
            if self.delay_file.exists():
                data = json.loads(self.delay_file.read_text(encoding="utf-8"))
                self.delay_reasons = data if isinstance(data, dict) else {}
        except Exception:
            self.delay_reasons = {}

    def _save_delay_reasons(self):
        try:
            self.delay_file.parent.mkdir(parents=True, exist_ok=True)
            self.delay_file.write_text(json.dumps(self.delay_reasons, ensure_ascii=False, indent=2), encoding="utf-8")
        except Exception:
            pass

    def _apply_delay_data(self, report):
        for row in report.get("rows", []):
            data = self.delay_reasons.get(delay_key(row), {})
            if isinstance(data, dict):
                row["Motivo Atraso"] = data.get("motivo", "")
                row["Observação Atraso"] = data.get("observacao", "")
                row["Responsável Atraso"] = data.get("responsavel", "")
                row["Ação Atraso"] = data.get("acao", "")
                row["Prazo Ação"] = data.get("prazo", "")
        return report

    def _read_queued_reports(self):
        reports = []
        for path in self.files:
            try:
                report = enrich(extract_report(path))
                self._apply_delay_data(report)
                for row in report.get("rows", []):
                    row["_ref_date"] = report.get("ref_date")
                reports.append(report)
            except Exception:
                continue
        return reports

    def _load_history(self):
        try:
            if self.history_file.exists():
                data=json.loads(self.history_file.read_text(encoding="utf-8"))
                self.generated_history=data if isinstance(data,list) else []
        except Exception:
            self.generated_history=[]

    def _save_history(self):
        try:
            self.history_file.parent.mkdir(parents=True,exist_ok=True)
            self.history_file.write_text(json.dumps(self.generated_history[-100:],ensure_ascii=False,indent=2),encoding="utf-8")
        except Exception:
            pass

    def _add_history(self, pdf_path, report):
        phases=report.get("phases",[])
        phase_text=" | ".join(f"{p.get('fase_num')} - {p.get('fase_nome')}" for p in phases) if phases else f"{report.get('fase_num','-')} - {report.get('fase_nome','-')}"
        rows = report.get("rows", [])
        time_values = [r.get("Dias na Fase") for r in rows if r.get("Dias na Fase") is not None]
        self.generated_history.append({
            "date":datetime.now().strftime("%d/%m/%Y %H:%M"),
            "file":os.path.basename(pdf_path),
            "mode":"Completo" if self.report_mode.get()=="completo" else "Resumido",
            "phases":phase_text,
            "total":len(rows),
            "overdue":sum(1 for r in rows if r.get("Status") == "Atrasada"),
            "critical":sum(1 for r in rows if r.get("Criticidade") == "Crítica"),
            "avg_phase":round(sum(time_values)/len(time_values),1) if time_values else 0,
            "path":os.path.dirname(os.path.abspath(pdf_path)),
            "pdf_path":os.path.abspath(pdf_path),
        })
        self._save_history()

    def _open_path(self, path):
        if not path: return
        try:
            os.startfile(os.path.abspath(path))
        except Exception:
            try:
                if sys.platform == "darwin": subprocess.Popen(["open", path])
                else: subprocess.Popen(["xdg-open", path])
            except Exception: pass

    def open_history_selected(self):
        if not hasattr(self,"history_tree"): return
        sel=self.history_tree.selection()
        if not sel: return
        item=self.generated_history[::-1][int(sel[0])]
        path=item.get("pdf_path","")
        if os.path.exists(path): self._open_path(path)
        else: messagebox.showwarning("Arquivo não encontrado", "O PDF não está mais no local salvo no histórico.")

    def open_history_folder(self):
        if not hasattr(self,"history_tree"): return
        sel=self.history_tree.selection()
        if not sel: return
        item=self.generated_history[::-1][int(sel[0])]
        folder=item.get("path","")
        if os.path.isdir(folder): self._open_path(folder)
        else: messagebox.showwarning("Pasta não encontrada", "A pasta registrada no histórico não existe mais.")

    def _invalidate_pages(self, *names):
        """Invalida páginas em cache que dependem do estado atualizado da aplicação."""
        for name in names:
            frame = self.pages.pop(name, None) if hasattr(self, "pages") else None
            if frame is not None:
                try:
                    frame.destroy()
                except Exception:
                    pass

    def refresh_history_page(self):
        self._invalidate_pages("historico")
        self.show_page("historico")

    def clear_history(self):
        if not self.generated_history: return
        if messagebox.askyesno("Limpar histórico", "Remover todos os registros do histórico? Os PDFs não serão apagados."):
            self.generated_history=[]; self._save_history(); self.refresh_history_page()

    def set_report_mode(self, mode):
        """Seleciona o modelo de PDF a ser gerado."""
        self.report_mode.set(mode)
        c = self.palette
        if mode == "resumido":
            self.mode_complete_btn.config(bg=c["surface2"], fg="#C6D0DF", text="○  Relatório completo")
            self.mode_summary_btn.config(bg=c["purple"], fg="#FFFFFF", text="●  Relatório resumido")
            self.footer_title.config(text="Modo resumido selecionado: visão executiva + atrasadas + base completa.")
        else:
            self.mode_complete_btn.config(bg=c["purple"], fg="#FFFFFF", text="●  Relatório completo")
            self.mode_summary_btn.config(bg=c["surface2"], fg="#C6D0DF", text="○  Relatório resumido")
            self.footer_title.config(text="Modo completo selecionado: análise PCP detalhada.")

    def _rounded_frame(self, frame, color):
        # Approximation using highlight border, keeping native widgets reliable.
        frame.configure(bg=color)

    def _action_card(self, parent, row, col, icon, title, subtitle, accent,
                     command, side="left", primary=False):
        c = self.palette
        card = tk.Frame(
            parent, bg=c["surface2"] if not primary else c["purple"],
            width=260, height=112,
            highlightbackground=accent if not primary else "#9C7FFF",
            highlightthickness=1
        )
        card.pack(side=side, fill="both", expand=True, padx=(0, 12 if col < 3 else 0))
        card.pack_propagate(False)

        emoji_img = self._get_emoji_image(icon, "small24")
        if emoji_img is not None:
            icon_label = tk.Label(card, image=emoji_img)
            icon_label.configure(bg=card["bg"], bd=0, highlightthickness=0, cursor="hand2")
        else:
            fallback = {"gerar": "📄", "inicio": "🏠", "historico": "🕒",
                        "atrasos": "⏱", "config": "⚙", "sobre": "ℹ"}.get(icon, "•")
            icon_label = tk.Label(card, text=fallback, bg=card["bg"],
                                  fg="#A887FF" if not primary else "#FFFFFF",
                                  font=("Segoe UI Emoji", 18), cursor="hand2")
        icon_label.pack(side="left", padx=(17, 13))

        text = tk.Frame(card, bg=card["bg"])
        text.pack(side="left", fill="both", expand=True)
        tk.Label(
            text, text=title, bg=card["bg"], fg="#FFFFFF",
            font=("Segoe UI", 11, "bold")
        ).pack(anchor="w", pady=(28, 2))
        tk.Label(
            text, text=subtitle, bg=card["bg"], fg="#C2CEE0" if not primary else "#E9E2FF",
            font=("Segoe UI", 9)
        ).pack(anchor="w")

        tk.Label(
            card, text="›", bg=card["bg"], fg="#A9B7C9" if not primary else "#FFFFFF",
            font=("Segoe UI", 22)
        ).pack(side="right", padx=14)

        # Entire card responds to click.
        for widget in (card, text):
            widget.bind("<Button-1>", lambda e, cb=command: cb())
        for widget in card.winfo_children():
            if isinstance(widget, tk.Label):
                widget.bind("<Button-1>", lambda e, cb=command: cb())

    def _toggle_maximize(self):
        try:
            self.state("zoomed")
        except tk.TclError:
            pass

    def _update_queue_count(self):
        count = len(self.files)
        label = "arquivo" if count == 1 else "arquivos"
        if hasattr(self, "queue_badge"):
            self.queue_badge.config(text=f"{count} {label}")

    def add_files(self):
        paths = filedialog.askopenfilenames(
            title="Selecione os relatórios PDF",
            filetypes=[("Arquivos PDF", "*.pdf")]
        )
        if not paths:
            return

        existing = set(self.files)
        for p in paths:
            if p not in existing:
                self.files.append(p)

        self.refresh_tree()
        self.status.set(f"{len(self.files)} arquivo(s) na fila.")
        self.footer_title.config(text="Arquivos carregados. Revise a fila antes de gerar.")
        self._update_queue_count()
        self._invalidate_pages("inicio", "atrasos")

    def clear_files(self):
        self.files = []
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.status.set("Fila limpa.")
        self.footer_title.config(text="Tudo pronto para gerar seu relatório!")
        self._update_queue_count()
        self._invalidate_pages("inicio", "atrasos")

    def choose_output(self):
        d = filedialog.askdirectory(title="Escolha a pasta de saída")
        if d:
            self.output_dir.set(d)
            self.status.set(f"Pasta de saída definida: {d}")

    def refresh_tree(self):
        if not hasattr(self, "tree"):
            return

        for item in self.tree.get_children():
            self.tree.delete(item)

        query = self.search_var.get().lower().strip()
        shown = 0

        for idx, p in enumerate(self.files, start=1):
            if query and query not in os.path.basename(p).lower():
                continue
            shown += 1
            self.tree.insert(
                "",
                "end",
                values=(shown, os.path.basename(p), "Lendo...", "-", "-", "-", "Pronto", "◉  🗑")
            )


    def generate_all(self):
        if pymupdf is None or colors is None:
            messagebox.showerror(
                "Dependências",
                "Instale as dependências primeiro:\n\n"
                'python -m pip install "reportlab>=4.0" "PyMuPDF>=1.24"'
            )
            return

        if not self.files:
            messagebox.showwarning(
                "Nenhum PDF",
                "Adicione pelo menos um relatório PDF."
            )
            return

        outdir = self.output_dir.get().strip()
        if not outdir:
            messagebox.showwarning(
                "Pasta de saída",
                "Escolha uma pasta de saída."
            )
            return

        os.makedirs(outdir, exist_ok=True)

        for item in self.tree.get_children():
            self.tree.delete(item)

        generated = []
        errors = []

        for p in self.files:
            try:
                self.status.set(f"Processando: {os.path.basename(p)}")
                self.footer_title.config(text=f"Processando {os.path.basename(p)}...")
                self.update_idletasks()

                report = enrich(extract_report(p))
                self._apply_delay_data(report)
                pdf_path = generate_pdf(report, outdir, mode=self.report_mode.get())
                if report.get("multi_fase"):
                    phase = " / ".join(f"{ph['fase_num']} - {ph['fase_nome']}" for ph in report.get("phases", []))
                    phase = f"{len(report.get('phases', []))} fases: " + phase
                else:
                    phase = f"{report['fase_num']} - {report['fase_nome']}"
                collab_count = len(set(
                    r["Colaborador"] for r in report["rows"] if r["Colaborador"]
                ))
                late_count = sum(
                    1 for r in report["rows"] if r["Status"] == "Atrasada"
                )

                self.tree.insert(
                    "", "end",
                    values=(
                        len(generated) + 1,
                        os.path.basename(p),
                        phase,
                        len(report["rows"]),
                        late_count,
                        collab_count,
                        "Pronto",
                        "◉  🗑",
                    ),
                )
                generated.append(pdf_path)

                self._add_history(pdf_path, report)
                # Auto-open generated PDF, when enabled in settings.
                if self.auto_open.get():
                    self._open_path(pdf_path)

            except Exception as e:
                self.tree.insert(
                    "", "end",
                    values=(len(generated) + 1, os.path.basename(p), "-", "-", "-", "-", "ERRO", "—")
                )
                errors.append(f"{os.path.basename(p)}: {e}")

        self._invalidate_pages("inicio", "historico", "atrasos")

        if errors:
            self.status.set(
                f"{len(generated)} relatório(s) gerado(s). {len(errors)} com erro."
            )
            self.footer_title.config(text="Processamento concluído com avisos.")
            messagebox.showwarning(
                "Processamento concluído com avisos",
                "Alguns arquivos não foram processados:\n\n" + "\n".join(errors),
            )
        else:
            self.status.set(
                f"{len(generated)} relatório(s) gerado(s) e aberto(s) automaticamente."
            )
            self.footer_title.config(text="Relatórios gerados e abertos automaticamente.")


def main():
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
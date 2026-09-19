import calendar
import html
import io
import json
import re
from datetime import datetime

import openpyxl
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
from github import Github
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from ortools.sat.python import cp_model


DEFAULT_TURNOS = [
    {"Responsabilidade": "Radiologia Convencional", "Turno": "M3", "Início": "08:00", "Fim": "16:00"},
    {"Responsabilidade": "Radiologia Convencional", "Turno": "T", "Início": "16:00", "Fim": "00:00"},
    {"Responsabilidade": "Radiologia Convencional", "Turno": "M12", "Início": "08:00", "Fim": "14:00"},
    {"Responsabilidade": "Radiologia Convencional", "Turno": "T24", "Início": "14:00", "Fim": "20:00"},
    {"Responsabilidade": "Tomografia Computorizada", "Turno": "M3", "Início": "08:00", "Fim": "16:00"},
    {"Responsabilidade": "Tomografia Computorizada", "Turno": "T", "Início": "16:00", "Fim": "00:00"},
    {"Responsabilidade": "Tomografia Computorizada", "Turno": "M12", "Início": "08:00", "Fim": "14:00"},
    {"Responsabilidade": "Tomografia Computorizada", "Turno": "T24", "Início": "14:00", "Fim": "20:00"},
    {"Responsabilidade": "Ressonância Magnética", "Turno": "M12", "Início": "08:00", "Fim": "14:00"},
    {"Responsabilidade": "Ressonância Magnética", "Turno": "T24", "Início": "14:00", "Fim": "20:00"},
    {"Responsabilidade": "Ecografia", "Turno": "M3", "Início": "08:00", "Fim": "16:00"},
    {"Responsabilidade": "Ecografia", "Turno": "M75", "Início": "08:00", "Fim": "18:00"},
    {"Responsabilidade": "Ecografia", "Turno": "M12", "Início": "08:00", "Fim": "14:00"},
    {"Responsabilidade": "Ecografia", "Turno": "T24", "Início": "14:00", "Fim": "20:00"},
]

# Cores de texto vivas por responsabilidade e hex para Excel
RESPONSABILIDADE_CORES = {
    "Radiologia Convencional": {"fundo": "#ffffff", "texto": "#2563eb", "hex": "2563EB", "borda": "#d1d5db"},     # Azul
    "Tomografia Computorizada": {"fundo": "#ffffff", "texto": "#16a34a", "hex": "16A34A", "borda": "#d1d5db"},    # Verde
    "Ecografia": {"fundo": "#ffffff", "texto": "#dc2626", "hex": "DC2626", "borda": "#d1d5db"},                   # Vermelho
    "Ressonância Magnética": {"fundo": "#ffffff", "texto": "#ea580c", "hex": "EA580C", "borda": "#d1d5db"},      # Laranja
}

RESPONSABILIDADE_INDICADORES = {
    "Radiologia Convencional": "🔵",
    "Tomografia Computorizada": "🟢",
    "Ecografia": "🔴",
    "Ressonância Magnética": "🟠",
}

LIMITE_JORNADAS_12H_PADRAO = 31
HORAS_CONTRATO_SEMANAL_PADRAO = 35.0
HORAS_DIA_FERIAS_LICENCA = 7.0

st.set_page_config(page_title="Gestor de Escalas de Trabalho", layout="wide")

# =============================================================================
# PERSISTÊNCIA VIA GITHUB API
# =============================================================================
def guardar_estado_no_github():
    if "GITHUB_TOKEN" not in st.secrets or "GITHUB_REPO" not in st.secrets:
        st.sidebar.warning("Secrets do GitHub não configuradas.")
        return False
    try:
        g = Github(st.secrets["GITHUB_TOKEN"])
        repo = g.get_repo(st.secrets["GITHUB_REPO"])
        
        turnos_dict = st.session_state.turnos.to_dict(orient="records") if isinstance(st.session_state.turnos, pd.DataFrame) else []
        df_nec_dict = st.session_state.df_necessidades.to_dict(orient="split") if isinstance(st.session_state.get("df_necessidades"), pd.DataFrame) else None
        pref_dict = st.session_state.preferencias_turnos.to_dict(orient="split") if isinstance(st.session_state.get("preferencias_turnos"), pd.DataFrame) else None

        dados = {
            "trabalhadores": st.session_state.trabalhadores,
            "trabalhadores_ativos": st.session_state.trabalhadores_ativos,
            "preferencias_jornadas_12h": st.session_state.preferencias_jornadas_12h,
            "limites_jornadas_12h": st.session_state.limites_jornadas_12h,
            "horas_contrato_semanal": st.session_state.horas_contrato_semanal,
            "competencias": st.session_state.competencias,
            "turnos": turnos_dict,
            "df_necessidades": df_nec_dict,
            "preferencias_turnos": pref_dict,
        }
        
        conteudo = json.dumps(dados, ensure_ascii=False, indent=2)
        caminho_ficheiro = "estado_escala.json"
        
        try:
            ficheiro_existente = repo.get_contents(caminho_ficheiro)
            repo.update_file(ficheiro_existente.path, "Atualizar estado da escala", conteudo, ficheiro_existente.sha)
        except Exception:
            repo.create_file(caminho_ficheiro, "Criar ficheiro de estado da escala", conteudo)
            
        st.sidebar.success("Guardado no GitHub com sucesso!")
        return True
    except Exception as e:
        st.sidebar.error(f"Erro ao guardar no GitHub: {e}")
        return False


def carregar_estado_do_github():
    if "GITHUB_TOKEN" not in st.secrets or "GITHUB_REPO" not in st.secrets:
        return False
    try:
        g = Github(st.secrets["GITHUB_TOKEN"])
        repo = g.get_repo(st.secrets["GITHUB_REPO"])
        ficheiro = repo.get_contents("estado_escala.json")
        dados = json.loads(ficheiro.decoded_content.decode("utf-8"))

        st.session_state.trabalhadores = dados.get("trabalhadores", ["Ana Silva", "Bruno Santos", "Carla Costa"])
        st.session_state.trabalhadores_ativos = dados.get("trabalhadores_ativos", {})
        st.session_state.preferencias_jornadas_12h = dados.get("preferencias_jornadas_12h", {})
        st.session_state.limites_jornadas_12h = dados.get("limites_jornadas_12h", {})
        st.session_state.horas_contrato_semanal = dados.get("horas_contrato_semanal", {})
        st.session_state.competencias = dados.get("competencias", {})
        
        if dados.get("turnos"):
            st.session_state.turnos = pd.DataFrame(dados["turnos"])
        if dados.get("df_necessidades"):
            raw_nec = dados["df_necessidades"]
            st.session_state.df_necessidades = pd.DataFrame(raw_nec["data"], index=raw_nec["index"], columns=raw_nec["columns"])
        if dados.get("preferencias_turnos"):
            raw_pref = dados["preferencias_turnos"]
            st.session_state.preferencias_turnos = pd.DataFrame(raw_pref["data"], index=raw_pref["index"], columns=raw_pref["columns"])
            
        return True
    except Exception:
        return False


if "estado_carregado_github" not in st.session_state:
    carregar_estado_do_github()
    st.session_state.estado_carregado_github = True

with st.sidebar:
    st.title("💾 Controlo de Dados")
    if st.button("☁️ Guardar Tudo no GitHub", use_container_width=True, type="primary"):
        guardar_estado_no_github()
    if st.button("🔄 Recarregar do GitHub", use_container_width=True):
        if carregar_estado_do_github():
            st.success("Dados recarregados!")
            st.rerun()

# =============================================================================
# INICIALIZAÇÃO DE ESTADOS BASE
# =============================================================================
if "trabalhadores" not in st.session_state:
    st.session_state.trabalhadores = ["Ana Silva", "Bruno Santos", "Carla Costa", "Daniel Rocha", "Eduarda Lima"]

if "trabalhadores_ativos" not in st.session_state:
    st.session_state.trabalhadores_ativos = {nome: True for nome in st.session_state.trabalhadores}

if "preferencias_jornadas_12h" not in st.session_state:
    st.session_state.preferencias_jornadas_12h = {nome: False for nome in st.session_state.trabalhadores}

if "limites_jornadas_12h" not in st.session_state:
    st.session_state.limites_jornadas_12h = {nome: LIMITE_JORNADAS_12H_PADRAO for nome in st.session_state.trabalhadores}

if "horas_contrato_semanal" not in st.session_state:
    st.session_state.horas_contrato_semanal = {nome: HORAS_CONTRATO_SEMANAL_PADRAO for nome in st.session_state.trabalhadores}

if "turnos" not in st.session_state:
    st.session_state.turnos = pd.DataFrame(DEFAULT_TURNOS)


def normalizar_turnos(df):
    colunas = ["Responsabilidade", "Turno", "Início", "Fim"]
    resultado = df.reindex(columns=colunas).copy()
    for coluna in colunas:
        resultado[coluna] = resultado[coluna].fillna("").astype(str).str.strip()
    resultado = resultado[(resultado["Responsabilidade"] != "") & (resultado["Turno"] != "")]
    return resultado.drop_duplicates(subset=["Responsabilidade", "Turno"], keep="last").reset_index(drop=True)


def assinatura_turnos(df):
    return tuple(tuple(linha) for linha in normalizar_turnos(df).itertuples(index=False, name=None))


def opcoes_turnos_por_responsabilidade(df):
    resultado = {}
    for linha in normalizar_turnos(df).itertuples(index=False):
        resultado.setdefault(linha.Responsabilidade, []).append(linha.Turno)
    return resultado


def chave_ordenacao_turno(codigo):
    correspondencia = re.match(r"^([A-Za-z]+)(\d*)$", codigo.strip())
    if not correspondencia:
        return (2, codigo.lower(), -1)
    tipo, numero = correspondencia.groups()
    ordem_tipo = {"M": 0, "T": 1}.get(tipo.upper(), 2)
    ordem_numero = int(numero) if numero else -1
    return (ordem_tipo, tipo.lower(), ordem_numero)


def turnos_ordenados(df):
    return sorted(
        normalizar_turnos(df).itertuples(index=False),
        key=lambda linha: (chave_ordenacao_turno(linha.Turno), linha.Responsabilidade.lower()),
    )


def codigos_turno_ordenados(df):
    codigos = []
    for slot in turnos_ordenados(df):
        if slot.Turno not in codigos:
            codigos.append(slot.Turno)
    return codigos


def chave_coluna_turno(slot):
    return f"{slot.Responsabilidade} | {slot.Turno}"


def indicador_responsabilidade(responsabilidade):
    return RESPONSABILIDADE_INDICADORES.get(responsabilidade, "⚪")


def e_fim_de_semana(dia, mes, ano):
    return datetime(ano, mes, dia).weekday() >= 5


def contar_dias_uteis_mes(mes, ano):
    """Calcula quantos dias úteis (Segunda a Sexta-feira) existem no mês."""
    num_dias = calendar.monthrange(ano, mes)[1]
    dias_uteis = 0
    for dia in range(1, num_dias + 1):
        if not e_fim_de_semana(dia, mes, ano):
            dias_uteis += 1
    return dias_uteis


def calcular_duracao_turno_horas(inicio, fim):
    try:
        h_inicio = datetime.strptime(inicio, "%H:%M")
        h_fim = datetime.strptime(fim, "%H:%M")
        if h_fim <= h_inicio:
            duracao = (datetime.strptime("24:00", "%H:%M") - h_inicio) + (h_fim - datetime.strptime("00:00", "%H:%M"))
        else:
            duracao = h_fim - h_inicio
        return round(duracao.total_seconds() / 3600.0, 2)
    except ValueError:
        return 8.0


def dia_a_partir_do_rotulo(rotulo):
    correspondencia = re.match(r"^\D*(\d+)", str(rotulo))
    return int(correspondencia.group(1)) if correspondencia else None


def estilizar_fins_de_semana(df, mes, ano, dias_nas_colunas):
    """Aplica estilo padronizado de bordas e fundo laranja para fins de semana."""
    preenchimento = "#fff7ed"
    borda_fim_de_semana = "#f59e0b"
    estilo_celula_fim_de_semana = f"background-color: {preenchimento}; border: 2px solid {borda_fim_de_semana};"

    if dias_nas_colunas:
        def estilo_coluna(coluna):
            dia = dia_a_partir_do_rotulo(coluna.name)
            fim_de_semana = dia is not None and e_fim_de_semana(dia, mes, ano)
            return [estilo_celula_fim_de_semana if fim_de_semana else "" for _ in coluna]

        def estilo_cabecalho_colunas(indice):
            return [
                f"font-weight: bold; background-color: #fff7ed; border-bottom: 3px solid {borda_fim_de_semana}; color: #9a3412;"
                if (dia_a_partir_do_rotulo(rotulo) is not None and e_fim_de_semana(dia_a_partir_do_rotulo(rotulo), mes, ano))
                else ""
                for rotulo in indice
            ]

        return df.style.apply(estilo_coluna, axis=0).apply_index(estilo_cabecalho_colunas, axis=1)

    def estilo_linha(linha):
        dia = dia_a_partir_do_rotulo(linha.name)
        fim_de_semana = dia is not None and e_fim_de_semana(dia, mes, ano)
        return [estilo_celula_fim_de_semana if fim_de_semana else "" for _ in linha]

    def estilo_indice_linhas(indice):
        return [
            f"font-weight: bold; background-color: #fff7ed; border-left: 4px solid {borda_fim_de_semana}; color: #9a3412;"
            if (dia_a_partir_do_rotulo(rotulo) is not None and e_fim_de_semana(dia_a_partir_do_rotulo(rotulo), mes, ano))
            else ""
            for rotulo in indice
        ]

    return df.style.apply(estilo_linha, axis=1).apply_index(estilo_indice_linhas, axis=0)


def renderizar_tabela_escala_html(df_resultado, df_meta_resps, mes, ano):
    """Tabela final HTML padronizada com CSS de impressão nativo (@media print)."""
    html_code = f"""
    <style>
        @media print {{
            @page {{
                size: A4 landscape;
                margin: 8mm;
            }}
            body * {{
                visibility: hidden;
            }}
            .printable-area, .printable-area * {{
                visibility: visible;
            }}
            .printable-area {{
                position: absolute;
                left: 0;
                top: 0;
                width: 100%;
            }}
            .no-print {{
                display: none !important;
            }}
            .escala-table {{
                font-size: 0.75rem !important;
            }}
            .escala-table th, .escala-table td {{
                padding: 4px 3px !important;
            }}
        }}
        .escala-table-container {{
            overflow-x: auto;
            margin-top: 10px;
            margin-bottom: 25px;
        }}
        .escala-table {{
            border-collapse: collapse;
            width: 100%;
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
            font-size: 0.88rem;
            text-align: center;
        }}
        .escala-table th, .escala-table td {{
            border: 1px solid #e5e7eb;
            padding: 8px 6px;
            white-space: nowrap;
        }}
        .escala-table th {{
            background-color: #f9fafb;
            font-weight: 600;
            color: #374151;
        }}
        .escala-table th.fds-header {{
            background-color: #fff7ed;
            border-bottom: 3px solid #f59e0b;
            color: #9a3412;
            font-weight: bold;
        }}
        .escala-table td.fds-cell {{
            background-color: #fffdfa;
        }}
        .escala-table td.nome-col {{
            text-align: left;
            font-weight: 600;
            background-color: #f9fafb;
            position: sticky;
            left: 0;
            z-index: 1;
            border-right: 2px solid #d1d5db;
        }}
        .txt-f {{ color: #9ca3af; font-weight: normal; }}
        .txt-l {{ color: #d97706; font-weight: bold; }}
        .res-col {{ font-weight: 600; background-color: #f3f4f6; }}
    </style>

    <div class="printable-area">
    <h2 style="margin-bottom: 5px;">Escala de Trabalho Mensal - {calendar.month_name[mes]} / {ano}</h2>
    <div class="escala-table-container">
    <table class="escala-table">
        <thead>
            <tr>
                <th class="nome-col">Trabalhador</th>
    """

    for col in df_resultado.columns:
        dia = dia_a_partir_do_rotulo(col)
        if dia is not None and e_fim_de_semana(dia, mes, ano):
            html_code += f'<th class="fds-header">{col}</th>'
        else:
            html_code += f'<th>{col}</th>'
    html_code += "</tr></thead><tbody>"

    for trab in df_resultado.index:
        html_code += f'<tr><td class="nome-col">{html.escape(str(trab))}</td>'
        for col in df_resultado.columns:
            val = df_resultado.loc[trab, col]
            dia = dia_a_partir_do_rotulo(col)
            is_fds = dia is not None and e_fim_de_semana(dia, mes, ano)
            td_class = ' class="fds-cell"' if is_fds else ''

            if col.isdigit():
                meta_item = df_meta_resps.loc[trab, col]
                if isinstance(meta_item, tuple):
                    cods, resps = meta_item
                    spans = []
                    for c_code, r_resp in zip(cods, resps):
                        cor = RESPONSABILIDADE_CORES.get(r_resp, {}).get("texto", "#111827")
                        spans.append(f'<span style="color:{cor}; font-weight:bold;">{html.escape(c_code)}</span>')
                    celula_html = " / ".join(spans)
                elif val in {"F", "L"}:
                    cls_txt = "txt-f" if val == "F" else "txt-l"
                    celula_html = f'<span class="{cls_txt}">{val}</span>'
                else:
                    cor = RESPONSABILIDADE_CORES.get(meta_item, {}).get("texto", "#111827")
                    celula_html = f'<span style="color:{cor}; font-weight:bold;">{html.escape(str(val))}</span>'

                html_code += f'<td{td_class}>{celula_html}</td>'
            else:
                html_code += f'<td class="res-col">{html.escape(str(val))}</td>'

        html_code += "</tr>"

    html_code += "tbody></table></div></div>"
    return html_code


def gerar_excel_escala_formatado(df_resultado, df_meta_resps, mes, ano):
    """Gera um ficheiro Excel (.xlsx) pré-formatado com cores por setor e destaque de fim de semana."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = f"Escala {calendar.month_name[mes]} {ano}"

    # Estilos Base
    fill_header = PatternFill(start_color="F3F4F6", end_color="F3F4F6", fill_type="solid")
    fill_header_fds = PatternFill(start_color="FFF7ED", end_color="FFF7ED", fill_type="solid")
    fill_cell_fds = PatternFill(start_color="FFFDF2", end_color="FFFDF2", fill_type="solid")
    fill_totals = PatternFill(start_color="F9FAFB", end_color="F9FAFB", fill_type="solid")

    font_header = Font(name="Segoe UI", size=10, bold=True, color="374151")
    font_header_fds = Font(name="Segoe UI", size=10, bold=True, color="9A3412")
    font_trab = Font(name="Segoe UI", size=10, bold=True, color="111827")
    font_f = Font(name="Segoe UI", size=10, bold=False, color="9CA3AF")
    font_l = Font(name="Segoe UI", size=10, bold=True, color="D97706")
    font_totals = Font(name="Segoe UI", size=10, bold=True, color="1F2937")

    thin_border = Border(
        left=Side(style="thin", color="E5E7EB"),
        right=Side(style="thin", color="E5E7EB"),
        top=Side(style="thin", color="E5E7EB"),
        bottom=Side(style="thin", color="E5E7EB")
    )
    align_center = Alignment(horizontal="center", vertical="center")
    align_left = Alignment(horizontal="left", vertical="center")

    # Headers
    headers = ["Trabalhador"] + list(df_resultado.columns)
    ws.append(headers)

    for col_idx, col_name in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_idx)
        dia = dia_a_partir_do_rotulo(col_name)
        if dia is not None and e_fim_de_semana(dia, mes, ano):
            cell.fill = fill_header_fds
            cell.font = font_header_fds
        else:
            cell.fill = fill_header
            cell.font = font_header
        cell.alignment = align_center if col_idx > 1 else align_left
        cell.border = thin_border

    # Data
    for row_idx, trab in enumerate(df_resultado.index, 2):
        ws.cell(row=row_idx, column=1, value=str(trab)).font = font_trab
        ws.cell(row=row_idx, column=1).alignment = align_left
        ws.cell(row=row_idx, column=1).border = thin_border

        for col_idx, col_name in enumerate(df_resultado.columns, 2):
            val = df_resultado.loc[trab, col_name]
            cell = ws.cell(row=row_idx, column=col_idx, value=str(val))
            cell.alignment = align_center
            cell.border = thin_border

            dia = dia_a_partir_do_rotulo(col_name)
            is_fds = dia is not None and e_fim_de_semana(dia, mes, ano)
            if is_fds:
                cell.fill = fill_cell_fds

            if col_name.isdigit():
                meta_item = df_meta_resps.loc[trab, col_name]
                if isinstance(meta_item, tuple):
                    cell.font = Font(name="Segoe UI", size=10, bold=True, color="111827")
                elif val == "F":
                    cell.font = font_f
                elif val == "L":
                    cell.font = font_l
                elif meta_item in RESPONSABILIDADE_CORES:
                    hex_color = RESPONSABILIDADE_CORES[meta_item]["hex"]
                    cell.font = Font(name="Segoe UI", size=10, bold=True, color=hex_color)
            else:
                cell.fill = fill_totals
                cell.font = font_totals

    # Widths
    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = openpyxl.utils.get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 3, 6)

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output


def competencias_iniciais(df):
    return {responsabilidade: list(turnos) for responsabilidade, turnos in opcoes_turnos_por_responsabilidade(df).items()}


def trabalhador_pode_fazer(nome, responsabilidade, turno):
    competencias = st.session_state.competencias.get(nome, {})
    return turno in competencias.get(responsabilidade, [])


def resumo_competencias_html(nome, numero=None):
    competencias = st.session_state.competencias.get(nome, {})
    etiquetas = []
    for responsabilidade, turnos in competencias.items():
        if not turnos:
            continue
        cores = RESPONSABILIDADE_CORES.get(responsabilidade, {"fundo": "#ffffff", "texto": "#374151", "borda": "#d1d5db"})
        responsabilidade_segura = html.escape(responsabilidade)
        turnos_seguro = html.escape(", ".join(turnos))
        etiquetas.append(
            f"<span style='display:inline-block; margin:3px 5px 3px 0; padding:4px 9px; border-radius:999px; background:#f9fafb; border:1px solid #d1d5db; color:{cores['texto']}; font-size:0.82rem; line-height:1.25;'><strong>{responsabilidade_segura}</strong>: {turnos_seguro}</span>"
        )

    nome_seguro = html.escape(nome)
    titulo = f"{numero}. {nome_seguro}" if numero is not None else nome_seguro
    if not etiquetas:
        etiquetas.append("<span style='display:inline-block; margin-top:4px; padding:4px 9px; border-radius:999px; background:#f3f4f6; color:#6b7280; font-size:0.82rem;'>Sem responsabilidades atribuídas</span>")
    return f"<div style='margin-bottom:8px;'><strong>{titulo}</strong><br>{''.join(etiquetas)}</div>"


def turno_atravessa_meia_noite(inicio, fim):
    try:
        hora_inicio = datetime.strptime(inicio, "%H:%M")
        hora_fim = datetime.strptime(fim, "%H:%M")
        return hora_fim <= hora_inicio
    except ValueError:
        return False


def e_turno_de_jornada_alargada(codigo):
    return codigo in {"M12", "T24"}


if "competencias" not in st.session_state:
    competencias_padrao = competencias_iniciais(st.session_state.turnos)
    st.session_state.competencias = {nome: dict(competencias_padrao) for nome in st.session_state.trabalhadores}
else:
    for nome in st.session_state.trabalhadores:
        st.session_state.competencias.setdefault(nome, {})

if "novo_trabalhador_form_version" not in st.session_state:
    st.session_state.novo_trabalhador_form_version = 0


st.title("🗓️ Gestor Inteligente de Escalas de Trabalho")

tabs = st.tabs(["1. Gestão da Equipa", "2. Responsabilidades e Turnos", "3. Necessidades Mensais", "4. Indisponibilidades", "5. Gerar Escala"])

# -----------------------------------------------------------------------------
# TAB 1: GESTÃO DA EQUIPA
# -----------------------------------------------------------------------------
with tabs[0]:
    st.header("Gestão de Trabalhadores")
    turnos_por_responsabilidade = opcoes_turnos_por_responsabilidade(st.session_state.turnos)
    responsabilidades_disponiveis = list(turnos_por_responsabilidade)
    st.caption("Associe a cada trabalhador as responsabilidades, contrato de horas e turnos que está habilitado a fazer.")
    col1, col2 = st.columns([1, 2])

    with col1:
        st.subheader("Adicionar Colaborador")
        form_version = st.session_state.novo_trabalhador_form_version
        novo_nome = st.text_input("Nome do Trabalhador", key=f"novo_trabalhador_nome_{form_version}")
        novas_horas_contrato = st.number_input("Horas de Contrato Semanal", min_value=1.0, max_value=60.0, value=HORAS_CONTRATO_SEMANAL_PADRAO, step=0.5, key=f"novo_trabalhador_horas_{form_version}")
        novas_responsabilidades = st.multiselect("Responsabilidades", responsabilidades_disponiveis, key=f"novo_trabalhador_responsabilidades_{form_version}")
        nova_preferencia_12h = st.checkbox("Prefere jornadas de 12 horas (M12 + T24)", value=False, key=f"novo_trabalhador_preferencia_12h_{form_version}")
        novo_limite_12h = st.number_input("Limite mensal de jornadas de 12 horas", min_value=0, max_value=31, value=LIMITE_JORNADAS_12H_PADRAO, step=1, help="Defina 0 se o trabalhador não puder fazer jornadas M12 + T24.", key=f"novo_trabalhador_limite_12h_{form_version}")
        novas_competencias = {}
        for responsabilidade in novas_responsabilidades:
            novas_competencias[responsabilidade] = st.multiselect(f"Turnos de {responsabilidade}", turnos_por_responsabilidade[responsabilidade], key=f"novo_trabalhador_turnos_{form_version}_{responsabilidade}")

        if st.button("Adicionar"):
            nome_normalizado = novo_nome.strip()
            if nome_normalizado and nome_normalizado not in st.session_state.trabalhadores:
                st.session_state.trabalhadores.append(nome_normalizado)
                st.session_state.trabalhadores_ativos[nome_normalizado] = True
                st.session_state.horas_contrato_semanal[nome_normalizado] = novas_horas_contrato
                st.session_state.preferencias_jornadas_12h[nome_normalizado] = nova_preferencia_12h
                st.session_state.limites_jornadas_12h[nome_normalizado] = novo_limite_12h
                st.session_state.competencias[nome_normalizado] = novas_competencias
                st.session_state.novo_trabalhador_form_version += 1
                guardar_estado_no_github()
                st.success(f"{nome_normalizado} adicionado com sucesso!")
                st.rerun()
            elif nome_normalizado in st.session_state.trabalhadores:
                st.warning("Este nome já existe na equipa.")
            else:
                st.warning("Introduza o nome do trabalhador.")

    with col2:
        st.subheader("Lista de Trabalhadores")
        st.caption("Use o seletor para definir quem participa na geração da escala. Os trabalhadores inativos continuam registados na equipa.")
        if st.session_state.trabalhadores:
            for i, nome in enumerate(st.session_state.trabalhadores):
                c1, c2, c3 = st.columns([3, 1, 1])
                c1.markdown(resumo_competencias_html(nome, i + 1), unsafe_allow_html=True)
                preferencia_12h = st.session_state.preferencias_jornadas_12h.get(nome, False)
                limite_12h = st.session_state.limites_jornadas_12h.get(nome, LIMITE_JORNADAS_12H_PADRAO)
                hrs_contrato = st.session_state.horas_contrato_semanal.get(nome, HORAS_CONTRATO_SEMANAL_PADRAO)
                c1.caption(f"Contrato: **{hrs_contrato}h/semana** · Jornadas 12h: {'prefere' if preferencia_12h else 'sem preferência'} (máx {limite_12h}/mês)")
                ativo = c2.checkbox("Participa na escala", value=st.session_state.trabalhadores_ativos.get(nome, True), key=f"trabalhador_ativo_{nome}")
                st.session_state.trabalhadores_ativos[nome] = ativo
                if c3.button("Remover", key=f"del_{nome}"):
                    st.session_state.trabalhadores.remove(nome)
                    st.session_state.trabalhadores_ativos.pop(nome, None)
                    st.session_state.horas_contrato_semanal.pop(nome, None)
                    st.session_state.competencias.pop(nome, None)
                    guardar_estado_no_github()
                    st.rerun()
                with st.expander("Editar horas de contrato e competências", expanded=False):
                    competencias_atuais = st.session_state.competencias.get(nome, {})
                    horas_editadas = st.number_input("Horas de Contrato Semanal", min_value=1.0, max_value=60.0, value=float(st.session_state.horas_contrato_semanal.get(nome, HORAS_CONTRATO_SEMANAL_PADRAO)), step=0.5, key=f"editar_horas_{nome}")
                    preferencia_12h_editada = st.checkbox("Prefere jornadas de 12 horas (M12 + T24)", value=st.session_state.preferencias_jornadas_12h.get(nome, False), key=f"editar_preferencia_12h_{nome}")
                    limite_12h_editado = st.number_input("Limite mensal de jornadas de 12 horas", min_value=0, max_value=31, value=int(st.session_state.limites_jornadas_12h.get(nome, LIMITE_JORNADAS_12H_PADRAO)), step=1, key=f"editar_limite_12h_{nome}")
                    responsabilidades_do_trabalhador = st.multiselect("Responsabilidades que pode fazer", responsabilidades_disponiveis, default=[resp for resp in responsabilidades_disponiveis if resp in competencias_atuais], key=f"editar_responsabilidades_{nome}")
                    competencias_editadas = {}
                    for responsabilidade in responsabilidades_do_trabalhador:
                        turnos_validos = turnos_por_responsabilidade[responsabilidade]
                        turnos_atuais = [turno for turno in competencias_atuais.get(responsabilidade, []) if turno in turnos_validos]
                        competencias_editadas[responsabilidade] = st.multiselect(f"Turnos de {responsabilidade}", turnos_validos, default=turnos_atuais, key=f"editar_turnos_{nome}_{responsabilidade}")
                    if st.button("Guardar alterações", key=f"guardar_competencias_{nome}"):
                        st.session_state.horas_contrato_semanal[nome] = horas_editadas
                        st.session_state.competencias[nome] = competencias_editadas
                        st.session_state.preferencias_jornadas_12h[nome] = preferencia_12h_editada
                        st.session_state.limites_jornadas_12h[nome] = limite_12h_editado
                        guardar_estado_no_github()
                        st.success(f"Dados de {nome} atualizados e guardados.")

# -----------------------------------------------------------------------------
# TAB 2: RESPONSABILIDADES E TURNOS
# -----------------------------------------------------------------------------
with tabs[1]:
    st.header("Responsabilidades e Turnos")
    st.caption("Defina os turnos disponíveis para cada responsabilidade. Pode editar os exemplos, remover linhas ou adicionar novas.")
    turnos_editados = st.data_editor(st.session_state.turnos, width="stretch", num_rows="dynamic", column_config={"Responsabilidade": st.column_config.TextColumn("Responsabilidade", required=True), "Turno": st.column_config.TextColumn("Código do turno", required=True), "Início": st.column_config.TextColumn("Início", help="Formato HH:MM, por exemplo 08:00"), "Fim": st.column_config.TextColumn("Fim", help="Formato HH:MM, por exemplo 16:00")}, key="editor_turnos")
    turnos_normalizados = normalizar_turnos(turnos_editados)
    if len(turnos_normalizados) != len(turnos_editados):
        st.info("As linhas sem responsabilidade ou turno não serão utilizadas.")
    if st.button("Guardar configuração dos turnos"):
        st.session_state.turnos = turnos_normalizados
        st.session_state.pop("df_necessidades", None)
        st.session_state.pop("assinatura_necessidades", None)
        guardar_estado_no_github()
        st.success("Responsabilidades e turnos guardados.")
        st.rerun()

# -----------------------------------------------------------------------------
# TAB 3: NECESSIDADES MENSAIS
# -----------------------------------------------------------------------------
with tabs[2]:
    st.header("Configuração de Necessidades do Mês")
    col_mes, col_ano = st.columns(2)
    mes_sel = col_mes.selectbox("Mês", list(range(1, 13)), index=datetime.now().month - 1)
    ano_sel = col_ano.number_input("Ano", min_value=2024, max_value=2030, value=datetime.now().year)
    num_dias = calendar.monthrange(ano_sel, mes_sel)[1]
    dias_uteis_mes = contar_dias_uteis_mes(mes_sel, ano_sel)
    
    st.info(f"📅 **{calendar.month_name[mes_sel]} {ano_sel}**: Este mês tem **{dias_uteis_mes} dias úteis** (alvo base de **{dias_uteis_mes * 7}h** para um contrato de 35h/semana).")

    turnos_ativos = normalizar_turnos(st.session_state.turnos)
    assinatura_atual = ("dias_nas_linhas_numericos", mes_sel, ano_sel, assinatura_turnos(turnos_ativos))

    if turnos_ativos.empty:
        st.warning("Adicione responsabilidades e turnos na Tab 2 primeiro.")
    else:
        slots_ordenados = turnos_ordenados(turnos_ativos)
        colunas_turnos = [chave_coluna_turno(slot) for slot in slots_ordenados]
        linhas_dias = [str(dia) for dia in range(1, num_dias + 1)]
        if "df_necessidades" not in st.session_state or st.session_state.get("assinatura_necessidades") != assinatura_atual:
            st.session_state.df_necessidades = pd.DataFrame(0, index=linhas_dias, columns=colunas_turnos)
            st.session_state.assinatura_necessidades = assinatura_atual

        st.caption("Cada linha representa um dia. Os fins de semana surgem destacados a dourado.")
        necessidades_com_fins_de_semana = estilizar_fins_de_semana(st.session_state.df_necessidades, mes_sel, ano_sel, dias_nas_colunas=False)
        
        df_editado_nec = st.data_editor(
            necessidades_com_fins_de_semana,
            width="stretch",
            num_rows="fixed",
            key="editor_necessidades_mensais",
            column_config={
                chave_coluna_turno(slot): st.column_config.NumberColumn(
                    f"{indicador_responsabilidade(slot.Responsabilidade)} {slot.Turno}",
                    help=slot.Responsabilidade,
                    min_value=0,
                    step=1,
                    format="%d"
                ) for slot in slots_ordenados
            }
        )
        if isinstance(df_editado_nec, pd.DataFrame):
            st.session_state.df_necessidades = df_editado_nec.copy()
            
        if st.button("💾 Guardar Necessidades Mensais"):
            guardar_estado_no_github()

# -----------------------------------------------------------------------------
# TAB 4: INDISPONIBILIDADES E PREFERÊNCIAS
# -----------------------------------------------------------------------------
with tabs[3]:
    st.header("Turnos, Folgas e Férias da Equipa")
    st.caption("Preencha diretamente a tabela por trabalhador e por dia. Célula vazia = disponível sem preferência, F = Folga, L = Licença/Férias (contabiliza 7h), ou um código de turno.")
    dias_do_mes = [str(dia) for dia in range(1, num_dias + 1)]
    codigos_turno = codigos_turno_ordenados(st.session_state.turnos)
    assinatura_preferencias = (mes_sel, ano_sel, tuple(st.session_state.trabalhadores), tuple(codigos_turno))
    preferencias_atuais = st.session_state.get("preferencias_turnos")

    if not isinstance(preferencias_atuais, pd.DataFrame) or st.session_state.get("assinatura_preferencias") != assinatura_preferencias:
        preferencias_turnos = pd.DataFrame("", index=st.session_state.trabalhadores, columns=dias_do_mes)
        st.session_state.preferencias_turnos = preferencias_turnos
        st.session_state.assinatura_preferencias = assinatura_preferencias

    opcoes_tabela = ["", "F", "L", *codigos_turno]
    preferencias_com_fins_de_semana = estilizar_fins_de_semana(st.session_state.preferencias_turnos, mes_sel, ano_sel, dias_nas_colunas=True)
    
    df_editado_pref = st.data_editor(
        preferencias_com_fins_de_semana,
        width="stretch",
        num_rows="fixed",
        key="editor_preferencias_turnos",
        column_config={dia: st.column_config.SelectboxColumn(dia, options=opcoes_tabela, required=False, help="Escolha F para folga, L para Licença/Férias (7h) ou um código de turno.") for dia in dias_do_mes}
    )
    if isinstance(df_editado_pref, pd.DataFrame):
        st.session_state.preferencias_turnos = df_editado_pref.copy()

    if st.button("💾 Guardar Indisponibilidades e Férias"):
        guardar_estado_no_github()

# -----------------------------------------------------------------------------
# TAB 5: GERADOR DE ESCALA AUTOMÁTICA
# -----------------------------------------------------------------------------
with tabs[4]:
    st.header("Gerador de Escala Automática")
    st.caption("Regras Ativas: Múltiplos turnos M12+T24; descanso noturno; máx. 5 dias seguidos; 1-2 fds livres; distribuição equitativa de turnos/responsabilidades; Turno T ao fim da sequência; folgas agrupadas (2 a 3) e distribuídas pelo mês.")
    
    if st.button("⚡ Gerar Escala Optimizada", type="primary"):
        trabalhadores = [nome for nome in st.session_state.trabalhadores if st.session_state.trabalhadores_ativos.get(nome, True)]
        turnos = normalizar_turnos(st.session_state.turnos)
        df_nec = st.session_state.get("df_necessidades")

        if not st.session_state.trabalhadores or not trabalhadores:
            st.error("Adicione e ative trabalhadores antes de gerar a escala.")
        elif turnos.empty or df_nec is None:
            st.error("Configure as responsabilidades, os turnos e as necessidades mensais.")
        else:
            slots = turnos_ordenados(turnos)
            duracoes_slots = [calcular_duracao_turno_horas(s.Início, s.Fim) for s in slots]
            dias_uteis_mes = contar_dias_uteis_mes(mes_sel, ano_sel)

            model = cp_model.CpModel()
            escala = {}
            trabalha_dia = {}

            # 1. Variáveis de decisão base
            for trabalhador in trabalhadores:
                for dia in range(1, num_dias + 1):
                    trabalha_dia[(trabalhador, dia)] = model.NewBoolVar(f"trabalha_{trabalhador}_{dia}")
                    for slot_index, _slot in enumerate(slots):
                        escala[(trabalhador, dia, slot_index)] = model.NewBoolVar(f"e_{trabalhador}_{dia}_{slot_index}")

            # Vincular trabalha_dia
            for trabalhador in trabalhadores:
                for dia in range(1, num_dias + 1):
                    turnos_do_dia = [escala[(trabalhador, dia, s_idx)] for s_idx in range(len(slots))]
                    model.Add(sum(turnos_do_dia) >= 1).OnlyEnforceIf(trabalha_dia[(trabalhador, dia)])
                    model.Add(sum(turnos_do_dia) == 0).OnlyEnforceIf(trabalha_dia[(trabalhador, dia)].Not())

            # 2. Restrição de Competências
            for trabalhador in trabalhadores:
                for dia in range(1, num_dias + 1):
                    for slot_index, slot in enumerate(slots):
                        if not trabalhador_pode_fazer(trabalhador, slot.Responsabilidade, slot.Turno):
                            model.Add(escala[(trabalhador, dia, slot_index)] == 0)

            # 3. Máximo de 5 Dias Consecutivos
            for trabalhador in trabalhadores:
                for d in range(1, num_dias - 4):
                    model.Add(sum(trabalha_dia[(trabalhador, d + k)] for k in range(6)) <= 5)

            # 4. Fins de Semana Livres (Mínimo 1, idealmente 2)
            fins_de_semana = []
            for d in range(1, num_dias):
                if datetime(ano_sel, mes_sel, d).weekday() == 5:
                    if d + 1 <= num_dias:
                        fins_de_semana.append((d, d + 1))

            fds_livre_var = {}
            for trabalhador in trabalhadores:
                for idx_fds, (sab, dom) in enumerate(fins_de_semana):
                    v_livre = model.NewBoolVar(f"fds_livre_{trabalhador}_{idx_fds}")
                    fds_livre_var[(trabalhador, idx_fds)] = v_livre
                    model.Add(trabalha_dia[(trabalhador, sab)] == 0).OnlyEnforceIf(v_livre)
                    model.Add(trabalha_dia[(trabalhador, dom)] == 0).OnlyEnforceIf(v_livre)
                    model.Add(trabalha_dia[(trabalhador, sab)] + trabalha_dia[(trabalhador, dom)] >= 1).OnlyEnforceIf(v_livre.Not())

                if fins_de_semana:
                    model.Add(sum(fds_livre_var[(trabalhador, idx_fds)] for idx_fds in range(len(fins_de_semana))) >= 1)

            # 5. Jornadas de 12 Horas (M12 + T24)
            pares_jornada = []
            pares_jornada_por_trabalhador = {trabalhador: [] for trabalhador in trabalhadores}
            pares_jornada_preferidos = []
            pares_responsabilidades_diferentes = []
            for trabalhador in trabalhadores:
                for dia in range(1, num_dias + 1):
                    variaveis_m12 = [escala[(trabalhador, dia, slot_index)] for slot_index, slot in enumerate(slots) if slot.Turno == "M12"]
                    variaveis_t24 = [escala[(trabalhador, dia, slot_index)] for slot_index, slot in enumerate(slots) if slot.Turno == "T24"]
                    variaveis_outros = [escala[(trabalhador, dia, slot_index)] for slot_index, slot in enumerate(slots) if not e_turno_de_jornada_alargada(slot.Turno)]

                    model.Add(sum(variaveis_m12) <= 1)
                    model.Add(sum(variaveis_t24) <= 1)
                    model.Add(sum(variaveis_outros) <= 1)
                    model.Add(sum(variaveis_m12) + sum(variaveis_outros) <= 1)
                    model.Add(sum(variaveis_t24) + sum(variaveis_outros) <= 1)

                    for m12_index, m12_slot in enumerate(slots):
                        if m12_slot.Turno != "M12": continue
                        for t24_index, t24_slot in enumerate(slots):
                            if t24_slot.Turno != "T24": continue
                            par = model.NewBoolVar(f"par_{trabalhador}_{dia}_{m12_index}_{t24_index}")
                            m12_atribuido = escala[(trabalhador, dia, m12_index)]
                            t24_atribuido = escala[(trabalhador, dia, t24_index)]
                            model.Add(par <= m12_atribuido)
                            model.Add(par <= t24_atribuido)
                            model.Add(par >= m12_atribuido + t24_atribuido - 1)
                            pares_jornada.append(par)
                            pares_jornada_por_trabalhador[trabalhador].append(par)
                            if st.session_state.preferencias_jornadas_12h.get(trabalhador, False):
                                pares_jornada_preferidos.append(par)
                            if m12_slot.Responsabilidade != t24_slot.Responsabilidade:
                                pares_responsabilidades_diferentes.append(par)

            for trabalhador in trabalhadores:
                limite_12h = max(0, int(st.session_state.limites_jornadas_12h.get(trabalhador, LIMITE_JORNADAS_12H_PADRAO)))
                model.Add(sum(pares_jornada_por_trabalhador[trabalhador]) <= limite_12h)

            # 6. Necessidades Diárias
            for dia in range(1, num_dias + 1):
                for slot_index, slot in enumerate(slots):
                    coluna = chave_coluna_turno(slot)
                    necessidade = int(df_nec.loc[str(dia), coluna])
                    model.Add(sum(escala[(trabalhador, dia, slot_index)] for trabalhador in trabalhadores) == necessidade)

            # 7. Preferências, Folgas 'F' e Férias/Licença 'L'
            preferencias_turnos = st.session_state.get("preferencias_turnos")
            dias_ferias_por_trabalhador = {t: 0 for t in trabalhadores}

            for trabalhador in trabalhadores:
                for dia in range(1, num_dias + 1):
                    preferencia = ""
                    coluna_dia = str(dia)
                    if isinstance(preferencias_turnos, pd.DataFrame) and trabalhador in preferencias_turnos.index and coluna_dia in preferencias_turnos.columns:
                        preferencia = str(preferencias_turnos.loc[trabalhador, coluna_dia]).strip()

                    if preferencia in {"F", "L"}:
                        for slot_index in range(len(slots)):
                            model.Add(escala[(trabalhador, dia, slot_index)] == 0)
                        if preferencia == "L":
                            dias_ferias_por_trabalhador[trabalhador] += 1
                    elif preferencia:
                        for slot_index, slot in enumerate(slots):
                            if slot.Turno != preferencia:
                                model.Add(escala[(trabalhador, dia, slot_index)] == 0)

            # 8. Descanso Noturno Mínimo
            for trabalhador in trabalhadores:
                for dia in range(1, num_dias):
                    for slot_index, slot in enumerate(slots):
                        if not turno_atravessa_meia_noite(slot.Início, slot.Fim): continue
                        for proximo_index, proximo_slot in enumerate(slots):
                            try:
                                hora_inicio = datetime.strptime(proximo_slot.Início, "%H:%M").hour
                            except ValueError:
                                continue
                            if hora_inicio < 12:
                                model.Add(escala[(trabalhador, dia, slot_index)] + escala[(trabalhador, dia + 1, proximo_index)] <= 1)

            # 9. Regras de Otimização e Sequenciamento
            penalizacoes_t_meio = []
            for trabalhador in trabalhadores:
                for dia in range(1, num_dias):
                    turnos_t_dia = [escala[(trabalhador, dia, s_idx)] for s_idx, s in enumerate(slots) if s.Turno == "T"]
                    if turnos_t_dia:
                        t_em_trabalho_seguido = model.NewBoolVar(f"t_in_middle_{trabalhador}_{dia}")
                        model.Add(sum(turnos_t_dia) + trabalha_dia[(trabalhador, dia + 1)] == 2).OnlyEnforceIf(t_em_trabalho_seguido)
                        model.Add(sum(turnos_t_dia) + trabalha_dia[(trabalhador, dia + 1)] < 2).OnlyEnforceIf(t_em_trabalho_seguido.Not())
                        penalizacoes_t_meio.append(t_em_trabalho_seguido)

            penalizacoes_folga_isolada = []
            for trabalhador in trabalhadores:
                for dia in range(2, num_dias):
                    folga_1_dia = model.NewBoolVar(f"folga_isolada_{trabalhador}_{dia}")
                    model.Add(trabalha_dia[(trabalhador, dia - 1)] + trabalha_dia[(trabalhador, dia)].Not() + trabalha_dia[(trabalhador, dia + 1)] == 3).OnlyEnforceIf(folga_1_dia)
                    model.Add(trabalha_dia[(trabalhador, dia - 1)] + trabalha_dia[(trabalhador, dia)].Not() + trabalha_dia[(trabalhador, dia + 1)] < 3).OnlyEnforceIf(folga_1_dia.Not())
                    penalizacoes_folga_isolada.append(folga_1_dia)

            desvios_equidade = []
            todas_resps = list({s.Responsabilidade for s in slots})
            for resp in todas_resps:
                contagens_resp = []
                for trabalhador in trabalhadores:
                    slots_resp = [escala[(trabalhador, d, s_idx)] for d in range(1, num_dias + 1) for s_idx, s in enumerate(slots) if s.Responsabilidade == resp]
                    var_c = model.NewIntVar(0, num_dias, f"count_resp_{resp}_{trabalhador}")
                    model.Add(var_c == sum(slots_resp))
                    contagens_resp.append(var_c)
                
                max_resp = model.NewIntVar(0, num_dias, f"max_resp_{resp}")
                min_resp = model.NewIntVar(0, num_dias, f"min_resp_{resp}")
                model.AddMaxEquality(max_resp, contagens_resp)
                model.AddMinEquality(min_resp, contagens_resp)
                diff_resp = model.NewIntVar(0, num_dias, f"diff_resp_{resp}")
                model.Add(diff_resp == max_resp - min_resp)
                desvios_equidade.append(diff_resp)

            todos_codigos = list({s.Turno for s in slots})
            for cod_t in todos_codigos:
                contagens_cod = []
                for trabalhador in trabalhadores:
                    slots_cod = [escala[(trabalhador, d, s_idx)] for d in range(1, num_dias + 1) for s_idx, s in enumerate(slots) if s.Turno == cod_t]
                    var_c = model.NewIntVar(0, num_dias, f"count_cod_{cod_t}_{trabalhador}")
                    model.Add(var_c == sum(slots_cod))
                    contagens_cod.append(var_c)
                
                max_cod = model.NewIntVar(0, num_dias, f"max_cod_{cod_t}")
                min_cod = model.NewIntVar(0, num_dias, f"min_cod_{cod_t}")
                model.AddMaxEquality(max_cod, contagens_cod)
                model.AddMinEquality(min_cod, contagens_cod)
                diff_cod = model.NewIntVar(0, num_dias, f"diff_cod_{cod_t}")
                model.Add(diff_cod == max_cod - min_cod)
                desvios_equidade.append(diff_cod)

            # 10. Cálculo do Balanço de Horas por Dia Útil
            duracoes_int = [int(round(d * 10)) for d in duracoes_slots]

            desvios_absolutos = []
            for trabalhador in trabalhadores:
                hrs_semanais = float(st.session_state.horas_contrato_semanal.get(trabalhador, HORAS_CONTRATO_SEMANAL_PADRAO))
                hrs_diarias_alvo = hrs_semanais / 5.0
                hrs_alvo_mes = hrs_diarias_alvo * dias_uteis_mes
                hrs_alvo_int = int(round(hrs_alvo_mes * 10))

                horas_ferias_int = int(round(dias_ferias_por_trabalhador[trabalhador] * HORAS_DIA_FERIAS_LICENCA * 10))

                expressao_horas = [horas_ferias_int]
                for dia in range(1, num_dias + 1):
                    for slot_index in range(len(slots)):
                        expressao_horas.append(escala[(trabalhador, dia, slot_index)] * duracoes_int[slot_index])

                total_hrs_var = model.NewIntVar(0, 4000, f"total_hrs_{trabalhador}")
                model.Add(total_hrs_var == sum(expressao_horas))

                desvio_var = model.NewIntVar(-4000, 4000, f"desvio_{trabalhador}")
                model.Add(desvio_var == total_hrs_var - hrs_alvo_int)

                desvio_abs = model.NewIntVar(0, 4000, f"desvio_abs_{trabalhador}")
                model.AddAbsEquality(desvio_abs, desvio_var)
                desvios_absolutos.append(desvio_abs)

            # 11. Função Objetivo Integrada
            objetivo = []
            if pares_jornada_preferidos:
                objetivo.append(1_000_000 * sum(pares_jornada_preferidos))
            if fins_de_semana:
                todos_fds_livres = [fds_livre_var[k] for k in fds_livre_var]
                objetivo.append(50_000 * sum(todos_fds_livres))
            if pares_responsabilidades_diferentes:
                objetivo.append(1_000 * sum(pares_responsabilidades_diferentes))

            if penalizacoes_t_meio:
                objetivo.append(-5_000 * sum(penalizacoes_t_meio))
            if penalizacoes_folga_isolada:
                objetivo.append(-3_000 * sum(penalizacoes_folga_isolada))
            if desvios_equidade:
                objetivo.append(-2_000 * sum(desvios_equidade))

            objetivo.append(-10 * sum(desvios_absolutos))
            model.Maximize(sum(objetivo))

            # Executar Solver
            solver = cp_model.CpSolver()
            solver.parameters.max_time_in_seconds = 30.0
            status = solver.Solve(model)

            if status in (cp_model.OPTIMAL, cp_model.FEASIBLE):
                st.success("Escala gerada com sucesso!")
                dados_escala = {}
                dados_meta_resps = {}
                totais_horas_realizadas = {}
                banco_horas = {}

                for trabalhador in trabalhadores:
                    linha_trabalhador = []
                    linha_meta_trabalhador = []
                    horas_realizadas_trab = dias_ferias_por_trabalhador[trabalhador] * HORAS_DIA_FERIAS_LICENCA

                    for dia in range(1, num_dias + 1):
                        pref_dia = ""
                        col_d = str(dia)
                        if isinstance(preferencias_turnos, pd.DataFrame) and trabalhador in preferencias_turnos.index and col_d in preferencias_turnos.columns:
                            pref_dia = str(preferencias_turnos.loc[trabalhador, col_d]).strip()

                        if pref_dia == "L":
                            linha_trabalhador.append("L")
                            linha_meta_trabalhador.append("L")
                            continue

                        codigos_do_dia = []
                        resps_do_dia = []
                        for slot_index, slot in enumerate(slots):
                            if solver.Value(escala[(trabalhador, dia, slot_index)]):
                                codigos_do_dia.append(slot.Turno)
                                resps_do_dia.append(slot.Responsabilidade)
                                horas_realizadas_trab += duracoes_slots[slot_index]

                        if len(codigos_do_dia) > 1:
                            atribuicao = " / ".join(codigos_do_dia)
                            meta_val = (codigos_do_dia, resps_do_dia)
                        elif len(codigos_do_dia) == 1:
                            atribuicao = codigos_do_dia[0]
                            meta_val = resps_do_dia[0]
                        else:
                            atribuicao = "F"
                            meta_val = ""

                        linha_trabalhador.append(atribuicao)
                        linha_meta_trabalhador.append(meta_val)

                    hrs_contrato_sem = float(st.session_state.horas_contrato_semanal.get(trabalhador, HORAS_CONTRATO_SEMANAL_PADRAO))
                    hrs_contrato_mes = round((hrs_contrato_sem / 5.0) * dias_uteis_mes, 1)
                    saldo_banco = round(horas_realizadas_trab - hrs_contrato_mes, 1)

                    totais_horas_realizadas[trabalhador] = horas_realizadas_trab
                    banco_horas[trabalhador] = saldo_banco

                    s_banco = f"+{saldo_banco}h" if saldo_banco > 0 else f"{saldo_banco}h"
                    
                    linha_trabalhador.extend([
                        f"{round(horas_realizadas_trab, 1)}h",
                        f"{hrs_contrato_mes}h",
                        s_banco
                    ])
                    linha_meta_trabalhador.extend(["", "", ""])
                    
                    dados_escala[trabalhador] = linha_trabalhador
                    dados_meta_resps[trabalhador] = linha_meta_trabalhador

                colunas_dias = [str(dia) for dia in range(1, num_dias + 1)]
                todas_colunas = colunas_dias + ["Horas Realizadas", "Alvo Contratual", "Banco de Horas"]

                df_resultado = pd.DataFrame.from_dict(dados_escala, orient="index", columns=todas_colunas)
                df_meta_resps = pd.DataFrame.from_dict(dados_meta_resps, orient="index", columns=todas_colunas)
                
                # Botões de Impressão e Exportação
                col_print, col_exp1, col_exp2 = st.columns([1, 1, 1])

                # Botão de Impressão com acionamento JavaScript nativo
                if col_print.button("🖨️ Imprimir / Guardar em PDF", use_container_width=True, type="primary"):
                    components.html("<script>window.parent.print();</script>", height=0, width=0)

                # Exportar Excel
                excel_bytes = gerar_excel_escala_formatado(df_resultado, df_meta_resps, mes_sel, ano_sel)
                col_exp1.download_button(
                    label="📊 Descarregar Excel (.xlsx)",
                    data=excel_bytes,
                    file_name=f"escala_{mes_sel}_{ano_sel}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

                # Exportar CSV
                csv = df_resultado.to_csv().encode("utf-8")
                col_exp2.download_button(
                    label="📥 Descarregar CSV",
                    data=csv,
                    file_name=f"escala_{mes_sel}_{ano_sel}.csv",
                    mime="text/csv",
                    use_container_width=True
                )

                # Renderizar Legenda das Cores do Texto por Responsabilidade
                st.markdown("### 🎨 Legenda dos Turnos por Setor")
                legenda_html = "<div style='display:flex; gap:15px; flex-wrap:wrap; margin-bottom:15px;'>"
                for resp, c_info in RESPONSABILIDADE_CORES.items():
                    legenda_html += f"<span style='color:{c_info['texto']}; font-weight:bold; font-size:0.95rem; background:#f9fafb; padding:4px 10px; border-radius:6px; border:1px solid #e5e7eb;'>● {resp}</span>"
                legenda_html += "</div>"
                st.markdown(legenda_html, unsafe_allow_html=True)

                # Renderizar Tabela HTML Padronizada
                tabela_html = renderizar_tabela_escala_html(df_resultado, df_meta_resps, mes_sel, ano_sel)
                st.markdown(tabela_html, unsafe_allow_html=True)

                st.subheader("📊 Resumo do Banco de Horas da Equipa")
                cols_met = st.columns(min(len(trabalhadores), 5))
                for idx_m, t_nome in enumerate(trabalhadores[:5]):
                    saldo = banco_horas[t_nome]
                    s_str = f"+{saldo}h" if saldo > 0 else f"{saldo}h"
                    cols_met[idx_m].metric(label=t_nome, value=f"{totais_horas_realizadas[t_nome]}h", delta=s_str)
            else:
                st.error("Não foi possível encontrar uma solução válida com as restrições impostas. Tente reduzir as necessidades mensais ou ajustar as folgas/férias solicitadas.")

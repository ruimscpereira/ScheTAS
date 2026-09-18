import calendar
import html
import re
from datetime import datetime

import pandas as pd
import streamlit as st
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

RESPONSABILIDADE_CORES = {
    "Radiologia Convencional": {
        "fundo": "#dbeafe",
        "texto": "#1d4ed8",
        "borda": "#93c5fd",
    },
    "Tomografia Computorizada": {
        "fundo": "#dcfce7",
        "texto": "#15803d",
        "borda": "#86efac",
    },
    "Ecografia": {
        "fundo": "#fee2e2",
        "texto": "#b91c1c",
        "borda": "#fca5a5",
    },
    "Ressonância Magnética": {
        "fundo": "#ffedd5",
        "texto": "#c2410c",
        "borda": "#fdba74",
    },
}

RESPONSABILIDADE_INDICADORES = {
    "Radiologia Convencional": "🔵",
    "Tomografia Computorizada": "🟢",
    "Ecografia": "🔴",
    "Ressonância Magnética": "🟠",
}

LIMITE_JORNADAS_12H_PADRAO = 31
TIPO_CONTRATO_PADRAO = "Tempo inteiro"
HORAS_DIARIAS_PADRAO = 7.0

st.set_page_config(page_title="Gestor de Escalas de Trabalho", layout="wide")

if "trabalhadores" not in st.session_state:
    st.session_state.trabalhadores = [
        "Ana Silva",
        "Bruno Santos",
        "Carla Costa",
        "Daniel Rocha",
        "Eduarda Lima",
    ]

if "trabalhadores_ativos" not in st.session_state:
    st.session_state.trabalhadores_ativos = {
        nome: True for nome in st.session_state.trabalhadores
    }
else:
    for nome in st.session_state.trabalhadores:
        st.session_state.trabalhadores_ativos.setdefault(nome, True)

if "preferencias_jornadas_12h" not in st.session_state:
    st.session_state.preferencias_jornadas_12h = {
        nome: False for nome in st.session_state.trabalhadores
    }
else:
    for nome in st.session_state.trabalhadores:
        st.session_state.preferencias_jornadas_12h.setdefault(nome, False)

if "limites_jornadas_12h" not in st.session_state:
    st.session_state.limites_jornadas_12h = {
        nome: LIMITE_JORNADAS_12H_PADRAO
        for nome in st.session_state.trabalhadores
    }
else:
    for nome in st.session_state.trabalhadores:
        st.session_state.limites_jornadas_12h.setdefault(
            nome, LIMITE_JORNADAS_12H_PADRAO
        )

if "tipos_contrato" not in st.session_state:
    st.session_state.tipos_contrato = {
        nome: TIPO_CONTRATO_PADRAO for nome in st.session_state.trabalhadores
    }
else:
    for nome in st.session_state.trabalhadores:
        st.session_state.tipos_contrato.setdefault(nome, TIPO_CONTRATO_PADRAO)

if "horas_diarias" not in st.session_state:
    st.session_state.horas_diarias = {
        nome: HORAS_DIARIAS_PADRAO for nome in st.session_state.trabalhadores
    }
else:
    for nome in st.session_state.trabalhadores:
        st.session_state.horas_diarias.setdefault(nome, HORAS_DIARIAS_PADRAO)

if "indisponibilidades" not in st.session_state:
    st.session_state.indisponibilidades = {}

if "turnos" not in st.session_state:
    st.session_state.turnos = pd.DataFrame(DEFAULT_TURNOS)


def normalizar_turnos(df):
    colunas = ["Responsabilidade", "Turno", "Início", "Fim"]
    resultado = df.reindex(columns=colunas).copy()
    for coluna in colunas:
        resultado[coluna] = resultado[coluna].fillna("").astype(str).str.strip()
    resultado = resultado[
        (resultado["Responsabilidade"] != "") & (resultado["Turno"] != "")
    ]
    return resultado.drop_duplicates(
        subset=["Responsabilidade", "Turno"], keep="last"
    ).reset_index(drop=True)


def assinatura_turnos(df):
    return tuple(
        tuple(linha)
        for linha in normalizar_turnos(df).itertuples(index=False, name=None)
    )


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
        key=lambda linha: (
            chave_ordenacao_turno(linha.Turno),
            linha.Responsabilidade.lower(),
        ),
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


def dias_uteis_no_mes(mes, ano):
    return sum(
        datetime(ano, mes, dia).weekday() < 5
        for dia in range(1, calendar.monthrange(ano, mes)[1] + 1)
    )


def horas_diarias_do_trabalhador(nome):
    try:
        return max(0.0, float(st.session_state.horas_diarias.get(nome, 0)))
    except (TypeError, ValueError):
        return 0.0


def formatar_horas(valor):
    return f"{float(valor):g}"


def dia_a_partir_do_rotulo(rotulo):
    correspondencia = re.match(r"^\D*(\d+)", str(rotulo))
    return int(correspondencia.group(1)) if correspondencia else None


def estilizar_fins_de_semana(df, mes, ano, dias_nas_colunas):
    preenchimento = "#fff7ed"
    borda_fim_de_semana = "#f59e0b"
    estilo_celula_fim_de_semana = (
        f"background-color: {preenchimento}; "
        f"border: 2px solid {borda_fim_de_semana};"
    )

    if dias_nas_colunas:
        def estilo_coluna(coluna):
            dia = dia_a_partir_do_rotulo(coluna.name)
            fim_de_semana = (
                dia is not None and e_fim_de_semana(dia, mes, ano)
            )
            return [
                estilo_celula_fim_de_semana if fim_de_semana else ""
                for _ in coluna
            ]

        def estilo_cabecalho_colunas(indice):
            return [
                f"font-weight: bold; border: 2px solid {borda_fim_de_semana};"
                if (
                    dia_a_partir_do_rotulo(rotulo) is not None
                    and e_fim_de_semana(
                        dia_a_partir_do_rotulo(rotulo), mes, ano
                    )
                )
                else ""
                for rotulo in indice
            ]

        return df.style.apply(estilo_coluna, axis=0).apply_index(
            estilo_cabecalho_colunas, axis=1
        )

    def estilo_linha(linha):
        dia = dia_a_partir_do_rotulo(linha.name)
        fim_de_semana = (
            dia is not None and e_fim_de_semana(dia, mes, ano)
        )
        return [
            estilo_celula_fim_de_semana if fim_de_semana else ""
            for _ in linha
        ]

    def estilo_indice_linhas(indice):
        return [
            f"font-weight: bold; border: 2px solid {borda_fim_de_semana};"
            if (
                dia_a_partir_do_rotulo(rotulo) is not None
                and e_fim_de_semana(dia_a_partir_do_rotulo(rotulo), mes, ano)
            )
            else ""
            for rotulo in indice
        ]

    return df.style.apply(estilo_linha, axis=1).apply_index(
        estilo_indice_linhas, axis=0
    )


def competencias_iniciais(df):
    return {
        responsabilidade: list(turnos)
        for responsabilidade, turnos in opcoes_turnos_por_responsabilidade(df).items()
    }


def trabalhador_pode_fazer(nome, responsabilidade, turno):
    competencias = st.session_state.competencias.get(nome, {})
    return turno in competencias.get(responsabilidade, [])


def resumo_competencias_html(nome, numero=None):
    competencias = st.session_state.competencias.get(nome, {})
    etiquetas = []
    for responsabilidade, turnos in competencias.items():
        if not turnos:
            continue
        cores = RESPONSABILIDADE_CORES.get(
            responsabilidade,
            {"fundo": "#f3f4f6", "texto": "#374151", "borda": "#d1d5db"},
        )
        responsabilidade_segura = html.escape(responsabilidade)
        turnos_seguro = html.escape(", ".join(turnos))
        etiquetas.append(
            "<span style='display:inline-block; margin:3px 5px 3px 0; "
            f"padding:4px 9px; border-radius:999px; background:{cores['fundo']}; "
            f"border:1px solid {cores['borda']}; color:{cores['texto']}; "
            "font-size:0.82rem; line-height:1.25;'>"
            f"<strong>{responsabilidade_segura}</strong>: {turnos_seguro}</span>"
        )

    nome_seguro = html.escape(nome)
    titulo = f"{numero}. {nome_seguro}" if numero is not None else nome_seguro
    if not etiquetas:
        etiquetas.append(
            "<span style='display:inline-block; margin-top:4px; padding:4px 9px; "
            "border-radius:999px; background:#f3f4f6; color:#6b7280; "
            "font-size:0.82rem;'>Sem responsabilidades atribuídas</span>"
        )
    return (
        f"<div style='margin-bottom:8px;'><strong>{titulo}</strong><br>"
        f"{''.join(etiquetas)}</div>"
    )


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
    st.session_state.competencias = {
        nome: dict(competencias_padrao)
        for nome in st.session_state.trabalhadores
    }
else:
    for nome in st.session_state.trabalhadores:
        st.session_state.competencias.setdefault(nome, {})

if "novo_trabalhador_form_version" not in st.session_state:
    st.session_state.novo_trabalhador_form_version = 0


st.title("🗓️ Gestor Inteligente de Escalas de Trabalho")

tabs = st.tabs(
    [
        "1. Gestão da Equipa",
        "2. Responsabilidades e Turnos",
        "3. Necessidades Mensais",
        "4. Indisponibilidades",
        "5. Gerar Escala",
    ]
)

# 1. Gestão da Equipa
with tabs[0]:
    st.header("Gestão de Trabalhadores")
    turnos_por_responsabilidade = opcoes_turnos_por_responsabilidade(
        st.session_state.turnos
    )
    responsabilidades_disponiveis = list(turnos_por_responsabilidade)
    st.caption(
        "Associe a cada trabalhador as responsabilidades e os turnos que "
        "está habilitado a fazer."
    )
    col1, col2 = st.columns([1, 2])

    with col1:
        st.subheader("Adicionar Colaborador")
        form_version = st.session_state.novo_trabalhador_form_version
        novo_nome = st.text_input(
            "Nome do Trabalhador",
            key=f"novo_trabalhador_nome_{form_version}",
        )
        novas_responsabilidades = st.multiselect(
            "Responsabilidades",
            responsabilidades_disponiveis,
            key=f"novo_trabalhador_responsabilidades_{form_version}",
        )
        nova_preferencia_12h = st.checkbox(
            "Prefere jornadas de 12 horas (M12 + T24)",
            value=False,
            key=f"novo_trabalhador_preferencia_12h_{form_version}",
        )
        novo_limite_12h = st.number_input(
            "Limite mensal de jornadas de 12 horas",
            min_value=0,
            max_value=31,
            value=LIMITE_JORNADAS_12H_PADRAO,
            step=1,
            help="Defina 0 se o trabalhador não puder fazer jornadas M12 + T24.",
            key=f"novo_trabalhador_limite_12h_{form_version}",
        )
        novas_competencias = {}
        for responsabilidade in novas_responsabilidades:
            novas_competencias[responsabilidade] = st.multiselect(
                f"Turnos de {responsabilidade}",
                turnos_por_responsabilidade[responsabilidade],
                key=(
                    f"novo_trabalhador_turnos_{form_version}_"
                    f"{responsabilidade}"
                ),
            )

        if st.button("Adicionar"):
            nome_normalizado = novo_nome.strip()
            if nome_normalizado and nome_normalizado not in st.session_state.trabalhadores:
                st.session_state.trabalhadores.append(nome_normalizado)
                st.session_state.trabalhadores_ativos[nome_normalizado] = True
                st.session_state.preferencias_jornadas_12h[nome_normalizado] = (
                    nova_preferencia_12h
                )
                st.session_state.limites_jornadas_12h[nome_normalizado] = (
                    novo_limite_12h
                )
                st.session_state.competencias[nome_normalizado] = novas_competencias
                st.session_state.novo_trabalhador_form_version += 1
                st.success(f"{nome_normalizado} adicionado com sucesso!")
                st.rerun()
            elif nome_normalizado in st.session_state.trabalhadores:
                st.warning("Este nome já existe na equipa.")
            else:
                st.warning("Introduza o nome do trabalhador.")

    with col2:
        st.subheader("Lista de Trabalhadores")
        st.caption(
            "Use o seletor para definir quem participa na geração da escala. "
            "Os trabalhadores inativos continuam registados na equipa."
        )
        if st.session_state.trabalhadores:
            for i, nome in enumerate(st.session_state.trabalhadores):
                c1, c2, c3 = st.columns([3, 1, 1])
                c1.markdown(
                    resumo_competencias_html(nome, i + 1),
                    unsafe_allow_html=True,
                )
                preferencia_12h = st.session_state.preferencias_jornadas_12h.get(
                    nome, False
                )
                limite_12h = st.session_state.limites_jornadas_12h.get(
                    nome, LIMITE_JORNADAS_12H_PADRAO
                )
                c1.caption(
                    f"Jornadas de 12h: "
                    f"{'tem preferência' if preferencia_12h else 'sem preferência'} "
                    f"· limite: {limite_12h}/mês"
                )
                ativo = c2.checkbox(
                    "Participa na escala",
                    value=st.session_state.trabalhadores_ativos.get(nome, True),
                    key=f"trabalhador_ativo_{nome}",
                )
                st.session_state.trabalhadores_ativos[nome] = ativo
                if c3.button("Remover", key=f"del_{nome}"):
                    st.session_state.trabalhadores.remove(nome)
                    st.session_state.trabalhadores_ativos.pop(nome, None)
                    st.session_state.indisponibilidades.pop(nome, None)
                    st.session_state.competencias.pop(nome, None)
                    st.rerun()
                with st.expander("Editar responsabilidades e turnos", expanded=False):
                    competencias_atuais = st.session_state.competencias.get(nome, {})
                    preferencia_12h_editada = st.checkbox(
                        "Prefere jornadas de 12 horas (M12 + T24)",
                        value=st.session_state.preferencias_jornadas_12h.get(
                            nome, False
                        ),
                        key=f"editar_preferencia_12h_{nome}",
                    )
                    limite_12h_editado = st.number_input(
                        "Limite mensal de jornadas de 12 horas",
                        min_value=0,
                        max_value=31,
                        value=int(
                            st.session_state.limites_jornadas_12h.get(
                                nome, LIMITE_JORNADAS_12H_PADRAO
                            )
                        ),
                        step=1,
                        help=(
                            "Defina 0 se o trabalhador não puder fazer "
                            "jornadas M12 + T24."
                        ),
                        key=f"editar_limite_12h_{nome}",
                    )
                    responsabilidades_do_trabalhador = st.multiselect(
                        "Responsabilidades que pode fazer",
                        responsabilidades_disponiveis,
                        default=[
                            responsabilidade
                            for responsabilidade in responsabilidades_disponiveis
                            if responsabilidade in competencias_atuais
                        ],
                        key=f"editar_responsabilidades_{nome}",
                    )
                    competencias_editadas = {}
                    for responsabilidade in responsabilidades_do_trabalhador:
                        turnos_validos = turnos_por_responsabilidade[
                            responsabilidade
                        ]
                        turnos_atuais = [
                            turno
                            for turno in competencias_atuais.get(responsabilidade, [])
                            if turno in turnos_validos
                        ]
                        competencias_editadas[responsabilidade] = st.multiselect(
                            f"Turnos de {responsabilidade}",
                            turnos_validos,
                            default=turnos_atuais,
                            key=f"editar_turnos_{nome}_{responsabilidade}",
                        )
                    if st.button(
                        "Guardar competências",
                        key=f"guardar_competencias_{nome}",
                    ):
                        st.session_state.competencias[nome] = competencias_editadas
                        st.session_state.preferencias_jornadas_12h[nome] = (
                            preferencia_12h_editada
                        )
                        st.session_state.limites_jornadas_12h[nome] = (
                            limite_12h_editado
                        )
                        st.success(f"Competências de {nome} atualizadas.")

# 2. Responsabilidades e turnos
with tabs[1]:
    st.header("Responsabilidades e Turnos")
    st.caption(
        "Defina os turnos disponíveis para cada responsabilidade. "
        "Pode editar os exemplos, remover linhas ou adicionar novas."
    )

    turnos_editados = st.data_editor(
        st.session_state.turnos,
        width="stretch",
        num_rows="dynamic",
        column_config={
            "Responsabilidade": st.column_config.TextColumn(
                "Responsabilidade", required=True
            ),
            "Turno": st.column_config.TextColumn("Código do turno", required=True),
            "Início": st.column_config.TextColumn(
                "Início", help="Formato HH:MM, por exemplo 08:00"
            ),
            "Fim": st.column_config.TextColumn(
                "Fim", help="Formato HH:MM, por exemplo 16:00"
            ),
        },
        key="editor_turnos",
    )

    turnos_normalizados = normalizar_turnos(turnos_editados)
    if len(turnos_normalizados) != len(turnos_editados):
        st.info("As linhas sem responsabilidade ou turno não serão utilizadas.")

    if st.button("Guardar configuração dos turnos"):
        st.session_state.turnos = turnos_normalizados
        st.session_state.pop("df_necessidades", None)
        st.session_state.pop("assinatura_necessidades", None)
        st.success("Responsabilidades e turnos guardados.")
        st.rerun()

    if not turnos_normalizados.empty:
        st.write(
            f"{len(turnos_normalizados)} turnos configurados em "
            f"{turnos_normalizados['Responsabilidade'].nunique()} responsabilidades."
        )
    else:
        st.warning("Configure pelo menos um turno antes de definir necessidades.")

# 3. Necessidades mensais
with tabs[2]:
    st.header("Configuração de Necessidades do Mês")
    col_mes, col_ano = st.columns(2)
    mes_sel = col_mes.selectbox(
        "Mês", list(range(1, 13)), index=datetime.now().month - 1
    )
    ano_sel = col_ano.number_input(
        "Ano", min_value=2024, max_value=2030, value=datetime.now().year
    )

    num_dias = calendar.monthrange(ano_sel, mes_sel)[1]
    turnos_ativos = normalizar_turnos(st.session_state.turnos)
    assinatura_atual = (
        "dias_nas_linhas_numericos",
        mes_sel,
        ano_sel,
        assinatura_turnos(turnos_ativos),
    )

    if turnos_ativos.empty:
        st.warning("Adicione responsabilidades e turnos na Tab 2 primeiro.")
    else:
        slots_ordenados = turnos_ordenados(turnos_ativos)
        colunas_turnos = [chave_coluna_turno(slot) for slot in slots_ordenados]
        linhas_dias = [str(dia) for dia in range(1, num_dias + 1)]
        if (
            "df_necessidades" not in st.session_state
            or st.session_state.get("assinatura_necessidades") != assinatura_atual
        ):
            st.session_state.df_necessidades = pd.DataFrame(
                0,
                index=linhas_dias,
                columns=colunas_turnos,
            )
            st.session_state.assinatura_necessidades = assinatura_atual

        st.caption(
            "Cada linha representa um dia. Introduza em cada coluna o número "
            "de pessoas necessárias para o turno indicado. O valor inicial é "
            "0 para evitar necessidades inventadas."
        )
        st.caption("Sábados e domingos aparecem com um preenchimento suave.")
        legenda = []
        responsabilidades_na_legenda = set()
        for slot in slots_ordenados:
            if slot.Responsabilidade in responsabilidades_na_legenda:
                continue
            responsabilidades_na_legenda.add(slot.Responsabilidade)
            cores = RESPONSABILIDADE_CORES.get(
                slot.Responsabilidade,
                {"fundo": "#f3f4f6", "texto": "#374151", "borda": "#d1d5db"},
            )
            indicador = indicador_responsabilidade(slot.Responsabilidade)
            legenda.append(
                f"<span style='display:inline-block; margin:2px 8px 2px 0; "
                f"padding:3px 8px; border-radius:999px; "
                f"background:{cores['fundo']}; color:{cores['texto']}; "
                f"border:1px solid {cores['borda']};'>"
                f"{indicador} {html.escape(slot.Responsabilidade)}</span>"
            )
        st.markdown(
            "**Legenda:** " + "".join(legenda),
            unsafe_allow_html=True,
        )
        necessidades_com_fins_de_semana = estilizar_fins_de_semana(
            st.session_state.df_necessidades,
            mes_sel,
            ano_sel,
            dias_nas_colunas=False,
        )
        st.session_state.df_necessidades = st.data_editor(
            necessidades_com_fins_de_semana,
            width="stretch",
            num_rows="fixed",
            column_config={
                chave_coluna_turno(slot): st.column_config.NumberColumn(
                    f"{indicador_responsabilidade(slot.Responsabilidade)} "
                    f"{slot.Turno}",
                    help=slot.Responsabilidade,
                    min_value=0,
                    step=1,
                    format="%d",
                )
                for slot in slots_ordenados
            },
        )

# 4. Indisponibilidades, turnos e folgas
with tabs[3]:
    st.header("Turnos e Folgas da Equipa")
    st.caption(
        "Preencha diretamente a tabela por trabalhador e por dia. Uma célula "
        "vazia significa disponibilidade sem preferência, F significa folga, "
        "e os restantes valores são apenas códigos de turno."
    )
    st.info(
        "A responsabilidade não é escolhida nesta tabela. Será atribuída pelo "
        "gerador de escala de acordo com as necessidades e as competências."
    )

    dias_do_mes = [str(dia) for dia in range(1, num_dias + 1)]
    codigos_turno = codigos_turno_ordenados(st.session_state.turnos)
    assinatura_preferencias = (
        mes_sel,
        ano_sel,
        tuple(st.session_state.trabalhadores),
        tuple(codigos_turno),
    )
    preferencias_atuais = st.session_state.get("preferencias_turnos")

    if (
        not isinstance(preferencias_atuais, pd.DataFrame)
        or st.session_state.get("assinatura_preferencias")
        != assinatura_preferencias
    ):
        preferencias_turnos = pd.DataFrame(
            "",
            index=st.session_state.trabalhadores,
            columns=dias_do_mes,
        )
        indisponibilidades_antigas = st.session_state.get("indisponibilidades", {})
        if isinstance(indisponibilidades_antigas, dict):
            for trabalhador, dias_folga in indisponibilidades_antigas.items():
                if trabalhador in preferencias_turnos.index:
                    for dia in dias_folga:
                        coluna_dia = str(dia)
                        if coluna_dia in preferencias_turnos.columns:
                            preferencias_turnos.loc[trabalhador, coluna_dia] = "F"
        st.session_state.preferencias_turnos = preferencias_turnos
        st.session_state.assinatura_preferencias = assinatura_preferencias

    opcoes_tabela = ["", "F", *codigos_turno]
    preferencias_com_fins_de_semana = estilizar_fins_de_semana(
        st.session_state.preferencias_turnos,
        mes_sel,
        ano_sel,
        dias_nas_colunas=True,
    )
    st.session_state.preferencias_turnos = st.data_editor(
        preferencias_com_fins_de_semana,
        width="stretch",
        num_rows="fixed",
        column_config={
            dia: st.column_config.SelectboxColumn(
                dia,
                options=opcoes_tabela,
                required=False,
                help="Escolha F para folga ou um código de turno.",
            )
            for dia in dias_do_mes
        },
        key="editor_preferencias_turnos",
    )
    st.caption(
        "Os sábados e domingos têm um preenchimento suave. A responsabilidade "
        "será atribuída apenas na geração da escala."
    )

# 5. Gerar escala
with tabs[4]:
    st.header("Gerador de Escala Automática")
    st.caption(
        "O gerador pode atribuir M12 e T24 ao mesmo trabalhador no mesmo dia, "
        "permitindo uma jornada das 08:00 às 20:00. Quando possível, privilegia "
        "responsabilidades diferentes e os trabalhadores que indicaram essa "
        "preferência, respeitando o limite mensal definido para cada pessoa."
    )
    if st.button("⚡ Gerar Escala Optimizada", type="primary"):
        trabalhadores = [
            nome
            for nome in st.session_state.trabalhadores
            if st.session_state.trabalhadores_ativos.get(nome, True)
        ]
        turnos = normalizar_turnos(st.session_state.turnos)
        df_nec = st.session_state.get("df_necessidades")

        if not st.session_state.trabalhadores:
            st.error("Adicione trabalhadores antes de gerar a escala.")
        elif not trabalhadores:
            st.error("Ative pelo menos um trabalhador antes de gerar a escala.")
        elif turnos.empty or df_nec is None:
            st.error("Configure as responsabilidades, os turnos e as necessidades mensais.")
        else:
            slots = turnos_ordenados(turnos)
            model = cp_model.CpModel()
            escala = {}

            for trabalhador in trabalhadores:
                for dia in range(1, num_dias + 1):
                    for slot_index, _slot in enumerate(slots):
                        escala[(trabalhador, dia, slot_index)] = model.NewBoolVar(
                            f"e_{trabalhador}_{dia}_{slot_index}"
                        )

            # Impedir atribuições a responsabilidades ou turnos não autorizados.
            for trabalhador in trabalhadores:
                for dia in range(1, num_dias + 1):
                    for slot_index, slot in enumerate(slots):
                        if not trabalhador_pode_fazer(
                            trabalhador, slot.Responsabilidade, slot.Turno
                        ):
                            model.Add(escala[(trabalhador, dia, slot_index)] == 0)

            # Um trabalhador pode fazer M12 e T24 no mesmo dia, mas não pode
            # acumular outros turnos nem repetir um dos turnos alargados.
            pares_jornada = []
            pares_jornada_por_trabalhador = {
                trabalhador: [] for trabalhador in trabalhadores
            }
            pares_jornada_preferidos = []
            pares_responsabilidades_diferentes = []
            for trabalhador in trabalhadores:
                for dia in range(1, num_dias + 1):
                    variaveis_m12 = [
                        escala[(trabalhador, dia, slot_index)]
                        for slot_index, slot in enumerate(slots)
                        if slot.Turno == "M12"
                    ]
                    variaveis_t24 = [
                        escala[(trabalhador, dia, slot_index)]
                        for slot_index, slot in enumerate(slots)
                        if slot.Turno == "T24"
                    ]
                    variaveis_outros_turnos = [
                        escala[(trabalhador, dia, slot_index)]
                        for slot_index, slot in enumerate(slots)
                        if not e_turno_de_jornada_alargada(slot.Turno)
                    ]

                    model.Add(sum(variaveis_m12) <= 1)
                    model.Add(sum(variaveis_t24) <= 1)
                    model.Add(sum(variaveis_outros_turnos) <= 1)
                    model.Add(sum(variaveis_m12) + sum(variaveis_outros_turnos) <= 1)
                    model.Add(sum(variaveis_t24) + sum(variaveis_outros_turnos) <= 1)

                    for m12_index, m12_slot in enumerate(slots):
                        if m12_slot.Turno != "M12":
                            continue
                        for t24_index, t24_slot in enumerate(slots):
                            if t24_slot.Turno != "T24":
                                continue
                            par = model.NewBoolVar(
                                f"par_m12_t24_{trabalhador}_{dia}_"
                                f"{m12_index}_{t24_index}"
                            )
                            m12_atribuido = escala[(trabalhador, dia, m12_index)]
                            t24_atribuido = escala[(trabalhador, dia, t24_index)]
                            model.Add(par <= m12_atribuido)
                            model.Add(par <= t24_atribuido)
                            model.Add(
                                par >= m12_atribuido + t24_atribuido - 1
                            )
                            pares_jornada.append(par)
                            pares_jornada_por_trabalhador[trabalhador].append(par)
                            if st.session_state.preferencias_jornadas_12h.get(
                                trabalhador, False
                            ):
                                pares_jornada_preferidos.append(par)
                            if (
                                m12_slot.Responsabilidade
                                != t24_slot.Responsabilidade
                            ):
                                pares_responsabilidades_diferentes.append(par)

            for trabalhador in trabalhadores:
                limite_12h = max(
                    0,
                    int(
                        st.session_state.limites_jornadas_12h.get(
                            trabalhador, LIMITE_JORNADAS_12H_PADRAO
                        )
                    ),
                )
                model.Add(
                    sum(pares_jornada_por_trabalhador[trabalhador])
                    <= limite_12h
                )

            # Cumprir a necessidade de cada responsabilidade e turno.
            for dia in range(1, num_dias + 1):
                for slot_index, slot in enumerate(slots):
                    coluna = chave_coluna_turno(slot)
                    necessidade = int(df_nec.loc[str(dia), coluna])
                    model.Add(
                        sum(
                            escala[(trabalhador, dia, slot_index)]
                            for trabalhador in trabalhadores
                        )
                        == necessidade
                    )

            # Respeitar folgas e turnos escolhidos na tabela da equipa.
            preferencias_turnos = st.session_state.get("preferencias_turnos")
            for trabalhador in trabalhadores:
                for dia in range(1, num_dias + 1):
                    preferencia = ""
                    coluna_dia = str(dia)
                    if (
                        isinstance(preferencias_turnos, pd.DataFrame)
                        and trabalhador in preferencias_turnos.index
                        and coluna_dia in preferencias_turnos.columns
                    ):
                        preferencia = str(
                            preferencias_turnos.loc[trabalhador, coluna_dia]
                        ).strip()

                    if preferencia == "F":
                        for slot_index in range(len(slots)):
                            model.Add(
                                escala[(trabalhador, dia, slot_index)] == 0
                            )
                    elif preferencia:
                        for slot_index, slot in enumerate(slots):
                            if slot.Turno != preferencia:
                                model.Add(
                                    escala[(trabalhador, dia, slot_index)] == 0
                                )

            # Evitar turno que atravessa a meia-noite seguido de um turno matinal.
            for trabalhador in trabalhadores:
                for dia in range(1, num_dias):
                    for slot_index, slot in enumerate(slots):
                        if not turno_atravessa_meia_noite(slot.Início, slot.Fim):
                            continue
                        for proximo_index, proximo_slot in enumerate(slots):
                            try:
                                hora_inicio = datetime.strptime(
                                    proximo_slot.Início, "%H:%M"
                                ).hour
                            except ValueError:
                                continue
                            if hora_inicio < 12:
                                model.Add(
                                    escala[(trabalhador, dia, slot_index)]
                                    + escala[(trabalhador, dia + 1, proximo_index)]
                                    <= 1
                                )

            if pares_jornada:
                model.Maximize(
                    1_000_000 * sum(pares_jornada_preferidos)
                    + 1_000 * sum(pares_jornada)
                    + sum(pares_responsabilidades_diferentes)
                )

            solver = cp_model.CpSolver()
            solver.parameters.max_time_in_seconds = 20.0
            status = solver.Solve(model)

            if status in (cp_model.OPTIMAL, cp_model.FEASIBLE):
                st.success("Escala gerada com sucesso!")
                dados_escala = {}

                for trabalhador in trabalhadores:
                    linha_trabalhador = []
                    for dia in range(1, num_dias + 1):
                        atribuicao = "F"
                        atribuicoes_do_dia = []
                        for slot_index, slot in enumerate(slots):
                            if solver.Value(escala[(trabalhador, dia, slot_index)]):
                                atribuicoes_do_dia.append(
                                    f"{slot.Responsabilidade}\n{slot.Turno}"
                                )
                        if atribuicoes_do_dia:
                            atribuicao = "\n\n".join(atribuicoes_do_dia)
                        linha_trabalhador.append(atribuicao)
                    dados_escala[trabalhador] = linha_trabalhador

                colunas_dias = [str(dia) for dia in range(1, num_dias + 1)]
                df_resultado = pd.DataFrame.from_dict(
                    dados_escala, orient="index", columns=colunas_dias
                )
                st.caption("Sábados e domingos aparecem com um preenchimento suave.")
                st.dataframe(
                    estilizar_fins_de_semana(
                        df_resultado,
                        mes_sel,
                        ano_sel,
                        dias_nas_colunas=True,
                    ),
                    width="stretch",
                )

                csv = df_resultado.to_csv().encode("utf-8")
                st.download_button(
                    "📥 Descarregar Escala (CSV)",
                    csv,
                    "escala_mensal.csv",
                    "text/csv",
                )
            else:
                st.error(
                    "Não foi possível encontrar uma solução válida com as "
                    "restrições, necessidades e trabalhadores disponíveis. "
                    "Reduza as necessidades ou adicione mais pessoal."
                )
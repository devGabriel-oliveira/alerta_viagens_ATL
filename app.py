"""
app.py - Torre de controle de viagens (interface web)

Uso:
    streamlit run app.py

Simulação: faça upload de uma planilha, o resultado é processado automaticamente.
Produção : lê o relatório atual do disco e processa quando você clicar em Processar.
"""

import io
import json
import os
from datetime import datetime
from pathlib import Path

import pandas as pd
import pydeck as pdk
import streamlit as st

import monitor_rotas as m

# ------------------------------------------------------------------ constantes
STATUS_LEGIVEL = {
    "PENDENTE": "Pendente",
    "CONCLUIDA": "Concluída",
    "PULADA": "Pulada",
    "CONCLUIDA_FORA_DE_ORDEM": "Concluída fora de ordem",
}
COR_SITUACAO = {
    "desvio": [192, 57, 43],     # vermelho
    "sentido": [217, 130, 43],   # laranja
    "normal": [31, 78, 121],     # azul
    "finalizada": [140, 150, 160],  # cinza
}
LEGENDA_MAPA = ("🔵 Em rota &nbsp;&nbsp; 🔴 Fora de rota &nbsp;&nbsp; "
                "🟠 Sentido errado &nbsp;&nbsp; ⚪ Finalizada")

st.set_page_config(page_title="Torre de controle de viagens", page_icon="🚛", layout="wide")
m.configurar_log()


def estado_sessao():
    if "estado_monitor" not in st.session_state:
        st.session_state.estado_monitor = {"_viagens": {}, "_avisos": {},
                                            "_alertas": [], "_historico": []}
    return st.session_state.estado_monitor


def limpar_estado():
    st.session_state.pop("estado_monitor", None)
    st.session_state.pop("_ultimo_upload", None)


@st.cache_resource
def geocodificador():
    return m.Geocodificador(m.ARQUIVO_MUNICIPIOS)


def processar_upload(arquivo):
    """Processa um upload sem manter arquivos temporários no disco."""
    # Grava em arquivo temporário porque pd.read_excel(sheet_name=None) precisa de path/buffer.
    # Usar BytesIO seria mais elegante, mas o monitor_rotas grava internamente para reaproveitar
    # o mesmo pipeline de produção; manter compatibilidade.
    tmp = Path(f"_upload_{os.getpid()}.xlsx")
    try:
        tmp.write_bytes(arquivo.getvalue())
        m.modo_simulacao(forcar_modo_teste=not st.session_state.get("envio_real", False))
        m.usar_estado_memoria(estado_sessao())
        return m.sim_processar_foto(tmp)
    finally:
        m.usar_estado_memoria(None)
        tmp.unlink(missing_ok=True)


def processar_producao():
    m.modo_producao()
    m.usar_estado_memoria(None)
    return m.processar()


def para_excel(abas):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        for nome, df in abas.items():
            df.to_excel(w, sheet_name=nome[:31], index=False)
    return buf.getvalue()


# ============================================================================
# BARRA LATERAL
# ============================================================================
with st.sidebar:
    st.title("🚛 Monitor de Rotas")
    modo = st.radio("Modo", ["Simulação", "Produção"], horizontal=True)
    st.divider()

    st.subheader("📧 E-mail de alerta")
    destinatario = st.text_input("Enviar para", value="gabrielos@vix.com.br")
    m.EMAIL_DESTINATARIOS = [destinatario]

    envio_real = st.toggle("Enviar e-mails reais", value=False,
                           help="Desligado: os e-mails ficam apenas registrados na tela.")
    st.session_state["envio_real"] = envio_real
    m.MODO_TESTE = not envio_real

    if envio_real:
        if m.EMAIL_METODO == "smtp":
            usuario = st.text_input("Conta", value=os.getenv("ALERTA_EMAIL_USUARIO", ""))
            senha = st.text_input("Senha", type="password",
                                  value=os.getenv("ALERTA_EMAIL_SENHA", ""))
            m.EMAIL_USUARIO, m.EMAIL_SENHA = usuario, senha
        else:
            st.caption("Envio pelo Outlook aberto na máquina.")

    if st.button("Enviar e-mail de teste", use_container_width=True):
        ok, msg = m.enviar_email_teste(destinatario)
        (st.success if ok else st.error)(msg)

    st.divider()

    if modo == "Simulação":
        st.subheader("📥 Testar com planilha")
        st.caption("Faça upload — o processamento é automático.")
        arquivo = st.file_uploader("Planilha", type=["xlsx"], label_visibility="collapsed")
        # Reprocessa sempre que o upload muda (nome+tamanho como chave); botão Limpar reseta
        upload_key = f"{arquivo.name}:{arquivo.size}" if arquivo else None
        if arquivo and upload_key != st.session_state.get("_ultimo_upload"):
            with st.spinner(f"Processando {arquivo.name}..."):
                try:
                    n = processar_upload(arquivo)
                    st.session_state["_ultimo_upload"] = upload_key
                    est = estado_sessao()
                    st.success(
                        f"✅ {n} posições processadas.\n\n"
                        f"{len(est['_alertas'])} alerta(s) no total."
                    )
                except (ValueError, KeyError, OSError) as e:
                    st.error(f"❌ Erro ao processar planilha: {e}")

        if st.button("🗑️ Limpar tudo", use_container_width=True):
            limpar_estado()
            st.rerun()
    else:
        st.subheader("🔄 Processar relatório")
        rel = Path(m.ARQUIVO_RELATORIO)
        if rel.exists():
            atualizado = datetime.fromtimestamp(rel.stat().st_mtime)
            st.caption(f"📄 `{rel.name}` — {atualizado:%d/%m/%Y %H:%M}")
            if st.button("Processar agora", use_container_width=True, type="primary"):
                with st.spinner("Processando..."):
                    n = processar_producao()
                st.success(f"✅ {n} posições novas.")
        else:
            st.warning(f"`{m.ARQUIVO_RELATORIO}` não encontrado.")


# ============================================================================
# CARREGAR DADOS
# ============================================================================
if modo == "Simulação":
    estado = estado_sessao()
    alertas = pd.DataFrame(estado.get("_alertas", []))
    historico = pd.DataFrame(estado.get("_historico", []))
else:
    p = Path(m.ARQUIVO_ESTADO)
    estado = json.loads(p.read_text(encoding="utf-8")) if p.exists() else {
        "_viagens": {}, "_avisos": {}}
    p_al = Path(m.ARQUIVO_ALERTAS)
    alertas = pd.read_csv(p_al, sep=";", encoding="utf-8-sig") if p_al.exists() else pd.DataFrame()
    p_h = Path(m.ARQUIVO_HISTORICO)
    historico = pd.read_csv(p_h, sep=";", encoding="utf-8-sig") if p_h.exists() else pd.DataFrame()

if not alertas.empty:
    alertas["Data/hora"] = pd.to_datetime(alertas["Data/hora"], errors="coerce")
    alertas = alertas.sort_values("Data/hora", ascending=False)

viagens = estado.get("_viagens", {})
situacao = m.montar_situacao(estado) if viagens else pd.DataFrame()

# ============================================================================
# CABEÇALHO
# ============================================================================
st.title("Torre de Controle de Viagens")

if not viagens:
    if modo == "Simulação":
        st.info("👈 Faça upload de uma planilha na barra lateral para começar.")
    else:
        st.info("👈 Clique em **Processar agora** na barra lateral.")
    st.stop()

em_andamento = situacao[situacao["Situação monitor"] == "EM ANDAMENTO"]
com_desvio = em_andamento[em_andamento["Em desvio de rota"] == "SIM"]
com_sentido = em_andamento[em_andamento["Em sentido incorreto"] == "SIM"]
finalizadas = situacao[situacao["Situação monitor"] == "FINALIZADA"]

c1, c2, c3, c4, c5 = st.columns(5)
c1.metric("🚛 Em andamento", len(em_andamento))
c2.metric("🚨 Alertas", len(alertas))
c3.metric("⚠️ Fora de rota", len(com_desvio))
c4.metric("↩️ Sentido errado", len(com_sentido))
c5.metric("✅ Finalizadas", len(finalizadas))

st.download_button(
    "📊 Exportar tudo (Excel)",
    data=para_excel({"Viagens": situacao, "Alertas": alertas}),
    file_name=f"monitor_{datetime.now():%Y%m%d_%H%M}.xlsx",
)

# ============================================================================
# ABAS
# ============================================================================
aba_alertas, aba_viagens, aba_mapa, aba_detalhe = st.tabs(
    ["🚨 Alertas", "🚛 Viagens", "🗺️ Mapa", "🔍 Detalhe"]
)

with aba_alertas:
    if alertas.empty:
        st.success("✅ Nenhum alerta registrado.")
    else:
        c1, c2 = st.columns([1, 3])
        tipos_disp = sorted(alertas["Tipo"].unique())
        tipos = c1.multiselect("Tipo", tipos_disp, default=tipos_disp)
        busca = c2.text_input("🔎 Buscar por placa ou motorista", "")

        vis = alertas[alertas["Tipo"].isin(tipos)]
        if busca:
            b = busca.upper()
            vis = vis[vis[["Placa Cavalo", "Motorista"]].astype(str).apply(
                lambda s: s.str.upper().str.contains(b, regex=False)).any(axis=1)]

        st.caption(f"{len(vis)} alerta(s)")
        st.dataframe(
            vis.drop(columns=["Latitude", "Longitude"], errors="ignore"),
            hide_index=True, use_container_width=True,
            column_config={
                "Google Maps": st.column_config.LinkColumn("Mapa", display_text="Ver 🗺️"),
                "Data/hora": st.column_config.DatetimeColumn(format="DD/MM HH:mm"),
            }
        )

with aba_viagens:
    c1, c2, c3 = st.columns([1, 1, 2])
    sits = sorted(situacao["Situação monitor"].unique())
    sit_sel = c1.multiselect("Situação", sits, default=["EM ANDAMENTO"])
    so_alerta = c2.toggle("Só com alertas", value=False)
    busca = c3.text_input("🔎 Buscar", "")

    vis = situacao[situacao["Situação monitor"].isin(sit_sel)]
    if so_alerta:
        vis = vis[vis["Qtde alertas"] > 0]
    if busca:
        b = busca.upper()
        vis = vis[vis[["Placa Cavalo", "Placa Carreta", "Motorista", "Destino final"]]
                  .astype(str).apply(lambda s: s.str.upper().str.contains(b, regex=False)).any(axis=1)]

    st.caption(f"{len(vis)} viagem(ns)")
    st.dataframe(
        vis.drop(columns=["Latitude", "Longitude"], errors="ignore"),
        hide_index=True, use_container_width=True,
        column_config={"Google Maps": st.column_config.LinkColumn("Mapa", display_text="Ver 🗺️")},
    )

with aba_mapa:
    pts = situacao.dropna(subset=["Latitude", "Longitude"]).copy()

    def cor(r):
        if r["Situação monitor"] != "EM ANDAMENTO":
            return COR_SITUACAO["finalizada"]
        if r["Em desvio de rota"] == "SIM":
            return COR_SITUACAO["desvio"]
        if r["Em sentido incorreto"] == "SIM":
            return COR_SITUACAO["sentido"]
        return COR_SITUACAO["normal"]

    if not pts.empty:
        pts["cor"] = pts.apply(cor, axis=1)
        pts["raio"] = pts["cor"].map(
            lambda c: 12000 if c in (COR_SITUACAO["desvio"], COR_SITUACAO["sentido"]) else 5000
        )
        st.markdown("🔵 Em rota &nbsp;&nbsp; 🔴 Fora de rota &nbsp;&nbsp; "
                    "🟠 Sentido errado &nbsp;&nbsp; ⚪ Finalizada")
        st.pydeck_chart(pdk.Deck(
            map_style="light",
            initial_view_state=pdk.ViewState(latitude=-18, longitude=-47, zoom=3.6),
            layers=[pdk.Layer("ScatterplotLayer", pts,
                              get_position=["Longitude", "Latitude"],
                              get_fill_color="cor", get_radius="raio",
                              radius_min_pixels=4, pickable=True)],
            tooltip={"text": "{Placa Cavalo} - {Motorista}\n{Situação monitor}\nPróxima: {Próxima parada}\n{Local}"},
        ), height=580)
    else:
        st.info("Sem posições válidas para exibir.")

with aba_detalhe:
    opcoes = situacao.sort_values("Qtde alertas", ascending=False)
    rotulos = {
        f"{r['Placa Cavalo']} · {str(r['Motorista'])[:30]} · {r['Destino final']} "
        f"({r['Qtde alertas']} alertas)": r["Viagem"]
        for _, r in opcoes.iterrows()
    }
    escolha = st.selectbox("Selecione a viagem", list(rotulos))
    vid = rotulos[escolha]
    est = viagens[vid]
    nomes = est.get("nomes_paradas") or est["assinatura"].split("|")

    col_a, col_b = st.columns([1, 2])
    with col_a:
        st.markdown(f"**Situação:** {est['situacao']}")
        st.markdown(f"**Viagem:** `{vid}`")
        st.markdown(f"**Carreta:** {est['info'].get('carreta', '-')}")
        paradas_df = pd.DataFrame({
            "Ordem": range(1, len(nomes) + 1),
            "Cidade": nomes,
            "Status": [STATUS_LEGIVEL.get(x, x) for x in est["status_paradas"]],
        })
        st.dataframe(paradas_df, hide_index=True, use_container_width=True)

        al_v = alertas[alertas["Viagem"] == vid] if not alertas.empty else pd.DataFrame()
        if not al_v.empty:
            st.markdown(f"**Alertas ({len(al_v)})**")
            st.dataframe(
                al_v[["Data/hora", "Ocorrência", "Detalhes"]],
                hide_index=True, use_container_width=True,
                column_config={"Data/hora": st.column_config.DatetimeColumn(format="DD/MM HH:mm")},
            )

    with col_b:
        hist_v = historico[historico["viagem"] == vid] if not historico.empty else pd.DataFrame()
        geo = geocodificador()
        stops = pd.DataFrame([
            {"nome": n, "status": STATUS_LEGIVEL.get(s, s), "lat": p[0], "lon": p[1]}
            for n, s in zip(nomes, est["status_paradas"], strict=False)
            if (p := geo.buscar(n))
        ])
        camadas = []
        if len(hist_v) > 1:
            camadas.append(pdk.Layer(
                "PathLayer",
                pd.DataFrame({"path": [[[r.lon, r.lat] for r in hist_v.itertuples()]]}),
                get_path="path", get_color=[31, 78, 121], width_min_pixels=3,
            ))
        if not stops.empty:
            camadas.append(pdk.Layer(
                "ScatterplotLayer", stops,
                get_position=["lon", "lat"], get_radius=12000,
                get_fill_color=[125, 60, 152, 90], get_line_color=[125, 60, 152],
                stroked=True, line_width_min_pixels=2, pickable=True,
            ))
        if not al_v.empty:
            camadas.append(pdk.Layer(
                "ScatterplotLayer", al_v,
                get_position=["Longitude", "Latitude"], get_radius=3000,
                radius_min_pixels=6, get_fill_color=[192, 57, 43], pickable=True,
            ))

        centro = (hist_v.iloc[-1] if len(hist_v) else
                  stops.iloc[0] if not stops.empty else None)
        st.pydeck_chart(pdk.Deck(
            map_style="light", layers=camadas,
            initial_view_state=pdk.ViewState(
                latitude=float(centro["lat"]) if centro is not None else -20,
                longitude=float(centro["lon"]) if centro is not None else -43,
                zoom=6,
            ),
            tooltip={"text": "{nome} {status}{Ocorrência}"},
        ), height=540)
        st.caption("Azul: trajeto · Roxo: paradas · Vermelho: alertas")

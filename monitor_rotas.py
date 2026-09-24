"""
monitor_rotas.py  (v3 - relatório Gestão de Viagens com paradas)

Entrada: relatório Excel no layout da aba "Consulta" (Gestão_Viagens.xlsx), com uma linha por
viagem e as colunas extras parada1, parada2, ... (aceita também "Parada 1", "PARADA 2") O destino final é a coluna "UF Destino".
Paradas e destino vêm como "Cidade/UF" (ex.: "Itaboraí/RJ").

O que faz a cada execução (rodar a cada 10 min):
  1. Lê as viagens "Em Trânsito" e a última posição de cada uma.
  2. Converte as cidades em coordenadas (tabela IBGE municipios_br.csv + CIDADES_EXTRA).
  3. Obtém o traçado da estrada (OSRM, uma vez por viagem, guardado em cache).
  4. Verifica desvio de rota, sentido incorreto e sequência das paradas.
  5. Quando o veículo para no destino final, a viagem é FINALIZADA e sai do monitoramento.
  6. Envia alertas por e-mail e registra tudo em arquivos lidos pela interface (app.py).

Uso:
    python monitor_rotas.py                      # processa o relatório atual
    python monitor_rotas.py --loop               # repete a cada INTERVALO_LOOP_MIN
    python monitor_rotas.py --reset              # zera o estado antes
    python monitor_rotas.py --simular snapshots  # teste com fotos simuladas (não usa internet)
"""

import argparse
import json
import logging
import math
import os
import re
import shutil
import smtplib
import sys
import time
import unicodedata
import urllib.request
from urllib.error import HTTPError, URLError
from datetime import datetime, timedelta
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path

import pandas as pd

# =============================================================================
# 1. CONFIGURAÇÃO
# =============================================================================
ARQUIVO_RELATORIO = "Gestão_Viagens.xlsx"
ABA_RELATORIO = "Consulta"
ARQUIVO_MUNICIPIOS = "municipios_br.csv"        # nome, uf, latitude, longitude (IBGE)

ARQUIVO_ESTADO = "estado_monitor.json"
ARQUIVO_CACHE_ROTAS = "rotas_cache.json"
ARQUIVO_HISTORICO = "historico_posicoes.csv"
ARQUIVO_ALERTAS = "alertas.csv"                  # histórico de alertas (lido pela interface)
ARQUIVO_SITUACAO = "situacao_viagens.xlsx"
GERAR_EXCEL_SITUACAO = False                    # a interface exporta Excel sob demanda
PASTA_ALERTAS = "alertas_gerados"
ARQUIVO_LOG = "monitor_rotas.log"

COL = {
    "placa": "Placa Cavalo",
    "proprietario": "Proprietário Cavalo",
    "frota": "Frota",
    "rastreador": "Nome rastreador",
    "viagem": "Última viagem",
    "status": "Status viagem",
    "carreta": "Placa Carreta",
    "motorista": "Motorista",
    "destino": "UF Destino",
    "datahora": "Data /Hora Posição",
    "local": "Local última posição",
    "latlong": "Lat long",
}
PADRAO_COLUNA_PARADA = r"^\s*parada\s*(\d+)\s*$"   # "Parada 1", "PARADA 2", "Parada3"...
STATUS_MONITORADOS = {"EM TRANSITO"}               # sem acento, maiúsculo

# Cidades fora da tabela do IBGE (exterior). Chave: "CIDADE/UF" sem acento, maiúsculo.
CIDADES_EXTRA = {
    "ZARATE/EX": (-34.0981, -59.0286),
    "SANTIAGO/EX": (-33.4489, -70.6693),
    "SANTA CRUZ DE LA SIERRA/EX": (-17.7833, -63.1821),
    "ASSUNCAO/EX": (-25.2637, -57.5759),
}

# Traçado da estrada (OSRM). O servidor público é de demonstração: para produção, avaliar servidor próprio.
USAR_OSRM = True
OSRM_URL = "https://router.project-osrm.org"
MAX_CONSULTAS_OSRM_POR_EXECUCAO = 60              # espalha as consultas da primeira carga
ESPACAMENTO_TRACADO_KM = 1.0                       # simplifica o traçado (1 ponto por km)

# Regras
RAIO_ROTA_KM = 10.0             # distância máxima do traçado
RAIO_CHEGADA_KM = 10.0          # raio em torno do centro da cidade (paradas vêm como cidade, não endereço)
PARADO_MAX_KM = 1.0             # deslocamento máximo entre leituras para considerar "parado"
LEITURAS_CHEGADA = 2            # leituras seguidas parado dentro do raio = parada confirmada
MIN_DESLOCAMENTO_KM = 0.5       # abaixo disso não avalia sentido
MIN_AFASTAMENTO_KM = 1.0
ANGULO_SENTIDO_ERRADO = 90.0
LEITURAS_SENTIDO = 2
COOLDOWN_ALERTA_MIN = 60
POSICAO_DESATUALIZADA_H = 6     # só para aviso no log

# E-mail
MODO_TESTE = True               # True = salva o HTML em PASTA_ALERTAS e não envia
EMAIL_METODO = "outlook"        # "outlook" = Outlook aberto na máquina | "smtp" = Office 365 com usuário e senha
SMTP_SERVIDOR = "smtp.office365.com"
SMTP_PORTA = 587
EMAIL_USUARIO = os.getenv("ALERTA_EMAIL_USUARIO", "monitoramento@empresa.com.br")
EMAIL_SENHA = os.getenv("ALERTA_EMAIL_SENHA", "")
EMAIL_DESTINATARIOS = ["gabrielos@vix.com.br"]

INTERVALO_LOOP_MIN = 10

log = logging.getLogger("monitor")


def configurar_log():
    if log.handlers:                    # evita duplicar (a interface importa este módulo)
        return
    log.setLevel(logging.INFO)
    log.propagate = False               # evita linha duplicada no log dentro do Streamlit
    fmt = logging.Formatter("%(asctime)s | %(levelname)-7s | %(message)s", "%Y-%m-%d %H:%M:%S")
    for h in (logging.StreamHandler(sys.stdout), logging.FileHandler(ARQUIVO_LOG, encoding="utf-8")):
        h.setFormatter(fmt)
        log.addHandler(h)

# =============================================================================
# 2. GEOMETRIA
# =============================================================================
R_TERRA_KM = 6371.0088


def distancia_km(a, b):
    la1, lo1, la2, lo2 = map(math.radians, (a[0], a[1], b[0], b[1]))
    h = math.sin((la2 - la1) / 2) ** 2 + math.cos(la1) * math.cos(la2) * math.sin((lo2 - lo1) / 2) ** 2
    return 2 * R_TERRA_KM * math.asin(math.sqrt(h))


def rumo_graus(a, b):
    la1, la2 = math.radians(a[0]), math.radians(b[0])
    dlo = math.radians(b[1] - a[1])
    x = math.sin(dlo) * math.cos(la2)
    y = math.cos(la1) * math.sin(la2) - math.sin(la1) * math.cos(la2) * math.cos(dlo)
    return (math.degrees(math.atan2(x, y)) + 360) % 360


def diferenca_angulo(a, b):
    d = abs(a - b) % 360
    return 360 - d if d > 180 else d


def distancia_ponto_segmento_km(p, a, b):
    kx = R_TERRA_KM * math.cos(math.radians(p[0])) * math.pi / 180
    ky = R_TERRA_KM * math.pi / 180
    ax, ay = (a[1] - p[1]) * kx, (a[0] - p[0]) * ky
    bx, by = (b[1] - p[1]) * kx, (b[0] - p[0]) * ky
    dx, dy = bx - ax, by - ay
    L2 = dx * dx + dy * dy
    t = 0.0 if L2 == 0 else max(0.0, min(1.0, -(ax * dx + ay * dy) / L2))
    return math.hypot(ax + t * dx, ay + t * dy)


def distancia_tracado_km(p, pontos):
    if len(pontos) == 1:
        return distancia_km(p, pontos[0])
    return min(distancia_ponto_segmento_km(p, pontos[i], pontos[i + 1]) for i in range(len(pontos) - 1))


def link_maps(lat, lon):
    return f"https://www.google.com/maps?q={lat:.6f},{lon:.6f}"

# =============================================================================
# 3. CIDADES -> COORDENADAS
# =============================================================================
def normalizar(txt):
    txt = unicodedata.normalize("NFKD", str(txt)).encode("ascii", "ignore").decode().upper()
    txt = txt.replace("-", " ").replace("'", " ")
    return " ".join(txt.split())


def chave_cidade(texto):
    """'Itaboraí/RJ' -> 'ITABORAI/RJ'. Devolve None se não estiver no formato Cidade/UF."""
    if texto is None or pd.isna(texto) or "/" not in str(texto):
        return None
    cidade, uf = str(texto).rsplit("/", 1)
    cidade, uf = normalizar(cidade), normalizar(uf)
    return f"{cidade}/{uf}" if cidade and uf else None


class Geocodificador:
    def __init__(self, arquivo):
        df = pd.read_csv(arquivo, encoding="utf-8")
        self.tabela = {f"{normalizar(n)}/{u}": (float(la), float(lo))
                       for n, u, la, lo in zip(df["nome"], df["uf"], df["latitude"], df["longitude"])}
        self.tabela.update(CIDADES_EXTRA)

    def buscar(self, texto):
        k = chave_cidade(texto)
        return self.tabela.get(k) if k else None

# =============================================================================
# 4. LEITURA DO RELATÓRIO
# =============================================================================
def parse_latlong(valor):
    if valor is None or pd.isna(valor):
        return None
    partes = str(valor).replace(";", ",").split(",")
    if len(partes) != 2:
        return None
    try:
        lat, lon = float(partes[0]), float(partes[1])
    except ValueError:
        return None
    if not (-90 <= lat <= 90 and -180 <= lon <= 180) or (lat == 0 and lon == 0):
        return None
    return lat, lon


def texto(v):
    return "" if v is None or pd.isna(v) else str(v).strip()


def ler_relatorio():
    for n in range(1, 4):
        try:
            df = pd.read_excel(ARQUIVO_RELATORIO, sheet_name=ABA_RELATORIO)
            break
        except PermissionError:
            log.warning(f"{ARQUIVO_RELATORIO} em uso (tentativa {n}/3)...")
            time.sleep(10)
    else:
        raise RuntimeError(f"Não foi possível abrir {ARQUIVO_RELATORIO}")

    faltando = [c for c in COL.values() if c not in df.columns]
    if faltando:
        raise ValueError(f"Colunas não encontradas no relatório: {faltando}")

    # colunas de parada, ordenadas pelo número
    cols_paradas = sorted([c for c in df.columns if re.match(PADRAO_COLUNA_PARADA, str(c), re.I)],
                          key=lambda c: int(re.match(PADRAO_COLUNA_PARADA, str(c), re.I).group(1)))

    df = df[df[COL["status"]].map(lambda s: normalizar(texto(s))).isin(STATUS_MONITORADOS)].copy()
    df["_pos"] = df[COL["latlong"]].map(parse_latlong)
    df["_ts"] = pd.to_datetime(df[COL["datahora"]], dayfirst=True, errors="coerce")
    df["_viagem"] = df[COL["viagem"]].map(texto)
    return df, cols_paradas


def montar_roteiro(r, cols_paradas, geo):
    """Lista de paradas na ordem (a última é o destino). Devolve (paradas, cidades_nao_encontradas)."""
    destino = texto(r[COL["destino"]])
    nomes = [texto(r[c]) for c in cols_paradas if texto(r[c])]
    nomes = [n for n in nomes if chave_cidade(n) != chave_cidade(destino)]  # parada igual ao destino
    nomes.append(destino)
    paradas, nao_encontradas = [], []
    for i, nome in enumerate(nomes):
        pos = geo.buscar(nome)
        if pos is None:
            nao_encontradas.append(nome or "(vazio)")
            continue
        paradas.append({"nome": nome, "tipo": "DESTINO" if i == len(nomes) - 1 else "PARADA", "pos": pos})
    return paradas, nao_encontradas


def assinatura_roteiro(nomes):
    return "|".join(chave_cidade(n) or "" for n in nomes)

# =============================================================================
# 5. TRAÇADO (OSRM + cache)
# =============================================================================
class Tracados:
    def __init__(self):
        p = Path(ARQUIVO_CACHE_ROTAS)
        self.cache = json.loads(p.read_text(encoding="utf-8")) if p.exists() else {}
        self.consultas = 0
        self.alterado = False

    def obter(self, vid, origem, paradas):
        ass = assinatura_roteiro([p["nome"] for p in paradas])
        c = self.cache.get(vid)
        if c and c["assinatura"] == ass:
            return [tuple(x) for x in c["pontos"]]
        if not USAR_OSRM or self.consultas >= MAX_CONSULTAS_OSRM_POR_EXECUCAO or origem is None:
            return None
        self.consultas += 1
        try:
            pontos = self._consultar_osrm([tuple(origem)] + [p["pos"] for p in paradas])
        except (URLError, HTTPError, TimeoutError, RuntimeError, json.JSONDecodeError) as e:
            log.warning(f"OSRM indisponível ({e}) - desvio de rota não avaliado nesta execução")
            self.consultas = MAX_CONSULTAS_OSRM_POR_EXECUCAO  # não insiste nesta execução
            return None
        self.cache[vid] = {"assinatura": ass, "origem": list(origem), "gerado_em": datetime.now().isoformat(),
                           "pontos": [list(p) for p in pontos]}
        self.alterado = True
        time.sleep(1)                                   # respeita o servidor público
        return pontos

    @staticmethod
    def _consultar_osrm(pontos):
        coords = ";".join(f"{lon:.5f},{lat:.5f}" for lat, lon in pontos)
        url = f"{OSRM_URL}/route/v1/driving/{coords}?overview=full&geometries=geojson"
        req = urllib.request.Request(url, headers={"User-Agent": "monitor-rotas/1.0"})
        with urllib.request.urlopen(req, timeout=30) as resp:
            dados = json.load(resp)
        if dados.get("code") != "Ok":
            raise RuntimeError(dados.get("message", dados.get("code")))
        brutos = [(lat, lon) for lon, lat in dados["routes"][0]["geometry"]["coordinates"]]
        # simplifica: mantém 1 ponto a cada ESPACAMENTO_TRACADO_KM
        simples = [brutos[0]]
        for p in brutos[1:-1]:
            if distancia_km(simples[-1], p) >= ESPACAMENTO_TRACADO_KM:
                simples.append(p)
        simples.append(brutos[-1])
        return simples

    def salvar(self):
        if self.alterado:
            Path(ARQUIVO_CACHE_ROTAS).write_text(json.dumps(self.cache), encoding="utf-8")

# =============================================================================
# 6. ESTADO
# =============================================================================
def estado_inicial(paradas, ctx, origem):
    return {
        "info": ctx,
        "assinatura": assinatura_roteiro([p["nome"] for p in paradas]),
        "nomes_paradas": [p["nome"] for p in paradas],
        "origem": list(origem),
        "ultima_leitura": None,
        "ultima_posicao": None,
        "idx_ativo": 0,
        "status_paradas": ["PENDENTE"] * len(paradas),
        "leituras_parado": [0] * len(paradas),
        "contador_afastamento": 0,
        "alertas": {},
        "qtd_alertas": 0,
        "ultimo_alerta": "",
        "em_desvio": {"ROTA": False, "SENTIDO": False},
        "situacao": "EM ANDAMENTO",
    }


# Estado pode ficar em memória (dict) ou em arquivo JSON.
# O Streamlit Cloud apaga arquivos entre execuções, então a interface usa memória.
_ESTADO_MEMORIA = None


def usar_estado_memoria(estado):
    """Faz o monitor ler/gravar num dict em memória em vez de arquivo. None volta ao arquivo."""
    global _ESTADO_MEMORIA
    _ESTADO_MEMORIA = estado


def carregar_estado(reset=False):
    if _ESTADO_MEMORIA is not None:
        if reset:
            _ESTADO_MEMORIA.clear()
        _ESTADO_MEMORIA.setdefault("_viagens", {})
        _ESTADO_MEMORIA.setdefault("_avisos", {})
        return _ESTADO_MEMORIA
    p = Path(ARQUIVO_ESTADO)
    if reset and p.exists():
        p.unlink()
    return json.loads(p.read_text(encoding="utf-8")) if p.exists() else {"_viagens": {}, "_avisos": {}}


def salvar_estado(estado):
    if _ESTADO_MEMORIA is not None:
        return
    tmp = Path(ARQUIVO_ESTADO + ".tmp")
    tmp.write_text(json.dumps(estado, ensure_ascii=False, indent=1, default=str), encoding="utf-8")
    tmp.replace(ARQUIVO_ESTADO)


def avisar_se_mudou(estado, chave, itens, mensagem, intervalo_min=0):
    """Loga avisos agregados só quando a lista muda (e no máximo 1 vez a cada intervalo_min)."""
    itens = sorted(set(itens))
    ultimo = estado["_avisos"].get(chave + "_hora")
    if ultimo and (datetime.now() - datetime.fromisoformat(ultimo)).total_seconds() < intervalo_min * 60:
        return
    if itens != estado["_avisos"].get(chave, []):
        estado["_avisos"][chave] = itens
        estado["_avisos"][chave + "_hora"] = datetime.now().isoformat()
        if itens:
            amostra = ", ".join(itens[:10]) + (" ..." if len(itens) > 10 else "")
            log.warning(f"{mensagem}: {len(itens)} | {amostra}")

# =============================================================================
# 7. E-MAIL
# =============================================================================
def montar_html(ctx, titulo, tipo, ts, pos, detalhes):
    cor = {"ROTA": "#c0392b", "SENTIDO": "#d35400"}.get(tipo, "#8e44ad")
    campos = {
        "Motorista": ctx["motorista"],
        "Cavalo / Carreta": f"{ctx['placa']} / {ctx['carreta']}",
        "Proprietário / Frota": f"{ctx['proprietario']} / {ctx['frota']}",
        "Viagem": ctx["viagem"],
        "Destino final": ctx["destino"],
        "Data/hora da posição": f"{ts:%d/%m/%Y %H:%M}",
        "Local (rastreador)": f"{ctx['local']} ({ctx['rastreador']})",
        "Coordenadas": f"{pos[0]:.5f}, {pos[1]:.5f}",
        **detalhes,
    }
    linhas = "".join(f"<tr><td style='padding:6px 10px;color:#555;white-space:nowrap'>{k}</td>"
                     f"<td style='padding:6px 10px'><b>{v}</b></td></tr>" for k, v in campos.items())
    return f"""<html><body style="font-family:Arial,sans-serif;font-size:14px;color:#222">
<div style="max-width:640px;border:1px solid #ddd">
  <div style="background:{cor};color:#fff;padding:14px 16px;font-size:18px"><b>{titulo}</b></div>
  <table style="border-collapse:collapse;width:100%">{linhas}</table>
  <div style="padding:16px"><a href="{link_maps(*pos)}" style="background:#1a73e8;color:#fff;padding:10px 16px;
     text-decoration:none;border-radius:4px">Ver localização no Google Maps</a></div>
  <div style="padding:10px 16px;color:#888;font-size:12px;border-top:1px solid #eee">
     Alerta automático - Monitor de Rotas</div>
</div></body></html>"""


def enviar_email(assunto, html, destinatarios=None):
    destinatarios = destinatarios or EMAIL_DESTINATARIOS
    if MODO_TESTE:
        Path(PASTA_ALERTAS).mkdir(exist_ok=True)
        nome = f"{datetime.now():%Y%m%d_%H%M%S_%f}.html"
        (Path(PASTA_ALERTAS) / nome).write_text(html, encoding="utf-8")
        log.info(f"   [MODO_TESTE] e-mail salvo em {PASTA_ALERTAS}/{nome}")
        return True
    if EMAIL_METODO == "outlook":
        return _enviar_outlook(assunto, html, destinatarios)
    return _enviar_smtp(assunto, html, destinatarios)


def _enviar_outlook(assunto, html, destinatarios):
    """Envia pelo Outlook instalado e logado nesta máquina (sem senha, sem SMTP)."""
    try:
        import pythoncom
        import win32com.client
        pythoncom.CoInitialize()          # necessário quando roda dentro do Streamlit
        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
            mail = outlook.CreateItem(0)
            mail.To = "; ".join(destinatarios)
            mail.Subject = assunto
            mail.HTMLBody = html
            mail.Send()
        finally:
            pythoncom.CoUninitialize()
        log.info(f"   e-mail enviado pelo Outlook para {', '.join(destinatarios)}")
        return True
    except Exception as e:
        log.error(f"   falha ao enviar pelo Outlook: {e}")
        return False


def _enviar_smtp(assunto, html, destinatarios):
    """Envia via SMTP Office 365 (exige SMTP AUTH liberado na conta)."""
    msg = MIMEMultipart("alternative")
    msg["Subject"], msg["From"], msg["To"] = assunto, EMAIL_USUARIO, ", ".join(destinatarios)
    msg.attach(MIMEText(html, "html", "utf-8"))
    try:
        with smtplib.SMTP(SMTP_SERVIDOR, SMTP_PORTA, timeout=30) as s:
            s.starttls()
            s.login(EMAIL_USUARIO, EMAIL_SENHA)
            s.sendmail(EMAIL_USUARIO, destinatarios, msg.as_string())
        log.info(f"   e-mail enviado por SMTP para {', '.join(destinatarios)}")
        return True
    except Exception as e:
        log.error(f"   falha ao enviar por SMTP: {e}")
        return False


def disparar_alerta(ctx, est, chave, tipo, titulo, ts, pos, detalhes, evento_unico=False):
    ultimo = est["alertas"].get(chave)
    if ultimo:
        if evento_unico:
            return
        if (ts - datetime.fromisoformat(ultimo)).total_seconds() / 60 < COOLDOWN_ALERTA_MIN:
            log.info(f"   ({ctx['placa']}: {titulo} continua - repetição suprimida pelo cooldown)")
            return
    log.warning(f"ALERTA [{ctx['placa']}] {titulo} | {ts:%d/%m %H:%M} | "
                + " | ".join(f"{k}: {v}" for k, v in detalhes.items()))
    html = montar_html(ctx, titulo, tipo, ts, pos, detalhes)
    enviado = enviar_email(f"[ALERTA] {titulo} - {ctx['placa']} - {ctx['motorista']}", html)
    registrar_alerta(ctx, tipo, titulo, ts, pos, detalhes, enviado)
    if enviado:
        est["alertas"][chave] = ts.isoformat()
    est["qtd_alertas"] += 1
    est["ultimo_alerta"] = f"{titulo} ({ts:%d/%m %H:%M})"


def registrar_alerta(ctx, tipo, titulo, ts, pos, detalhes, enviado):
    linha = {"Data/hora": ts, "Tipo": {"ROTA": "Desvio de rota", "SENTIDO": "Sentido incorreto"}.get(tipo, "Sequência"),
             "Ocorrência": titulo, "Placa Cavalo": ctx["placa"], "Placa Carreta": ctx["carreta"],
             "Motorista": ctx["motorista"], "Viagem": ctx["viagem"], "Destino final": ctx["destino"],
             "Local": ctx["local"], "Detalhes": " | ".join(f"{k}: {v}" for k, v in detalhes.items()),
             "Latitude": pos[0], "Longitude": pos[1], "Google Maps": link_maps(*pos),
             "E-mail": ("salvo (modo teste)" if MODO_TESTE else "enviado") if enviado else "FALHOU"}
    if _ESTADO_MEMORIA is not None:
        _ESTADO_MEMORIA.setdefault("_alertas", []).append(linha)
        return
    novo_arq = not Path(ARQUIVO_ALERTAS).exists()
    pd.DataFrame([linha]).to_csv(ARQUIVO_ALERTAS, mode="a", header=novo_arq, index=False, sep=";", encoding="utf-8-sig")


def enviar_email_teste(destinatario):
    """Dispara um alerta fictício para validar o SMTP. Devolve (ok, mensagem)."""
    ctx = {"motorista": "MOTORISTA TESTE", "placa": "TST0A00", "carreta": "TST0B00", "proprietario": "VIX LOGISTICA S/A",
           "frota": "PRÓPRIO", "viagem": "TESTE", "destino": "Ubá/MG", "local": "CARIACICA / ES",
           "rastreador": "Rastreador OnixSat"}
    pos = (-20.2576, -40.3766)
    html = montar_html(ctx, "E-mail de teste do Monitor de Rotas", "ROTA", datetime.now(), pos,
                       {"Observação": "Se você recebeu esta mensagem, o envio de alertas está funcionando."})
    ok = enviar_email("[TESTE] Monitor de Rotas", html, [destinatario])
    if not ok:
        return False, "falha no envio - veja monitor_rotas.log"
    return True, f"salvo em {PASTA_ALERTAS} (modo teste ligado)" if MODO_TESTE else f"enviado para {destinatario}"

# =============================================================================
# 8. VALIDAÇÕES
# =============================================================================
def seq_prevista(paradas):
    return " > ".join(p["nome"] for p in paradas)


def avancar(ctx, est, paradas, ts):
    """Próxima parada pendente. Se não houver (destino atingido), finaliza a viagem."""
    est["contador_afastamento"] = 0
    nxt = est["idx_ativo"] + 1
    while nxt < len(paradas) and est["status_paradas"][nxt] != "PENDENTE":
        nxt += 1
    if nxt >= len(paradas) or est["status_paradas"][-1] != "PENDENTE":
        for i, s in enumerate(est["status_paradas"]):      # destino atingido: o que ficou pendente foi pulado
            if s == "PENDENTE":
                est["status_paradas"][i] = "PULADA"
        est["situacao"] = "FINALIZADA"
        est["finalizada_em"] = ts.isoformat()
        resumo = ", ".join(f"{p['nome']}: {s}" for p, s in zip(paradas, est["status_paradas"]))
        log.info(f"[{ctx['placa']}] VIAGEM FINALIZADA às {ts:%d/%m %H:%M} (destino atingido). {resumo}")
    else:
        est["idx_ativo"] = nxt
        log.info(f"[{ctx['placa']}] próxima parada: {paradas[nxt]['nome']}")


def checar_paradas(ctx, est, paradas, ts, pos):
    """Parada confirmada = LEITURAS_CHEGADA leituras seguidas PARADO dentro do raio da cidade.
    Passar direto pela cidade não conta."""
    ant = est["ultima_posicao"]
    parado = ant is not None and distancia_km(tuple(ant), pos) <= PARADO_MAX_KM
    for i, p in enumerate(paradas):
        dentro = distancia_km(pos, p["pos"]) <= RAIO_CHEGADA_KM
        est["leituras_parado"][i] = est["leituras_parado"][i] + 1 if (parado and dentro) else 0

    def confirmou(i):
        return est["leituras_parado"][i] >= LEITURAS_CHEGADA

    idx = est["idx_ativo"]

    # a) parou na parada ativa
    if confirmou(idx):
        est["status_paradas"][idx] = "CONCLUIDA"
        log.info(f"[{ctx['placa']}] parada em {paradas[idx]['nome']} ({paradas[idx]['tipo']}) confirmada às {ts:%H:%M}")
        avancar(ctx, est, paradas, ts)
        return

    # b) parou numa parada posterior -> pulou as anteriores
    for j in range(idx + 1, len(paradas)):
        if confirmou(j):
            puladas = [i for i in range(idx, j) if est["status_paradas"][i] == "PENDENTE"]
            for i in puladas:
                est["status_paradas"][i] = "PULADA"
            est["status_paradas"][j] = "CONCLUIDA"
            disparar_alerta(ctx, est, f"SEQ_{j}", "SEQUENCIA", "Quebra de Sequência de Entregas", ts, pos, {
                "Parou em": f"{paradas[j]['nome']} ({paradas[j]['tipo'].lower()} {j + 1})",
                "Parada(s) pulada(s)": ", ".join(paradas[i]["nome"] for i in puladas),
                "Sequência prevista": seq_prevista(paradas),
            }, evento_unico=True)
            est["idx_ativo"] = j
            avancar(ctx, est, paradas, ts)
            return

    # c) voltou a uma parada pulada -> inversão
    for i, p in enumerate(paradas):
        if est["status_paradas"][i] == "PULADA" and confirmou(i):
            est["status_paradas"][i] = "CONCLUIDA_FORA_DE_ORDEM"
            disparar_alerta(ctx, est, f"INV_{i}", "SEQUENCIA", "Entrega Fora de Ordem (Inversão)", ts, pos, {
                "Parou em": f"{p['nome']} (parada {i + 1})",
                "Situação": "Parada atendida depois de uma parada posterior",
                "Sequência prevista": seq_prevista(paradas),
            }, evento_unico=True)
            return


def checar_rota(ctx, est, tracado, paradas, ts, pos):
    if not tracado:
        return
    d = distancia_tracado_km(pos, tracado)
    fora = d > RAIO_ROTA_KM
    if fora:
        disparar_alerta(ctx, est, "ROTA", "ROTA", "Desvio de Rota Detectado", ts, pos, {
            "Distância do traçado": f"{d:.1f} km (tolerância {RAIO_ROTA_KM:.0f} km)",
            "Próxima parada": paradas[est["idx_ativo"]]["nome"],
        })
    elif est["em_desvio"]["ROTA"]:
        log.info(f"[{ctx['placa']}] voltou para a rota às {ts:%H:%M} ({d:.1f} km do traçado)")
        est["alertas"].pop("ROTA", None)
    est["em_desvio"]["ROTA"] = fora


def checar_sentido(ctx, est, paradas, ts, pos):
    if not est["ultima_posicao"]:
        return
    ant = tuple(est["ultima_posicao"])
    alvo = paradas[est["idx_ativo"]]
    if distancia_km(ant, pos) < MIN_DESLOCAMENTO_KM:
        return
    d_ant, d_atual = distancia_km(ant, alvo["pos"]), distancia_km(pos, alvo["pos"])
    angulo = diferenca_angulo(rumo_graus(ant, pos), rumo_graus(ant, alvo["pos"]))
    afastando = (d_atual - d_ant) >= MIN_AFASTAMENTO_KM and angulo > ANGULO_SENTIDO_ERRADO
    est["contador_afastamento"] = est["contador_afastamento"] + 1 if afastando else 0

    if est["contador_afastamento"] >= LEITURAS_SENTIDO:
        est["em_desvio"]["SENTIDO"] = True
        disparar_alerta(ctx, est, "SENTIDO", "SENTIDO", "Veículo em Sentido Incorreto", ts, pos, {
            "Próxima parada": alvo["nome"],
            "Distância até a parada": f"{d_atual:.1f} km (era {d_ant:.1f} km na leitura anterior)",
            "Afastamento": f"+{d_atual - d_ant:.1f} km desde a última leitura "
                           f"({est['contador_afastamento']} leituras seguidas)",
            "Rumo do veículo x rumo da parada": f"{angulo:.0f}° de diferença",
        })
    elif not afastando and est["em_desvio"]["SENTIDO"]:
        log.info(f"[{ctx['placa']}] voltou a se aproximar de {alvo['nome']} às {ts:%H:%M}")
        est["em_desvio"]["SENTIDO"] = False
        est["alertas"].pop("SENTIDO", None)

# =============================================================================
# 9. PROCESSAMENTO
# =============================================================================
def processar(reset=False):
    df, cols_paradas = ler_relatorio()
    geo = Geocodificador(ARQUIVO_MUNICIPIOS)
    tracados = Tracados()
    estado = carregar_estado(reset)
    viagens = estado["_viagens"]

    sem_pos, sem_cidade, sem_tracado, desatualizadas, novas = [], [], [], [], 0
    ref = df["_ts"].max()
    processadas = 0
    historico = []

    for _, r in df.iterrows():
        vid = r["_viagem"]
        ctx = {k: texto(r[c]) for k, c in COL.items()}
        pos, ts = r["_pos"], r["_ts"]

        if pos is None or pd.isna(ts):
            sem_pos.append(ctx["placa"])
            continue
        ts = ts.to_pydatetime()
        if pd.notna(ref) and ref - ts > timedelta(hours=POSICAO_DESATUALIZADA_H):
            desatualizadas.append(f"{ctx['placa']} ({ts:%d/%m %H:%M})")

        paradas, nao_achadas = montar_roteiro(r, cols_paradas, geo)
        if nao_achadas:
            sem_cidade.append(f"{ctx['placa']}: {', '.join(nao_achadas)}")
            continue

        est = viagens.get(vid)
        ass = assinatura_roteiro([p["nome"] for p in paradas])
        if est and est.get("assinatura") != ass and est["situacao"] == "EM ANDAMENTO":
            log.info(f"[{ctx['placa']}] roteiro alterado no relatório - reiniciando controle das paradas")
            est = None
        if est is None:
            est = viagens[vid] = estado_inicial(paradas, ctx, pos)
            novas += 1
        est["info"] = ctx
        if est["situacao"] != "EM ANDAMENTO":
            continue
        if est["ultima_leitura"] and ts <= datetime.fromisoformat(est["ultima_leitura"]):
            continue                                    # posição não atualizou

        tracado = tracados.obter(vid, est["origem"], paradas)
        if not tracado:
            sem_tracado.append(ctx["placa"])

        processadas += 1
        historico.append({"viagem": vid, "placa": ctx["placa"], "datahora": ts,
                          "lat": pos[0], "lon": pos[1], "local": ctx["local"]})

        checar_rota(ctx, est, tracado, paradas, ts, pos)
        checar_sentido(ctx, est, paradas, ts, pos)
        checar_paradas(ctx, est, paradas, ts, pos)

        est["ultima_posicao"] = list(pos)
        est["ultima_leitura"] = ts.isoformat()

    # viagens que saíram de "Em Trânsito" no relatório
    ativas = set(df["_viagem"])
    for vid, est in viagens.items():
        if est["situacao"] == "EM ANDAMENTO" and vid not in ativas:
            est["situacao"] = "ENCERRADA NO RELATÓRIO"
            log.info(f"[{est['info'].get('placa', vid)}] saiu de 'Em Trânsito' no relatório - monitoramento encerrado")

    if novas:
        log.info(f"{novas} viagem(ns) nova(s) em monitoramento")
    avisar_se_mudou(estado, "sem_pos", sem_pos, "Em trânsito sem coordenada")
    avisar_se_mudou(estado, "sem_cidade", sem_cidade, "Cidade de parada/destino não localizada (incluir em CIDADES_EXTRA)")
    avisar_se_mudou(estado, "sem_tracado", sem_tracado, "Sem traçado ainda (desvio de rota não avaliado)", intervalo_min=60)
    avisar_se_mudou(estado, "desatualizadas", desatualizadas, f"Posição com mais de {POSICAO_DESATUALIZADA_H}h", intervalo_min=60)

    if historico:
        if _ESTADO_MEMORIA is not None:
            _ESTADO_MEMORIA.setdefault("_historico", []).extend(historico)
        else:
            novo = not Path(ARQUIVO_HISTORICO).exists()
            pd.DataFrame(historico).to_csv(ARQUIVO_HISTORICO, mode="a", header=novo, index=False,
                                           sep=";", encoding="utf-8-sig")
    tracados.salvar()
    salvar_estado(estado)
    if GERAR_EXCEL_SITUACAO:
        try:
            montar_situacao(estado).to_excel(ARQUIVO_SITUACAO, index=False)
        except PermissionError:
            log.warning(f"{ARQUIVO_SITUACAO} aberto no Excel - resumo não atualizado nesta execução")
    return processadas


def montar_situacao(estado):
    """Resumo por viagem (usado pela interface e, opcionalmente, gravado em Excel)."""
    linhas = []
    for vid, est in estado["_viagens"].items():
        i = est["info"]
        nomes = est.get("nomes_paradas") or est["assinatura"].split("|")
        pos = est["ultima_posicao"]
        linhas.append({
            "Viagem": vid, "Placa Cavalo": i.get("placa"), "Placa Carreta": i.get("carreta"),
            "Motorista": i.get("motorista"), "Destino final": i.get("destino"),
            "Situação monitor": est["situacao"],
            "Próxima parada": nomes[est["idx_ativo"]] if est["situacao"] == "EM ANDAMENTO" else "",
            "Paradas": " | ".join(f"{n}: {s}" for n, s in zip(nomes, est["status_paradas"])),
            "Em desvio de rota": "SIM" if est["em_desvio"]["ROTA"] else "",
            "Em sentido incorreto": "SIM" if est["em_desvio"]["SENTIDO"] else "",
            "Qtde alertas": est["qtd_alertas"], "Último alerta": est["ultimo_alerta"],
            "Última posição": est["ultima_leitura"], "Local": i.get("local"),
            "Latitude": pos[0] if pos else None, "Longitude": pos[1] if pos else None,
            "Google Maps": link_maps(*pos) if pos else "",
        })
    ordem = {"EM ANDAMENTO": 0, "FINALIZADA": 1}
    out = pd.DataFrame(linhas)
    if not out.empty:
        out = out.sort_values(["Situação monitor", "Qtde alertas"], ascending=[True, False],
                              key=lambda s: s.map(ordem).fillna(2) if s.name == "Situação monitor" else s)
    return out

# =============================================================================
# 10. SIMULAÇÃO E EXECUÇÃO
# =============================================================================
_PADRAO = {k: globals()[k] for k in ("ARQUIVO_RELATORIO", "ARQUIVO_ESTADO", "ARQUIVO_CACHE_ROTAS",
                                      "ARQUIVO_HISTORICO", "ARQUIVO_ALERTAS", "ARQUIVO_SITUACAO",
                                      "PASTA_ALERTAS", "USAR_OSRM")}


def modo_producao():
    globals().update(_PADRAO)


def modo_simulacao(forcar_modo_teste=True):
    """Arquivos separados para não misturar com a operação real, e sem internet."""
    global ARQUIVO_RELATORIO, ARQUIVO_ESTADO, ARQUIVO_CACHE_ROTAS, ARQUIVO_HISTORICO, \
        ARQUIVO_ALERTAS, ARQUIVO_SITUACAO, PASTA_ALERTAS, USAR_OSRM, MODO_TESTE
    ARQUIVO_ALERTAS = "sim_alertas.csv"
    ARQUIVO_RELATORIO = "sim_relatorio_atual.xlsx"
    ARQUIVO_ESTADO = "sim_estado.json"
    ARQUIVO_CACHE_ROTAS = "sim_rotas_cache.json"
    ARQUIVO_HISTORICO = "sim_historico.csv"
    ARQUIVO_SITUACAO = "sim_situacao_viagens.xlsx"
    PASTA_ALERTAS = "sim_alertas"
    USAR_OSRM = False
    if forcar_modo_teste:
        MODO_TESTE = True


def sim_reiniciar():
    """Apaga os arquivos da simulação (chamar com modo_simulacao ativo)."""
    for arq in (ARQUIVO_HISTORICO, ARQUIVO_ALERTAS, ARQUIVO_ESTADO, ARQUIVO_RELATORIO):
        Path(arq).unlink(missing_ok=True)
    shutil.rmtree(PASTA_ALERTAS, ignore_errors=True)


# Constantes de configuração da simulação
_MIN_LINHAS_POR_SNAPSHOT = 10  # blocos menores que isso no _historico são ignorados


def sim_processar_foto(caminho):
    """Processa uma foto do relatório em modo simulação.

    Se a planilha tiver aba `_historico` com snapshots anteriores empilhados,
    cada bloco é processado em ordem cronológica antes da foto principal.
    O arquivo de origem NÃO é modificado — usa-se um arquivo temporário isolado.

    Args:
        caminho: Path para o arquivo Excel a processar.

    Returns:
        Total de posições processadas em todos os snapshots.
    """
    try:
        abas = pd.read_excel(caminho, sheet_name=None, dtype=str)
    except (ValueError, OSError) as e:
        log.error(f"Não foi possível ler {caminho}: {e}")
        return 0

    aba_principal = ABA_RELATORIO if ABA_RELATORIO in abas else next(iter(abas))
    df_principal = abas[aba_principal]

    total = 0
    # Processa histórico embutido, se houver
    hist = abas.get("_historico")
    if hist is not None and not hist.empty:
        n_por_snap = len(df_principal)
        for inicio in range(0, len(hist), n_por_snap):
            bloco = hist.iloc[inicio:inicio + n_por_snap]
            if len(bloco) < _MIN_LINHAS_POR_SNAPSHOT:
                continue
            total += _processar_dataframe(bloco)

    total += _processar_dataframe(df_principal)
    return total


def _processar_dataframe(df):
    """Processa um DataFrame como se fosse o relatório atual, sem tocar em arquivos externos."""
    tmp = Path(ARQUIVO_RELATORIO)
    with pd.ExcelWriter(tmp, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name=ABA_RELATORIO, index=False)
    return processar()


def rodar_simulacao(pasta):
    modo_simulacao()
    sim_reiniciar()
    fotos = sorted(Path(pasta).glob("*.xlsx"))
    log.info(f"===== SIMULAÇÃO: {len(fotos)} fotos do relatório em '{pasta}' =====")
    for f in fotos:
        sim_processar_foto(f)
    log.info("===== FIM DA SIMULAÇÃO =====")


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--reset", action="store_true")
    ap.add_argument("--loop", action="store_true")
    ap.add_argument("--simular", metavar="PASTA")
    a = ap.parse_args()
    configurar_log()

    if a.simular:
        rodar_simulacao(a.simular)
        sys.exit()

    n = processar(a.reset)
    log.info(f"{n} posição(ões) nova(s) processada(s)")
    while a.loop:
        time.sleep(INTERVALO_LOOP_MIN * 60)
        try:
            n = processar()
            log.info(f"{n} posição(ões) nova(s) processada(s)")
        except Exception as e:
            log.exception("Erro no ciclo")

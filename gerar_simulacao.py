"""
gerar_simulacao.py
Gera "fotos" FICTÍCIAS do relatório futuro (layout da aba Consulta + colunas parada1, parada2),
partindo do Gestão_Viagens.xlsx real. Cada foto representa uma atualização de 10 min.

- As 964 linhas reais entram em todas as fotos (quem não é simulado fica com a posição parada).
- 6 viagens Em Trânsito recebem paradas e trajetos fictícios:
    GBY2H84 -> Manhuaçu > Ubá          : sai no sentido errado (BR-101 norte) e corrige -> ROTA + SENTIDO, depois finaliza
    GGF8J43 -> Itaboraí > Duque de Caxias > Bauru : pula Itaboraí e volta           -> SENTIDO + SEQUÊNCIA + INVERSÃO
    GCK8J52 -> Vacaria > Caxias do Sul  : normal                                      -> finaliza
    GGN3J74 -> Contagem                 : normal                                      -> finaliza
    GGW4H43 -> Rio de Janeiro           : já está no destino                          -> finaliza
    GHW5F11 -> Barueri                  : parado no pátio                             -> sem alerta
- Grava sim_rotas_cache.json com o traçado fictício dessas viagens (a simulação não usa internet).
"""
import json
import math
import random
from datetime import datetime, timedelta
from pathlib import Path

import pandas as pd

from monitor_rotas import assinatura_roteiro

random.seed(7)
ARQUIVO_BASE = "Gestão_Viagens.xlsx"
N_FOTOS = 60                                   # 60 x 10 min = 10 h
INICIO = datetime(2026, 9, 8, 15, 50)          # logo depois da extração real

C = {  # coordenadas aproximadas para o trajeto simulado
    "PÁTIO CARIACICA": (-20.2576, -40.3766), "CARIACICA": (-20.3393, -40.4049), "SERRA": (-20.1286, -40.3078),
    "FUNDÃO": (-19.9322, -40.4061), "VIANA": (-20.3900, -40.4958), "VENDA NOVA DO IMIGRANTE": (-20.3397, -41.1350),
    "IBATIBA": (-20.2346, -41.5100), "MANHUAÇU": (-20.2583, -42.0339), "MURIAÉ": (-21.1306, -42.3664),
    "UBÁ": (-21.1204, -42.9427), "GUARAPARI": (-20.6667, -40.4975), "CAMPOS DOS GOYTACAZES": (-21.7545, -41.3244),
    "MACAÉ": (-22.3708, -41.7869), "CASIMIRO DE ABREU": (-22.4779, -42.0920), "RIO BONITO": (-22.7086, -42.6089),
    "ITABORAÍ": (-22.7447, -42.8594), "DUQUE DE CAXIAS": (-22.7858, -43.3117), "NITERÓI": (-22.8832, -43.1034),
    "RIO DE JANEIRO": (-22.8860, -43.2150), "SEROPÉDICA": (-22.7443, -43.7076), "PIRAÍ": (-22.6290, -43.8980),
    "RESENDE": (-22.4705, -44.4509), "SÃO JOSÉ DOS CAMPOS": (-23.1791, -45.8872), "SÃO PAULO": (-23.5505, -46.6333),
    "BARUERI": (-23.5057, -46.8790), "BAURU": (-22.3246, -49.0871), "JOÃO MONLEVADE": (-19.8126, -43.1735),
    "CAETÉ": (-19.7545, -43.6305), "BELO HORIZONTE": (-19.9167, -43.9345), "CONTAGEM": (-19.9317, -44.0536),
    "CURITIBA": (-25.4284, -49.2733), "LAGES": (-27.8161, -50.3261), "VACARIA": (-28.5122, -50.9339),
    "CAXIAS DO SUL": (-29.1678, -51.1794),
}
UF = {"PÁTIO CARIACICA": "ES", "CARIACICA": "ES", "SERRA": "ES", "FUNDÃO": "ES", "VIANA": "ES",
      "VENDA NOVA DO IMIGRANTE": "ES", "IBATIBA": "ES", "GUARAPARI": "ES", "MANHUAÇU": "MG", "MURIAÉ": "MG",
      "UBÁ": "MG", "JOÃO MONLEVADE": "MG", "CAETÉ": "MG", "BELO HORIZONTE": "MG", "CONTAGEM": "MG",
      "SÃO JOSÉ DOS CAMPOS": "SP", "SÃO PAULO": "SP", "BARUERI": "SP", "BAURU": "SP", "CURITIBA": "PR",
      "LAGES": "SC", "VACARIA": "RS", "CAXIAS DO SUL": "RS"}

# placa: (paradas antes do destino, traçado previsto, trajeto real [(local, leituras parado)])
CENARIOS = {
    "GBY2H84": (["Manhuaçu/MG"],
                ["PÁTIO CARIACICA", "CARIACICA", "VIANA", "VENDA NOVA DO IMIGRANTE", "IBATIBA", "MANHUAÇU", "MURIAÉ", "UBÁ"],
                [("SERRA", 0), ("FUNDÃO", 0), ("SERRA", 0), ("CARIACICA", 0), ("VIANA", 0), ("VENDA NOVA DO IMIGRANTE", 0),
                 ("IBATIBA", 0), ("MANHUAÇU", 3), ("MURIAÉ", 0), ("UBÁ", 3)]),
    "GGF8J43": (["Itaboraí/RJ", "Duque de Caxias/RJ"],
                ["PÁTIO CARIACICA", "GUARAPARI", "CAMPOS DOS GOYTACAZES", "MACAÉ", "CASIMIRO DE ABREU", "RIO BONITO",
                 "ITABORAÍ", "DUQUE DE CAXIAS", "SEROPÉDICA", "PIRAÍ", "RESENDE", "SÃO JOSÉ DOS CAMPOS", "SÃO PAULO", "BAURU"],
                [("RIO BONITO", 0), ("ITABORAÍ", 0), ("DUQUE DE CAXIAS", 3), ("ITABORAÍ", 3), ("DUQUE DE CAXIAS", 0),
                 ("SEROPÉDICA", 0), ("PIRAÍ", 0), ("RESENDE", 0), ("SÃO JOSÉ DOS CAMPOS", 0)]),
    "GCK8J52": (["Vacaria/RS"],
                ["PÁTIO CARIACICA", "CAMPOS DOS GOYTACAZES", "RIO DE JANEIRO", "SÃO PAULO", "CURITIBA", "LAGES", "VACARIA", "CAXIAS DO SUL"],
                [("VACARIA", 3), ("CAXIAS DO SUL", 4)]),
    "GGN3J74": ([],
                ["PÁTIO CARIACICA", "VENDA NOVA DO IMIGRANTE", "MANHUAÇU", "JOÃO MONLEVADE", "CAETÉ", "BELO HORIZONTE", "CONTAGEM"],
                [("BELO HORIZONTE", 0), ("CONTAGEM", 4)]),
    "GGW4H43": ([],
                ["PÁTIO CARIACICA", "GUARAPARI", "CAMPOS DOS GOYTACAZES", "MACAÉ", "RIO BONITO", "NITERÓI", "RIO DE JANEIRO"],
                [("RIO DE JANEIRO", 6)]),
    "GHW5F11": ([],
                ["PÁTIO CARIACICA", "GUARAPARI", "CAMPOS DOS GOYTACAZES", "MACAÉ", "RIO BONITO", "NITERÓI", "RIO DE JANEIRO",
                 "PIRAÍ", "RESENDE", "SÃO JOSÉ DOS CAMPOS", "SÃO PAULO", "BARUERI"],
                [("PÁTIO CARIACICA", 80)]),
}


def hav(a, b):
    la1, lo1, la2, lo2 = map(math.radians, (*a, *b))
    h = math.sin((la2 - la1) / 2) ** 2 + math.cos(la1) * math.cos(la2) * math.sin((lo2 - lo1) / 2) ** 2
    return 2 * 6371.0088 * math.asin(math.sqrt(h))


def simular(pontos, inicio, vel_kmh=70):
    passo = vel_kmh / 6
    out, pos, sobra = [inicio], inicio, 0.0
    for nome, parado in pontos:
        alvo = C[nome]
        seg = hav(pos, alvo)
        d = passo - sobra
        while d < seg:
            f = d / seg
            out.append((pos[0] + (alvo[0] - pos[0]) * f + random.uniform(-.002, .002),
                        pos[1] + (alvo[1] - pos[1]) * f + random.uniform(-.002, .002)))
            d += passo
        sobra = seg - (d - passo)
        pos = alvo
        for _ in range(parado):
            out.append((alvo[0] + random.uniform(-.0008, .0008), alvo[1] + random.uniform(-.0008, .0008)))
            sobra = 0.0
    while len(out) < N_FOTOS:
        out.append((pos[0] + random.uniform(-.0005, .0005), pos[1] + random.uniform(-.0005, .0005)))
    return out[:N_FOTOS]


def cidade_proxima(p):
    n = min((k for k in C if k != "PÁTIO CARIACICA"), key=lambda k: hav(p, C[k]))
    return f"{n} / {UF.get(n, 'RJ')}"


base = pd.read_excel(ARQUIVO_BASE, sheet_name="Consulta")
pos_dest = base.columns.get_loc("UF Destino") + 1
base.insert(pos_dest, "parada1", "")
base.insert(pos_dest + 1, "parada2", "")

trajetos, cache = {}, {}
for placa, (paradas, tracado, real) in CENARIOS.items():
    i = base.index[base["Placa Cavalo"] == placa][0]
    for k, nome in enumerate(paradas, 1):
        base.at[i, f"parada{k}"] = nome
    la, lo = map(float, str(base.at[i, "Lat long"]).split(","))
    trajetos[i] = simular(real, (la, lo))
    vid = str(base.at[i, "Última viagem"]).strip()
    cache[vid] = {"assinatura": assinatura_roteiro(paradas + [base.at[i, "UF Destino"]]),
                  "origem": list(C["PÁTIO CARIACICA"]), "gerado_em": "simulacao",
                  "pontos": [list(C[n]) for n in tracado]}

Path("sim_rotas_cache.json").write_text(json.dumps(cache), encoding="utf-8")
Path("snapshots").mkdir(exist_ok=True)
for f in Path("snapshots").glob("*.xlsx"):
    f.unlink()

for k in range(N_FOTOS):
    foto = base.copy()
    ts = INICIO + timedelta(minutes=10 * k)
    for i, traj in trajetos.items():
        la, lo = traj[k]
        foto.at[i, "Data /Hora Posição"] = ts + timedelta(seconds=random.randint(0, 59))
        foto.at[i, "Local última posição"] = cidade_proxima((la, lo))
        foto.at[i, "Lat long"] = f"{la:.4f},{lo:.4f}"
    foto.to_excel(f"snapshots/relatorio_{k:03d}.xlsx", sheet_name="Consulta", index=False)

print(f"OK: {N_FOTOS} fotos em ./snapshots ({len(base)} linhas cada), traçado de {len(cache)} viagens em sim_rotas_cache.json")

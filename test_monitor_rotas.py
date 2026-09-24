"""
test_monitor_rotas.py — Testes unitários e de integração do motor de monitoramento.

Rodar: pytest test_monitor_rotas.py -v
"""

import math
from pathlib import Path

import pandas as pd
import pytest

import monitor_rotas as m


# =============================================================================
# 1. GEOMETRIA (unitários puros)
# =============================================================================

class TestGeometria:
    """Funções de cálculo geodésico devem ser exatas para casos conhecidos."""

    def test_distancia_km_ponto_igual(self):
        assert m.distancia_km((-20.3, -40.3), (-20.3, -40.3)) == pytest.approx(0, abs=0.001)

    def test_distancia_km_rio_sao_paulo(self):
        # Distância real entre RJ e SP: ~360 km em linha reta
        rj = (-22.9068, -43.1729)
        sp = (-23.5505, -46.6333)
        d = m.distancia_km(rj, sp)
        assert 350 < d < 370, f"esperado ~360 km, obtido {d:.1f}"

    def test_rumo_graus_norte(self):
        # De (0,0) para (1,0) → norte
        assert m.rumo_graus((0, 0), (1, 0)) == pytest.approx(0, abs=0.1)

    def test_rumo_graus_leste(self):
        # De (0,0) para (0,1) → leste
        assert m.rumo_graus((0, 0), (0, 1)) == pytest.approx(90, abs=0.1)

    def test_diferenca_angulo_limites(self):
        assert m.diferenca_angulo(10, 350) == pytest.approx(20, abs=0.1)
        assert m.diferenca_angulo(180, 0) == pytest.approx(180, abs=0.1)
        assert m.diferenca_angulo(90, 270) == pytest.approx(180, abs=0.1)

    def test_distancia_tracado_ponto_sobre_rota(self):
        rota = [(0, 0), (1, 0), (2, 0)]
        assert m.distancia_tracado_km((1, 0), rota) == pytest.approx(0, abs=0.001)

    def test_distancia_tracado_ponto_fora(self):
        rota = [(0, 0), (0, 1)]
        # Ponto a ~111 km ao sul da rota
        d = m.distancia_tracado_km((-1, 0.5), rota)
        assert 100 < d < 125

    def test_link_maps_formato(self):
        url = m.link_maps(-20.3393, -40.4049)
        assert url.startswith("https://www.google.com/maps?q=")
        assert "-20.339300" in url
        assert "-40.404900" in url


# =============================================================================
# 2. NORMALIZAÇÃO E PARSING
# =============================================================================

class TestParsing:

    def test_normalizar_remove_acentos(self):
        assert m.normalizar("São Paulo") == "SAO PAULO"
        assert m.normalizar("cariacica / es") == "CARIACICA / ES"

    def test_normalizar_none(self):
        assert m.normalizar(None) == "NONE"  # str(None) = "None"

    def test_chave_cidade_valida(self):
        assert m.chave_cidade("Itaboraí/RJ") == "ITABORAI/RJ"
        assert m.chave_cidade("São Paulo / SP ") == "SAO PAULO/SP"

    def test_chave_cidade_invalida(self):
        assert m.chave_cidade(None) is None
        assert m.chave_cidade("") is None
        assert m.chave_cidade("apenas cidade") is None

    def test_parse_latlong_valido(self):
        assert m.parse_latlong("-20.3393,-40.4049") == (-20.3393, -40.4049)
        assert m.parse_latlong("-20.3393;-40.4049") == (-20.3393, -40.4049)

    def test_parse_latlong_invalido(self):
        assert m.parse_latlong(None) is None
        assert m.parse_latlong("") is None
        assert m.parse_latlong(" / ") is None
        assert m.parse_latlong("abc,def") is None
        assert m.parse_latlong("0,0") is None  # ponto zero absoluto tratado como inválido
        assert m.parse_latlong("100,0") is None  # fora da faixa


# =============================================================================
# 3. GEOCODIFICAÇÃO
# =============================================================================

@pytest.fixture(scope="module")
def geo():
    return m.Geocodificador(m.ARQUIVO_MUNICIPIOS)


class TestGeocodificador:

    def test_cidade_conhecida(self, geo):
        pos = geo.buscar("São Paulo/SP")
        assert pos is not None
        assert -24 < pos[0] < -23 and -47 < pos[1] < -46

    def test_cidade_com_acento(self, geo):
        assert geo.buscar("Itaboraí/RJ") is not None

    def test_cidade_inexistente(self, geo):
        assert geo.buscar("Cidade que não existe/XX") is None

    def test_cidade_extra_exterior(self, geo):
        # Zarate está na lista CIDADES_EXTRA
        assert geo.buscar("Zarate/EX") is not None


# =============================================================================
# 4. ESTADO EM MEMÓRIA
# =============================================================================

class TestEstadoMemoria:

    def teardown_method(self):
        """Garante que cada teste começa com estado em disco."""
        m.usar_estado_memoria(None)

    def test_estado_memoria_isola_de_arquivo(self, tmp_path, monkeypatch):
        monkeypatch.setattr(m, "ARQUIVO_ESTADO", str(tmp_path / "estado.json"))
        mem = {}
        m.usar_estado_memoria(mem)
        estado = m.carregar_estado()
        assert estado is mem
        assert "_viagens" in estado
        assert not Path(m.ARQUIVO_ESTADO).exists()  # não gravou em disco

    def test_estado_memoria_reset(self):
        mem = {"_viagens": {"X": {}}, "_avisos": {}}
        m.usar_estado_memoria(mem)
        m.carregar_estado(reset=True)
        assert mem["_viagens"] == {}


# =============================================================================
# 5. INTEGRAÇÃO — Cenários end-to-end
# =============================================================================

class TestCenarios:
    """Testes de integração com as planilhas de exemplo."""

    def _rodar_cenario(self, planilha):
        sessao = {"_viagens": {}, "_avisos": {}, "_alertas": [], "_historico": []}
        m.modo_simulacao(forcar_modo_teste=True)
        m.sim_reiniciar()
        m.usar_estado_memoria(sessao)
        try:
            m.sim_processar_foto(planilha)
        finally:
            m.usar_estado_memoria(None)
        return sessao

    def test_cenario_1_gera_4_alertas(self):
        sessao = self._rodar_cenario("cenario_1_150min.xlsx")
        assert len(sessao["_alertas"]) == 4
        tipos = {a["Tipo"] for a in sessao["_alertas"]}
        assert "Sentido incorreto" in tipos
        assert "Sequência" in tipos

    def test_cenario_2_gera_6_alertas(self):
        sessao = self._rodar_cenario("cenario_2_300min.xlsx")
        assert len(sessao["_alertas"]) == 6
        # No cenário 2 devemos ter os 3 tipos
        tipos = {a["Tipo"] for a in sessao["_alertas"]}
        assert tipos == {"Sentido incorreto", "Sequência"}
        # Inversão vem como "Sequência"
        assert any("Inversão" in a["Ocorrência"] for a in sessao["_alertas"])

    def test_cenario_2_finaliza_viagens(self):
        sessao = self._rodar_cenario("cenario_2_300min.xlsx")
        finalizadas = [v for v in sessao["_viagens"].values()
                       if v.get("situacao") == "FINALIZADA"]
        assert len(finalizadas) >= 3

    def test_cenario_nao_modifica_arquivo_original(self, tmp_path):
        """CRÍTICO: o script não pode alterar o arquivo enviado pelo usuário."""
        import shutil
        copia = tmp_path / "cenario.xlsx"
        shutil.copy("cenario_1_150min.xlsx", copia)
        hash_antes = copia.read_bytes()
        self._rodar_cenario(copia)
        hash_depois = copia.read_bytes()
        assert hash_antes == hash_depois, "sim_processar_foto MODIFICOU o arquivo de entrada!"


# =============================================================================
# 6. VALIDAÇÕES DE REGRA
# =============================================================================

class TestValidacoes:

    def test_cooldown_evita_alerta_duplicado(self):
        """Um mesmo alerta em <60 min deve ser suprimido."""
        sessao = {"_viagens": {}, "_avisos": {}, "_alertas": [], "_historico": []}
        m.modo_simulacao(forcar_modo_teste=True)
        m.sim_reiniciar()
        m.usar_estado_memoria(sessao)
        try:
            m.sim_processar_foto("cenario_2_300min.xlsx")
        finally:
            m.usar_estado_memoria(None)

        # Contar alertas por (placa, tipo) — dentro de 1h só deve ter 1
        from collections import Counter
        pares = Counter((a["Placa Cavalo"], a["Ocorrência"])
                        for a in sessao["_alertas"])
        # Nenhum par deve ter mais que 2 (2h ou mais entre eles)
        for par, n in pares.items():
            assert n <= 3, f"{par}: {n} alertas — cooldown falhando"

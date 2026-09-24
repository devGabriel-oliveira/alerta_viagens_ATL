"""
test_app.py — Testes de integração da interface Streamlit.

Rodar: pytest test_app.py -v
"""

from pathlib import Path

import pytest
from streamlit.testing.v1 import AppTest


APP = "app.py"
CENARIO_1 = "cenario_1_150min.xlsx"
CENARIO_2 = "cenario_2_300min.xlsx"


@pytest.fixture(autouse=True)
def _limpa_estado(monkeypatch):
    """Remove arquivos temporários entre testes e restaura o monitor."""
    import monitor_rotas as m
    for f in ["sim_estado.json", "sim_alertas.csv", "sim_historico.csv",
              "sim_relatorio_atual.xlsx"]:
        Path(f).unlink(missing_ok=True)
    m.usar_estado_memoria(None)
    m.modo_producao()  # garante configuração de produção como default
    yield
    m.usar_estado_memoria(None)


def _upload_planilha(at, caminho, nome=None):
    """Simula upload de uma planilha via file_uploader."""
    nome = nome or Path(caminho).name
    with open(caminho, "rb") as f:
        conteudo = f.read()
    return at.sidebar.file_uploader[0].set_value(
        (nome, conteudo,
         "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    ).run()


def _metricas(at):
    """Devolve dict de métrica → valor da tela."""
    return {x.label: x.value for x in at.metric}


class TestSimulacao:
    """A tela em modo Simulação deve reagir corretamente aos uploads."""

    def test_tela_inicial_pede_upload(self):
        at = AppTest.from_file(APP, default_timeout=180).run()
        assert any("upload" in i.value.lower() for i in at.info)
        assert not at.exception

    def test_upload_cenario_1_mostra_4_alertas(self):
        at = AppTest.from_file(APP, default_timeout=180).run()
        at = _upload_planilha(at, CENARIO_1)
        assert not at.exception
        assert _metricas(at)["🚨 Alertas"] == "4"

    def test_upload_cenario_2_mostra_6_alertas(self):
        at = AppTest.from_file(APP, default_timeout=180).run()
        at = _upload_planilha(at, CENARIO_2)
        assert not at.exception
        assert _metricas(at)["🚨 Alertas"] == "6"
        assert _metricas(at)["✅ Finalizadas"] == "3"

    def test_botao_limpar_reseta_estado(self):
        at = AppTest.from_file(APP, default_timeout=180).run()
        at = _upload_planilha(at, CENARIO_1)
        assert _metricas(at)["🚨 Alertas"] == "4"

        # Limpar o uploader e clicar em limpar tudo
        at.sidebar.file_uploader[0].set_value(None)
        btn = next(b for b in at.sidebar.button if "Limpar" in b.label)
        at = btn.click().run()
        # Sem métricas na tela (estado zerado)
        assert not at.metric or _metricas(at).get("🚨 Alertas") == "0"

    def test_modo_producao_muda_ui(self):
        at = AppTest.from_file(APP, default_timeout=180).run()
        at = at.sidebar.radio[0].set_value("Produção").run()
        assert not at.exception
        assert at.sidebar.radio[0].value == "Produção"
        # Botão de processar deve aparecer
        labels = [b.label for b in at.sidebar.button]
        assert any("Processar" in lbl for lbl in labels)

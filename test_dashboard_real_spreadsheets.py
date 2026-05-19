#!/usr/bin/env python3
"""Smoke test da correlacao do dashboard com as planilhas reais fornecidas."""

from pathlib import Path
import sys
import types

import pandas as pd


REPO_ROOT = Path(__file__).resolve().parents[1]
LOCAL_SITE_PACKAGES = REPO_ROOT / "site-packages"
if LOCAL_SITE_PACKAGES.exists():
    sys.path.insert(0, str(LOCAL_SITE_PACKAGES))


class _SidebarStub:
    def __getattr__(self, _name):
        return lambda *args, **kwargs: None


def _install_streamlit_stub():
    fake_st = types.SimpleNamespace(
        session_state={},
        set_page_config=lambda *args, **kwargs: None,
        markdown=lambda *args, **kwargs: None,
        error=lambda *args, **kwargs: None,
        warning=lambda *args, **kwargs: None,
        info=lambda *args, **kwargs: None,
        success=lambda *args, **kwargs: None,
        sidebar=_SidebarStub(),
    )
    sys.modules.setdefault("streamlit", fake_st)


def _install_plotly_stub_if_needed():
    try:
        import plotly.graph_objects  # noqa: F401
        import plotly.express  # noqa: F401
        from plotly.subplots import make_subplots  # noqa: F401
        return
    except Exception:
        pass

    plotly_module = types.ModuleType("plotly")
    graph_objects_module = types.ModuleType("plotly.graph_objects")
    express_module = types.ModuleType("plotly.express")
    subplots_module = types.ModuleType("plotly.subplots")
    subplots_module.make_subplots = lambda *args, **kwargs: None
    graph_objects_module.Figure = object

    sys.modules.setdefault("plotly", plotly_module)
    sys.modules.setdefault("plotly.graph_objects", graph_objects_module)
    sys.modules.setdefault("plotly.express", express_module)
    sys.modules.setdefault("plotly.subplots", subplots_module)


def _input_paths():
    status_candidates = list(Path("/Users/henrique_guenka/Downloads/MV DATA").glob("Relat*SJ 240426.xls"))
    if not status_candidates:
        raise FileNotFoundError("Planilha de status SJ 240426 nao encontrada.")

    achados_path = Path("/Users/henrique_guenka/Downloads/Relatorio Resultado Critico V2 jan abr 2026.xls")
    if not achados_path.exists():
        raise FileNotFoundError("Planilha de resultado critico jan-abr 2026 nao encontrada.")

    return status_candidates[0], achados_path


def run_test():
    _install_streamlit_stub()
    _install_plotly_stub_if_needed()
    from dashboard_achados_criticos import DashboardAchadosCriticos

    status_path, achados_path = _input_paths()

    dashboard = DashboardAchadosCriticos()
    dashboard.df_status = pd.read_excel(status_path, engine="xlrd").dropna(how="all")
    dashboard.df_achados = pd.read_excel(achados_path, engine="xlrd").dropna(how="all")

    assert dashboard.correlate_data(), "correlacao falhou"
    summary = dashboard.match_summary
    assert summary["alta"] == 57, summary
    assert summary["sem_same"] == 5, summary
    assert summary["sem_confianca"] == 1, summary

    correlated = dashboard.df_correlacionado
    assert (correlated["match_status"] == "correlacionado").sum() == 57

    gilda = correlated[correlated["Nome_Paciente"].astype(str).str.contains("GILDA PERDIGAO", na=False)].iloc[0]
    assert gilda["DESCRICAO_PROCEDIMENTO"] == "TC ABDOME TOTAL"

    claudio = correlated[correlated["Nome_Paciente"].astype(str).str.contains("CLAUDIO APFEL", na=False)].iloc[0]
    assert claudio["DESCRICAO_PROCEDIMENTO"] == "TC CRANIO OU ENCEFALO"

    assert dashboard.calculate_times(), "calculo de tempos falhou"
    assert len(dashboard.df_correlacionado) == 57
    assert len(dashboard.df_revisao_correlacao) == 6
    assert dashboard.df_correlacionado["match_procedure_similarity"].min() >= 0.99
    assert int(dashboard.df_correlacionado["tempo_negativo"].sum()) == 6

    full_df = dashboard.df_correlacionado.copy()
    full_review_df = dashboard.df_revisao_correlacao.copy()
    dashboard.apply_date_filter(2026, 1)
    assert len(dashboard.df_correlacionado) == 16

    dashboard.df_correlacionado = full_df.copy()
    dashboard.df_revisao_correlacao = full_review_df.copy()
    sys.modules["streamlit"].session_state["current_filter"] = {
        "ano": 2026,
        "mes": None,
        "registros_filtrados": len(full_df),
        "registros_totais": len(full_df),
    }
    pdf_report = dashboard.create_pdf_report()
    assert pdf_report is not None
    assert pdf_report.getvalue().startswith(b"%PDF")

    print("OK - correlacao real validada com 57 registros calculados.")


if __name__ == "__main__":
    run_test()

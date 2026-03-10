#!/usr/bin/env python3
"""
Script para inicializar o Dashboard de Achados Críticos
"""

import subprocess
import sys
import os
from pathlib import Path

def run_dashboard():
    """Executa o dashboard Streamlit"""
    dashboard_path = Path(__file__).parent / "dashboard_achados_criticos.py"

    print("🚀 Iniciando Dashboard de Achados Críticos...")
    print("📱 O dashboard será aberto no seu navegador automaticamente")
    print("🔗 URL: http://localhost:8501")
    print("⏹️  Para parar: Ctrl+C\n")

    try:
        # Executar streamlit
        subprocess.run([
            sys.executable, "-m", "streamlit", "run",
            str(dashboard_path),
            "--theme.base=dark",
            "--theme.primaryColor=#667eea",
            "--theme.backgroundColor=#0e1117",
            "--theme.secondaryBackgroundColor=#262730"
        ])
    except KeyboardInterrupt:
        print("\n👋 Dashboard encerrado pelo usuário")
    except Exception as e:
        print(f"❌ Erro ao executar dashboard: {e}")

if __name__ == "__main__":
    run_dashboard()
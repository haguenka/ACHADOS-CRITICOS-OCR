#!/usr/bin/env python3
"""
Script para executar a Interface GUI de Achados Críticos
"""

import subprocess
import sys
import os
from pathlib import Path

def run_gui():
    """Executa a interface GUI"""
    gui_path = Path(__file__).parent / "gui_achados_criticos.py"

    print("🚀 Iniciando Interface GUI - Achados Críticos...")
    print("🖥️  A interface gráfica será exibida em uma nova janela")
    print("⏹️  Para fechar: use o botão X da janela ou Ctrl+C no terminal\n")

    try:
        subprocess.run([sys.executable, str(gui_path)])
        print("\n👋 Interface GUI encerrada")
    except KeyboardInterrupt:
        print("\n👋 Interface GUI encerrada pelo usuário")
    except Exception as e:
        print(f"❌ Erro ao executar interface GUI: {e}")

if __name__ == "__main__":
    run_gui()
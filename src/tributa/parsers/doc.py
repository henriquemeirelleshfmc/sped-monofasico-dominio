"""
Conversor de documentos .doc para texto puro (.txt).
Suporta Word COM (Windows) e LibreOffice CLI (cross-platform).
"""
import os
import sys
import subprocess
import shutil

from ..config import console


def extract_txt_from_doc(doc_path):
    """
    Converte .doc para .txt. Prioriza Word COM no Windows, LibreOffice como fallback.
    Retorna caminho do .txt gerado ou None.
    """
    txt_path = os.path.abspath("tabela_temp.txt")
    abs_doc = os.path.abspath(doc_path)
    out_dir = os.path.dirname(abs_doc)

    # --- Tentativa 1: Microsoft Word COM (Windows) ---
    if sys.platform == "win32":
        try:
            import win32com.client as win32
            word = win32.Dispatch("Word.Application")
            word.Visible = False
            doc = word.Documents.Open(abs_doc)
            doc.SaveAs(txt_path, FileFormat=2)
            doc.Close()
            console.print("[dim]  Convertido via Microsoft Word[/dim]")
            return txt_path
        except Exception as e:
            console.print(f"[yellow]![/yellow] Word COM falhou: {e}. Tentando LibreOffice...")
        finally:
            try:
                word.Quit()
            except Exception:
                pass

    # --- Tentativa 2: LibreOffice CLI (cross-platform) ---
    soffice_paths = [
        "soffice",
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        "/usr/bin/soffice",
        "/usr/local/bin/soffice",
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    ]
    for soffice in soffice_paths:
        try:
            subprocess.run(
                [soffice, "--headless", "--convert-to", "txt:Text(utf8)", "--outdir", out_dir, abs_doc],
                capture_output=True, text=True, timeout=60,
            )
            lo_output = abs_doc.rsplit(".", 1)[0] + ".txt"
            if os.path.exists(lo_output):
                if lo_output != txt_path:
                    shutil.move(lo_output, txt_path)
                console.print("[dim]  Convertido via LibreOffice[/dim]")
                return txt_path
        except (FileNotFoundError, subprocess.TimeoutExpired, OSError):
            continue

    console.print("[red]x[/red] Nenhum conversor disponivel (instale LibreOffice ou Microsoft Word)")
    return None
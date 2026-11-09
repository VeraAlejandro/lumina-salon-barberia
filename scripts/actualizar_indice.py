"""Actualiza el índice y los campos del documento con Microsoft Word y exporta PDF."""

from __future__ import annotations

from pathlib import Path

import win32com.client as win32

BASE = Path(__file__).resolve().parent.parent
DOCX = BASE / "docs" / "PIA_Lumina.docx"
PDF = BASE / "docs" / "PIA_Lumina.pdf"


def main() -> None:
    if not DOCX.exists():
        raise FileNotFoundError(DOCX)

    word = win32.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    documento = None
    try:
        documento = word.Documents.Open(str(DOCX))
        documento.TrackRevisions = False

        # Se repite el ciclo porque al crecer el índice cambian los números de página.
        for _ in range(3):
            for indice in documento.TablesOfContents:
                indice.Update()
            documento.Repaginate()

        documento.Fields.Update()
        documento.Repaginate()
        for indice in documento.TablesOfContents:
            indice.Update()
        documento.Repaginate()

        documento.Save()
        documento.ExportAsFixedFormat(
            str(PDF),
            17,  # wdExportFormatPDF
            False,
            0,  # wdExportOptimizeForPrint
            0,
            1,
            1,
            0,
            True,
            True,
            False,
            True,
            False,
            False,
        )
        paginas = documento.ComputeStatistics(2)  # wdStatisticPages
        palabras = documento.ComputeStatistics(0)  # wdStatisticWords
        print(f"Índice actualizado. Páginas: {paginas} | Palabras: {palabras:,}")
        print(f"DOCX: {DOCX}")
        print(f"PDF:  {PDF}")
    finally:
        if documento is not None:
            documento.Close(False)
        word.Quit()


if __name__ == "__main__":
    main()

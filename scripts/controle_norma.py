# -*- coding: utf-8 -*-
"""Controle positivo da conferencia de pagina da Norma.

A conferencia procura no PDF folha em branco e referencia de campo quebrada.
Este controle desenha PDFs com e sem esses defeitos, e exige que ela acuse os
que tem e se cale nos que nao tem. Nao precisa de Word.

    python scripts/controle_norma.py
"""
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import norma as N                                                    # noqa: E402


def desenhar(caminho, paginas):
    import fitz
    d = fitz.open()
    for texto in paginas:
        pg = d.new_page(width=595, height=842)
        if texto:
            pg.insert_text((72, 100), texto, fontsize=12)
    d.save(str(caminho))


def main():
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass
    tmp = Path(tempfile.mkdtemp(prefix="controle-norma-"))
    casos = [
        ("limpo", ["Capitulo um.", "Capitulo dois."], lambda a: not a),
        ("folha em branco", ["Capitulo um.", "", "Capitulo dois."],
         lambda a: any("branco" in x and "2" in x for x in a)),
        ("referencia quebrada", ["Sumario ... Erro! Indicador nao definido.", "Texto."],
         lambda a: any("quebrada" in x and "1" in x for x in a)),
        ("referencia quebrada, em ingles", ["Contents ... Error! Bookmark not defined."],
         lambda a: any("quebrada" in x for x in a)),
    ]
    falhas = 0
    print("\nCONTROLE DA CONFERENCIA DE PAGINA DA NORMA")
    for nome, paginas, exigido in casos:
        pdf = tmp / (nome.replace(" ", "-").replace(",", "") + ".pdf")
        desenhar(pdf, paginas)
        _, avisos = N.conferir_pdf(pdf)
        ok = exigido(avisos)
        falhas += not ok
        print("  %s  %-32s %s" % ("PASSA  " if ok else "REPROVA", nome, avisos or "nenhum aviso"))
    print()
    if falhas:
        print("  %d caso(s) REPROVADO(S): a conferencia de pagina nao serve." % falhas)
        return 1
    print("  todos passaram: a conferencia acusa o que tem de acusar e se cala no resto.")
    return 0


if __name__ == "__main__":
    sys.exit(main())

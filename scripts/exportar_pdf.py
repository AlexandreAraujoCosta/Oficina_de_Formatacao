# -*- coding: utf-8 -*-
"""Exporta .docx para PDF pelo Word, com a marcacao desligada.

POR QUE A MARCACAO TEM DE SER DESLIGADA

Basta uma alteracao controlada no arquivo, mesmo de formatacao (`w:rPrChange`),
para o Word exportar em modo de marcacao: ele reserva a coluna de balao a
direita em TODAS as paginas, encolhe o texto e o empurra para a esquerda. A
margem, a mancha e a densidade medidas nesse PDF sao as do balao, e nao as do
documento. Desligada a marcacao, a geometria do PDF volta a ser a que o .docx
declara.

O ARQUIVO DE ENTRADA NAO E TOCADO

Abre com ReadOnly=True e AddToRecentFiles=False. A marcacao e desligada na
janela, que e estado de exibicao e nao conteudo, e o documento e fechado sem
salvar.

CONTROLE

O programa confere que o PDF saiu, que tem pagina, e que o numero de paginas do
PDF e o mesmo que o Word conta no .docx. Divergencia ai quer dizer que a
exportacao cortou alguma coisa, e o programa recusa terminar em silencio.

    python scripts/exportar_pdf.py --saida DIR trabalho.docx [outro.docx ...]
"""
import argparse
import sys
from pathlib import Path

try:
    import win32com.client as win32
except ImportError:
    sys.exit("pywin32 nao esta instalado. `pip install pywin32`")

try:
    import fitz
except ImportError:
    sys.exit("PyMuPDF nao esta instalado. `pip install pymupdf`")

PDF = 17                  # wdFormatPDF
SEM_MARCACAO = 0          # wdRevisionsMarkupNone
ESTATISTICA_PAGINAS = 2   # wdStatisticPages


def exportar(alvos, saida):
    saida.mkdir(parents=True, exist_ok=True)
    resultados = []
    # Word proprio, e nao o do usuario. `Dispatch` se anexa ao Word que estiver
    # aberto, e o programa esconderia e fecharia a janela e os documentos de
    # quem esta trabalhando. `DispatchEx` abre uma instancia nova, e so ela e
    # fechada.
    w = win32.DispatchEx("Word.Application")
    w.Visible = False
    try:
        for a in alvos:
            destino = saida / (a.stem + ".pdf")
            d = w.Documents.Open(str(a.resolve()), ReadOnly=True,
                                 AddToRecentFiles=False)
            try:
                v = d.ActiveWindow.View
                antes = (bool(v.ShowRevisionsAndComments),
                         int(v.RevisionsFilter.Markup))
                v.ShowRevisionsAndComments = False
                v.RevisionsFilter.Markup = SEM_MARCACAO
                d.SaveAs(str(destino), FileFormat=PDF)
                paginas_word = int(d.ComputeStatistics(ESTATISTICA_PAGINAS))
            finally:
                d.Close(False)
            resultados.append((a, destino, antes, paginas_word))
    finally:
        w.Quit()
    return resultados


def conferir(resultados):
    falhas = 0
    for origem, destino, antes, paginas_word in resultados:
        if not destino.exists():
            print("  REPROVA  nao saiu: %s" % destino.name)
            falhas += 1
            continue
        with fitz.open(destino) as doc:
            n = doc.page_count
        if n < 1:
            print("  REPROVA  PDF sem pagina: %s" % destino.name)
            falhas += 1
            continue
        if n != paginas_word:
            print("  REPROVA  %s: o Word conta %d paginas e o PDF tem %d"
                  % (destino.name, paginas_word, n))
            falhas += 1
            continue
        marca = "estava ligada" if antes[0] or antes[1] else "ja estava desligada"
        print("  PASSA    %-34s %4d paginas   (marcacao %s)"
              % (destino.name, n, marca))
    return falhas


def main():
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    ap.add_argument("--saida", required=True, help="diretorio dos PDFs")
    ap.add_argument("docx", nargs="+")
    a = ap.parse_args()

    alvos = [Path(p) for p in a.docx]
    for p in alvos:
        if not p.exists():
            sys.exit("nao encontrei %s" % p)
        if p.suffix.lower() not in (".docx", ".doc"):
            sys.exit("nao e documento do Word: %s" % p)

    print()
    print("EXPORTACAO PARA PDF, com a marcacao desligada")
    resultados = exportar(alvos, Path(a.saida).resolve())
    falhas = conferir(resultados)
    print()
    if falhas:
        sys.exit("%d exportacao(oes) reprovada(s)." % falhas)


if __name__ == "__main__":
    main()

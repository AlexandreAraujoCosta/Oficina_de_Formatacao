# -*- coding: utf-8 -*-
"""Cria a copia com controle de alteracoes, e quem escreve a marcacao e o Word.

POR QUE NAO ESCREVEMOS A MARCACAO

Escrever `w:ins` e `w:del` no XML sai correto e o resultado nao e: marcas de
paragrafo excluidas, todas em paragrafo vazio, fundiram paragrafos de texto
quando aceitas no Word, e a causa nao foi isolada. Como nao se conserta o que
nao se diagnosticou, removeu-se a dependencia: quem escreve a marcacao passa a ser o Word, comparando o original
com a copia ja formatada. Aceitar o resultado devolve a copia formatada por
construcao, porque foi dela que a marcacao nasceu.

O QUE ESTE PROGRAMA CONFERE SOZINHO

Aceita as alteracoes numa copia descartavel e coteja com a copia formatada,
paragrafo a paragrafo. Se nao bater, ele recusa certificar e mostra a diferenca.
E essa conferencia que pega paragrafo fundido antes da entrega.

UMA FALHA CONHECIDA DA COMPARACAO DO WORD

Quando o estilo `Normal` muda, a comparacao reorganiza paragrafos em regioes que
ninguem tocou, e o pre-textual e onde isso aparece. O programa nao conserta isso;
ele mede e diz.

    python scripts/revisao.py original.docx formatado.docx --saida revisao.docx
"""
import argparse
import difflib
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import ooxml as O                                                    # noqa: E402

NOVO_DOC, POR_PALAVRA, FORMATO_DOCX = 2, 1, 16
AUTOR = "Formatação"


def word():
    try:
        import win32com.client.dynamic
    except ImportError:
        sys.exit("precisa do pywin32 e do Word: pip install pywin32")
    import win32com.client
    # Word proprio, e nao o do usuario. `Dispatch` se anexa ao Word que estiver
    # aberto, e o programa esconderia e fecharia a janela e os documentos de
    # quem esta trabalhando. `DispatchEx` abre uma instancia nova, e so ela e
    # fechada.
    w = win32com.client.dynamic.Dispatch(
        win32com.client.DispatchEx("Word.Application")._oleobj_)
    w.Visible = False
    w.DisplayAlerts = False
    return w


def textos(caminho):
    """O texto de cada paragrafo com texto, na ordem, por varredura com
    profundidade. Regex ingenua de `<w:p>` quebra em caixa de texto."""
    doc = O.ler(caminho)["word/document.xml"]
    saida = []
    for a, b in O.paragrafos(doc):
        t = " ".join(O.texto(doc[a:b]).split())
        if t:
            saida.append(t)
    return saida


def atualizar_sumario(w, caminho):
    """Faz o Word preencher o campo de sumario e grava. Devolve quantas entradas."""
    d = w.Documents.Open(str(caminho), ReadOnly=False, AddToRecentFiles=False)
    try:
        n = int(d.TablesOfContents.Count)
        n_il = int(d.TablesOfFigures.Count)
        # Os campos SEQ primeiro, depois as listas: a lista de ilustracoes le os
        # SEQ, e o sumario e a lista leem a paginacao que os campos definem.
        d.Fields.Update()
        for i in range(1, n + 1):
            d.TablesOfContents(i).Update()
        for i in range(1, n_il + 1):
            d.TablesOfFigures(i).Update()
        entradas = 0
        if n:
            entradas = len(str(d.TablesOfContents(1).Range.Text).strip().splitlines())
        for i in range(1, n_il + 1):
            linhas = str(d.TablesOfFigures(i).Range.Text).strip().splitlines()
            print("  lista de ilustracoes %d: %d entrada(s) atualizada(s) pelo Word"
                  % (i, len(linhas)))
        d.Save()
    finally:
        d.Close(0)
    return n, entradas


def comparar(w, original, formatado, destino):
    d = w.CompareDocuments(
        w.Documents.Open(str(original), ReadOnly=True, AddToRecentFiles=False),
        w.Documents.Open(str(formatado), ReadOnly=True, AddToRecentFiles=False),
        NOVO_DOC, POR_PALAVRA,
        True,    # formatacao: e o que este programa muda, entao tem de entrar
        False,   # caixa
        True,    # espaco em branco
        True,    # tabelas
        True,    # cabecalhos e rodapes
        True,    # notas
        True,    # caixas de texto
        True,    # campos
        True,    # comentarios
        True,    # movimentos
        AUTOR,
        True)
    d.SaveAs2(str(destino), FileFormat=FORMATO_DOCX)
    n = int(d.Revisions.Count)
    d.Close(0)
    for outro in list(w.Documents):
        try:
            outro.Close(0)
        except Exception:
            pass
    return n


def aceitar(w, origem, destino):
    d = w.Documents.Open(str(origem), ReadOnly=False, AddToRecentFiles=False)
    try:
        d.Revisions.AcceptAll()
        d.SaveAs2(str(destino), FileFormat=FORMATO_DOCX)
    finally:
        d.Close(0)


def main():
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("original")
    ap.add_argument("formatado")
    ap.add_argument("--saida", required=True)
    ap.add_argument("--sem-sumario", action="store_true",
                    help="nao manda o Word preencher o campo de sumario antes")
    a = ap.parse_args()

    orig = Path(a.original).resolve()
    fmt = Path(a.formatado).resolve()
    saida = Path(a.saida).resolve()
    for p in (orig, fmt):
        if not p.exists():
            sys.exit("nao encontrei %s" % p)

    print()
    print("COPIA COM CONTROLE DE ALTERACOES")
    w = word()
    prova = saida.with_name(saida.stem + "-prova.docx")
    try:
        if not a.sem_sumario:
            n, entradas = atualizar_sumario(w, fmt)
            print("  sumario: %d campo(s) atualizado(s) pelo Word, %d linhas"
                  % (n, entradas))
        n = comparar(w, orig, fmt, saida)
        print("  %d alteracoes, escritas pelo Word e assinadas por %s" % (n, AUTOR))
        aceitar(w, saida, prova)
    finally:
        w.Quit()

    esperado, obtido = textos(fmt), textos(prova)
    print("  conferencia: aceitar a revisao da %d paragrafos com texto, contra "
          "%d da copia formatada" % (len(obtido), len(esperado)))
    if obtido == esperado:
        prova.unlink(missing_ok=True)
        print("  CONFERE: aceitar esta revisao devolve exatamente a copia formatada.")
        print()
        print("  gravado em %s" % saida)
        print()
        return 0

    print("  NAO CONFERE. A revisao nao devolve a copia formatada, e por isso")
    print("  este arquivo nao deve ser entregue. A prova ficou em %s" % prova.name)
    linhas = list(difflib.unified_diff(obtido, esperado, "aceita", "formatada",
                                       lineterm="", n=0))
    for l in linhas[:40]:
        print("    %s" % l[:110])
    if len(linhas) > 40:
        print("    ... (%d linhas de diferenca ao todo)" % len(linhas))
    return 1


if __name__ == "__main__":
    sys.exit(main())

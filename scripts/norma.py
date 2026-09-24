# -*- coding: utf-8 -*-
"""Norma: formata um trabalho na norma do PMPD, com um comando so.

O QUE ELA ENTREGA, AO LADO DO ARQUIVO DE ENTRADA

  <nome>-norma.docx                          a copia formatada
  <nome>-norma-com-controle-de-alteracoes.docx  a mesma, com cada mudanca
                                             marcada, para aceitar ou recusar
  <nome>-norma-relatorio.md                  o que mudou, e o que conferir
  <nome>-norma.pdf                           a copia formatada em PDF

O arquivo de entrada nao e alterado.

O QUE ELA FAZ, EM ORDEM

  1 formatar.py   grava a copia formatada, e confere o texto, a estrutura, o
                  XML e que o Word abre o arquivo;
  2 revisao.py    atualiza sumario e listas, faz a copia com controle de
                  alteracoes e confere que aceita-la devolve a copia formatada;
  3 exportar_pdf  exporta a copia formatada, com a marcacao desligada;
  4 a pagina      procura no PDF o que as conferencias de arquivo nao veem:
                  folha em branco e referencia de campo quebrada.

Se uma etapa reprova, a Norma para ali e diz qual. O que ja foi gravado fica,
para ser examinado, mas nao deve ser entregue.

O QUE PRECISA

Windows com Word, Python 3, PyMuPDF e pywin32. O Word usado e uma instancia
propria, separada da que estiver aberta.

    python scripts/norma.py trabalho.docx
    python scripts/norma.py trabalho.docx --saida PASTA
"""
import argparse
import subprocess
import sys
from pathlib import Path

AQUI = Path(__file__).resolve().parent


def rodar(programa, *args):
    """Roda um dos programas da oficina, mostrando o que ele diz. True se passou."""
    r = subprocess.run([sys.executable, str(AQUI / programa)] + [str(a) for a in args])
    return r.returncode == 0


def conferir_pdf(caminho):
    """O que so a pagina mostra: folha em branco e referencia de campo quebrada.

    Folha em branco e a sem texto e sem imagem. Referencia quebrada e o texto que
    o Word escreve no lugar de um campo que aponta para marcador inexistente
    ("Erro! Indicador nao definido", ou "Error! Bookmark not defined").
    Devolve (numero de paginas, lista de avisos).
    """
    import fitz
    d = fitz.open(str(caminho))
    avisos = []
    brancas = [i + 1 for i, pg in enumerate(d)
               if not pg.get_text().strip() and not pg.get_images()
               and not pg.get_drawings()]
    if brancas:
        avisos.append("folha(s) em branco: %s" % ", ".join(map(str, brancas)))
    quebradas = [i + 1 for i, pg in enumerate(d)
                 if "Erro! Indicador" in pg.get_text() or "Error! Bookmark" in pg.get_text()
                 or "Erro! A origem" in pg.get_text() or "Error! Reference" in pg.get_text()]
    if quebradas:
        avisos.append("referencia de campo quebrada nas paginas: %s"
                      % ", ".join(map(str, quebradas)))
    return len(d), avisos


def main():
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("trabalho")
    ap.add_argument("--saida", help="pasta dos arquivos gerados (padrao: a do trabalho)")
    a = ap.parse_args()

    origem = Path(a.trabalho).resolve()
    if not origem.exists():
        sys.exit("nao encontrei %s" % origem)
    if origem.suffix.lower() != ".docx":
        sys.exit("a Norma le .docx; salve o trabalho nesse formato no Word.")
    pasta = Path(a.saida).resolve() if a.saida else origem.parent
    pasta.mkdir(parents=True, exist_ok=True)
    base = origem.stem + "-norma"
    formatado = pasta / (base + ".docx")
    com_controle = pasta / (base + "-com-controle-de-alteracoes.docx")
    pdf = pasta / (base + ".pdf")
    if formatado.resolve() == origem:
        sys.exit("o arquivo de saida seria o proprio trabalho; use --saida.")

    print("NORMA: %s" % origem.name)

    print("\n[1/4] formatacao")
    if not rodar("formatar.py", origem, "--saida", formatado):
        sys.exit("\nA Norma parou na formatacao. Nada do que foi gravado deve ser entregue.")

    print("\n[2/4] copia com controle de alteracoes")
    if not rodar("revisao.py", origem, formatado, "--saida", com_controle):
        sys.exit("\nA Norma parou na copia com controle de alteracoes.")

    print("\n[3/4] PDF")
    if not rodar("exportar_pdf.py", "--saida", pasta, formatado):
        sys.exit("\nA Norma parou na exportacao para PDF.")

    print("\n[4/4] a pagina")
    paginas, avisos = conferir_pdf(pdf)
    relatorio = pasta / (base + "-relatorio.md")
    linhas = ["", "## A pagina, conferida no PDF", "",
              "%d paginas." % paginas]
    if avisos:
        linhas += ["", "**Conferir antes de entregar:**", ""] + ["- %s" % x for x in avisos]
    else:
        linhas += ["Nenhuma folha em branco e nenhuma referencia de campo quebrada."]
    linhas += ["", "O resto se confere olhando: capa, sumario, a abertura de cada "
               "capitulo, as ilustracoes."]
    with open(relatorio, "a", encoding="utf-8") as f:
        f.write("\n".join(linhas) + "\n")
    print("  %d paginas" % paginas)
    for x in avisos:
        print("  CONFERIR: %s" % x)
    if not avisos:
        print("  nenhuma folha em branco, nenhuma referencia de campo quebrada")

    print("\nENTREGUE, na pasta %s:" % pasta)
    for p in (formatado, com_controle, relatorio, pdf):
        print("  %s" % p.name)
    return 1 if avisos else 0


if __name__ == "__main__":
    sys.exit(main())

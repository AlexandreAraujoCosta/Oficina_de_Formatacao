# -*- coding: utf-8 -*-
"""Controle positivo do medir_pagina.py: paginas quebradas de proposito.

POR QUE ISTO EXISTE, E POR QUE ELE VEM JUNTO E NAO DEPOIS

Um conferidor que nunca reprovou nada nao informa nada: quando ele se cala,
nao da para distinguir pagina boa de conferidor quebrado. A acusacao de um
teste e hipotese ate que se mostre que ele reprova o que tem de reprovar.

Este programa desenha, com o proprio PyMuPDF, paginas cuja geometria e conhecida
porque foi escrita aqui, e exige que o medir_pagina.py devolva esses numeros.
Nao ha .docx, nao ha Word e nao ha arquivo de fora: o gabarito e o codigo.

OS CINCO CASOS, E O QUE CADA UM PEGA

  1 margem normal      pega erro de conversao de ponto para centimetro
  2 margem estreita    pega medidor que devolve a margem do papel, e nao a do
                       texto: se ele reportasse 3 cm numa pagina de 0,5, uma
                       capa desmontada passaria despercebida
  3 linha longa/curta  pega contagem de caracteres feita sobre linha curta
  4 rodape sozinho     pega mancha que engole o numero de pagina, que e o que
                       faria toda margem inferior parecer menor do que e
  5 duas linhas no pe  pega o exagero do caso 4: texto que desceu ate o pe nao
                       e rodape, e descarta-lo esconderia o defeito

    python scripts/controle_medir_pagina.py
"""
import sys
import tempfile
from pathlib import Path

try:
    import fitz
except ImportError:
    sys.exit("PyMuPDF nao esta instalado. `pip install pymupdf`")

sys.path.insert(0, str(Path(__file__).resolve().parent))
import medir_pagina as M                                             # noqa: E402

PT = M.PT_POR_CM
FONTE = "cour"        # monoespacada: a largura da linha e previsivel
CORPO = 11.0
LARGURA_CHAR = 0.6 * CORPO
A4_L, A4_A = 595.276, 841.89

# O comprimento das linhas nao e livre: `insert_text` corta na borda da pagina
# e NAO avisa, e a linha cortada faria o controle acusar o medidor de contar
# errado. Na A4 com margem de 3 cm dos dois lados, a largura util e de 510 pt,
# onde cabem 68 caracteres de cour 11; 65 deixa folga.
N_LONGA = 65
N_CURTA = 30
TOL_H = 0.05          # cm, na horizontal a posicao e a que se pediu
TOL_V = 0.45          # cm, na vertical cabe uma linha entre a base e o topo


def linha(n):
    """Uma linha de exatamente n caracteres, sem espaco no fim."""
    return ("abcdefghij" * (n // 10 + 1))[:n]


def desenhar(caminho):
    doc = fitz.open()

    # 1: margem esquerda 3 cm, superior 3 cm, 20 linhas longas.
    p = doc.new_page(width=A4_L, height=A4_A)
    for i in range(20):
        p.insert_text((3 * PT, 3 * PT + (i + 1) * 16), linha(N_LONGA),
                      fontname=FONTE, fontsize=CORPO)

    # 2: a mesma coisa com margem de 0,5 cm. E a pagina quebrada.
    p = doc.new_page(width=A4_L, height=A4_A)
    for i in range(20):
        p.insert_text((0.5 * PT, 0.5 * PT + (i + 1) * 16), linha(N_LONGA),
                      fontname=FONTE, fontsize=CORPO)

    # 3: linhas curtas, para que a mediana caia de N_LONGA para N_CURTA.
    p = doc.new_page(width=A4_L, height=A4_A)
    for i in range(20):
        p.insert_text((3 * PT, 3 * PT + (i + 1) * 16), linha(N_CURTA),
                      fontname=FONTE, fontsize=CORPO)

    # 4: como a 1, mais um numero de pagina sozinho a 1 cm do pe.
    p = doc.new_page(width=A4_L, height=A4_A)
    for i in range(20):
        p.insert_text((3 * PT, 3 * PT + (i + 1) * 16), linha(N_LONGA),
                      fontname=FONTE, fontsize=CORPO)
    p.insert_text((A4_L / 2, A4_A - 1 * PT), "42", fontname=FONTE,
                  fontsize=CORPO)

    # 5: como a 1, mais DUAS linhas de texto a 1 cm do pe. Nao e rodape.
    p = doc.new_page(width=A4_L, height=A4_A)
    for i in range(20):
        p.insert_text((3 * PT, 3 * PT + (i + 1) * 16), linha(N_LONGA),
                      fontname=FONTE, fontsize=CORPO)
    p.insert_text((3 * PT, A4_A - 1.5 * PT), linha(N_LONGA), fontname=FONTE,
                  fontsize=CORPO)
    p.insert_text((3 * PT, A4_A - 1.0 * PT), linha(N_LONGA), fontname=FONTE,
                  fontsize=CORPO)

    doc.save(caminho)
    doc.close()


class Placar:
    def __init__(self):
        self.falhas = 0

    def exigir(self, nome, ok, obtido, esperado):
        marca = "PASSA " if ok else "REPROVA"
        if not ok:
            self.falhas += 1
        print("  %s  %-46s obtido %-14s esperado %s"
              % (marca, nome, obtido, esperado))


def main():
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

    tmp = Path(tempfile.mkdtemp(prefix="controle-medir-"))
    pdf = tmp / "controle.pdf"
    desenhar(pdf)
    m = M.medir(pdf)
    p = {d["pagina"]: d for d in m["paginas"]}
    mancha_esperada = LARGURA_CHAR * N_LONGA / PT      # em cm

    print()
    print("CONTROLE POSITIVO DE medir_pagina.py")
    print("  gabarito escrito neste programa, PDF em %s" % pdf)
    print()
    z = Placar()

    # 1
    z.exigir("1 margem esquerda de 3 cm",
             abs(p[1]["margem_esq_cm"] - 3.0) < TOL_H,
             "%.2f cm" % p[1]["margem_esq_cm"], "3,00 cm +- 0,05")
    z.exigir("1 margem superior de 3 cm",
             abs(p[1]["margem_sup_cm"] - 3.0) < TOL_V,
             "%.2f cm" % p[1]["margem_sup_cm"], "3,00 cm +- 0,45")
    z.exigir("1 mancha do numero de caracteres pedido",
             abs(p[1]["mancha_larg_cm"] - mancha_esperada) < TOL_H,
             "%.2f cm" % p[1]["mancha_larg_cm"],
             "%.2f cm +- 0,05" % mancha_esperada)
    z.exigir("1 vinte linhas na pagina",
             p[1]["linhas"] == 20, p[1]["linhas"], 20)

    # 2, a pagina quebrada
    z.exigir("2 margem estreita e acusada como 0,5 cm",
             abs(p[2]["margem_esq_cm"] - 0.5) < TOL_H,
             "%.2f cm" % p[2]["margem_esq_cm"], "0,50 cm +- 0,05")
    z.exigir("2 a margem estreita difere da normal",
             abs(p[2]["margem_esq_cm"] - p[1]["margem_esq_cm"]) > 2.0,
             "%.2f cm de diferenca"
             % abs(p[2]["margem_esq_cm"] - p[1]["margem_esq_cm"]),
             "mais de 2 cm")

    # 3
    c1 = M.statistics.median(p[1]["caracteres_por_linha"])
    c3 = M.statistics.median(p[3]["caracteres_por_linha"])
    z.exigir("3 linha longa e contada com o numero certo",
             c1 == N_LONGA, c1, N_LONGA)
    z.exigir("3 linha curta e contada com o numero certo",
             c3 == N_CURTA, c3, N_CURTA)

    # 4, o rodape
    z.exigir("4 o rodape sozinho sai da mancha",
             p[4].get("descartadas") == 1, p[4].get("descartadas"), 1)
    z.exigir("4 a margem inferior nao muda por causa do rodape",
             abs(p[4]["margem_inf_cm"] - p[1]["margem_inf_cm"]) < TOL_H,
             "%.2f cm contra %.2f" % (p[4]["margem_inf_cm"],
                                      p[1]["margem_inf_cm"]),
             "a mesma, +- 0,05")

    # 5, o exagero do 4
    z.exigir("5 duas linhas no pe nao sao rodape",
             p[5].get("descartadas") == 0, p[5].get("descartadas"), 0)
    z.exigir("5 o texto que desceu ate o pe encurta a margem inferior",
             p[5]["margem_inf_cm"] < p[1]["margem_inf_cm"] - 1.0,
             "%.2f cm contra %.2f" % (p[5]["margem_inf_cm"],
                                      p[1]["margem_inf_cm"]),
             "menor que a da pagina 1")

    print()
    if z.falhas:
        print("  %d exigencia(s) REPROVADA(S). O medidor nao esta certificado,"
              % z.falhas)
        print("  e nada que ele disser sobre um trabalho real vale enquanto isso.")
        sys.exit(1)
    print("  todas as exigencias passaram. O medidor reprova o que tem de")
    print("  reprovar, e por isso o silencio dele passa a informar.")
    print()


if __name__ == "__main__":
    main()

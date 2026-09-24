# -*- coding: utf-8 -*-
"""Mede a geometria das paginas de um PDF, e nao julga nenhuma delas.

POR QUE ISTO EXISTE

Contagem de paragrafo e contagem de estilo nao veem o que a pagina parece: uma
capa pode estar desmontada com todo o texto dela no lugar, na ordem certa. Os
defeitos de formatacao que importam aparecem na pagina.

POR QUE ELE MEDE EM VEZ DE JULGAR

Metade do que se quer saber sobre uma pagina e geometria de caixa de texto, e
sai do PDF em numero: margem, mancha, comprimento de linha em caracteres,
densidade. A outra metade e juizo, e o juizo de um modelo nao se reproduz: a
mesma pagina lida duas vezes da dois relatorios, e nao ha como saber se mudou a
pagina ou o leitor. Este programa faz a metade que se conta. O que se julga vem
depois, e sobre as paginas que o numero marcou.

ALCANCE, E ELE E ESTREITO DE PROPOSITO

Mede o que esta desenhado na pagina. Nao sabe o que e titulo, o que e legenda e
o que e tabela, porque isso nao esta no PDF: esta no .docx que o gerou. Orfa,
viuva, titulo isolado no pe e tabela partida exigem essa estrutura e nao estao
aqui. Nao se adivinha papel por tamanho de fonte.

O QUE ELE CHAMA DE CORPO

Linha sozinha na faixa de 8% do alto ou do pe da pagina e tratada como
cabecalho ou rodape e sai da conta da mancha. O programa diz quantas sairam.

O comprimento de linha e medido so nas linhas cheias, que sao as de largura
igual ou maior que 90% da largura mediana. Ultima linha de paragrafo e curta por
natureza, e incluir todas rebaixaria a medida sem dizer nada sobre a coluna.

    python scripts/medir_pagina.py trabalho.pdf
    python scripts/medir_pagina.py trabalho.pdf --json medidas.json
"""
import argparse
import json
import statistics
import sys
from pathlib import Path

try:
    import fitz  # PyMuPDF
except ImportError:
    sys.exit("PyMuPDF nao esta instalado. `pip install pymupdf`")

PT_POR_CM = 72.0 / 2.54
FAIXA_BORDA = 0.08      # faixa do alto e do pe em que linha sozinha e rodape
CHEIA = 0.90            # fracao da largura mediana que define linha cheia


def cm(pontos):
    return pontos / PT_POR_CM


def linhas_da_pagina(pagina):
    """Devolve [(x0, y0, x1, y1, texto)] das linhas de texto, na ordem de y."""
    saida = []
    for bloco in pagina.get_text("dict")["blocks"]:
        if bloco.get("type") != 0:          # 0 e texto, 1 e imagem
            continue
        for linha in bloco["lines"]:
            texto = "".join(s["text"] for s in linha["spans"])
            if not texto.strip():
                continue
            x0, y0, x1, y1 = linha["bbox"]
            saida.append((x0, y0, x1, y1, texto))
    saida.sort(key=lambda l: (round(l[1], 1), l[0]))
    return saida


def separar_corpo(linhas, altura_pagina):
    """Tira a linha que esta sozinha na faixa do alto ou do pe.

    Sozinha quer dizer unica linha dentro da faixa. Duas linhas na faixa nao
    sao rodape, sao texto que chegou ate la, e isso e o que se quer medir.
    """
    alto = altura_pagina * FAIXA_BORDA
    pe = altura_pagina * (1 - FAIXA_BORDA)
    no_alto = [l for l in linhas if l[3] <= alto]
    no_pe = [l for l in linhas if l[1] >= pe]
    fora = []
    if len(no_alto) == 1:
        fora.append(no_alto[0])
    if len(no_pe) == 1:
        fora.append(no_pe[0])
    corpo = [l for l in linhas if l not in fora]
    return corpo, fora


def medir(caminho):
    doc = fitz.open(caminho)
    paginas = []
    descartadas = 0
    for n, pagina in enumerate(doc, start=1):
        L, A = pagina.rect.width, pagina.rect.height
        linhas = linhas_da_pagina(pagina)
        corpo, fora = separar_corpo(linhas, A) if linhas else ([], [])
        descartadas += len(fora)
        if not corpo:
            paginas.append({"pagina": n, "vazia": True, "linhas": 0,
                            "largura_pt": L, "altura_pt": A})
            continue
        x0 = min(l[0] for l in corpo)
        x1 = max(l[2] for l in corpo)
        y0 = min(l[1] for l in corpo)
        y1 = max(l[3] for l in corpo)
        larguras = [l[2] - l[0] for l in corpo]
        mediana_larg = statistics.median(larguras)
        cheias = [l for l in corpo if (l[2] - l[0]) >= CHEIA * mediana_larg]
        alturas = sum(l[3] - l[1] for l in corpo)
        paginas.append({
            "pagina": n,
            "vazia": False,
            "largura_pt": L, "altura_pt": A,
            "linhas": len(corpo),
            "descartadas": len(fora),
            "margem_esq_cm": cm(x0),
            "margem_dir_cm": cm(L - x1),
            "margem_sup_cm": cm(y0),
            "margem_inf_cm": cm(A - y1),
            "mancha_larg_cm": cm(x1 - x0),
            "mancha_alt_cm": cm(y1 - y0),
            "linhas_cheias": len(cheias),
            "caracteres_por_linha": [len(l[4]) for l in cheias],
            "ocupacao": alturas / (y1 - y0) if y1 > y0 else 0.0,
        })
    doc.close()
    return {"arquivo": str(caminho), "paginas": paginas,
            "linhas_de_borda_descartadas": descartadas}


def resumo(m):
    uteis = [p for p in m["paginas"] if not p["vazia"]]
    chars = [c for p in uteis for c in p["caracteres_por_linha"]]

    def med(chave):
        v = [p[chave] for p in uteis]
        return statistics.median(v) if v else 0.0

    return {
        "paginas": len(m["paginas"]),
        "paginas_com_texto": len(uteis),
        "margem_esq_cm": med("margem_esq_cm"),
        "margem_dir_cm": med("margem_dir_cm"),
        "margem_sup_cm": med("margem_sup_cm"),
        "margem_inf_cm": med("margem_inf_cm"),
        "mancha_larg_cm": med("mancha_larg_cm"),
        "linhas_por_pagina": med("linhas"),
        "caracteres_por_linha": statistics.median(chars) if chars else 0.0,
        "caracteres_min": min(chars) if chars else 0,
        "caracteres_max": max(chars) if chars else 0,
        "linhas_de_borda_descartadas": m["linhas_de_borda_descartadas"],
    }


def imprimir(m):
    r = resumo(m)
    print()
    print("MEDIDA DE PAGINA: %s" % Path(m["arquivo"]).name)
    print("  %d paginas, %d com texto" % (r["paginas"], r["paginas_com_texto"]))
    print()
    print("  MARGEM (mediana, em cm)")
    print("    esquerda %.2f   direita %.2f   superior %.2f   inferior %.2f"
          % (r["margem_esq_cm"], r["margem_dir_cm"],
             r["margem_sup_cm"], r["margem_inf_cm"]))
    print("    mancha, largura %.2f cm" % r["mancha_larg_cm"])
    print()
    print("  COMPRIMENTO DE LINHA (so as linhas cheias)")
    print("    mediana %.0f caracteres  (menor %d, maior %d)"
          % (r["caracteres_por_linha"], r["caracteres_min"], r["caracteres_max"]))
    print()
    print("  DENSIDADE")
    print("    mediana de %.0f linhas por pagina" % r["linhas_por_pagina"])
    print()
    print("  ALCANCE")
    print("    %d linhas sozinhas na borda sairam da mancha, tratadas como"
          % r["linhas_de_borda_descartadas"])
    print("    cabecalho ou rodape.")
    print("    Nao se mediu orfa, viuva, titulo isolado nem tabela partida:")
    print("    isso exige a estrutura do .docx, que este programa nao le.")
    print()


def main():
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    ap.add_argument("pdf")
    ap.add_argument("--json", help="grava as medidas por pagina neste arquivo")
    a = ap.parse_args()
    alvo = Path(a.pdf)
    if not alvo.exists():
        sys.exit("nao encontrei %s" % alvo)
    m = medir(alvo)
    imprimir(m)
    if a.json:
        Path(a.json).write_text(json.dumps(m, ensure_ascii=False, indent=1),
                                encoding="utf-8")
        print("  medidas por pagina em %s" % a.json)


if __name__ == "__main__":
    main()

# -*- coding: utf-8 -*-
"""Mostra as formas que o .docx usa, e nao escolhe nenhuma. Nao escreve nada.

POR QUE ELE MOSTRA EM VEZ DE ORGANIZAR

Um programa que parte do que o autor fez, acha o padrao dominante e aproxima o
resto dele precisa distinguir padrao de ruido, e erra onde o padrao nao e
padrao: numa capa, que e diagramada a mao, ele a desmonta.
Contar, ele faz bem. Escolher qual forma vale e decisao de quem tem o arquivo, e
o que este programa entrega e a informacao para tomar essa decisao.

    python scripts/mostrador.py trabalho.docx
    python scripts/mostrador.py trabalho.docx --titulos     (lista os titulos, um a um)
"""
import argparse
import re
import sys
from collections import Counter
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import ooxml as O                                                    # noqa: E402

ESTRUTURAL = re.compile(
    r"^(RESUMO|ABSTRACT|RESUMEN|SUM[ÁA]RIO|REFER[ÊE]NCIAS?|BIBLIOGRAFIA"
    r"|AP[ÊE]NDICES?|ANEXOS?|LISTA DE |AGRADECIMENTOS?|DEDICAT[ÓO]RIA"
    r"|EP[ÍI]GRAFE|INTRODU[ÇC][ÃA]O|CONCLUS[ÃA]O|DECLARA[ÇC][ÃA]O)", re.I)


def descrever(tag):
    return tag.decode() if tag else "(herdado)"


def levantar(caminho):
    partes = O.ler(caminho)
    doc = partes["word/document.xml"]
    est = partes.get("word/styles.xml", b"")
    pars = O.paragrafos(doc)
    tabs = O.tabelas(doc)
    esta_em_tabela = O.dentro_de_tabela(tabs)

    itens = []
    for i, (a, b) in enumerate(pars):
        bloco = doc[a:b]
        t = O.texto(bloco).strip()
        num, nivel = O.numero_de_titulo(t)
        itens.append({
            "i": i, "ini": a, "fim": b, "texto": t,
            "tabela": esta_em_tabela(a),
            "estilo": O.estilo_de(bloco),
            "pPr": O.ppr(bloco),
            "vazio": O.vazio(bloco),
            "travado": O.travado(bloco),
            "negrito": O.negrito_inteiro(bloco),
            "sz": O.corpo_de_letra(bloco),
            "num": num, "nivel": nivel,
        })
    return {"doc": doc, "estilos_xml": est, "itens": itens,
            "tabelas": tabs, "partes": partes}


def imprimir(lev, caminho, listar_titulos=False):
    itens = lev["itens"]
    est = lev["estilos_xml"]
    defs = O.estilos_definidos(est)
    corpo = [d for d in itens if not d["tabela"]]
    tab = [d for d in itens if d["tabela"]]
    com_texto = [d for d in corpo if d["texto"]]

    print()
    print("=" * 74)
    print("O QUE HA EM %s" % Path(caminho).name)
    print("=" * 74)
    print("  %d paragrafos: %d fora de tabela (%d com texto), %d dentro de %d tabelas"
          % (len(itens), len(corpo), len(com_texto), len(tab), len(lev["tabelas"])))
    print("  %d paragrafos vazios, dos quais %d carregam quebra, ancora, campo"
          % (sum(1 for d in itens if d["vazio"]),
             sum(1 for d in itens if d["vazio"] and d["travado"])))
    print("  ou marcacao, e por isso nao sao espaco em branco.")

    print()
    print("A PAGINA")
    for s in O.secoes(lev["doc"]):
        print("    %.1f x %.1f cm, margens sup %.1f  esq %.1f  dir %.1f  inf %.1f"
              % (s.get("larg", 0), s.get("alt", 0), s.get("top", 0),
                 s.get("left", 0), s.get("right", 0), s.get("bottom", 0)))

    print()
    print("O QUE O CORPO HERDA, isto e, o estilo Normal")
    n = defs.get("Normal")
    if not n:
        print("    nao ha estilo Normal definido")
    else:
        for tag in (b"w:rFonts", b"w:sz", b"w:spacing", b"w:jc", b"w:ind"):
            m = re.search(rb"<" + tag + rb"(?:\s[^>]*)?/>", n["xml"])
            print("    %s" % (m.group(0).decode() if m else "%s: nao declara"
                              % tag.decode()))

    print()
    print("USO DE ESTILO DE PARAGRAFO")
    uso = Counter(d["estilo"] or "(Normal, sem estilo declarado)" for d in itens)
    for sid, q in uso.most_common(12):
        nome = defs.get(sid, {}).get("nome") or ""
        print("    %5d  %-30s %s" % (q, sid, nome))

    print()
    print("FORMAS DO CORPO DO TEXTO (fora de tabela, so paragrafos com texto)")
    for rotulo, chave in (("alinhamento", b"w:jc"), ("recuo", b"w:ind"),
                          ("entrelinha", b"w:spacing")):
        c = Counter(descrever(O.prop(d["pPr"], chave)) for d in com_texto)
        print("    %s" % rotulo)
        dominante, quantos = c.most_common(1)[0]
        for k, q in c.most_common(5):
            marca = "  <== dominante" if k == dominante else ""
            print("      %5d  %s%s" % (q, k, marca))
        if quantos < len(com_texto) / 2:
            print("      nenhuma forma reune metade: nao ha padrao assentado aqui")

    print()
    print("CORPO DE LETRA DECLARADO EM CORRIDA")
    c = Counter(d["sz"] for d in itens if d["texto"])
    for sz, q in sorted(c.items(), key=lambda x: (x[0] is not None, x[0] or 0)):
        rotulo = ("herdado do estilo" if sz is None
                  else "%g pt" % (sz / O.MEIO_PONTO))
        print("    %5d paragrafos  %s" % (q, rotulo))

    # --- titulos
    numerados = [d for d in corpo if d["nivel"] and len(d["texto"]) < 200]
    estruturais = [d for d in corpo
                   if not d["nivel"] and d["texto"] and len(d["texto"]) < 90
                   and ESTRUTURAL.match(d["texto"])]
    com_estilo_titulo = [d for d in itens if d["estilo"]
                         and re.match(r"(Ttulo|T[íi]tulo|Heading)\d",
                                      d["estilo"] or "")]

    print()
    print("TITULOS")
    print("    %d paragrafos numerados fora de tabela (1, 1.1, 2.3.4)" % len(numerados))
    print("    %d titulos estruturais sem numero (RESUMO, REFERENCIAS, ...)"
          % len(estruturais))
    print("    %d paragrafos usam estilo de titulo" % len(com_estilo_titulo))
    sem_estilo = [d for d in numerados + estruturais
                  if not (d["estilo"] or "").startswith(("Ttulo", "Título",
                                                         "Heading"))]
    print("    %d dos %d titulos identificados NAO usam estilo de titulo"
          % (len(sem_estilo), len(numerados) + len(estruturais)))
    if com_estilo_titulo:
        print()
        print("    quem usa estilo de titulo hoje:")
        for d in com_estilo_titulo:
            print("      [%4d] %-8s %s" % (d["i"], d["estilo"],
                                           d["texto"][:52] or "(sem texto)"))

    if listar_titulos:
        print()
        print("    OS TITULOS, UM A UM (confira antes de mandar aplicar estilo)")
        for d in numerados:
            print("      [%4d] nivel %d  %s" % (d["i"], d["nivel"], d["texto"][:70]))
        for d in estruturais:
            print("      [%4d] estrutural  %s" % (d["i"], d["texto"][:70]))

    # --- sumario
    campos, sdt = O.tem_sumario_automatico(lev["doc"])
    print()
    print("SUMARIO")
    if campos or sdt:
        print("    ha sumario automatico: %d campo(s) TOC, %d bloco(s) de conteudo"
              % (len(campos), sdt))
        for c in campos[:3]:
            print("      %s" % c.strip()[:70])
    else:
        print("    NAO ha sumario automatico: nenhum campo TOC e nenhum bloco de")
        print("    conteudo de sumario. O que houver de sumario foi digitado.")
        alvo = [d for d in itens if d["texto"].upper().startswith("SUM")]
        for d in alvo[:2]:
            print("      o paragrafo [%d] diz %r%s"
                  % (d["i"], d["texto"][:30],
                     ", e esta dentro de tabela" if d["tabela"] else ""))

    print()
    print("QUEBRAS DE PAGINA: %d"
          % sum(1 for d in itens
                if re.search(rb'<w:br[^>]*w:type="page"', lev["doc"][d["ini"]:d["fim"]])))
    print()


def main():
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("trabalho")
    ap.add_argument("--titulos", action="store_true",
                    help="lista cada titulo identificado, um por linha")
    a = ap.parse_args()
    alvo = Path(a.trabalho)
    if not alvo.exists():
        sys.exit("nao encontrei %s" % alvo)
    if alvo.suffix.lower() != ".docx":
        sys.exit("so funciona com .docx: num PDF a forma ja esta achatada.")
    imprimir(levantar(alvo), alvo, a.titulos)


if __name__ == "__main__":
    main()

# -*- coding: utf-8 -*-
"""Controle do medir_pagina.py sobre material real: desmonta uma capa de proposito.

POR QUE ISTO EXISTE

O controle_medir_pagina.py desenha paginas sinteticas e prova que o medidor le
geometria. Ele nao prova que o medidor acha o defeito que motivou a oficina.

Escrever a forma do corpo dentro do estilo `Normal` desmonta a capa: tudo o que
herda do Normal muda junto, e a capa e feita de paragrafos sem estilo proprio.
Uma contagem de paragrafo e de estilo diz que a capa nao foi tocada, porque o
texto dela continua la; so a pagina mostra o defeito.

Um trabalho qualquer nao serve de prova como esta: se a capa dele tem formatacao
direta que a protege, escrever no Normal nao a move, e o controle mediria so a
ausencia de alarme falso, que e a metade errada.

Por isso o programa constroi as DUAS pontas a partir do mesmo arquivo:

  A  a capa desprotegida, tirando dos paragrafos do pre-textual a formatacao
     direta de espacamento e recuo. E a capa comum, a que so depende do Normal.
  B  a mesma coisa de A, mais a forma de corpo escrita dentro do Normal.

A unica diferenca entre A e B e a escrita no Normal, entao qualquer mudanca na
pagina 1 vem dali.

O QUE ELE EXIGE

Que o medir_pagina.py acuse diferenca na pagina 1 entre A e B. Se ele nao
acusar, ele nao serve para o que a oficina precisa, e o silencio dele sobre
qualquer outro trabalho nao vale nada.

    python scripts/controle_capa.py trabalho.docx --saida DIR

Precisa do Word instalado, porque so ele sabe paginar um .docx como o autor o ve.
"""
import argparse
import shutil
import sys
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import medir_pagina as M                                             # noqa: E402
import exportar_pdf as X                                             # noqa: E402

import re                                                            # noqa: E402

# A forma tipica de corpo de trabalho academico brasileiro. Escrita no Normal,
# ela alcanca todo paragrafo que nao se defenda com formatacao direta.
FORMA_DE_CORPO = (b'<w:pPr><w:jc w:val="both"/>'
                  b'<w:ind w:firstLine="708"/>'
                  b'<w:spacing w:line="360" w:lineRule="auto" w:after="0"/>'
                  b"</w:pPr>")

RE_NORMAL = rb'<w:style [^>]*w:styleId="Normal".*?</w:style>'

# Quantos paragrafos do inicio se considera pre-textual para efeito do controle.
# Nao e uma regra sobre .docx: e o alcance deste teste, e so precisa cobrir a
# capa.
PARAGRAFOS_DA_CAPA = 15


def paragrafos(doc):
    """Fatias (inicio, fim) de cada <w:p>, com profundidade.

    Nao se usa `<w:p>.*?</w:p>`: caixa de texto guarda paragrafo dentro de
    paragrafo, e a busca preguicosa fecha no </w:p> de dentro, deixando um
    </w:p> orfao que impede o arquivo de abrir."""
    fatias = []
    profundidade = 0
    inicio = None
    for m in re.finditer(rb"<w:p(?:\s[^>]*)?>|</w:p>|<w:p(?:\s[^>]*)?/>", doc):
        t = m.group(0)
        if t.endswith(b"/>"):
            if profundidade == 0:
                fatias.append((m.start(), m.end()))
            continue
        if t.startswith(b"</"):
            profundidade -= 1
            if profundidade == 0 and inicio is not None:
                fatias.append((inicio, m.end()))
                inicio = None
        else:
            if profundidade == 0:
                inicio = m.start()
            profundidade += 1
    return fatias


def desproteger_capa(doc, quantos=PARAGRAFOS_DA_CAPA):
    """Tira espacamento e recuo diretos dos primeiros paragrafos.

    O alinhamento fica: sem ele a capa deixaria de ser capa, e o que se quer e
    uma capa que dependa do Normal, nao uma pagina qualquer."""
    fatias = paragrafos(doc)[:quantos]
    tirados = 0
    saida = []
    fim_anterior = 0
    for ini, fim in fatias:
        bloco = doc[ini:fim]
        novo, n1 = re.subn(rb"<w:spacing(?: [^>]*)?/>", b"", bloco)
        novo, n2 = re.subn(rb"<w:ind(?: [^>]*)?/>", b"", novo)
        tirados += n1 + n2
        saida.append(doc[fim_anterior:ini])
        saida.append(novo)
        fim_anterior = fim
    saida.append(doc[fim_anterior:])
    return b"".join(saida), len(fatias), tirados


def reescrever(origem, destino, transformar):
    """Copia o .docx aplicando transformar(nome, dados) em cada item."""
    shutil.copyfile(origem, destino)
    with zipfile.ZipFile(destino) as z:
        itens = [(i, z.read(i.filename)) for i in z.infolist()]
    saida = [(i, transformar(i.filename, d)) for i, d in itens]
    with zipfile.ZipFile(destino, "w", zipfile.ZIP_DEFLATED) as z:
        for info, dados in saida:
            z.writestr(info, dados)
    return destino


def escrever_no_normal(estilos_xml):
    """Poe a forma de corpo dentro do estilo Normal. Devolve (xml, mudou)."""
    m = re.search(RE_NORMAL, estilos_xml, re.S)
    if not m:
        sys.exit("este .docx nao tem estilo Normal, e o controle nao se aplica")
    bloco = m.group(0)
    if re.search(rb"<w:pPr>.*?</w:pPr>", bloco, re.S):
        bloco2 = re.sub(rb"<w:pPr>.*?</w:pPr>", FORMA_DE_CORPO, bloco,
                        count=1, flags=re.S)
    else:
        mm = re.search(rb"<w:name [^>]*/>", bloco)
        pos = mm.end() if mm else bloco.find(b">") + 1
        bloco2 = bloco[:pos] + FORMA_DE_CORPO + bloco[pos:]
    novo = estilos_xml[:m.start()] + bloco2 + estilos_xml[m.end():]
    return novo, novo != estilos_xml


def difere(a, b, tol_cm=0.20):
    """Duas medidas de pagina divergem de modo visivel."""
    if a["vazia"] != b["vazia"]:
        return True
    if a["vazia"]:
        return False
    if a["linhas"] != b["linhas"]:
        return True
    for chave in ("margem_sup_cm", "margem_esq_cm", "mancha_alt_cm",
                  "mancha_larg_cm"):
        if abs(a[chave] - b[chave]) > tol_cm:
            return True
    return False


def main():
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("trabalho")
    ap.add_argument("--saida", required=True)
    a = ap.parse_args()

    origem = Path(a.trabalho)
    if not origem.exists():
        sys.exit("nao encontrei %s" % origem)
    saida = Path(a.saida).resolve()
    saida.mkdir(parents=True, exist_ok=True)

    print()
    print("CONTROLE SOBRE MATERIAL REAL: capa que depende do Normal")
    print("  de %s" % origem.name)
    print()

    conta = {}

    # A, a linha de base: a capa desprotegida, e nada mais.
    def so_capa(nome, dados):
        if nome != "word/document.xml":
            return dados
        novo, vistos, tirados = desproteger_capa(dados)
        conta["paragrafos"] = vistos
        conta["tirados"] = tirados
        return novo

    A = reescrever(origem, saida / (origem.stem + "-A-capa-exposta.docx"), so_capa)
    print("  A  %d paragrafos do inicio examinados, %d propriedades diretas de"
          % (conta.get("paragrafos", 0), conta.get("tirados", 0)))
    print("     espacamento ou recuo tiradas da capa")
    if not conta.get("tirados"):
        print()
        print("  a capa deste arquivo nao tinha espacamento nem recuo direto para")
        print("  tirar. O controle nao se aplica a ele, e dizer que o medidor")
        print("  passou seria dizer nada.")
        sys.exit(2)

    # B, a unica diferenca: a forma de corpo dentro do Normal.
    conta["normal"] = False

    def so_normal(nome, dados):
        if nome != "word/styles.xml":
            return dados
        novo, mudou = escrever_no_normal(dados)
        conta["normal"] = mudou
        return novo

    B = reescrever(A, saida / (origem.stem + "-B-normal-trocado.docx"), so_normal)
    print("  B  forma de corpo escrita no Normal: %s"
          % ("sim" if conta["normal"] else "NAO"))
    if not conta["normal"]:
        print()
        print("  o Normal nao mudou, entao A e B sao o mesmo arquivo e o controle")
        print("  nao testa nada.")
        sys.exit(2)
    print()

    resultados = X.exportar([A, B], saida)
    if X.conferir(resultados):
        sys.exit("a exportacao falhou, e sem PDF nao ha o que medir")
    print()

    pa = M.medir(saida / (A.stem + ".pdf"))["paginas"]
    pb = M.medir(saida / (B.stem + ".pdf"))["paginas"]

    mudou_capa = difere(pa[0], pb[0])
    print("  %s  a pagina 1 e acusada como diferente"
          % ("PASSA " if mudou_capa else "REPROVA"))
    for chave, rotulo in (("linhas", "linhas"),
                          ("margem_sup_cm", "margem superior"),
                          ("margem_esq_cm", "margem esquerda"),
                          ("mancha_alt_cm", "altura da mancha")):
        va, vb = pa[0].get(chave), pb[0].get(chave)
        if isinstance(va, float):
            print("           %-18s %6.2f -> %6.2f" % (rotulo, va, vb))
        else:
            print("           %-18s %6s -> %6s" % (rotulo, va, vb))

    n = min(len(pa), len(pb))
    mudadas = [i + 1 for i in range(n) if difere(pa[i], pb[i])]
    print()
    print("  %d de %d paginas comparaveis mudaram; a primeira foi a %s"
          % (len(mudadas), n, mudadas[0] if mudadas else "nenhuma"))
    print("  paginas: %d -> %d" % (len(pa), len(pb)))
    print()
    if not mudou_capa:
        print("  REPROVADO: o medidor nao viu a capa mudar. Ele nao esta")
        print("  certificado para material real, e nada que ele disser sobre um")
        print("  trabalho vale enquanto isso.")
        sys.exit(1)
    print("  o medidor ve na capa o defeito que a contagem de paragrafo nao viu.")
    print()


if __name__ == "__main__":
    main()

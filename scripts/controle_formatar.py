# -*- coding: utf-8 -*-
"""Controle positivo do formatar.py: exige que ele acerte e que a conferencia reprove.

POR QUE ISTO EXISTE

A conferencia do `formatar.py` diz que o texto continua o mesmo e que o XML esta
bem formado. Uma conferencia que nunca reprovou nada nao informa nada: quando ela
se cala, nao da para distinguir arquivo bom de conferidor quebrado.

SAO DUAS METADES, E AS DUAS SAO PRECISAS

  efeito       o programa fez o que diz que faz: o Normal carrega a norma, os
               titulos passaram a usar estilo de titulo, quem herda ficou
               blindado, o campo de sumario entrou.
  sensibilidade a conferencia REPROVA um arquivo quebrado de proposito. Sem esta
               metade, a primeira so mostra que o programa nao falhou hoje.

O .docx de teste e desenhado aqui, com a estrutura minima de uma dissertacao:
capa, sumario digitado em tabela, titulos numerados, corpo e tabela de dados.
O gabarito e o codigo, e nao ha arquivo de fora.

    python scripts/controle_formatar.py
"""
import re
import shutil
import sys
import tempfile
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import ooxml as O                                                    # noqa: E402
import formatar as F                                                 # noqa: E402

NS = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"')


def p(texto, ppr=""):
    t = ("<w:r><w:t xml:space=\"preserve\">%s</w:t></w:r>" % texto) if texto else ""
    pp = "<w:pPr>%s</w:pPr>" % ppr if ppr else ""
    return "<w:p>%s%s</w:p>" % (pp, t)


def celula(texto):
    return ("<w:tc><w:tcPr><w:tcW w:w=\"2000\" w:type=\"dxa\"/></w:tcPr>%s</w:tc>"
            % p(texto, '<w:jc w:val="center"/>'))


def tabela(linhas):
    trs = "".join("<w:tr>%s</w:tr>" % "".join(celula(c) for c in linha)
                  for linha in linhas)
    return "<w:tbl><w:tblPr><w:tblW w:w=\"0\" w:type=\"auto\"/></w:tblPr>%s</w:tbl>" % trs


DOCUMENTO = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document {ns}><w:body>{corpo}</w:body></w:document>"""

ESTILOS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles {ns}>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/>
<w:qFormat/><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>
<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Ttulo1"><w:name w:val="heading 1"/>
<w:basedOn w:val="Normal"/><w:qFormat/></w:style>
<w:style w:type="paragraph" w:styleId="Ttulo2"><w:name w:val="heading 2"/>
<w:basedOn w:val="Normal"/><w:qFormat/>
<w:pPr><w:jc w:val="center"/><w:outlineLvl w:val="1"/></w:pPr>
<w:rPr><w:b/><w:bCs/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Textodenotaderodap"><w:name w:val="footnote text"/>
<w:basedOn w:val="Normal"/><w:rPr><w:sz w:val="20"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Recuodecorpodetexto"><w:name w:val="Body Text Indent"/>
<w:basedOn w:val="Normal"/><w:pPr><w:ind w:left="3402"/><w:jc w:val="both"/></w:pPr></w:style>
<w:style w:type="paragraph" w:customStyle="1" w:styleId="msonormal0"><w:name w:val="msonormal"/>
<w:basedOn w:val="Normal"/><w:pPr><w:spacing w:before="100" w:beforeAutospacing="1"/></w:pPr></w:style>
<w:style w:type="paragraph" w:customStyle="1" w:styleId="Derivado"><w:name w:val="Derivado"/>
<w:basedOn w:val="msonormal0"/></w:style>
</w:styles>"""

CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
<Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>
</Types>"""

RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

DOC_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>
<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>
</Relationships>"""

SETTINGS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings {ns}><w:defaultTabStop w:val="708"/><w:footnotePr><w:footnote w:id="-1"/><w:footnote w:id="0"/></w:footnotePr></w:settings>"""

# O separador das notas sem recuo proprio, e com um paragrafo vazio a mais, como
# o Word costuma gravar.
NOTAS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:footnotes {ns}><w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p><w:p/></w:footnote><w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote></w:footnotes>"""

HEADER = r"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr {ns}><w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r>
<w:r><w:instrText>PAGE   \* MERGEFORMAT</w:instrText></w:r>
<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>2</w:t></w:r>
<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p></w:hdr>"""


def desenhar(caminho):
    """Um .docx minimo com a estrutura de uma dissertacao."""
    corpo = []
    # capa: centrada, com um paragrafo usando estilo de titulo sem ser titulo
    corpo.append(p("UNIVERSIDADE FEDERAL", '<w:jc w:val="center"/>'))
    # Sem `jc` proprio: o centro e o negrito vem do estilo de titulo, e tirar o
    # estilo os leva junto se nada for congelado antes.
    corpo.append(p("Fulana de Tal", '<w:pStyle w:val="Ttulo2"/>'))
    # Paragrafo vazio autofechado, que e como o Word grava linha em branco.
    corpo.append("<w:p/>")
    corpo.append(p(""))
    corpo.append(p("TITULO DO TRABALHO", '<w:jc w:val="center"/>'))
    corpo.append(p("Cidade", '<w:jc w:val="center"/>'))
    # Segunda linha de bloco da capa com entrelinha escrita a mao.
    corpo.append(p("2026", '<w:spacing w:after="40" w:line="360" w:lineRule="auto"/>'
                   '<w:jc w:val="center"/>'))
    # Folha de rosto: usa estilo proprio, que carrega o recuo e a justificacao,
    # e declara `hanging`, que um `firstLine` vindo de estilo cancelaria.
    corpo.append(p("Dissertacao apresentada ao Programa de Pos-Graduacao.",
                   '<w:pStyle w:val="Recuodecorpodetexto"/>'))
    corpo.append(p("_________________ (Orientador)",
                   '<w:pStyle w:val="Recuodecorpodetexto"/>'
                   '<w:ind w:hanging="1098"/>'))
    corpo.append(p("_______________ (Orientador)",
                   '<w:pStyle w:val="Recuodecorpodetexto"/>'))
    corpo.append(p("Dedicatoria", '<w:jc w:val="center"/>'))
    corpo.append(p("A quem me ensinou.", '<w:jc w:val="center"/>'))
    # Epigrafe logo depois da dedicatoria: as duas secoes no pe ficam coladas, e
    # e o caso em que uma corrida encontra a secao da outra ja aberta.
    corpo.append(p("Epigrafe", '<w:jc w:val="center"/>'))
    corpo.append(p("Uma frase de alguem.", '<w:jc w:val="right"/>'))
    corpo.append(p("LISTA DE GRAFICOS E TABELAS", '<w:jc w:val="center"/>'))
    corpo.append(p("Grafico 1 - Uma legenda qualquer ..... 9"))
    corpo.append(p("Tabela 1 - Uma tabela ..... 10"))
    corpo.append(p("LISTA DE ABREVIATURAS", '<w:jc w:val="center"/>'))
    corpo.append(p("Siglas", '<w:jc w:val="left"/>'))
    corpo.append(p("ABNT - Associacao Brasileira de Normas Tecnicas"))
    # Espaco repetido no pre-textual: la ele fica.
    corpo.append(p("UNB - Universidade  de Brasilia"))
    corpo.append(p("AGRADECIMENTOS", '<w:jc w:val="center"/>'))
    # texto corrido do pre-textual, sem recuo e a esquerda: ganha o do corpo
    corpo.append(p("Agradeco a quem leu este trabalho.",
                   '<w:ind w:firstLine="0"/><w:jc w:val="left"/>'))
    corpo.append(p("RESUMO", '<w:jc w:val="center"/>'))
    corpo.append(p("Texto do resumo.", '<w:jc w:val="both"/>'))
    # Vazio autofechado que sobrevive: nao esta na capa nem antes de titulo. Sem
    # ele, o teste de protecao abaixo nao teria o que testar.
    corpo.append("<w:p/>")
    # Palavras-chave com corpo escrito a mao em cada corrida, que nao pode
    # vencer o corpo do resumo.
    corpo.append('<w:p><w:pPr><w:jc w:val="both"/></w:pPr>'
                 '<w:r><w:rPr><w:b/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr>'
                 '<w:t xml:space="preserve">Palavras-chave: </w:t></w:r>'
                 '<w:r><w:rPr><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr>'
                 '<w:t>Primeira expressao; Universidade de Brasilia. Segunda '
                 'expressao; IBGE</w:t></w:r></w:p>')
    # Lista de ilustracoes ja automatica, com entradas numeradas e com
    # dois-pontos: nao podem virar titulo, nem ser divididas, nem ser trocadas.
    corpo.append(p("LISTA DE FIGURAS", '<w:jc w:val="center"/>'))
    corpo.append('<w:p><w:pPr><w:pStyle w:val="ndicedeilustraes"/></w:pPr>'
                 '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
                 '<w:r><w:instrText xml:space="preserve"> TOC \\h \\z \\c "Figura" </w:instrText></w:r>'
                 '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
                 '<w:r><w:t>1 Primeira figura: detalhe</w:t></w:r></w:p>')
    corpo.append('<w:p><w:pPr><w:pStyle w:val="ndicedeilustraes"/></w:pPr>'
                 '<w:r><w:t>2 Segunda figura</w:t></w:r>'
                 '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>')
    corpo.append(p("SUMARIO", '<w:jc w:val="center"/>'))
    corpo.append(tabela([["1", "INTRODUCAO", "8"], ["2", "DESENVOLVIMENTO", "12"]]))
    # parte textual
    corpo.append(p("1 INTRODUCAO", '<w:jc w:val="left"/>'
                                   '<w:spacing w:line="360" w:lineRule="auto"/>'))
    for k in range(4):
        corpo.append(p("Paragrafo de corpo numero %d, com texto suficiente." % k,
                       '<w:jc w:val="both"/><w:ind w:firstLine="709"/>'
                       '<w:spacing w:line="360" w:lineRule="auto"/>'))
        if k == 1:
            # vazio depois de paragrafo de corpo, que tem recuo: sai
            corpo.append(p(""))
    # texto colado da web, no estilo msonormal: passa ao Normal
    corpo.append(p("Paragrafo colado da web, em msonormal.",
                   '<w:pStyle w:val="msonormal0"/>'))
    # o corpo com lista desligada e o recuo escrito a mao: o Word zeraria o
    # recuo se a lista ficasse e o recuo saisse
    corpo.append(p("Paragrafo de corpo com lista desligada.",
                   '<w:numPr><w:ilvl w:val="1"/><w:numId w:val="0"/></w:numPr>'
                   '<w:spacing w:line="360" w:lineRule="auto"/>'
                   '<w:ind w:firstLine="709"/><w:jc w:val="both"/>'))
    corpo.append(p("Paragrafo  com  espaco  duplo, no corpo."))
    corpo.append(p("METODOLOGIA DE ANALISE"))
    corpo.append(p("Paragrafo empurrado para baixo com espaco posto a mao.",
                   '<w:spacing w:before="1440"/>'))
    # Corrida de vazios antes de um titulo que quebra pagina, com um vazio
    # guardado no meio: os outros saem, esse fica, e a varredura nao para nele.
    corpo.append(p(""))
    corpo.append(p(""))
    corpo.append('<w:p><w:bookmarkStart w:id="1" w:name="marca"/>'
                 '<w:bookmarkEnd w:id="1"/></w:p>')
    corpo.append(p(""))
    corpo.append(p(""))
    # Com o marcador que o sumario do Word poe no titulo: a divisao no
    # dois-pontos nao pode perde-lo.
    corpo.append('<w:p><w:pPr><w:jc w:val="left"/></w:pPr>'
                 '<w:bookmarkStart w:id="7" w:name="_Toc777"/>'
                 '<w:r><w:rPr><w:b/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr>'
                 '<w:t xml:space="preserve">2 DESENVOLVIMENTO: o que explica o '
                 'capitulo</w:t></w:r><w:bookmarkEnd w:id="7"/></w:p>')
    # Com a lista desligada que o Word deixa quando o autor tira a numeracao:
    # ela desfaz o recuo pendente do titulo.
    corpo.append(p("2.1 Primeira parte", '<w:numPr><w:ilvl w:val="1"/>'
                   '<w:numId w:val="0"/></w:numPr><w:jc w:val="left"/>'))
    # vazio encostado em titulo: sai
    corpo.append(p(""))
    # Titulo que ja traz o ponto: nao pode sair com dois.
    corpo.append(p("2.2. Segunda parte", '<w:jc w:val="left"/>'))
    # Titulo de terceiro nivel, que tem de virar Titulo 3.
    corpo.append(p("2.2.1 Terceiro nivel", '<w:jc w:val="left"/>'))
    # Titulo de quarto nivel: com a numeracao automatica, passa a letra.
    corpo.append(p("2.2.1.1 Quarto nivel", '<w:pStyle w:val="Ttulo4"/>'))
    corpo.append('<w:p><w:pPr><w:ind w:left="2268"/></w:pPr><w:r><w:rPr>'
                 '<w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr>'
                 '<w:t>Citacao longa recuada, de mais de tres linhas.</w:t></w:r></w:p>')
    # Vazios do autor em volta da ilustracao: antes do titulo e entre o titulo e
    # a figura, onde quebram a cadeia do keepNext.
    corpo.append(p(""))
    corpo.append(p(""))
    # Titulo de grafico com negrito posto a mao.
    corpo.append('<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/>'
                 '<w:bCs/></w:rPr><w:t>Grafico 1 - Uma legenda qualquer</w:t></w:r></w:p>')
    corpo.append(p(""))
    corpo.append('<w:p><w:r><w:drawing/></w:r></w:p>')
    corpo.append('<w:p><w:r><w:rPr><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr>'
                 '<w:t>Fonte: elaboracao propria.</w:t></w:r></w:p>')
    corpo.append(p("Texto comum depois da fonte, que nao e legenda."))
    corpo.append(tabela([["Categoria", "N"], ["Uma", "10"], ["Outra", "20"]]))
    # Titulo de referencias numerado, que a faixa tem de reconhecer.
    corpo.append(p("3 REFERENCIAS", '<w:jc w:val="left"/>'))
    corpo.append(p("AUTOR, Um. Obra. Cidade: Editora, 2020.",
                   '<w:jc w:val="left"/><w:ind w:firstLine="709"/>'))
    corpo.append(p(""))
    corpo.append(p("BUTOR, Dois. Outra obra. Cidade: Editora, 2021.",
                   '<w:jc w:val="left"/>'))
    corpo.append(p("APENDICES", '<w:jc w:val="left"/>'))
    corpo.append(p("APENDICE A - Base de dados", '<w:pStyle w:val="Ttulo2"/>'))
    corpo.append('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
                 '<w:pgMar w:top="1701" w:right="1134" w:bottom="1134" '
                 'w:left="1701"/></w:sectPr>')
    doc = DOCUMENTO.format(ns=NS, corpo="".join(corpo))

    with zipfile.ZipFile(caminho, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CONTENT_TYPES)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/_rels/document.xml.rels", DOC_RELS)
        z.writestr("word/document.xml", doc)
        z.writestr("word/styles.xml", ESTILOS.format(ns=NS))
        z.writestr("word/settings.xml", SETTINGS.format(ns=NS))
        z.writestr("word/header1.xml", HEADER.format(ns=NS))
        z.writestr("word/footnotes.xml", NOTAS.format(ns=NS))
    return caminho


class Placar:
    def __init__(self):
        self.falhas = 0

    def exigir(self, nome, ok, obtido="", esperado=""):
        if not ok:
            self.falhas += 1
        print("  %s  %-52s %s" % ("PASSA " if ok else "REPROVA", nome,
                                  ("obtido %s, esperado %s" % (obtido, esperado))
                                  if esperado else obtido))


def sem_a_regra(origem, tmp, trecho, troca, estilos=False, destino_de=None):
    """Formata com uma copia do programa em que `trecho` virou `troca`, e
    devolve o document.xml (ou o styles.xml). O trecho tem de existir uma vez so."""
    fonte = Path(F.__file__).read_text(encoding="utf-8")
    assert fonte.count(trecho) == 1, "trecho do controle nao casa: %r" % trecho[:50]
    import types
    G = types.ModuleType("F_sem_regra")
    G.__file__ = F.__file__
    exec(compile(fonte.replace(trecho, troca), F.__file__, "exec"), G.__dict__)
    G.NORMA["numero_de_titulo"]["automatica"] = False
    destino = tmp / "sem-regra.docx"
    G.formatar(origem, destino, corpo_pt=13.0, entrelinha=1.3, fonte="Times New Roman")
    if destino_de:
        return destino
    return O.ler(destino)["word/styles.xml" if estilos else "word/document.xml"]


def com_lista_do_word(origem, tmp, desfazer=False):
    """O trabalho em que um titulo ja e numerado pelo Word e outro tem o numero
    digitado. Devolve (numeracao automatica ligada?, o titulo manteve a lista?)."""
    import types
    fonte = Path(F.__file__).read_text(encoding="utf-8")
    if desfazer:
        trecho = "\n                    and not any(numera_pelo_word(f) for f in titulos_))"
        assert fonte.count(trecho) == 1
        fonte = fonte.replace(trecho, ")")
    G = types.ModuleType("F_lista_word")
    G.__file__ = F.__file__
    exec(compile(fonte, F.__file__, "exec"), G.__dict__)
    G.NORMA["numero_de_titulo"]["automatica"] = True
    d = O.ler(origem)["word/document.xml"]
    alvo = b'<w:t xml:space="preserve">2.2. Segunda parte</w:t>'
    i = d.index(alvo)
    a = d.rindex(b"<w:p>", 0, i)
    d = d[:a] + (b'<w:p><w:pPr><w:pStyle w:val="Ttulo2"/>'
                 b'<w:numPr><w:ilvl w:val="1"/><w:numId w:val="7"/></w:numPr>'
                 b'</w:pPr><w:r><w:t xml:space="preserve">Segunda parte</w:t></w:r></w:p>') \
        + d[d.index(b"</w:p>", i) + len(b"</w:p>"):]
    fonte_docx = tmp / "lista-do-word.docx"
    with zipfile.ZipFile(origem) as zi, \
            zipfile.ZipFile(fonte_docx, "w", zipfile.ZIP_DEFLATED) as zo:
        for info in zi.infolist():
            zo.writestr(info, d if info.filename == "word/document.xml"
                        else zi.read(info.filename))
    saida = tmp / "lista-do-word-formatado.docx"
    res = G.formatar(fonte_docx, saida, corpo_pt=13.0, entrelinha=1.3,
                     fonte="Times New Roman")
    novo = O.ler(saida)["word/document.xml"]
    seg = next((novo[a:b] for a, b in O.paragrafos(novo)
                if O.texto(novo[a:b]).strip() == "Segunda parte"), b"")
    return res["numeracao_auto"], b'<w:numId w:val="7"/>' in seg


def variante_de_numero(origem, tmp, nome, refs=b"", cabecalho=None, rodape=None):
    """O arquivo de teste com a secao final apontando para `refs`; `cabecalho`
    troca o header1.xml, e `rodape`, se dado, entra como footer1.xml."""
    saida = tmp / ("numero-%s.docx" % nome)
    with zipfile.ZipFile(origem) as zi, zipfile.ZipFile(saida, "w", zipfile.ZIP_DEFLATED) as zo:
        for info in zi.infolist():
            dados = zi.read(info.filename)
            if info.filename == "word/document.xml":
                k = dados.rindex(b"<w:sectPr>") + len(b"<w:sectPr>")
                dados = dados[:k] + refs + dados[k:]
            elif info.filename == "word/header1.xml" and cabecalho is not None:
                dados = cabecalho
            elif info.filename == "word/_rels/document.xml.rels" and rodape is not None:
                dados = dados.replace(b"</Relationships>", (
                    b'<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/'
                    b'officeDocument/2006/relationships/footer" Target="footer1.xml"/>'
                    b"</Relationships>"))
            zo.writestr(info, dados)
        if rodape is not None:
            zo.writestr("word/footer1.xml", rodape)
    return saida


def numero_na_parte_textual(destino):
    """(a secao final mostra numero, o pre-textual nao mostra, partes novas)"""
    with zipfile.ZipFile(destino) as zz:
        nomes = zz.namelist()
        doc = zz.read("word/document.xml")
        rels = zz.read("word/_rels/document.xml.rels").decode()
        partes = {n: zz.read(n) for n in nomes if n.startswith("word/")}
    alvo = dict(re.findall(r'Id="([^"]+)"[^>]*Target="([^"]+)"', rels))
    secs = re.findall(rb"<w:sectPr\b.*?</w:sectPr>", doc, re.S)

    def mostra(sp):
        ids = re.findall(rb'<w:(?:header|footer)Reference w:type="default" r:id="([^"]+)"', sp)
        return any(b"PAGE" in partes.get("word/" + alvo.get(i.decode(), ""), b"") for i in ids)
    return mostra(secs[-1]), not mostra(secs[0]), "word/header-numero.xml" in nomes


def como_esta(destino):
    """As exigencias da norma como esta, sobre um arquivo formatado: cabecalho
    so com o numero de pagina, e titulo com dois-pontos inteiro."""
    with zipfile.ZipFile(destino) as zz:
        cabs = [zz.read(n) for n in zz.namelist() if re.match(r"word/header\d*\.xml$", n)]
        conf = zz.read("word/settings.xml")
    doc = O.ler(destino)["word/document.xml"]
    com_page = [c for c in cabs if b"PAGE" in c]
    par = next((c[a:b] for c in com_page for a, b in O.paragrafos(c) if b"PAGE" in c[a:b]), b"")
    textos = [O.texto(doc[a:b]).strip() for a, b in O.paragrafos(doc)]
    return [
        ("o cabecalho traz so o numero de pagina, sem titulo",
         bool(com_page) and not any(b"STYLEREF" in c for c in cabs),
         "com STYLEREF" if any(b"STYLEREF" in c for c in cabs) else "so PAGE"),
        ("...a direita e no estilo Cabealho, que da o italico",
         O.estilo_de(par) == "Cabealho" and b'<w:jc w:val="right"/>' in O.ppr(par),
         "%s, %s" % (O.estilo_de(par), O.atributo(O.ppr(par), b"<w:jc", b"w:val"))),
        ("sem hifenizacao: ela so vai na versao de impressao",
         b'<w:autoHyphenation w:val="true"/>' not in conf,
         "ligada" if b'<w:autoHyphenation w:val="true"/>' in conf else "desligada"),
        ("o titulo com dois-pontos fica inteiro",
         any(t.endswith("DESENVOLVIMENTO: o que explica o capitulo") for t in textos),
         [t for t in textos if "DESENVOLVIMENTO" in t][:2]),
    ]


def numeracao_automatica(origem, tmp, pular=None):
    """Formata com a numeracao automatica e devolve as exigencias, sem julgar.

    `pular` desfaz uma peca do programa, para o controle positivo.
    """
    import importlib.util
    spec = importlib.util.spec_from_file_location("F_num", F.__file__)
    G = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(G)
    G.NORMA["numero_de_titulo"]["automatica"] = True
    G.NORMA["cabecalho"]["titulo_no_cabecalho"] = True
    G.NORMA["cabecalho"]["dividir_no_dois_pontos"] = True
    if pular == "com_lista":
        G.com_lista = lambda x, a, l: x or (b'<?xml version="1.0" encoding="UTF-8" '
                                            b'standalone="yes"?><w:numbering xmlns:w="http://'
                                            b'schemas.openxmlformats.org/wordprocessingml/'
                                            b'2006/main"></w:numbering>')
    if pular == "titulo_no_cabecalho":
        sem_numero = G.titulo_no_cabecalho
        G.titulo_no_cabecalho = lambda x, larg, com_numero=False: sem_numero(x, larg)
    if pular == "com_numpr":
        G.com_numpr = lambda pp, ilvl, n: G.sem_numpr(pp)
    destino = tmp / ("auto-%s.docx" % (pular or "inteiro"))
    res = G.formatar(origem, destino, corpo_pt=13.0, entrelinha=1.3,
                     fonte="Times New Roman")
    with zipfile.ZipFile(destino) as zz:
        nomes = zz.namelist()
        doc = zz.read("word/document.xml")
        est = zz.read("word/styles.xml")
        num = zz.read("word/numbering.xml") if "word/numbering.xml" in nomes else b""
        rels = zz.read("word/_rels/document.xml.rels")
        ct = zz.read("[Content_Types].xml")
        hdr = zz.read("word/header1.xml")
    defs = O.estilos_definidos(est)
    nid = res.get("num_id")
    pars = [doc[a:b] for a, b in O.paragrafos(doc)]
    por_texto = {O.texto(x).strip(): x for x in pars}
    out = []

    tits = {"INTRODUCAO": 0, "DESENVOLVIMENTO": 0, "Primeira parte": 1,
            "Segunda parte": 1, "Terceiro nivel": 2, "Quarto nivel": 3,
            "REFERENCIAS": 0}
    presentes = [t for t in tits if t in por_texto]
    out.append(("o numero digitado sai do texto do titulo",
                len(presentes) == len(tits)
                and not any(re.match(r"\d", O.texto(x).strip()) for x in pars
                            if O.estilo_de(x) and O.estilo_de(x).startswith("Ttulo")),
                presentes, list(tits)))
    abst = re.search(rb"<w:abstractNum .*?</w:abstractNum>", num, re.S)
    abst = abst.group(0) if abst else b""
    ligados = re.findall(rb'<w:pStyle w:val="(Ttulo\d)"/><w:lvlText w:val="([^"]*)"', abst)
    out.append(("a lista de quatro niveis ligada aos estilos Titulo 1 a 4",
                ligados == [(b"Ttulo1", b"%1"), (b"Ttulo2", b"%1.%2"),
                            (b"Ttulo3", b"%1.%2.%3"), (b"Ttulo4", b"%4)")],
                ligados, "Ttulo1..3 com numero, Ttulo4 com %4)"))
    out.append(("...o Titulo 4 com letra",
                b'<w:numFmt w:val="lowerLetter"/><w:pStyle w:val="Ttulo4"/>' in abst))
    out.append(("...numero sem ponto, e o Titulo 1 com o recuo pendente do estilo",
                b'<w:ind w:left="709" w:hanging="709"/></w:pPr></w:lvl><w:lvl w:ilvl="1"' in abst
                and not re.search(rb'lvlText w:val="[^"]*\."', abst)))
    out.append(("...e a parte existe para o Word: relacao e tipo",
                b"relationships/numbering" in rels and b"/word/numbering.xml" in ct
                or b'PartName="/word/numbering.xml"' in ct))
    faltam = [sid for sid in ("Ttulo1", "Ttulo2", "Ttulo3", "Ttulo4")
              if b'<w:numId w:val="%d"/>' % (nid or -1) not in defs.get(sid, {}).get("xml", b"")]
    out.append(("os estilos Titulo 1 a 4 levam a lista", not faltam, faltam, "nenhum"))
    t4 = defs.get("Ttulo4", {}).get("xml", b"")
    out.append(("...com o numPr antes do spacing, na ordem do esquema",
                0 <= t4.find(b"<w:numPr>") < t4.find(b"<w:spacing")))
    errados = []
    for t, k in tits.items():
        x = por_texto.get(t, b"")
        m = re.search(rb'<w:numPr><w:ilvl w:val="(\d)"/><w:numId w:val="(\d+)"/>', O.ppr(x))
        if not m or int(m.group(1)) != k or int(m.group(2)) != nid:
            errados.append(t)
    out.append(("os titulos numerados levam a lista, no nivel certo", not errados, errados, "nenhum"))
    sem = [t for t in ("APENDICES",) if b'<w:numId w:val="0"/>' not in O.ppr(por_texto.get(t, b""))]
    out.append(("o titulo que nao tinha numero leva a lista desligada", not sem, sem, "nenhum"))
    sub = por_texto.get("o que explica o capitulo", b"")
    out.append(("o subtitulo nao e numerado", bool(sub) and b"<w:numPr>" not in sub))
    out.append(("o cabecalho traz o numero do capitulo pelo campo",
                rb"STYLEREF 1 \n" in hdr, "presente" if rb"STYLEREF 1 \n" in hdr else "ausente"))
    out.append(("a conferencia passa com os numeros tirados",
                not G.conferir(res, destino), "; ".join(G.conferir(res, destino))[:60] or "nada"))
    out.append(("o relatorio avisa so o titulo cujo numero muda: o de quarto nivel",
                [(a, b) for _, a, b in res.get("avisos_numero", [])] == [("2.2.1.1", "a)")],
                res.get("avisos_numero"), "[('2.2.1.1', 'a)')]"))
    return out


def main():
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass
    tmp = Path(tempfile.mkdtemp(prefix="controle-formatar-"))
    origem = desenhar(tmp / "teste.docx")
    destino = tmp / "teste-formatado.docx"

    print()
    print("CONTROLE POSITIVO DE formatar.py")
    print("  .docx de teste desenhado em %s" % tmp)
    print()
    # A primeira metade exercita o numero digitado, que ganha tabulacao; a
    # numeracao automatica tem a sua parte adiante, sobre o mesmo arquivo.
    # Ela tambem liga o titulo no cabecalho e a divisao no dois-pontos, que a
    # norma desligou: o codigo continua, e so se sabe que funciona exercitando.
    # A norma como esta tem a sua parte, adiante.
    F.NORMA["numero_de_titulo"]["automatica"] = False
    F.NORMA["cabecalho"]["titulo_no_cabecalho"] = True
    F.NORMA["cabecalho"]["dividir_no_dois_pontos"] = True
    F.NORMA["hifenizar"] = True
    res = F.formatar(origem, destino, corpo_pt=13.0, entrelinha=1.3,
                     fonte="Times New Roman")
    F.NORMA["hifenizar"] = False
    F.NORMA["numero_de_titulo"]["automatica"] = True
    F.NORMA["cabecalho"]["titulo_no_cabecalho"] = False
    F.NORMA["cabecalho"]["dividir_no_dois_pontos"] = False
    z = Placar()

    novo = O.ler(destino)
    doc, est = novo["word/document.xml"], novo["word/styles.xml"]
    defs = O.estilos_definidos(est)

    print("  PRIMEIRA METADE: o programa faz o que diz")
    normal = defs["Normal"]["xml"]
    z.exigir("o Normal passa a 13 pt",
             b'<w:sz w:val="26"/>' in normal,
             O.atributo(normal, b"<w:sz", b"w:val"), "26")
    z.exigir("o Normal passa a entrelinha 1,3 (312 de 240)",
             b'w:line="312"' in normal,
             O.atributo(normal, b"<w:spacing", b"w:line"), "312")
    z.exigir("o Normal passa a Times New Roman",
             b'w:ascii="Times New Roman"' in normal,
             O.atributo(normal, b"<w:rFonts", b"w:ascii"), "Times New Roman")

    for sid in ("PreTextual", "TabelaTrabalho", "Legenda",
                "Ttulo0", "FiguraGrafico", "Referencia"):
        z.exigir("o estilo %s foi criado" % sid, sid in defs)

    # A tabela continua protegida; o pre-textual, por decisao da norma, passou a
    # seguir o corpo e a entrelinha do texto.
    tab = defs.get("TabelaTrabalho", {}).get("xml", b"")
    z.exigir("TabelaTrabalho devolve o corpo de 12 pt a quem herda",
             b'<w:sz w:val="24"/>' in tab)
    z.exigir("TabelaTrabalho devolve a entrelinha simples a quem herda",
             b'w:line="240"' in tab)
    pt_ = defs.get("PreTextual", {}).get("xml", b"")
    z.exigir("PreTextual leva o corpo do texto, e nao o antigo",
             b"<w:sz " not in pt_, "declara sz" if b"<w:sz " in pt_ else "nao declara",
             "nao declara")
    z.exigir("PreTextual leva entrelinha simples, por decisao da norma",
             b'w:line="240"' in pt_,
             O.atributo(pt_, b"<w:spacing", b"w:line"), "240")

    nota = defs.get("Textodenotaderodap", {}).get("xml", b"")
    z.exigir("a nota de rodape deixa de herdar a entrelinha do corpo",
             b"<w:spacing" in nota)

    pars = O.paragrafos(doc)

    def tt(bloco):
        """O texto com a tabulacao, que `O.texto` nao mostra."""
        return "".join(
            "\t" if x == b"<w:tab/>" else
            re.sub(rb"<[^>]+>", b"", x).decode("utf-8")
            for x in re.findall(rb"<w:t(?:\s[^>]*)?>.*?</w:t>|<w:tab/>", bloco, re.S))

    papel = {}
    for a, b in pars:
        bloco = doc[a:b]
        papel[tt(bloco).strip()] = O.estilo_de(bloco)
    z.exigir("a Introducao recebe Ttulo1", papel.get("1\tINTRODUCAO") == "Ttulo1",
             papel.get("1\tINTRODUCAO"), "Ttulo1")
    z.exigir("'2.1<tab>Primeira parte' recebe Ttulo2",
             papel.get("2.1\tPrimeira parte") == "Ttulo2",
             papel.get("2.1\tPrimeira parte"), "Ttulo2")

    textos_ = [tt(doc[a:b]).strip() for a, b in pars]
    z.exigir("a Introducao e numerada, como as demais secoes primarias",
             "1\tINTRODUCAO" in textos_,
             [x for x in textos_ if x.endswith("INTRODUCAO")], "1<tab>INTRODUCAO")
    z.exigir("numero sem ponto e seguido de tabulacao, nos dois niveis",
             "2\tDESENVOLVIMENTO" in textos_ and "2.1\tPrimeira parte" in textos_,
             [x for x in textos_ if x[:1] == "2"], "2<tab> e 2.1<tab>")
    z.exigir("o ponto que o autor pos depois do numero sai (NBR 6024)",
             "2.2\tSegunda parte" in textos_,
             [x for x in textos_ if x.startswith("2.2")], "2.2<tab>Segunda parte")
    z.exigir("nenhum capitulo e renumerado",
             [x for x in textos_ if re.match(r"\d+\t[A-Z]", x)]
             == ["1\tINTRODUCAO", "2\tDESENVOLVIMENTO", "3\tREFERENCIAS"],
             [x for x in textos_ if re.match(r"\d+\t[A-Z]", x)],
             ["1<tab>INTRODUCAO", "2<tab>DESENVOLVIMENTO", "3<tab>REFERENCIAS"])
    z.exigir("'3 REFERENCIAS' recebe Ttulo1", papel.get("3\tREFERENCIAS") == "Ttulo1",
             papel.get("3\tREFERENCIAS"), "Ttulo1")
    z.exigir("'Fulana de Tal' PERDE o estilo de titulo que tinha na capa",
             papel.get("Fulana de Tal") == "PreTextual",
             papel.get("Fulana de Tal"), "PreTextual")

    # A capa nao pode perder o centro nem o negrito que vinham do estilo de
    # titulo que se tira dela.
    capa = next((doc[a:b] for a, b in pars
                 if O.texto(doc[a:b]).strip() == "Fulana de Tal"), b"")
    z.exigir("...e mantem o centro que vinha do estilo",
             b'<w:jc w:val="center"/>' in O.ppr(capa),
             str(O.prop(O.ppr(capa), b"w:jc")), '<w:jc w:val="center"/>')
    z.exigir("...e mantem o negrito que vinha do estilo",
             all(re.search(rb"<w:b(?:\s[^>]*)?/>", r) for r in O.corridas(capa)),
             "corridas: %d, em negrito: %d"
             % (len(O.corridas(capa)),
                sum(1 for r in O.corridas(capa)
                    if re.search(rb"<w:b(?:\s[^>]*)?/>", r))))

    # Trocar o estilo proprio da folha de rosto por PreTextual levaria junto o
    # recuo e a justificacao que vinham dele.
    z.exigir("o paragrafo do pre-textual MANTEM o estilo proprio que ja tinha",
             papel.get("Dissertacao apresentada ao Programa de Pos-Graduacao.")
             == "Recuodecorpodetexto",
             papel.get("Dissertacao apresentada ao Programa de Pos-Graduacao."),
             "Recuodecorpodetexto")
    rec = defs.get("Recuodecorpodetexto", {}).get("xml", b"")
    # Estilo usado so no pre-textual solta o corpo, para seguir o Normal; o que
    # ele declara de proprio, como o recuo, fica.
    z.exigir("...e esse estilo solta o corpo, para seguir o Normal",
             b"<w:sz " not in rec,
             "declara sz" if b"<w:sz " in rec else "nao declara", "nao declara")
    z.exigir("...e o recuo que ele declarava continua la",
             b'w:left="3402"' in rec, str(O.prop(rec, b"w:ind")), "left 3402")
    assin = next((doc[a:b] for a, b in pars
                  if O.texto(doc[a:b]).strip().startswith("____")), b"")
    z.exigir("o recuo negativo (hanging) da assinatura sobrevive",
             b'w:hanging="1098"' in (O.prop(O.ppr(assin), b"w:ind") or b"")
             and b"firstLine" not in (O.prop(O.ppr(assin), b"w:ind") or b""),
             str(O.prop(O.ppr(assin), b"w:ind")), 'hanging sem firstLine')
    for sid in ("PreTextual", "TabelaTrabalho"):
        z.exigir("%s devolve a ausencia de recuo a quem herda" % sid,
                 b'<w:ind w:firstLine="0"/>' in defs.get(sid, {}).get("xml", b""))

    # So o pre-textual: o vazio da parte textual nao leva estilo de protecao,
    # porque nada la herda forma que precise ser devolvida.
    corte = next((k for k, (a, b) in enumerate(pars)
                  if tt(doc[a:b]).strip() == "1\tINTRODUCAO"), len(pars))
    vazios = [doc[a:b] for a, b in pars[:corte] if not O.texto(doc[a:b]).strip()]
    protegidos = [v for v in vazios if O.estilo_de(v)]
    z.exigir("o paragrafo vazio autofechado tambem recebe estilo de protecao",
             len(vazios) >= 1 and len(protegidos) == len(vazios),
             "%d de %d vazios do pre-textual protegidos"
             % (len(protegidos), len(vazios)), "todos")
    z.exigir("'RESUMO' vai para Ttulo0, que tem a forma do Ttulo1",
             papel.get("RESUMO") == "Ttulo0", papel.get("RESUMO"), "Ttulo0")
    z.exigir("'Categoria', dentro de tabela, recebe TabelaTrabalho",
             papel.get("Categoria") == "TabelaTrabalho",
             papel.get("Categoria"), "TabelaTrabalho")
    z.exigir("'Grafico 1 - ...' recebe a legenda do Word",
             papel.get("Grafico 1 - Uma legenda qualquer") == "Legenda",
             papel.get("Grafico 1 - Uma legenda qualquer"), "Legenda")
    leg_p = next((doc[a:b] for a, b in pars
                  if O.texto(doc[a:b]).strip() == "Grafico 1 - Uma legenda qualquer"
                  and O.estilo_de(doc[a:b]) == "Legenda"), b"")
    z.exigir("o numero do grafico vai num campo SEQ, e o texto nao muda",
             "SEQ Gráfico" in leg_p.decode("utf-8"),
             "com SEQ" if b"SEQ" in leg_p else "sem SEQ", "com SEQ")
    z.exigir("o titulo da ilustracao perde o negrito posto a mao",
             not any(re.search(rb"<w:b(?:\s[^>]*)?/>", r) for r in O.corridas(leg_p)))
    z.exigir("o titulo da ilustracao fica com a figura (keepNext)",
             b"<w:keepNext/>" in O.ppr(leg_p))
    k_leg = next(k for k, (a, b) in enumerate(pars)
                 if O.texto(doc[a:b]).strip() == "Grafico 1 - Uma legenda qualquer"
                 and O.estilo_de(doc[a:b]) == "Legenda")
    viz = [O.texto(doc[a:b]).strip() or ("[figura]" if b"<w:drawing" in doc[a:b] else "[vazio]")
           for a, b in pars[k_leg - 1:k_leg + 3]]
    z.exigir("os vazios em volta da ilustracao saem",
             viz[1:] == ["Grafico 1 - Uma legenda qualquer", "[figura]",
                         "Fonte: elaboracao propria."] and viz[0] != "[vazio]",
             viz, "texto, titulo, figura, fonte")
    z.exigir("o titulo da ilustracao leva 12 pt depois",
             b'w:after="240"' in O.ppr(leg_p),
             O.atributo(O.prop(O.ppr(leg_p), b"w:spacing") or b"", b"<w:spacing", b"w:after"),
             "240")
    refs = [O.texto(doc[a:b]).strip() for a, b in pars]
    i_ref = refs.index("AUTOR, Um. Obra. Cidade: Editora, 2020.")
    z.exigir("o vazio entre as referencias sai",
             refs[i_ref + 1] == "BUTOR, Dois. Outra obra. Cidade: Editora, 2021.",
             repr(refs[i_ref + 1]), "BUTOR...")
    cit = next((doc[a:b] for a, b in pars
                if O.texto(doc[a:b]).strip().startswith("Citacao longa")), b"")
    z.exigir("a citacao larga o corpo escrito a mao",
             O.estilo_de(cit) == "Citao" and b"<w:sz " not in cit,
             O.estilo_de(cit), "Citao, sem sz")
    fonte_p = next((doc[a:b] for a, b in pars
                    if O.texto(doc[a:b]).strip() == "Fonte: elaboracao propria."), b"")
    z.exigir("a legenda larga o corpo escrito a mao, e segue o estilo",
             b"<w:sz " not in leg_p + fonte_p)
    tof = defs.get("ndicedeilustraes", {}).get("xml", b"")
    z.exigir("o estilo da lista de ilustracoes e o do Word, sem customStyle",
             bool(tof) and b"customStyle" not in est[est.find(b'w:styleId="ndicedeilustraes"') - 80:
                                                    est.find(b'w:styleId="ndicedeilustraes"')],
             "ausente" if not tof else "presente", "presente, embutido")
    lg = defs.get("Legenda", {}).get("xml", b"")
    z.exigir("Legenda: 11 pt, em italico",
             b'<w:sz w:val="22"/>' in lg and b"<w:i/>" in lg,
             "%s, %s" % (O.atributo(lg, b"<w:sz", b"w:val"),
                         "italico" if b"<w:i/>" in lg else "sem italico"),
             "22, italico")

    corpo1 = next((doc[a:b] for a, b in pars
                   if O.texto(doc[a:b]).strip().startswith("Paragrafo de corpo")),
                  b"")
    z.exigir("o corpo perde a entrelinha direta, para seguir o Normal",
             O.prop(O.ppr(corpo1), b"w:spacing") is None,
             str(O.prop(O.ppr(corpo1), b"w:spacing")), "None")
    z.exigir("o corpo nao escreve alinhamento nem recuo: vem do Normal",
             O.prop(O.ppr(corpo1), b"w:jc") is None
             and O.prop(O.ppr(corpo1), b"w:ind") is None,
             "jc=%s ind=%s" % (O.prop(O.ppr(corpo1), b"w:jc"),
                               O.prop(O.ppr(corpo1), b"w:ind")), "os dois None")

    print()
    print("  A NORMA, item por item")
    n = defs["Normal"]["xml"]
    z.exigir("Normal: 3 pt depois", b'w:after="60"' in n,
             O.atributo(n, b"<w:spacing", b"w:after"), "60")
    z.exigir("Normal: recuo de primeira linha de 1,25 cm",
             b'w:firstLine="709"' in n,
             O.atributo(n, b"<w:ind", b"w:firstLine"), "709")
    z.exigir("Normal: justificado", b'<w:jc w:val="both"/>' in n)

    t1 = defs["Ttulo1"]["xml"]
    z.exigir("Titulo 1: a esquerda (NBR 6024, 4.1, c)", b'<w:jc w:val="left"/>' in t1,
             O.atributo(t1, b"<w:jc", b"w:val"), "left")
    z.exigir("Titulo 1: entrelinha simples, como todos os titulos",
             all(b'w:line="240"' in defs.get(sid, {}).get("xml", b"")
                 for sid in ("Ttulo0", "Ttulo1", "Ttulo2", "Ttulo3", "SubtituloCapitulo")),
             [(sid, O.atributo(defs.get(sid, {}).get("xml", b""), b"<w:spacing", b"w:line"))
              for sid in ("Ttulo0", "Ttulo1", "Ttulo2", "Ttulo3", "SubtituloCapitulo")],
             "240 em todos")
    z.exigir("Titulo 1: corpo 16", b'<w:sz w:val="32"/>' in t1,
             O.atributo(t1, b"<w:sz", b"w:val"), "32")
    z.exigir("Titulo 1: negrito", b"<w:b/>" in t1)
    z.exigir("Titulo 1: no alto da mancha, 36 pt depois",
             b'w:before="0"' in t1 and b'w:after="720"' in t1,
             "before=%s after=%s" % (O.atributo(t1, b"<w:spacing", b"w:before"),
                                     O.atributo(t1, b"<w:spacing", b"w:after")),
             "0 e 720")
    z.exigir("Titulo 1: parada de 1,25 cm, a coluna do titulo depois do numero",
             b'<w:tab w:val="left" w:pos="709"/>' in t1,
             O.atributo(t1, b"<w:tab", b"w:pos"), "709")
    z.exigir("Titulo 1: quebra pagina antes", b"<w:pageBreakBefore/>" in t1)
    z.exigir("Titulo 1: recuo pendente de 1,25 cm, o do corpo",
             b'<w:ind w:left="709" w:hanging="709"/>' in t1,
             str(O.prop(O.ppr(t1) if b"<w:p>" in t1 else t1, b"w:ind")), "left 709 hanging 709")
    apx = next((doc[a:b] for a, b in pars if O.texto(doc[a:b]).strip() == "APENDICES"), b"")
    z.exigir("o Titulo 1 sem numero fica no centro e sem recuo (NBR 14724, 5.2.3)",
             b'<w:jc w:val="center"/>' in O.ppr(apx)
             and b'<w:ind w:left="0" w:firstLine="0"/>' in O.ppr(apx),
             O.ppr(apx)[:90], "jc center, ind 0")
    z.exigir("Titulo 1: sem hifenizacao", b"<w:suppressAutoHyphens/>" in t1)

    t0 = defs["Ttulo0"]["xml"]
    def forma(x):
        import re as _re
        y = _re.search(rb"<w:pPr>(.*?)</w:pPr>", x, _re.S)
        y = y.group(1) if y else b""
        # A parada do Titulo 1 e do numero, e o Titulo 0 nao tem numero.
        y = _re.sub(rb"<w:tabs>.*?</w:tabs>", b"", y, flags=_re.S)
        # E o alinhamento e o recuo: o Titulo 0, sem numero, fica no centro.
        y = _re.sub(rb"<w:jc [^>]*/>|<w:ind [^>]*/>", b"", y)
        return _re.sub(rb'<w:outlineLvl w:val="\d"/>', b"", y)
    z.exigir("Titulo 0: centralizado", b'<w:jc w:val="center"/>' in t0)
    z.exigir("Titulo 0: no mais, a forma do Titulo 1", forma(t0) == forma(t1),
             "iguais" if forma(t0) == forma(t1) else "diferentes")
    z.exigir("Titulo 0: fora do sumario (outlineLvl 9)",
             b'<w:outlineLvl w:val="9"/>' in t0,
             O.atributo(t0, b"<w:outlineLvl", b"w:val"), "9")

    if "CabealhodoSumrio" in defs:
        z.exigir("Cabecalho do Sumario: forma do Titulo 0",
                 forma(defs["CabealhodoSumrio"]["xml"]) == forma(t0))

    t2 = defs["Ttulo2"]["xml"]
    t21 = next((doc[a:b] for a, b in pars if tt(doc[a:b]).strip() == "2.1	Primeira parte"), b"")
    z.exigir("o titulo perde a lista desligada, que desfaz o recuo pendente",
             bool(t21) and b"<w:numPr>" not in O.ppr(t21))
    z.exigir("Titulo 2: parada explicita na coluna do titulo",
             b'<w:tab w:val="left" w:pos="709"/>' in t2,
             O.atributo(t2, b"<w:tab", b"w:pos"), "709")
    z.exigir("Titulo 2: 18 pt antes e 12 pt depois",
             b'w:before="360"' in t2 and b'w:after="240"' in t2,
             "before=%s after=%s" % (O.atributo(t2, b"<w:spacing", b"w:before"),
                                     O.atributo(t2, b"<w:spacing", b"w:after")), "360 e 240")
    z.exigir("Titulo 2: recuo pendente de 1,25 cm e sem hifenizacao",
             b'<w:ind w:left="709" w:hanging="709"/>' in t2
             and b"<w:suppressAutoHyphens/>" in t2,
             str(O.prop(t2, b"w:ind")), "left 709 hanging 709")

    ct = defs["Citao"]["xml"]
    z.exigir("Citacao: um ponto menor que o corpo", b'<w:sz w:val="24"/>' in ct,
             O.atributo(ct, b"<w:sz", b"w:val"), "24")
    z.exigir("Citacao: recuo esquerdo de 2,5 cm e sem recuo de primeira",
             b'w:left="1417"' in ct and b'w:firstLine="0"' in ct,
             str(O.prop(ct, b"w:ind")), 'left 1417 e firstLine 0')
    z.exigir("Citacao: 12 pt antes e depois",
             b'w:before="240"' in ct and b'w:after="240"' in ct)
    z.exigir("o paragrafo recuado recebe Citacao",
             papel.get("Citacao longa recuada, de mais de tres linhas.") == "Citao",
             papel.get("Citacao longa recuada, de mais de tres linhas."), "Citao")

    fg = defs["FiguraGrafico"]["xml"]
    z.exigir("Figura: centralizada e sem recuo",
             b'<w:jc w:val="center"/>' in fg and b'<w:ind w:firstLine="0"/>' in fg)
    fig = next((doc[a:b] for a, b in pars if b"<w:drawing" in doc[a:b]), b"")
    z.exigir("o paragrafo da figura recebe FiguraGrafico",
             O.estilo_de(fig) == "FiguraGrafico", O.estilo_de(fig), "FiguraGrafico")
    z.exigir("o texto logo depois da figura vira legenda",
             papel.get("Fonte: elaboracao propria.") == "Legenda",
             papel.get("Fonte: elaboracao propria."), "Legenda")
    z.exigir("o texto seguinte, que nao e legenda, NAO vira legenda",
             papel.get("Texto comum depois da fonte, que nao e legenda.") is None,
             papel.get("Texto comum depois da fonte, que nao e legenda."), "None")

    z.exigir("os vazios antes de titulo que quebra pagina sao apagados",
             len(res["apagados"]) == 4, len(res["apagados"]), 4)
    guardado = [a for a, b in pars if b"<w:bookmarkStart" in doc[a:b]
                and not O.texto(doc[a:b]).strip()]
    z.exigir("o vazio que carrega marcador sobrevive, e nao protege os de tras",
             len(guardado) == 1, "%d vazio(s) com marcador" % len(guardado), 1)

    res_ = defs["ResumoTrabalho"]["xml"]
    z.exigir("Resumo: entrelinha simples e sem recuo de primeira linha",
             b'w:line="240"' in res_ and b'<w:ind w:firstLine="0"/>' in res_)
    z.exigir("o corpo do resumo recebe ResumoTrabalho",
             papel.get("Texto do resumo.") == "ResumoTrabalho",
             papel.get("Texto do resumo."), "ResumoTrabalho")
    z.exigir("a capa perde os paragrafos vazios e ganha espaco antes",
             all(O.texto(doc[a:b]).strip() or O.estilo_de(doc[a:b])
                 for a, b in pars[:6]),
             "", "")

    capa_sub = next((doc[a:b] for a, b in pars
                     if O.texto(doc[a:b]).strip() == "TITULO DO TRABALHO"), b"")
    capa_1 = next((doc[a:b] for a, b in pars
                   if O.texto(doc[a:b]).strip() == "UNIVERSIDADE FEDERAL"), b"")
    z.exigir("a capa vai em entrelinha simples",
             b'w:line="240"' in (O.prop(O.ppr(capa_1), b"w:spacing") or b""),
             str(O.prop(O.ppr(capa_1), b"w:spacing")), "line 240")
    z.exigir("a capa nao se hifeniza",
             b"<w:suppressAutoHyphens/>" in O.ppr(capa_sub))

    s1 = defs["Sumrio1"]["xml"]
    s2 = defs["Sumrio2"]["xml"]
    z.exigir("Sumario nivel 1: negrito", b"<w:b/>" in s1)
    z.exigir("Sumario nivel 1: recuo pendente, para a linha que dobra",
             b'w:hanging="510"' in s1, str(O.prop(s1, b"w:ind")), "hanging 510")
    z.exigir("Sumario nivel 2: sem negrito, numero na margem (NBR 6027)",
             b"<w:b/>" not in s2 and b'<w:ind w:left="510" w:hanging="510"/>' in s2,
             str(O.prop(s2, b"w:ind")), "left 510 hanging 510")
    z.exigir("Sumario: tabulacao pontilhada na largura da coluna",
             b'w:leader="dot"' in s1 and abs(int(O.atributo(s1, b"<w:tab", b"w:pos")) - 8165) <= 2,
             O.atributo(s1, b"<w:tab", b"w:pos"), "8165 +- 2")

    lst = defs["ListaPreTextual"]["xml"]
    z.exigir("Entrada de lista: recuo pendente, para alinhar a continuacao",
             b'w:hanging="680"' in lst, str(O.prop(lst, b"w:ind")), "hanging 680")
    z.exigir("a entrada da lista recebe ListaPreTextual",
             papel.get("ABNT - Associacao Brasileira de Normas Tecnicas")
             == "ListaPreTextual",
             papel.get("ABNT - Associacao Brasileira de Normas Tecnicas"),
             "ListaPreTextual")
    z.exigir("o subtitulo da lista NAO e entrada",
             papel.get("Siglas") != "ListaPreTextual",
             papel.get("Siglas"), "outra coisa")

    apr = next((doc[a:b] for a, b in pars
                if O.texto(doc[a:b]).strip().startswith("____")), b"")
    z.exigir("a folha de aprovacao abre em folha propria",
             b"<w:pageBreakBefore/>" in O.ppr(apr))

    textos = [O.texto(doc[a:b]).strip() for a, b in pars]
    z.exigir("a dedicatoria perde o titulo, que a NBR nao quer",
             "Dedicatoria" not in textos,
             "ainda la" if "Dedicatoria" in textos else "saiu", "saiu")
    ded = next((doc[a:b] for a, b in pars
                if O.texto(doc[a:b]).strip() == "A quem me ensinou."), b"")
    ded_def = defs.get("Dedicatoria", {}).get("xml", b"")
    larg = int(round((21.0 - 3.3 - 3.3) * O.TWIPS_POR_CM))
    z.exigir("Dedicatoria: sem recuo de primeira linha, texto a partir do meio",
             b'w:firstLine="0"' in ded_def
             and abs(int(O.atributo(ded_def, b"<w:ind", b"w:left") or 0)
                     - larg // 2) <= 2,
             str(O.prop(ded_def, b"w:ind")), "left %d" % (larg // 2))
    z.exigir("o texto da dedicatoria recebe o estilo Dedicatoria",
             O.estilo_de(ded) == "Dedicatoria", O.estilo_de(ded), "Dedicatoria")
    z.exigir("...e fecha uma secao alinhada ao pe da pagina",
             b'<w:vAlign w:val="bottom"/>' in O.ppr(ded))
    k_ded = next(k for k, (a, b) in enumerate(pars)
                 if O.texto(doc[a:b]).strip() == "A quem me ensinou.")
    ant = doc[pars[k_ded - 1][0]:pars[k_ded - 1][1]]
    z.exigir("o paragrafo antes dela fecha a secao anterior, que nao vai ao pe",
             b"<w:sectPr" in O.ppr(ant) and b"<w:vAlign" not in O.ppr(ant))
    secs = re.findall(rb"<w:sectPr\b.*?</w:sectPr>", doc, re.S)
    com_inicio = [k for k, x in enumerate(secs) if b"<w:pgNumType" in x]
    z.exigir("so a primeira secao diz onde a contagem comeca",
             com_inicio == [0], com_inicio, [0])
    z.exigir("a dedicatoria nao guarda quebra de pagina, que a secao ja da",
             b"<w:pageBreakBefore/>" not in O.ppr(ded))

    pc = next((doc[a:b] for a, b in pars
               if O.texto(doc[a:b]).strip().startswith("Palavras-chave")), b"")
    z.exigir("as palavras-chave ficam no corpo do resumo",
             O.estilo_de(pc) == "ResumoTrabalho"
             and not any(b"<w:sz " in r for r in O.corridas(pc)),
             "%s, %d corrida(s) com corpo a mao"
             % (O.estilo_de(pc), sum(b"<w:sz " in r for r in O.corridas(pc))),
             "ResumoTrabalho, 0")
    z.exigir("...e o negrito do rotulo fica",
             any(b"<w:b/>" in r for r in O.corridas(pc)))

    z.exigir("Referencia: sem recuo de primeira linha",
             b'<w:ind w:firstLine="0"/>' in defs["Referencia"]["xml"])
    z.exigir("a entrada da bibliografia recebe Referencia",
             papel.get("AUTOR, Um. Obra. Cidade: Editora, 2020.") == "Referencia",
             papel.get("AUTOR, Um. Obra. Cidade: Editora, 2020."), "Referencia")

    z.exigir("margens 3,3 / 3,3 / 3,3 / 2,5 cm",
             all(abs(O.secoes(doc)[0][k] - v) < 0.02 for k, v in
                 (("left", 3.3), ("right", 3.3), ("top", 3.3), ("bottom", 2.5))),
             " ".join("%s=%.2f" % (k, O.secoes(doc)[0][k])
                      for k in ("left", "right", "top", "bottom")),
             "3,30 3,30 3,30 2,50")

    import zipfile as _zip
    with _zip.ZipFile(destino) as _z:
        conf = _z.read("word/settings.xml")
        hdr = _z.read("word/header1.xml")
    z.exigir("hifenizacao do documento ligada",
             b'<w:autoHyphenation w:val="true"/>' in conf)
    z.exigir("hifenizacao logo depois do defaultTabStop, como o esquema pede",
             conf.index(b"<w:autoHyphenation")
             > conf.index(b"<w:defaultTabStop"))
    z.exigir("no maximo 2 hifens em linhas seguidas",
             b'<w:consecutiveHyphenLimit w:val="2"/>' in conf)
    # Com o titulo corrente, o numero vai para a direita por tabulacao, e nao
    # por alinhamento: o paragrafo passou a ter conteudo a esquerda.
    z.exigir("numero de pagina a direita, por tabulacao",
             b'<w:tab w:val="right"' in hdr and b"<w:tab/>" in hdr)
    z.exigir("o cabecalho le o titulo vivo, pelo campo STYLEREF",
             b"STYLEREF 1" in hdr,
             "STYLEREF 1" if b"STYLEREF 1" in hdr else "ausente", "STYLEREF 1")
    z.exigir("o cabecalho NAO guarda copia do titulo",
             b"TITULO DO CAPITULO" not in hdr)

    # O titulo com dois-pontos vira titulo e subtitulo, e o subtitulo fica fora
    # do sumario.
    z.exigir("o titulo com dois-pontos e dividido",
             papel.get("2\tDESENVOLVIMENTO") == "Ttulo1"
             and papel.get("o que explica o capitulo") == "SubtituloCapitulo",
             "%s / %s" % (papel.get("2\tDESENVOLVIMENTO"),
                          papel.get("o que explica o capitulo")),
             "Ttulo1 / SubtituloCapitulo")
    sub = defs.get("SubtituloCapitulo", {}).get("xml", b"")
    z.exigir("o subtitulo fica fora do sumario",
             b'<w:outlineLvl w:val="9"/>' in sub)
    z.exigir("o estilo do subtitulo nao declara negrito",
             b"<w:b/>" not in sub)
    # E o resultado, que e o que importa: a direta do paragrafo vence o estilo,
    # e conferir so o estilo nao basta.
    sub_p = next((doc[a:b] for a, b in pars
                  if O.estilo_de(doc[a:b]) == "SubtituloCapitulo"), b"")
    z.exigir("e o subtitulo sai mesmo sem negrito, no paragrafo",
             not any(re.search(rb"<w:b(?:\s[^>]*)?/>", r)
                     for r in O.corridas(sub_p)),
             "%d de %d corridas em negrito"
             % (sum(1 for r in O.corridas(sub_p)
                    if re.search(rb"<w:b(?:\s[^>]*)?/>", r)),
                len(O.corridas(sub_p))), "nenhuma")

    # A lista de ilustracoes digitada sai, e entra o campo do Word.
    z.exigir("a lista de ilustracoes vira campo, pelo rotulo do SEQ",
             'TOC \\h \\z \\c "Gráfico"' in doc.decode("utf-8"))
    z.exigir("...e as entradas digitadas saem, a da tabela inclusive",
             not any(x.startswith(("Grafico 1 - Uma legenda qualquer .",
                                   "Tabela 1 - Uma tabela")) for x in textos_))
    z.exigir("...e o titulo perde o ' E TABELAS'",
             "LISTA DE GRAFICOS" in textos_,
             [x for x in textos_ if x.startswith("LISTA DE G")], "LISTA DE GRAFICOS")

    pc2 = next((tt(doc[a:b]) for a, b in pars
                if O.texto(doc[a:b]).strip().startswith("Palavras-chave")), "")
    z.exigir("palavras-chave como a NBR 6028: minuscula, ponto e virgula, ponto final",
             pc2.strip() == "Palavras-chave: primeira expressao; Universidade de "
                            "Brasilia; segunda expressao; IBGE.",
             pc2.strip(), "...; IBGE.")

    ref_d = defs.get("Referencia", {}).get("xml", b"")
    z.exigir("Referencia: espaco simples e 12 pt depois",
             b'w:line="240"' in ref_d and b'w:after="240"' in ref_d,
             str(O.prop(ref_d, b"w:spacing")), "line 240 after 240")
    z.exigir("nota de rodape em 12 pt",
             b'<w:sz w:val="24"/>' in defs.get("Textodenotaderodap", {}).get("xml", b""))
    z.exigir("cabecalho e numero de pagina em 10 pt",
             b'<w:sz w:val="20"/>' in defs.get("Cabealho", {}).get("xml", b""),
             O.atributo(defs.get("Cabealho", {}).get("xml", b""), b"<w:sz", b"w:val"),
             "20")
    with zipfile.ZipFile(destino) as _zc:
        _hdr = _zc.read("word/header1.xml")
    _p_num = next((_hdr[a:b] for a, b in O.paragrafos(_hdr) if b"PAGE" in _hdr[a:b]), b"")
    z.exigir("...e o paragrafo do cabecalho usa o estilo Cabealho",
             O.estilo_de(_p_num) == "Cabealho", O.estilo_de(_p_num), "Cabealho")

    secs = re.findall(rb"<w:sectPr\b.*?</w:sectPr>", doc, re.S)
    # 3,3 cm sao 1871 twips; a linha de 10 pt ocupa 230: o alto dela a 820.
    z.exigir("a linha do numero de pagina no meio da margem superior",
             all(b'w:header="820"' in x for x in secs),
             [O.atributo(x, b"<w:pgMar", b"w:header") for x in secs], "820 em todas")
    z.exigir("a capa e a folha 0: a contagem comeca na folha de rosto",
             b'<w:pgNumType w:start="0"/>' in secs[0], O.atributo(secs[0], b"<w:pgNumType", b"w:start"), "0")

    nat = next((doc[a:b] for a, b in pars
                if O.texto(doc[a:b]).strip().startswith("Dissertacao apresentada")), b"")
    z.exigir("...em espaco simples",
             b'w:line="240"' in (O.prop(O.ppr(nat), b"w:spacing") or b""),
             str(O.prop(O.ppr(nat), b"w:spacing")), "line 240")
    z.exigir("...sem recuo de primeira linha",
             b'w:firstLine="0"' in (O.prop(O.ppr(nat), b"w:ind") or b""),
             str(O.prop(O.ppr(nat), b"w:ind")), "firstLine 0")
    z.exigir("a natureza do trabalho comeca no meio da mancha",
             O.atributo(O.prop(O.ppr(nat), b"w:ind") or b"", b"<w:ind", b"w:left") == "4082",
             O.atributo(O.prop(O.ppr(nat), b"w:ind") or b"", b"<w:ind", b"w:left"), "4082")

    epi = next((doc[a:b] for a, b in pars
                if O.texto(doc[a:b]).strip() == "Uma frase de alguem."), b"")
    z.exigir("a epigrafe perde o titulo e recebe a forma da dedicatoria",
             "Epigrafe" not in textos_ and O.estilo_de(epi) == "Epigrafe",
             O.estilo_de(epi), "Epigrafe")
    z.exigir("...e fecha a sua propria secao, no pe",
             b'<w:vAlign w:val="bottom"/>' in O.ppr(epi))
    z.exigir("a dedicatoria, colada nela, continua no pe",
             b'<w:vAlign w:val="bottom"/>' in O.ppr(ded))
    z.exigir("o apendice tem a forma da secao primaria",
             papel.get("APENDICE A - Base de dados") == "Ttulo1",
             papel.get("APENDICE A - Base de dados"), "Ttulo1")

    print()
    print("  O QUE VEIO DA NORMA ANTIGA")
    with zipfile.ZipFile(destino) as _zn:
        notas = _zn.read("word/footnotes.xml")
    seps = re.findall(rb'<w:footnote w:type="(?:separator|continuationSeparator)".*?</w:footnote>',
                      notas, re.S)
    z.exigir("1. o separador das notas comeca na margem",
             len(seps) == 2 and all(b'<w:ind w:left="0" w:firstLine="0"/>' in x for x in seps),
             [bool(b'w:firstLine="0"' in x) for x in seps], "[True, True]")
    z.exigir("1. ...e perde o paragrafo vazio a mais",
             len(re.findall(rb"<w:p[ >/]", seps[0])) == 1,
             len(re.findall(rb"<w:p[ >/]", seps[0])), 1)

    # 2. a porta das alteracoes: o mesmo documento, com uma insercao de quem orienta
    rev = tmp / "com-revisao.docx"
    with zipfile.ZipFile(origem) as _zo, zipfile.ZipFile(rev, "w") as _zr:
        for it in _zo.infolist():
            dd = _zo.read(it.filename)
            if it.filename == "word/document.xml":
                dd = dd.replace(
                    b"Texto comum depois da fonte, que nao e legenda.</w:t></w:r>",
                    b"Texto comum depois da fonte, que nao e legenda.</w:t></w:r>"
                    b'<w:ins w:id="901" w:author="Orientador" w:date="2026-09-01T00:00:00Z">'
                    b"<w:r><w:t>X</w:t></w:r></w:ins>", 1)
            _zr.writestr(it, dd)
    try:
        F.formatar(rev, tmp / "com-revisao-formatado.docx")
        recusou = ""
    except SystemExit as e:
        recusou = str(e.code)
    z.exigir("2. recusa arquivo com alteracao controlada, e diz de quem e",
             "Orientador" in recusou, recusou.strip().splitlines()[0][:50] if recusou else "rodou",
             "recusa, com o autor")

    # 3. as travas novas, uma a uma, e o vazio sem nada, que nao e travado
    casos = [("quebra de pagina antes", b"<w:p><w:pPr><w:pageBreakBefore/></w:pPr></w:p>", True),
             ("quebra de coluna", b'<w:p><w:r><w:br w:type="column"/></w:r></w:p>', True),
             ("fim de marcador", b'<w:p><w:bookmarkEnd w:id="3"/></w:p>', True),
             ("referencia de comentario", b'<w:p><w:r><w:commentReference w:id="1"/></w:r></w:p>', True),
             ("texto apagado sob revisao", b"<w:p><w:r><w:delText>x</w:delText></w:r></w:p>", True),
             ("quebra de pagina desligada", b'<w:p><w:pPr><w:pageBreakBefore w:val="0"/></w:pPr></w:p>', False),
             ("vazio sem nada", b"<w:p><w:pPr><w:jc w:val=\"left\"/></w:pPr></w:p>", False)]
    erradas = [n for n, x, esp in casos if O.travado(x) != esp]
    z.exigir("3. as travas da Norma antiga protegem, e o vazio sem nada nao",
             not erradas, erradas or "todas certas", "todas certas")
    z.exigir("3. ...mas a quebra nao protege o vazio antes de titulo que ja quebra",
             not O.travado(casos[0][1], aceitar_quebra=True))

    z.exigir("4. o titulo de referencias numerado abre a faixa",
             papel.get("AUTOR, Um. Obra. Cidade: Editora, 2020.") == "Referencia",
             papel.get("AUTOR, Um. Obra. Cidade: Editora, 2020."), "Referencia")

    z.exigir("5. espaco repetido vira um so no corpo",
             "Paragrafo com espaco duplo, no corpo." in textos_,
             [x for x in textos_ if x.startswith("Paragrafo com") or x.startswith("Paragrafo  com")],
             "Paragrafo com espaco duplo, no corpo.")
    z.exigir("5. ...e fica no pre-textual",
             "UNB - Universidade  de Brasilia" in textos_)

    i1 = textos_.index("Paragrafo de corpo numero 1, com texto suficiente.")
    p1 = doc[pars[i1][0]:pars[i1][1]]
    esp1 = O.atributo(O.prop(O.ppr(p1), b"w:spacing") or b"", b"<w:spacing", b"w:after")
    z.exigir("6. vazio depois de paragrafo de corpo sai, sem virar espaco",
             textos_[i1 + 1].startswith("Paragrafo de corpo numero 2") and esp1 is None,
             "seguinte: %r, after=%s" % (textos_[i1 + 1][:28], esp1), "numero 2, after None")
    pl = next((doc[a:b] for a, b in pars
               if O.texto(doc[a:b]).strip() == "Paragrafo de corpo com lista desligada."), b"")
    z.exigir("o corpo perde a lista desligada junto com o recuo escrito",
             bool(pl) and b"<w:numPr>" not in O.ppr(pl) and b"<w:ind" not in O.ppr(pl),
             O.ppr(pl)[:80], "sem numPr e sem ind")
    amostra = (b'<w:p><w:pPr><w:spacing w:before="0" w:after="480"/></w:pPr><w:r><w:rPr>'
               b'<w:spacing w:val="15"/></w:rPr><w:t>2.1 Titulo</w:t></w:r></w:p>')
    limpa = F.sem_espacejamento(amostra)
    z.exigir("o titulo perde o espacejamento de letras escrito na corrida",
             b'w:val="15"' not in limpa)
    z.exigir("...e o espacamento do paragrafo fica",
             b'<w:spacing w:before="0" w:after="480"/>' in limpa)
    # Controle positivo das duas regras: desfeitas numa copia do programa,
    # cada exigencia tem de reprovar.
    dl = sem_a_regra(origem, tmp, "                pc = sem_numpr(pc)", "                pass")
    pl2 = next((dl[a:b] for a, b in O.paragrafos(dl)
                if O.texto(dl[a:b]).strip() == "Paragrafo de corpo com lista desligada."), b"")
    z.exigir("controle: sem a regra, a lista desligada fica", b"<w:numPr>" in O.ppr(pl2))
    dc = sem_a_regra(origem, tmp, "                    and not numera_pelo_word(f)):",
                     "                    and False):")
    apx2 = next((dc[a:b] for a, b in O.paragrafos(dc)
                 if O.texto(dc[a:b]).strip() == "APENDICES"), b"")
    z.exigir("controle: sem a regra, o titulo sem numero nao vai ao centro",
             b'<w:jc w:val="center"/>' not in O.ppr(apx2))
    e_ent = sem_a_regra(origem, tmp, 'linha_tit = int(round(240 * NORMA["titulos_entrelinha"]))',
                        "linha_tit = linha", estilos=True)
    z.exigir("controle: sem a regra, o Titulo 1 volta a entrelinha do corpo",
             b'w:line="240"' not in O.estilos_definidos(e_ent)["Ttulo1"]["xml"])
    dv = sem_a_regra(origem, tmp, 'or pd in ENCOSTA or pa == "corpo"):', "or pd in ENCOSTA):")
    tv = [O.texto(dv[a:b]).strip() for a, b in O.paragrafos(dv)]
    iv = tv.index("Paragrafo de corpo numero 1, com texto suficiente.")
    z.exigir("controle: sem a regra, o vazio depois do corpo fica", tv[iv + 1] == "",
             repr(tv[iv + 1][:28]), "''")
    i21 = textos_.index("2.1\tPrimeira parte")
    z.exigir("6. vazio encostado em titulo sai",
             textos_[i21 + 1] == "2.2\tSegunda parte", repr(textos_[i21 + 1]), "2.2<tab>Segunda parte")

    tpm = [x for _, x in res["titulos_por_marcar"]]
    z.exigir("8. aponta o que parece titulo e nao recebeu estilo",
             "METODOLOGIA DE ANALISE" in tpm
             and not any(x.startswith("Paragrafo de corpo") for x in tpm),
             tpm, "[METODOLOGIA DE ANALISE]")
    emp = [x for _, x, _v, _p in res["empurroes"]]
    z.exigir("9. aponta o espaco de 36 pt ou mais posto a mao",
             any(x.startswith("Paragrafo empurrado") for x in emp)
             and not any(x.startswith("Paragrafo de corpo") for x in emp),
             emp, "[Paragrafo empurrado...]")

    print()
    print("  O QUE O TESTE EM OUTROS TRABALHOS ENSINOU")
    ind_ps = [doc[a:b] for a, b in pars if O.texto(doc[a:b]).strip()
              in ("1 Primeira figura: detalhe", "2 Segunda figura")]
    z.exigir("1. entrada de indice automatico nao vira titulo nem e dividida",
             len(ind_ps) == 2 and all(O.estilo_de(x) == "ndicedeilustraes"
                                      and b"pageBreakBefore" not in x for x in ind_ps),
             [O.estilo_de(x) for x in ind_ps], "2 x ndicedeilustraes")
    z.exigir("1. ...e nao decide onde comeca a parte textual",
             res["inicio_textual"] > 0 and textos_[0] != "1 Primeira figura: detalhe"
             and papel.get("1\tINTRODUCAO") == "Ttulo1",
             papel.get("1\tINTRODUCAO"), "Ttulo1")

    # 2. a divisao, pedaco a pedaco
    def bl(x):
        return ("<w:p>" + x + "</w:p>").encode()

    def dividir(x, pos):
        # erro de programa conta como reprovacao, e nao derruba o controle
        try:
            return F.dividir_no_caractere(x, pos)
        except Exception as e:
            return ("erro: %s" % e,)
    r_marc = dividir(bl(
        '<w:bookmarkStart w:id="1" w:name="m"/><w:r><w:t>AB: CD</w:t></w:r>'
        '<w:bookmarkEnd w:id="1"/>'), 2)
    z.exigir("2. a divisao leva o marcador, cada ponta para o seu lado",
             bool(r_marc) and len(r_marc) == 2 and b"bookmarkStart" in r_marc[0]
             and b"bookmarkEnd" in r_marc[1])
    r_link = dividir(bl(
        '<w:hyperlink w:anchor="x"><w:r><w:t>AB: CD</w:t></w:r></w:hyperlink>'), 2)
    z.exigir("2. nao divide dentro de hiperlink", r_link is None,
             "dividiu" if r_link else "nao dividiu", "nao dividiu")
    r_campo = dividir(bl(
        '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> REF x </w:instrText></w:r>'
        '<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>AB: CD</w:t></w:r>'
        '<w:r><w:fldChar w:fldCharType="end"/></w:r>'), 2)
    z.exigir("2. nao divide dentro de campo", r_campo is None,
             "dividiu" if r_campo else "nao dividiu", "nao dividiu")
    z.exigir("2. o titulo dividido guarda o marcador do sumario",
             b'w:name="_Toc777"' in doc and doc.count(b'<w:bookmarkEnd w:id="7"/>') == 1)

    # 3. a conferencia da estrutura, e o controle positivo dela
    z.exigir("3. a conferencia conta a estrutura, e aqui nada se perdeu",
             not F.conferir(res, destino), "; ".join(F.conferir(res, destino))[:60] or "nada",
             "nada")
    res_sem = dict(res)
    res_sem["doc_novo"] = res["doc_novo"].replace(b'<w:bookmarkStart w:id="7" w:name="_Toc777"/>', b"", 1)
    acusou = F.conferir(res_sem, destino)
    z.exigir("3. ...e acusa o marcador que some",
             any("marcador" in x for x in acusou), "; ".join(acusou)[:60] or "nao acusou",
             "acusa marcador")

    # 4. capitulo marcado so por estilo
    mini = ('<w:body>'
            '<w:p><w:pPr><w:pStyle w:val="Ttulo1"/></w:pPr><w:r><w:t>RESUMO</w:t></w:r></w:p>'
            '<w:p><w:r><w:t>Texto do resumo.</w:t></w:r></w:p>'
            '<w:p><w:pPr><w:pStyle w:val="Ttulo1"/></w:pPr><w:r><w:t>INTRODUCAO</w:t></w:r></w:p>'
            '<w:p><w:r><w:t>Texto.</w:t></w:r></w:p>'
            '<w:p><w:pPr><w:pStyle w:val="Ttulo2"/></w:pPr><w:r><w:t>Metodologia</w:t></w:r></w:p>'
            '</w:body>').encode()
    try:
        fm, pm, im, _, _ = F.planejar(mini, 26, 312)
    except SystemExit:
        pm, im = {2: "parou", 4: "parou"}, None
    z.exigir("4. o capitulo sem numero, em estilo de titulo, abre a parte textual",
             im == 2 and pm[2] == "titulo1" and pm[4] == "titulo2",
             "inicio %s, papeis %s" % (im, [pm[k] for k in (2, 4)]), "inicio 2, titulo1 e titulo2")

    z.exigir("5. o titulo de terceiro nivel vira Titulo 3",
             papel.get("2.2.1\tTerceiro nivel") == "Ttulo3",
             papel.get("2.2.1\tTerceiro nivel"), "Ttulo3")
    t3 = defs.get("Ttulo3", {}).get("xml", b"")
    z.exigir("5. ...com a tabulacao depois de keepNext, que e a ordem do esquema",
             b"<w:tabs>" in t3 and t3.find(b"<w:keepNext/>") < t3.find(b"<w:tabs>"),
             "keepNext %d, tabs %d" % (t3.find(b"<w:keepNext/>"), t3.find(b"<w:tabs>")),
             "keepNext antes")

    ingles = [("Figure 1 - Um grafico", F.LEGENDA), ("REFERENCES", F.ESTRUTURAL_TEXTUAL),
              ("APPENDIX A - Data", F.ESTRUTURAL_SUB), ("ACKNOWLEDGEMENTS", F.TITULO_PRE),
              ("LIST OF FIGURES", F.TITULO_PRE), ("Source: the author.", F.FONTE_APOS_FIGURA)]
    falta = [x for x, r in ingles if not r.match(x)]
    z.exigir("6. o vocabulario em ingles e reconhecido", not falta, falta or "todos", "todos")
    z.exigir("6. ...e o rotulo do SEQ nao troca Graph por Grafico",
             F.rotulo_seq("Graph") == "Graph" and F.rotulo_seq("Grafico") == "Gráfico",
             (F.rotulo_seq("Graph"), F.rotulo_seq("Grafico")), ("Graph", "Gráfico"))

    z.exigir("7. a lista ja automatica fica como estava, com o campo",
             doc.count(b'TOC \\h \\z \\c "Figura"') == 1,
             doc.count(b'TOC \\h \\z \\c "Figura"'), 1)
    z.exigir("7. ...e a lista digitada continua a ser trocada",
             'TOC \\h \\z \\c "Gráfico"' in doc.decode("utf-8"))

    ano = next((doc[a:b] for a, b in pars if O.texto(doc[a:b]).strip() == "2026"), b"")
    z.exigir("a linha de bloco da capa perde a entrelinha escrita a mao",
             bool(ano) and b'w:line="360"' not in O.ppr(ano),
             str(O.prop(O.ppr(ano), b"w:spacing")), "sem line 360")
    t2 = next((doc[a:b] for a, b in pars if tt(doc[a:b]).strip() == "2	DESENVOLVIMENTO"), b"")
    z.exigir("o titulo perde o corpo escrito a mao, e segue o estilo",
             bool(t2) and not any(b"<w:sz " in r for r in O.corridas(t2)))

    print()
    print("  O RESTO")
    campos, sdt = O.tem_sumario_automatico(doc)
    z.exigir("entrou um campo de sumario automatico", bool(campos),
             "%d campo(s)" % len(campos), "pelo menos 1")
    z.exigir("o sumario digitado saiu",
             "DESENVOLVIMENTO" not in [O.texto(doc[a:b]).strip() for a, b in pars],
             "", "")

    print()
    print("  A NUMERACAO AUTOMATICA DOS TITULOS")
    for linha in numeracao_automatica(origem, tmp, pular=None):
        z.exigir(*linha)
    # Controle positivo: cada peca desfeita tem de reprovar a sua exigencia.
    for peca, nome in (("com_lista", "a lista de quatro niveis ligada aos estilos"),
                       ("com_numpr", "os titulos numerados levam a lista"),
                       ("titulo_no_cabecalho", "o cabecalho traz o numero do capitulo")):
        linhas = numeracao_automatica(origem, tmp, pular=peca)
        caiu = [l for l in linhas if l[0].startswith(nome) and not l[1]]
        z.exigir("controle: sem `%s`, reprova '%s'" % (peca, nome), bool(caiu))

    print()
    print("  A NORMA COMO ESTA: so o numero no cabecalho, titulo inteiro")
    padrao = tmp / "como-esta.docx"
    F.formatar(origem, padrao, corpo_pt=13.0, entrelinha=1.3, fonte="Times New Roman")
    for linha in como_esta(padrao):
        z.exigir(*linha)
    sem_estilo = sem_a_regra(
        origem, tmp, 'p = com_pstyle(tirar(O.ppr(bloco), b"w:jc"), "Cabealho")',
        'p = tirar(O.ppr(bloco), b"w:jc")', destino_de=True)
    z.exigir("controle: sem a regra, o numero de pagina fica fora do estilo Cabealho",
             not como_esta(sem_estilo)[1][1])
    ag = O.ler(padrao)
    agp = next((ag["word/document.xml"][a:b] for a, b in O.paragrafos(ag["word/document.xml"])
                if O.texto(ag["word/document.xml"][a:b]).strip().startswith("Agradeco")), b"")
    agd = O.estilos_definidos(ag["word/styles.xml"]).get("TextoPreTextual", {}).get("xml", b"")
    z.exigir("os agradecimentos com o recuo e o alinhamento do corpo, pelo estilo",
             O.estilo_de(agp) == "TextoPreTextual" and b"<w:ind" not in O.ppr(agp)
             and b'w:firstLine="709"' in agd and b'<w:jc w:val="both"/>' in agd,
             "%s, %s" % (O.estilo_de(agp), O.ppr(agp)[:60]), "TextoPreTextual, sem ind")
    sem_ag = sem_a_regra(origem, tmp, 'else "TextoPreTextual" if f["i"] in dos_agradecimentos',
                         'else "PreTextual" if False')
    agp2 = next((sem_ag[a:b] for a, b in O.paragrafos(sem_ag)
                 if O.texto(sem_ag[a:b]).strip().startswith("Agradeco")), b"")
    z.exigir("controle: sem a regra, os agradecimentos ficam fora do estilo",
             O.estilo_de(agp2) != "TextoPreTextual", O.estilo_de(agp2))
    ligado = (b'<w:settings><w:defaultTabStop w:val="708"/>'
              b'<w:autoHyphenation w:val="true"/></w:settings>')
    des, era = F.desligar_hifenizacao(ligado)
    z.exigir("a hifenizacao que o autor ligou e desligada, e o relatorio sabe",
             b"autoHyphenation" not in des and era, "%s, %s" % (des[-60:], era))
    des2, era2 = F.desligar_hifenizacao(b'<w:settings><w:defaultTabStop w:val="708"/></w:settings>')
    z.exigir("...e o arquivo sem ela fica como estava", not era2
             and des2 == b'<w:settings><w:defaultTabStop w:val="708"/></w:settings>')

    pm = O.ler(padrao)
    est_m = O.estilos_definidos(pm["word/styles.xml"])
    par_m = next((pm["word/document.xml"][a:b] for a, b in O.paragrafos(pm["word/document.xml"])
                  if O.texto(pm["word/document.xml"][a:b]).strip().startswith("Paragrafo colado")), b"")
    z.exigir("o paragrafo em msonormal passa ao Normal",
             bool(par_m) and O.estilo_de(par_m) in (None, "Normal"), O.estilo_de(par_m), "Normal")
    z.exigir("...o estilo sai do arquivo",
             not any("msonormal" in sid.lower() for sid in est_m), [x for x in est_m if "mso" in x.lower()])
    z.exigir("...e o estilo que se baseava nele passa a se basear no Normal",
             b'<w:basedOn w:val="Normal"/>' in est_m.get("Derivado", {}).get("xml", b""),
             O.atributo(est_m.get("Derivado", {}).get("xml", b""), b"<w:basedOn", b"w:val"), "Normal")
    e_mso = sem_a_regra(origem, tmp, "    doc, est, n_mso, ids_mso = sem_msonormal(doc, est)",
                        "    n_mso, ids_mso = 0, []", estilos=True)
    z.exigir("controle: sem a regra, o msonormal fica",
             "msonormal0" in O.estilos_definidos(e_mso))

    print()
    print("  O NUMERO DE PAGINA QUE FALTA")
    # 1. a secao final sem cabecalho: herdaria o vazio do pre-textual
    textual, pre_vazio, novo = numero_na_parte_textual(padrao)
    z.exigir("secao sem cabecalho: ganha um cabecalho novo com o numero",
             textual and novo, "numero=%s, parte nova=%s" % (textual, novo))
    z.exigir("...e o pre-textual continua sem numero", pre_vazio)
    # 2. cabecalho proprio, vazio: o numero entra nele, sem parte nova
    v2 = variante_de_numero(origem, tmp, "vazio", b'<w:headerReference w:type="default" r:id="rId3"/>',
                            cabecalho=b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                            b'<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/'
                            b'2006/main"><w:p/><w:p/></w:hdr>')
    F.formatar(v2, tmp / "numero-vazio-f.docx", corpo_pt=13.0, entrelinha=1.3,
               fonte="Times New Roman")
    t2, p2, n2 = numero_na_parte_textual(tmp / "numero-vazio-f.docx")
    with zipfile.ZipFile(tmp / "numero-vazio-f.docx") as zz:
        h2 = zz.read("word/header1.xml")
    z.exigir("cabecalho proprio sem numero: o numero entra nele",
             t2 and p2 and not n2 and h2.count(b"<w:p>") + h2.count(b"<w:p ") == 1,
             "numero=%s, parte nova=%s, paragrafos=%d" % (t2, n2, h2.count(b"<w:p>")))
    # 3. numero no rodape: nao falta, e nada entra
    v3 = variante_de_numero(origem, tmp, "rodape", b'<w:footerReference w:type="default" r:id="rId9"/>',
                            rodape=HEADER.format(ns=NS).replace("w:hdr", "w:ftr").encode())
    F.formatar(v3, tmp / "numero-rodape-f.docx", corpo_pt=13.0, entrelinha=1.3,
               fonte="Times New Roman")
    t3, _, n3 = numero_na_parte_textual(tmp / "numero-rodape-f.docx")
    z.exigir("numero no rodape: nao falta, e nenhum cabecalho novo entra",
             t3 and not n3, "numero=%s, parte nova=%s" % (t3, n3))
    # 4. o cabecalho ja tem o numero: nao ganha um segundo
    v4 = variante_de_numero(origem, tmp, "com-numero",
                            b'<w:headerReference w:type="default" r:id="rId3"/>')
    F.formatar(v4, tmp / "numero-com-f.docx", corpo_pt=13.0, entrelinha=1.3,
               fonte="Times New Roman")
    with zipfile.ZipFile(tmp / "numero-com-f.docx") as zz:
        h4 = zz.read("word/header1.xml")
        n4 = "word/header-numero.xml" in zz.namelist()
    z.exigir("cabecalho que ja tem o numero: nao ganha um segundo",
             h4.count(b"PAGE") == 1 and not n4,
             "PAGE %d vez(es), parte nova=%s" % (h4.count(b"PAGE"), n4), "1, False")
    # controle: sem a regra, a secao do caso 1 fica sem numero
    sem_num = sem_a_regra(origem, tmp, '"inicio": 0, "inserir_se_faltar": True',
                          '"inicio": 0, "inserir_se_faltar": False', destino_de=True)
    z.exigir("controle: sem a regra, a parte textual fica sem numero",
             not numero_na_parte_textual(sem_num)[0])

    # Controle: o arquivo da primeira metade, com as opcoes ligadas, reprova.
    for nome, ok, *_ in como_esta(destino):
        if nome.startswith("..."):
            continue
        z.exigir("controle: com a opcao ligada, reprova '%s'" % nome, not ok)

    ligada, manteve = com_lista_do_word(origem, tmp)
    z.exigir("com um titulo numerado pelo Word, a numeracao do autor fica",
             not ligada and manteve, "automatica=%s, lista mantida=%s" % (ligada, manteve),
             "automatica=False, lista mantida=True")
    ligada2, manteve2 = com_lista_do_word(origem, tmp, desfazer=True)
    z.exigir("controle: sem essa regra, a lista do autor se perde", ligada2 and not manteve2,
             "automatica=%s, lista mantida=%s" % (ligada2, manteve2))

    print()
    print("  SEGUNDA METADE: a conferencia reprova o que tem de reprovar")
    # 1: um paragrafo a menos
    quebrado = tmp / "quebrado-paragrafo.docx"
    shutil.copyfile(destino, quebrado)
    d2 = O.ler(destino)["word/document.xml"]
    a, b = O.paragrafos(d2)[-1]
    res_falso = dict(res)
    res_falso["doc_novo"] = d2[:a] + d2[b:]
    z.exigir("acusa quando um paragrafo desaparece",
             bool(F.conferir(res_falso, destino)),
             "; ".join(F.conferir(res_falso, destino))[:40] or "nao acusou")

    # 2: uma palavra trocada
    res_falso2 = dict(res)
    res_falso2["doc_novo"] = d2.replace(b"<w:t xml:space=\"preserve\">REFERENCIAS<",
                                        b"<w:t xml:space=\"preserve\">REFERENCIA<", 1)
    z.exigir("acusa quando uma palavra muda",
             bool(F.conferir(res_falso2, destino)),
             "; ".join(F.conferir(res_falso2, destino))[:40] or "nao acusou")

    # 3: XML quebrado
    ruim = tmp / "quebrado-xml.docx"
    shutil.copyfile(destino, ruim)
    with zipfile.ZipFile(ruim) as zz:
        itens = [(i, zz.read(i.filename)) for i in zz.infolist()]
    with zipfile.ZipFile(ruim, "w", zipfile.ZIP_DEFLATED) as zz:
        for info, dados in itens:
            if info.filename == "word/document.xml":
                dados = dados.replace(b"</w:body>", b"", 1)
            zz.writestr(info, dados)
    z.exigir("acusa quando o XML deixa de fechar",
             any("bem formado" in f for f in F.conferir(res, ruim)),
             "; ".join(F.conferir(res, ruim))[:44] or "nao acusou")

    print()
    if z.falhas:
        print("  %d exigencia(s) REPROVADA(S). O formatador nao esta certificado,"
              % z.falhas)
        print("  e nada que ele produzir sobre um trabalho real vale enquanto isso.")
        sys.exit(1)
    print("  todas passaram: o programa faz o que diz, e a conferencia dele")
    print("  reprova arquivo quebrado de proposito.")
    print()


if __name__ == "__main__":
    main()

# -*- coding: utf-8 -*-
"""Leitura de OOXML. Nenhuma funcao daqui escreve no documento.

POR QUE ESTE MODULO EXISTE SEPARADO

A Oficina de Orientacao le e relata; esta escreve no arquivo de quem escreveu o
trabalho. Enquanto a leitura morar dentro do transformador, a fronteira entre as
duas e a promessa escrita no cabecalho, e nao a estrutura. Aqui mora so a
leitura, e por isso importar este modulo nao traz nenhuma funcao que altere
arquivo.

A VARREDURA DE PARAGRAFO E O UNICO PONTO DELICADO

`<w:p>.*?</w:p>` esta errado e o erro e silencioso: caixa de texto guarda
paragrafo dentro de paragrafo, e a busca preguicosa fecha no `</w:p>` de dentro.
Apagar o que ela devolve deixa um `</w:p>` orfao e o arquivo para de abrir, e
o defeito volta facil em codigo novo. `blocos()` conta profundidade e nao tem
esse defeito; e ela que todo o resto usa.
"""
import re
import zipfile
from collections import Counter

TWIPS_POR_CM = 566.9291338582677      # 1 cm = 567 twips
MEIO_PONTO = 2                        # w:sz vem em meio-ponto


def blocos(doc, tag):
    """Fatias (inicio, fim) de cada <tag> de primeiro nivel, contando profundidade.

    Serve para `w:p`, `w:tbl`, `w:tc` e qualquer outro elemento que possa
    aninhar-se dentro de si mesmo.
    """
    tag = tag if isinstance(tag, bytes) else tag.encode()
    padrao = re.compile(rb"<" + tag + rb"(?:\s[^>]*)?/>"
                        rb"|<" + tag + rb"(?:\s[^>]*)?>"
                        rb"|</" + tag + rb">")
    fatias, profundidade, inicio = [], 0, None
    for m in padrao.finditer(doc):
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


def paragrafos(doc):
    return blocos(doc, b"w:p")


def tabelas(doc):
    return blocos(doc, b"w:tbl")


def dentro_de_tabela(faixas_de_tabela):
    """Devolve uma funcao que diz se uma posicao cai dentro de alguma tabela."""
    faixas = sorted(faixas_de_tabela)

    def esta(pos):
        for a, b in faixas:
            if a <= pos < b:
                return True
            if pos < a:
                return False
        return False
    return esta


def ppr(bloco):
    """O pPr DO PROPRIO paragrafo, e nao o do que estiver dentro dele.

    Num paragrafo com caixa de texto, o primeiro `<w:pPr>` do trecho e o do
    paragrafo de dentro. O pPr do proprio vem antes de qualquer `<w:r>` e antes
    de qualquer `<w:p>` aninhado.
    """
    corte = len(bloco)
    for m in re.finditer(rb"<w:r(?:\s[^>]*)?>|<w:p(?:\s[^>]*)?>", bloco[1:]):
        corte = m.start() + 1
        break
    m = re.search(rb"<w:pPr>(.*?)</w:pPr>", bloco[:corte], re.S)
    return m.group(1) if m else b""


def texto(bloco):
    """O texto do paragrafo, incluindo o que esteja em caixa de texto dentro dele."""
    return "".join(t.decode("utf-8", "replace") for t in
                   re.findall(rb"<w:t(?:\s[^>]*)?>(.*?)</w:t>", bloco, re.S))


def estilo_de(bloco):
    m = re.search(rb'<w:pStyle w:val="([^"]+)"', ppr(bloco))
    return m.group(1).decode() if m else None


def prop(pprop, tag):
    """A tag inteira, como esta escrita, ou None. `tag` sem os sinais."""
    tag = tag if isinstance(tag, bytes) else tag.encode()
    m = re.search(rb"<" + tag + rb"(?:\s[^>]*)?/>", pprop)
    return m.group(0) if m else None


def atributo(xml, tag, attr):
    """O valor do atributo, com o espaco depois do nome da tag.

    O espaco nao e decoracao: sem ele, `<w:sz` casa tambem `<w:szCs`, e a
    contagem de corpo de letra soma o tamanho de escrita complexa, que nao tem
    efeito no texto latino, e a conta sai varias vezes maior.
    """
    tag = tag if isinstance(tag, bytes) else tag.encode()
    attr = attr if isinstance(attr, bytes) else attr.encode()
    m = re.search(tag + rb'(?:\s[^>]*)?\s' + attr + rb'="([^"]*)"', xml)
    return m.group(1).decode() if m else None


def corridas(bloco):
    """Os rPr das corridas que tem texto, na ordem."""
    saida = []
    for m in re.finditer(rb"<w:r(?:\s[^>]*)?>(.*?)</w:r>", bloco, re.S):
        r = m.group(1)
        if not re.search(rb"<w:t(?:\s[^>]*)?>", r):
            continue
        pr = re.search(rb"<w:rPr>(.*?)</w:rPr>", r, re.S)
        saida.append(pr.group(1) if pr else b"")
    return saida


def negrito_inteiro(bloco):
    rs = corridas(bloco)
    return bool(rs) and all(re.search(rb"<w:b(?:\s[^>]*)?/>", r) for r in rs)


def corpo_de_letra(bloco):
    """O tamanho mais frequente declarado nas corridas, em meio-pontos, ou None."""
    tam = [int(m.group(1)) for m in re.finditer(rb'<w:sz\s[^>]*w:val="(\d+)"', bloco)]
    return Counter(tam).most_common(1)[0][0] if tam else None


def forma(bloco):
    """A assinatura de forma do paragrafo: (estilo, jc, ind, spacing)."""
    p = ppr(bloco)
    return (estilo_de(bloco),
            prop(p, b"w:jc"), prop(p, b"w:ind"), prop(p, b"w:spacing"))


def vazio(bloco):
    """Sem texto nenhum. Paragrafo suprimido com controle de alteracoes conta
    como vazio para `w:t`, e por isso tambem se olha `w:delText`."""
    return not texto(bloco).strip() and not re.search(rb"<w:delText", bloco)


QUEBRA = (rb'<w:br[^>]*w:type="(?:page|column)"'
          rb'|<w:pageBreakBefore(?:\s+w:val="(?:1|true|on)")?\s*/>')


def travado(bloco, aceitar_quebra=False):
    """Carrega coisa que se perde se o paragrafo for apagado.

    A lista reune as travas deste programa e as da Norma antiga, que tinha a
    quebra de pagina antes, a quebra de coluna, o fim de marcador, a referencia
    de comentario e o texto apagado sob revisao, que aqui faltavam.

    `aceitar_quebra` e para o vazio que fica logo antes de um titulo que ja
    quebra pagina: ali a quebra do vazio repete a do titulo, e mante-lo abre uma
    folha em branco com a linha vazia no alto.
    """
    if not aceitar_quebra and re.search(QUEBRA, bloco):
        return True
    return bool(re.search(
        rb"<w:sectPr|<w:drawing|<w:pict|<w:object"
        rb"|<w:bookmarkStart|<w:bookmarkEnd|<w:commentRangeStart|<w:commentRangeEnd"
        rb"|<w:commentReference|<w:fldChar|<w:instrText|<w:footnoteReference"
        rb"|<w:endnoteReference|<w:ins\b|<w:del\b|<w:delText", bloco))


NUMERO_DE_TITULO = re.compile(r"^(\d+(?:\.\d+)*)\.?\s+\S")


def numero_de_titulo(t):
    """('2.1', 2) para '2.1 Origens...', ou (None, 0)."""
    m = NUMERO_DE_TITULO.match(t.strip())
    if not m:
        return None, 0
    n = m.group(1)
    return n, n.count(".") + 1


def ler(caminho):
    """Devolve {nome: bytes} das partes que interessam."""
    with zipfile.ZipFile(caminho) as z:
        nomes = set(z.namelist())
        partes = {}
        for n in ("word/document.xml", "word/styles.xml", "word/settings.xml",
                  "word/numbering.xml"):
            if n in nomes:
                partes[n] = z.read(n)
    return partes


def estilos_definidos(styles_xml):
    """{styleId: {'nome':..., 'tipo':..., 'xml': bytes}}"""
    saida = {}
    for m in re.finditer(rb"<w:style\s[^>]*>.*?</w:style>", styles_xml, re.S):
        bloco = m.group(0)
        sid = atributo(bloco[:300], b"<w:style", b"w:styleId")
        if not sid:
            continue
        saida[sid] = {
            "nome": atributo(bloco, b"<w:name", b"w:val"),
            "tipo": atributo(bloco[:300], b"<w:style", b"w:type"),
            "xml": bloco,
        }
    return saida


def tem_sumario_automatico(document_xml):
    """(campos TOC, blocos de conteudo de sumario)."""
    instr = [m.group(1).decode("utf-8", "replace") for m in
             re.finditer(rb"<w:instrText[^>]*>(.*?)</w:instrText>",
                         document_xml, re.S)]
    campos = [i for i in instr if re.search(r"\bTOC\b", i, re.I)]
    sdt = len(re.findall(rb'<w:docPartGallery w:val="Table of Contents"',
                         document_xml))
    return campos, sdt


def secoes(document_xml):
    """Os sectPr do documento, com tamanho de pagina e margem em cm."""
    saida = []
    for m in re.finditer(rb"<w:sectPr\b.*?</w:sectPr>", document_xml, re.S):
        s = m.group(0)
        d = {}
        for attr, chave in (("w:w", "larg"), ("w:h", "alt")):
            v = atributo(s, b"<w:pgSz", attr.encode())
            if v:
                d[chave] = int(v) / TWIPS_POR_CM
        for attr in ("w:top", "w:right", "w:bottom", "w:left"):
            v = atributo(s, b"<w:pgMar", attr.encode())
            if v:
                d[attr.split(":")[1]] = int(v) / TWIPS_POR_CM
        saida.append(d)
    return saida

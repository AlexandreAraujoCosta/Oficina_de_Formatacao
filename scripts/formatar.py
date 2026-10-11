# -*- coding: utf-8 -*-
"""Escreve uma copia formatada do .docx, na norma do PMPD. Nao toca na entrada.

A norma esta em `NORMA`, logo abaixo, e descrita por extenso em NORMA-PMPD.md.

O QUE ELE FAZ, E A ORDEM IMPORTA

  1 blinda quem herda do Normal, ANTES de mexer no Normal;
  2 escreve no Normal a forma do corpo do texto;
  3 reescreve os estilos de titulo, legenda, citacao, referencia, sumario;
  4 da a cada paragrafo o estilo do papel que ele cumpre, e tira a forma escrita
    a mao que o estilo passa a dar;
  5 diagrama a capa e o pre-textual, e poe o pre-textual numa secao contada e
    nao numerada;
  6 troca o sumario e a lista de ilustracoes digitados por campos do Word.

POR QUE O PASSO 1 VEM PRIMEIRO

Escrever a forma do corpo dentro do `Normal` alcanca tudo o que herda dele: a
capa, as tabelas, as notas, o cabecalho. Quem herda sem se defender muda junto,
e a capa se desmonta sem que o texto dela mude. A defesa e um estilo por papel,
e nao formatacao direta em cada paragrafo: a forma mora no estilo, onde se muda
uma vez.

O QUE ELE NAO FAZ

Nao muda texto, a nao ser o que a norma manda e o relatorio lista: o numero do
titulo, o titulo dividido no dois-pontos, as palavras-chave, o titulo da
dedicatoria e da epigrafe. Nao reescreve prosa, nem as remissoes do texto a
capitulos. Nao aceita nem recusa alteracao controlada de ninguem: arquivo com
alteracao pendente e recusado.

    python scripts/formatar.py trabalho.docx --saida novo.docx
    python scripts/formatar.py trabalho.docx --saida novo.docx --corpo 13 --entrelinha 1.3
"""
import argparse
import re
import shutil
import sys
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import ooxml as O                                                    # noqa: E402

# O separador e o que a NBR pede, mas ha quem o esqueca: `Grafico 1 Evolucao`
# tambem e titulo de ilustracao, e o que o distingue do texto corrido e a
# maiuscula logo depois do numero.
# Em portugues e em ingles: o programa recebe tese escrita em ingles.
LEGENDA = re.compile(r"^(Gr[áa]fico|Tabela|Figura|Quadro|Imagem"
                     r"|Figure|Chart|Graph|Table|Picture)\s*\d+"
                     r"(\s*[–\-—.:]|\s+(?=[A-ZÁÉÍÓÚÂÊÔÃÕÇ]))")
ROTULO_DE_ILUSTRACAO = re.compile(r"^(Gr[áa]fico|Figura|Quadro|Imagem"
                                  r"|Figure|Chart|Graph|Picture)\s*(\d+)")
ESTRUTURAL_TEXTUAL = re.compile(
    r"^(REFER[ÊE]NCIAS?|BIBLIOGRAFIA|AP[ÊE]NDICES?|ANEXOS?|DECLARA[ÇC][ÃA]O"
    r"|REFERENCES|BIBLIOGRAPHY|APPENDI(?:X|CES)|ANNEX(?:ES)?|DECLARATION)\b", re.I)
ESTRUTURAL_SUB = re.compile(r"^(AP[ÊE]NDICE|ANEXO|APPENDIX|ANNEX)\s+[A-Z0-9]", re.I)
# Os titulos do pre-textual, num lugar so: antes havia tres copias desta lista,
# que ja divergiam.
TITULO_PRE = re.compile(
    r"^(RESUMO|ABSTRACT|RESUMEN|SUM[ÁA]RIO|CONTENTS|TABLE OF CONTENTS"
    r"|LISTA DE|LIST OF|[ÍI]NDICE DE|AGRADECIMENTOS?|ACKNOWLEDG(?:E)?MENTS?"
    r"|DEDICAT[ÓO]RIA|DEDICATION|EP[ÍI]GRAFE|EPIGRAPH)", re.I)
LISTA_DE = re.compile(r"^(LISTA DE|LIST OF|[ÍI]NDICE DE) ", re.I)


# O que o paragrafo carrega alem do texto, e que a conferencia de texto nao ve.
ESTRUTURA = {
    "marcador": rb"<w:bookmarkStart\b",
    "fim de marcador": rb"<w:bookmarkEnd\b",
    "hiperlink": rb"<w:hyperlink\b",
    "chamada de nota": rb"<w:footnoteReference\b|<w:endnoteReference\b",
    "comentario": rb"<w:commentReference\b|<w:commentRangeStart\b",
}


def contar_estrutura(xml):
    return {k: len(re.findall(v, xml)) for k, v in ESTRUTURA.items()}


def rotulo_seq(palavra):
    """O identificador do campo SEQ: `Gráfico` para as grafias do portugues, e a
    palavra como veio para o resto. Antes, `Graph` virava `Gráfico` porque o
    teste era o comeco `gr`."""
    return "Gráfico" if re.match(r"(?i)gr[áa]fico$", palavra) else palavra.capitalize()


FONTE_APOS_FIGURA = re.compile(
    r"^(Fonte|Fontes|Nota|Notas|Elabora[çc][ãa]o|Fuente|Source|Sources|Note|Notes)\s*:",
    re.I)

CM = 566.929                # twips por centimetro
PT = 20                     # twips por ponto


def tw(cm):
    return int(round(cm * CM))


def pt(p):
    return int(round(p * PT))


# A NORMA. Um lugar so, e e daqui que sai tudo o que o programa escreve.
# Cada linha e uma decisao de quem coordena o programa, e nao do programa.
NORMA = {
    "fonte": "Times New Roman",
    "corpo_pt": 13.0,
    "entrelinha": 1.3,
    "normal_depois_pt": 3,
    "recuo_primeira_cm": 1.25,
    # A hifenizacao e a ultima providencia, so na versao final de impressao
    # (`--impressao`), por decisao da coordenacao: no arquivo de trabalho ela
    # fica desligada, mesmo que o autor a tenha ligado.
    "hifenizar": False,
    "hifens_seguidos": 2,       # linhas seguidas que podem terminar em hifen
    "margens_cm": {"esq": 3.3, "dir": 3.3, "sup": 3.3, "inf": 2.5},
    # Antes 0: a NBR 14724 (5.2.2) quer o titulo da secao primaria no alto da
    # mancha.
    # A esquerda, como o Titulo 2 (NBR 6024, 4.1, alinea c). O que nao tem
    # numero, como REFERENCIAS e APENDICES, fica centralizado (NBR 14724,
    # 5.2.3; NBR 6024, 4.1, alinea h).
    "titulo1": {"corpo_pt": 16, "antes_pt": 0, "depois_pt": 36,
                "centralizado": False, "quebra_pagina": True},
    # Todos os titulos em entrelinha simples, por decisao da coordenacao.
    "titulos_entrelinha": 1.0,
    "titulo2": {"antes_pt": 18, "depois_pt": 12},
    "numero_de_titulo": {
        # A NBR 6024:2012 (4.1, alinea d) nao admite sinal entre o indicativo e
        # o titulo.
        "ponto_depois": False,
        # Entre o numero e o titulo vai uma tabulacao, e nao um espaco. E ela que
        # alinha a segunda linha do titulo sob a primeira letra dele (NBR 6024,
        # 4.1, alinea i) e, no sumario, os titulos numa coluna so (NBR 6027,
        # 5.2). Cada estilo numerado declara a sua parada: sem ela, o Word usa as
        # paradas padrao, o titulo centralizado ganha um vao largo, e no
        # cabecalho corrente o titulo e empurrado para a parada do numero de
        # pagina.
        "tabulacao": True,
        "parada_t1_cm": 0.5,            # o vao do numero no cabecalho corrente
        # A coluna do titulo, em todos os niveis, e a do recuo de primeira linha
        # do corpo: o numero fica na margem e o titulo comeca onde comeca o
        # paragrafo recuado, por decisao da coordenacao.
        "pendente_cm": {1: 1.25, 2: 1.25, 3: 1.25, 4: 1.25},
        # O numero digitado sai e o Word passa a numerar, por uma lista ligada
        # aos estilos Titulo 1 a 4: acrescentar ou tirar um titulo renumera o
        # resto sozinho. A numeracao progressiva vai ate o Titulo 3; o Titulo 4
        # leva so uma letra, "a)", que recomeca a cada Titulo 3, por decisao da
        # coordenacao (a forma e a da alinea, NBR 6024, 4.2). Trabalho que ja numera pelo Word fica com a
        # numeracao dele.
        "automatica": True,
    },
    "citacao": {"menor_pt": 1, "recuo_esq_cm": 2.5,
                "antes_pt": 12, "depois_pt": 12},
    # A legenda usa o estilo de legenda do proprio Word, o mesmo que "Inserir
    # legenda" aplica, e nao um estilo criado pelo programa. O titulo da
    # ilustracao leva o campo SEQ, e e por ele, e nao pelo estilo, que o Word
    # monta a lista de ilustracoes: a fonte pode dividir o estilo sem entrar
    # na lista.
    # O bloco da ilustracao (titulo, figura, fonte) se separa do texto por 12 pt.
    # O titulo leva 12 pt depois, por decisao da coordenacao, escritos no
    # proprio paragrafo, porque o estilo e um so para o titulo e para a fonte.
    # Legenda em italico e 11 pt, por decisao da coordenacao; o titulo e a fonte
    # da ilustracao dividem o estilo, e os dois saem assim.
    "legenda": {"corpo_pt": 11, "italico": True, "antes_pt": 12, "depois_pt": 12,
                "titulo_depois_pt": 12},
    # A NBR 14724 (5.1) quer citacao, notas, legendas, fontes e paginacao em
    # tamanho menor e uniforme. O PMPD fixou 12 para citacao e notas, 11 para a
    # legenda e 10 para o cabecalho.
    "notas": {"corpo_pt": 12},
    "referencia": {"entrelinha": 1.0, "depois_pt": 12},
    # A lista de ilustracoes sai por campo, com numero de pagina. O recuo
    # pendente e a largura de `Grafico 5 – ` em Times 13, medida pela metrica
    # da fonte: a continuacao de uma entrada longa alinha sob o titulo.
    "lista_de_ilustracoes": {"automatica": True, "pendente_cm": 2.2,
                             "sem_tabelas": True},
    # NBR 6028:2021 (4.1.7): separadas por ponto e virgula, terminadas por ponto,
    # iniciais minusculas salvo nome proprio.
    "palavras_chave_abnt": True,
    # A contagem comeca na folha de rosto (NBR 14724, 5.3): a capa e a folha 0.
    # O trabalho sem numero de pagina na parte textual ganha o numero no
    # cabecalho (NBR 14724, 5.3).
    "numeracao": {"mostra_do_textual": True, "inicio": 0, "inserir_se_faltar": True},
    "cabecalho": {
        # O titulo de capitulo e curto, e o que explica desce para o subtitulo.
        # Com isso o cabecalho corrente le o titulo VIVO, pelo campo STYLEREF, e
        # nao uma copia que envelhece quando o autor edita.
        # Desligados por decisao da coordenacao: o cabecalho traz so o numero
        # de pagina, e o titulo longo e o autor que encurta. O codigo fica, e o
        # controle o exercita ligado.
        "titulo_no_cabecalho": False,
        "dividir_no_dois_pontos": False,
        # Titulo corrente e numero de pagina em 10 pt, por decisao da
        # coordenacao. O titulo divide a linha com o numero, e num corpo maior
        # um titulo longo dobra para duas linhas.
        "corpo_pt": 10,
        # Em italico, titulo e numero, por decisao da coordenacao.
        "italico": True,
        # A linha do numero fica no meio da margem superior, por decisao da
        # coordenacao (a NBR 14724, 5.3, a poe a 2 cm da borda). A distancia da
        # borda direita continua a da margem, para que o numero e o titulo
        # corrente fiquem na coluna do texto.
        "no_meio_da_margem": True,
    },
    "pretextual": {
        "segue_o_normal": True,      # o corpo e o do texto, e nao o antigo
        "entrelinha": 1.0,           # a entrelinha, porem, e simples
        "capa_por_espacamento": True,
        "capa_fracoes": (0.25, 0.45),   # onde comecam os blocos do meio
        "capa_folga_pe_cm": 0.5,        # o ultimo bloco termina acima do pe
        "resumo_entrelinha": 1.0,
        # A NBR 14724 poe a folha de aprovacao em folha propria, depois da folha
        # de rosto, e nao quer titulo na dedicatoria.
        "aprovacao_em_folha_propria": True,
        "dedicatoria_sem_titulo": True,
        "recuo_da_lista_cm": 1.2,
        # A dedicatoria fica no pe da folha, com o texto comecando no meio da
        # largura da mancha. O pe vem do alinhamento vertical de uma secao
        # propria, e nao de espaco calculado: o autor pode reescrever a
        # dedicatoria, e ela continua no pe.
        "dedicatoria_na_base": True,
        "dedicatoria_recuo_fracao": 0.5,    # da largura da mancha
        # A epigrafe tem o tratamento da dedicatoria (NBR 14724, 5.2.4), e a
        # natureza do trabalho, na folha de rosto e na de aprovacao, o mesmo
        # recuo (5.2): um so eixo vertical para o pre-textual.
        "natureza_no_meio": True,
    },
    "sumario": {
        # O recuo pendente e o que faz a segunda linha de uma entrada longa
        # alinhar sob a primeira, em vez de voltar a margem.
        "nivel1": {"negrito": True, "antes_pt": 6, "depois_pt": 0,
                   "recuo_cm": 0.0, "pendente_cm": 0.9},
        # Os numeros de todos os niveis na margem esquerda, e os titulos numa
        # coluna so, pela tabulacao (NBR 6027, 5.1 e 5.2).
        "nivel2": {"negrito": False, "antes_pt": 0, "depois_pt": 0,
                   "recuo_cm": 0.0, "pendente_cm": 0.9},
    },
    "limiar_recuo_tw": 1000,    # acima disto, paragrafo recuado e citacao
}


# ----------------------------------------------------------------- estilos
def bloco_do_estilo(styles_xml, sid):
    m = re.search(rb'<w:style\s[^>]*w:styleId="' + re.escape(sid.encode())
                  + rb'"[^>]*>.*?</w:style>', styles_xml, re.S)
    return m


def por_ppr_rpr(bloco, ppr, rpr):
    """Escreve pPr e rPr num bloco de estilo, na ordem que o esquema exige.

    pPr vem antes de rPr, e os dois vem depois de name, basedOn, qFormat e
    companhia. Inserir pPr imediatamente antes de rPr, e rPr antes do
    fechamento, respeita a ordem sem ter de conhecer tudo o que vem antes.
    """
    if ppr is not None:
        if re.search(rb"<w:pPr>.*?</w:pPr>", bloco, re.S):
            bloco = re.sub(rb"<w:pPr>.*?</w:pPr>", ppr, bloco, count=1, flags=re.S)
        elif re.search(rb"<w:rPr>", bloco):
            bloco = re.sub(rb"<w:rPr>", ppr + b"<w:rPr>", bloco, count=1)
        else:
            bloco = bloco.replace(b"</w:style>", ppr + b"</w:style>")
    if rpr is not None:
        if re.search(rb"<w:rPr>.*?</w:rPr>", bloco, re.S):
            bloco = re.sub(rb"<w:rPr>.*?</w:rPr>", rpr, bloco, count=1, flags=re.S)
        else:
            bloco = bloco.replace(b"</w:style>", rpr + b"</w:style>")
    return bloco


def reescrever_estilo(styles_xml, sid, ppr=None, rpr=None):
    m = bloco_do_estilo(styles_xml, sid)
    if not m:
        return styles_xml, False
    novo = por_ppr_rpr(m.group(0), ppr, rpr)
    return styles_xml[:m.start()] + novo + styles_xml[m.end():], True


def criar_estilo(styles_xml, sid, nome, ppr, rpr, base=b"Normal",
                 embutido=False):
    """Cria o estilo. `embutido` e para os estilos do proprio Word, como
    `caption` e `table of figures`: esses nao podem levar `customStyle`.
    Marcado como personalizado, o Word nao o reconhece como o dele, renomeia-o e
    cria outro, e o que ele gera sai com a forma do outro."""
    if bloco_do_estilo(styles_xml, sid):
        return styles_xml, False
    bloco = (b'<w:style w:type="paragraph"'
             + (b"" if embutido else b' w:customStyle="1"')
             + b' w:styleId="'
             + sid.encode() + b'"><w:name w:val="' + nome.encode() + b'"/>'
             b'<w:basedOn w:val="' + base + b'"/><w:qFormat/>'
             + (ppr or b"") + (rpr or b"") + b"</w:style>")
    if b"</w:styles>" not in styles_xml:
        return styles_xml, False
    return styles_xml.replace(b"</w:styles>", bloco + b"</w:styles>", 1), True


def blindar_estilo(styles_xml, sid, espacamento, corpo, recuo=None):
    """Escreve num estilo a entrelinha e o corpo que ele hoje herda do Normal.

    Um estilo baseado no Normal so se defende no que declara. `Recuodecorpodetexto`
    declara recuo e alinhamento e nao declara corpo nem entrelinha: mudado o
    Normal, os paragrafos dele mudam de tamanho e de altura de linha sem que
    ninguem os tenha tocado. Isto fecha essa porta, e so onde ela esta aberta.
    """
    m = bloco_do_estilo(styles_xml, sid)
    if not m:
        return styles_xml, False
    bloco, mexeu = m.group(0), False

    pp = re.search(rb"<w:pPr>(.*?)</w:pPr>", bloco, re.S)
    if not (pp and re.search(rb"<w:spacing(?:\s[^>]*)?/>", pp.group(1))):
        if pp:
            bloco = (bloco[:pp.start()] + b"<w:pPr>" + espacamento + pp.group(1)
                     + b"</w:pPr>" + bloco[pp.end():])
        else:
            bloco = por_ppr_rpr(bloco, b"<w:pPr>" + espacamento + b"</w:pPr>", None)
        mexeu = True

    # O recuo de primeira linha tambem passa a vir do Normal. O estilo que ja
    # declara um w:ind com outra coisa dentro, um w:left por exemplo, precisa
    # que o w:firstLine entre NAQUELE mesmo elemento: dois w:ind no mesmo pPr
    # sao invalidos, e o Word fica com um deles sem dizer qual.
    #
    # Estilo que declara w:hanging fica como esta. O w:hanging do paragrafo
    # vence o w:firstLine do estilo, mas dentro do MESMO
    # w:ind os dois brigam, e o esquema manda ignorar o w:firstLine.
    if recuo is not None:
        pp = re.search(rb"<w:pPr>(.*?)</w:pPr>", bloco, re.S)
        dentro = pp.group(1) if pp else b""
        ind = re.search(rb"<w:ind(?:\s[^>]*)?/>", dentro)
        if ind is None:
            if pp:
                bloco = (bloco[:pp.start()] + b"<w:pPr>" + dentro + recuo
                         + b"</w:pPr>" + bloco[pp.end():])
            else:
                bloco = por_ppr_rpr(bloco, b"<w:pPr>" + recuo + b"</w:pPr>", None)
            mexeu = True
        elif b"w:firstLine=" not in ind.group(0) and b"w:hanging=" not in ind.group(0):
            novo_ind = ind.group(0)[:-2] + b' w:firstLine="0"/>'
            dentro2 = dentro[:ind.start()] + novo_ind + dentro[ind.end():]
            bloco = (bloco[:pp.start()] + b"<w:pPr>" + dentro2 + b"</w:pPr>"
                     + bloco[pp.end():])
            mexeu = True

    rr = re.search(rb"<w:rPr>(.*?)</w:rPr>", bloco, re.S)
    if not (rr and re.search(rb"<w:sz(?:\s[^>]*)?/>", rr.group(1))):
        if rr:
            bloco = (bloco[:rr.start()] + b"<w:rPr>" + corpo + rr.group(1)
                     + b"</w:rPr>" + bloco[rr.end():])
        else:
            bloco = por_ppr_rpr(bloco, None, b"<w:rPr>" + corpo + b"</w:rPr>")
        mexeu = True

    if not mexeu:
        return styles_xml, False
    return styles_xml[:m.start()] + bloco + styles_xml[m.end():], True


# ------------------------------------------------------------- paragrafos
def sem_pstyle(pPr_bytes):
    return re.sub(rb'<w:pStyle w:val="[^"]*"/>', b"", pPr_bytes)


def escrever_ppr(bloco, novo_ppr_interno):
    """Troca o pPr DO PROPRIO paragrafo, sem tocar no de caixa de texto dentro."""
    corte = len(bloco)
    for m in re.finditer(rb"<w:r(?:\s[^>]*)?>|<w:p(?:\s[^>]*)?>", bloco[1:]):
        corte = m.start() + 1
        break
    cabeca, cauda = bloco[:corte], bloco[corte:]
    novo = b"<w:pPr>" + novo_ppr_interno + b"</w:pPr>" if novo_ppr_interno else b""

    # Paragrafo vazio vem autofechado, `<w:p .../>`, e nao tem onde por o pPr.
    # Inserido depois do `/>`, isto e, fora do paragrafo, o Word o descarta em
    # silencio, e o vazio fica sem protecao.
    if not cauda and re.fullmatch(rb"<w:p(?:\s[^>]*)?/>", cabeca):
        abertura = cabeca[:-2] + b">"
        return abertura + novo + b"</w:p>"

    m = re.search(rb"<w:pPr>.*?</w:pPr>", cabeca, re.S)
    if m:
        cabeca = cabeca[:m.start()] + novo + cabeca[m.end():]
    else:
        fim_abertura = cabeca.find(b">") + 1
        cabeca = cabeca[:fim_abertura] + novo + cabeca[fim_abertura:]
    return cabeca + cauda


def com_pstyle(pPr_bytes, sid):
    """pPr com o pStyle pedido em primeiro lugar, que e onde o esquema o quer."""
    resto = sem_pstyle(pPr_bytes)
    return b'<w:pStyle w:val="' + sid.encode() + b'"/>' + resto


def tirar(pPr_bytes, *tags):
    for t in tags:
        pPr_bytes = re.sub(rb"<" + t + rb"(?:\s[^>]*)?/>", b"", pPr_bytes)
    return pPr_bytes


def do_estilo(styles_xml, sid, dentro, tags):
    """As tags pedidas do pPr (ou rPr) de um estilo, como estao escritas."""
    m = bloco_do_estilo(styles_xml, sid)
    if not m:
        return []
    caixa = re.search(rb"<" + dentro + rb">(.*?)</" + dentro + rb">",
                      m.group(0), re.S)
    if not caixa:
        return []
    saida = []
    for t in tags:
        x = re.search(rb"<" + t + rb"(?:\s[^>]*)?/>", caixa.group(1))
        if x:
            saida.append(x.group(0))
    return saida


def nome_da_tag(t):
    """b'<w:sz w:val="28"/>' -> b'w:sz', e b'<w:b/>' -> b'w:b'.

    Cortar no primeiro espaco nao basta: tag sem atributo nao tem espaco, e
    `<w:b/>`.split(b' ')[0] devolve a tag inteira. Isso fazia toda tag sem
    atributo parecer ausente, e a congelacao duplicava `<w:bCs/>`.
    """
    m = re.match(rb"<([^\s/>]+)", t)
    return m.group(1) if m else t


def ja_declara(xml, tag):
    return bool(re.search(rb"<" + re.escape(nome_da_tag(tag))
                          + rb"(?:\s[^>]*)?/>", xml))


def congelar(bloco, do_ppr, do_rpr):
    """Escreve no paragrafo o que ele herdava do estilo, so o que falta.

    Serve ao paragrafo do pre-textual que usa estilo de titulo sem ser titulo:
    tirar o estilo dele leva junto o centro e o negrito que vinham de la. O que
    o paragrafo ja declara fica como esta, porque ja o protege.
    """
    p = O.ppr(bloco)
    faltam = b"".join(t for t in do_ppr if not ja_declara(p, t))
    if faltam:
        # antes do rPr, que no pPr e o penultimo elemento do esquema
        m = re.search(rb"<w:rPr>", p)
        p = p[:m.start()] + faltam + p[m.start():] if m else p + faltam
        bloco = escrever_ppr(bloco, p)

    if not do_rpr:
        return bloco

    def completar(m):
        r = m.group(1)
        falta = b"".join(t for t in do_rpr if not ja_declara(r, t))
        return b"<w:rPr>" + falta + r + b"</w:rPr>"

    bloco = re.sub(rb"<w:rPr>(.*?)</w:rPr>", completar, bloco, flags=re.S)
    bloco = re.sub(rb"(<w:r(?:\s[^>]*)?>)(?!<w:rPr>)",
                   rb"\1<w:rPr>" + b"".join(do_rpr) + rb"</w:rPr>", bloco)
    return bloco


def distancia_do_cabecalho(sup_cm):
    """Da borda ao alto da linha do cabecalho, para que o meio dela caia no meio
    da margem superior."""
    linha = int(round(NORMA["cabecalho"]["corpo_pt"] * FOLGA_DE_LINHA * PT))
    return (tw(sup_cm) - linha) // 2


def escrever_margens(doc, margens_cm):
    """Troca o pgMar de toda secao. Nao mexe no tamanho do papel."""
    novo = (b'<w:pgMar w:top="%d" w:right="%d" w:bottom="%d" w:left="%d" '
            b'w:header="%d" w:footer="708" w:gutter="0"/>'
            % (tw(margens_cm["sup"]), tw(margens_cm["dir"]),
               tw(margens_cm["inf"]), tw(margens_cm["esq"]),
               distancia_do_cabecalho(margens_cm["sup"])))
    doc, n = re.subn(rb"<w:pgMar[^>]*/>", novo, doc)
    return doc, n


def desligar_hifenizacao(settings_xml):
    """Desliga a hifenizacao do documento, se o arquivo a trazia ligada."""
    if settings_xml is None:
        return None, False
    m = re.search(rb'<w:autoHyphenation(?:\s[^>]*)?/>', settings_xml)
    if not m:
        return settings_xml, False
    ligada = not re.search(rb'w:val="(?:false|0|off)"', m.group(0))
    return settings_xml[:m.start()] + settings_xml[m.end():], ligada


def sem_msonormal(doc, est):
    """O estilo `msonormal`, que o Word cria quando se cola texto de pagina da
    web ou de e-mail, passa ao Normal: o paragrafo que o usa perde o estilo, e o
    estilo sai do arquivo. Estilo que se baseava nele passa a se basear no
    Normal. Devolve o documento, os estilos e quantos paragrafos o usavam."""
    ids = [sid for sid, x in O.estilos_definidos(est).items()
           if re.fullmatch(r"(?i)msonormal\d*", sid)
           or re.fullmatch(r"(?i)msonormal\d*", x.get("nome") or "")]
    n = 0
    for sid in ids:
        b_sid = sid.encode()
        doc, k = re.subn(rb'<w:pStyle w:val="%s"/>' % re.escape(b_sid), b"", doc)
        n += k
        est = re.sub(rb'<w:style\b[^>]*w:styleId="%s"[^>]*>.*?</w:style>' % re.escape(b_sid),
                     b"", est, flags=re.S)
        est = re.sub(rb'<w:(basedOn|next) w:val="%s"/>' % re.escape(b_sid),
                     rb'<w:\1 w:val="Normal"/>', est)
    # O pPr que so tinha o estilo fica vazio, e sai.
    doc = doc.replace(b"<w:pPr></w:pPr>", b"")
    return doc, est, n, ids


def ligar_hifenizacao(settings_xml):
    """Liga a hifenizacao do documento.

    A posicao importa: no esquema, `autoHyphenation` vem logo depois de
    `defaultTabStop`. Posto fora de ordem, o Word reclama do arquivo.
    """
    if settings_xml is None:
        return None, False
    if re.search(rb"<w:autoHyphenation(?:\s[^>]*)?/>", settings_xml):
        settings_xml = re.sub(rb"<w:autoHyphenation(?:\s[^>]*)?/>",
                              b'<w:autoHyphenation w:val="true"/>', settings_xml)
        return settings_xml, True
    # `consecutiveHyphenLimit` vai logo depois, na ordem do esquema. Sem ele o
    # Word hifeniza linha apos linha.
    tag = (b'<w:autoHyphenation w:val="true"/>'
           b'<w:consecutiveHyphenLimit w:val="%d"/>'
           % NORMA["hifens_seguidos"])
    settings_xml = re.sub(rb"<w:consecutiveHyphenLimit(?:\s[^>]*)?/>", b"",
                          settings_xml)
    m = re.search(rb"<w:defaultTabStop(?:\s[^>]*)?/>", settings_xml)
    if m:
        return settings_xml[:m.end()] + tag + settings_xml[m.end():], True
    if b"</w:settings>" in settings_xml:
        return settings_xml.replace(b"</w:settings>", tag + b"</w:settings>", 1), True
    return settings_xml, False


def titulo_no_cabecalho(parte_xml, largura_tw, com_numero=False):
    """Poe o titulo do capitulo a esquerda do numero de pagina, no cabecalho.

    O campo e `STYLEREF 1`, que le o paragrafo em estilo Titulo 1 daquela pagina,
    ou o ultimo antes dela. O argumento numerico evita o nome do estilo, que e
    localizado: num documento em portugues o Word chama `heading 1` de
    `Titulo 1`, e um campo escrito com um dos dois falha no outro.

    O campo le o titulo VIVO. Guardar uma copia curta num paragrafo oculto
    resolveria o comprimento e criaria um defeito pior: copia de texto que o
    autor edita envelhece em silencio, e o cabecalho passa a mentir sem que nada
    acuse. Por isso o titulo e que encurta, na divisao no dois-pontos.

    O numero fica onde estava, a direita, agora por tabulacao e nao por
    alinhamento, porque o paragrafo passou a ter conteudo a esquerda.
    """
    alvo = None
    for a, b in O.paragrafos(parte_xml):
        if re.search(rb"PAGE", parte_xml[a:b]):
            alvo = (a, b)
            break
    if alvo is None:
        return parte_xml, False
    a, b = alvo
    bloco = parte_xml[a:b]
    if b"STYLEREF" in bloco:
        return parte_xml, False

    # A ordem do pPr nao e livre: `w:pStyle` e o primeiro elemento e `w:tabs`
    # vem depois dele. Posto antes, o Word ignora a tabulacao e o numero de
    # pagina sai no meio da linha, colado no titulo.
    pp = tirar(O.ppr(bloco), b"w:jc")
    pp = re.sub(rb"<w:tabs>.*?</w:tabs>", b"", pp, flags=re.S)
    # O corpo vem do estilo `Cabealho`, e o paragrafo tem de usa-lo.
    pp = com_pstyle(pp, "Cabealho")
    # A parada da esquerda e a do vao entre o numero e o titulo, que o campo
    # traz com a tabulacao do titulo. Sem ela, essa tabulacao salta para a parada do numero de pagina e empurra o titulo para
    # a direita.
    tabs = (b'<w:tabs><w:tab w:val="left" w:pos="%d"/>'
            b'<w:tab w:val="right" w:pos="%d"/></w:tabs>'
            % (tw(NORMA["numero_de_titulo"]["parada_t1_cm"]), largura_tw))
    mst = re.match(rb'<w:pStyle w:val="[^"]*"/>', pp)
    pp = (pp[:mst.end()] + tabs + pp[mst.end():]) if mst else (tabs + pp)
    bloco = escrever_ppr(bloco, pp)

    m = re.search(rb"<w:r(?:\s[^>]*)?>", bloco)
    if not m:
        return parte_xml, False
    def styleref(chave):
        return (b'<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
                b'<w:r><w:instrText xml:space="preserve"> STYLEREF 1 ' + chave
                + rb"\* MERGEFORMAT </w:instrText></w:r>"
                b'<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
                b"<w:r><w:t> </w:t></w:r>"
                b'<w:r><w:fldChar w:fldCharType="end"/></w:r>')
    def instr(texto):
        return b'<w:r><w:instrText xml:space="preserve">' + texto + b"</w:instrText></w:r>"

    # Com a numeracao automatica, o numero do capitulo nao e texto do titulo, e
    # o STYLEREF simples o perde: `\n` traz o numero da lista. No capitulo sem
    # numero (REFERENCIAS, APENDICES), porem, `\n` devolve "0", e o cabecalho
    # sairia "0 REFERENCIAS": o IF so poe o numero, e a tabulacao dele, quando
    # ele nao e 0.
    numero = (b'<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
              + instr(b' IF "') + styleref(rb"\n ") + instr(b'" = "0" "" "')
              + styleref(rb"\n ") + b"<w:r><w:tab/></w:r>" + instr(b'" ')
              + b'<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
              b"<w:r><w:t> </w:t></w:r>"
              b'<w:r><w:fldChar w:fldCharType="end"/></w:r>')
    campo = (numero if com_numero else b"") + styleref(b"") + b"<w:r><w:tab/></w:r>"
    # O corpo do cabecalho vem do estilo `Cabealho`, e nao de cada corrida: o
    # resultado do campo e gerado pelo Word a cada atualizacao e nao herda o que
    # se escreveu nas corridas do campo.
    resto = re.sub(rb"<w:szCs(?:\s[^>]*)?/>", b"",
                   re.sub(rb"<w:sz(?:\s[^>]*)?/>", b"", bloco[m.start():]))
    return parte_xml[:a] + bloco[:m.start()] + campo + resto + parte_xml[b:], True


def numero_de_pagina_a_direita(parte_xml):
    """Alinha a direita o paragrafo do cabecalho que carrega o campo PAGE, e o
    poe no estilo `Cabealho`, de onde vem o corpo e o italico: o que estiver
    escrito nas corridas sai, porque o venceria."""
    mudou = [0]

    def um(m):
        bloco = m.group(0)
        if not re.search(rb"PAGE", bloco):
            return bloco
        p = com_pstyle(tirar(O.ppr(bloco), b"w:jc"), "Cabealho")
        p = no_ppr(p, b'<w:jc w:val="right"/>',
                   rb"textDirection|textAlignment|textboxTightWrap|outlineLvl|divId"
                   rb"|cnfStyle|rPr|sectPr|pPrChange")
        novo = sem_corpo_de_letra(escrever_ppr(bloco, p))
        if novo != bloco:
            mudou[0] += 1
        return novo

    saida, pos = [], 0
    for a, b in O.paragrafos(parte_xml):
        saida.append(parte_xml[pos:a])
        saida.append(um(re.match(rb".*", parte_xml[a:b], re.S)))
        pos = b
    saida.append(parte_xml[pos:])
    return b"".join(saida), mudou[0]

CABECALHO_COM_NUMERO = "header-numero.xml"

PARAGRAFO_DO_NUMERO = (
    b'<w:p><w:pPr><w:pStyle w:val="Cabealho"/><w:jc w:val="right"/></w:pPr>'
    b'<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
    b'<w:r><w:instrText xml:space="preserve"> PAGE \\* MERGEFORMAT </w:instrText></w:r>'
    b'<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
    b"<w:r><w:t>1</w:t></w:r>"
    b'<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>')

CABECALHO_NUMERADO = (
    b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
    b'<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    + PARAGRAFO_DO_NUMERO + b"</w:hdr>")


def com_numero_de_pagina(parte_xml):
    """Poe o numero de pagina num cabecalho que nao o tem. Cabecalho sem nada
    escrito nem desenhado perde os paragrafos vazios, que empurrariam o numero
    para baixo; o que tem conteudo ganha o numero depois dele."""
    corpo = re.search(rb"(<w:hdr\b[^>]*>)(.*)(</w:hdr>)", parte_xml, re.S)
    if not corpo:
        return parte_xml
    dentro = corpo.group(2)
    vazio = (not re.search(rb"<w:t[ >]", dentro)
             and not re.search(rb"<w:(?:drawing|pict|object|fldChar|fldSimple)\b", dentro))
    novo = PARAGRAFO_DO_NUMERO if vazio else dentro + PARAGRAFO_DO_NUMERO
    return parte_xml[:corpo.start(2)] + novo + parte_xml[corpo.end(2):]


def numero_onde_falta(doc, partes, rels_xml, id_pretextual, id_novo):
    """Acha as secoes da parte textual que nao mostram numero de pagina, nem no
    cabecalho nem no rodape, e decide onde ele entra.

    A secao sem cabecalho proprio herda o da anterior, e a primeira da parte
    textual herdaria o cabecalho vazio do pre-textual: ela recebe uma referencia
    a um cabecalho novo, com o numero, e as seguintes que tambem nao tem
    cabecalho herdam esse. A que tem cabecalho proprio sem numero recebe o
    numero nele. Devolve o documento, as partes de cabecalho que recebem o
    numero, se o cabecalho novo foi usado, e quantas secoes faltavam.
    """
    alvos = {}
    for m in re.finditer(rb"<Relationship\b[^>]*>", rels_xml):
        rid = re.search(rb'Id="([^"]*)"', m.group(0))
        alvo = re.search(rb'Target="([^"]*)"', m.group(0))
        if rid and alvo:
            alvos[rid.group(1).decode()] = "word/" + alvo.group(1).decode().lstrip("/")
    pares = b"<w:evenAndOddHeaders" in partes.get("word/settings.xml", b"")

    def tem_numero(parte):
        # A parte que ja recebeu o numero nesta passada tambem conta.
        return parte is not None and (parte in receber
                                      or parte == "word/" + CABECALHO_COM_NUMERO
                                      or b"PAGE" in partes.get(parte, b""))

    secoes = list(re.finditer(rb"<w:sectPr\b.*?</w:sectPr>", doc, re.S))
    ultimo_pre = max((k for k, m in enumerate(secoes)
                      if ('r:id="%s"' % id_pretextual).encode() in m.group(0)), default=None)
    if ultimo_pre is None:
        return doc, set(), False, 0
    herdado = {}
    receber, novo_usado, faltavam, inserir = set(), False, 0, []
    for k, m in enumerate(secoes):
        sp = m.group(0)
        refs = {}
        for r in re.finditer(rb'<w:(header|footer)Reference w:type="(\w+)" r:id="([^"]*)"/>', sp):
            refs[(r.group(1).decode(), r.group(2).decode())] = alvos.get(r.group(3).decode())
        tipos = ["default"] + (["even"] if pares else []) + (
            ["first"] if b"<w:titlePg/>" in sp or b'<w:titlePg w:val="1"/>' in sp else [])
        efetivo = {}
        for t in ("default", "even", "first"):
            for lado in ("header", "footer"):
                if (lado, t) in refs:
                    herdado[(lado, t)] = refs[(lado, t)]
                efetivo[(lado, t)] = herdado.get((lado, t))
        if k <= ultimo_pre:
            continue
        if any(tem_numero(efetivo[(lado, t)]) for t in tipos for lado in ("header", "footer")):
            continue
        faltavam += 1
        faltam_ref = []
        for t in tipos:
            parte = efetivo[("header", t)]
            proprio = ("header", t) in refs
            if proprio and parte and parte != "word/" + CABECALHO_SEM_NUMERO:
                receber.add(parte)
            else:
                faltam_ref.append(t)
                herdado[("header", t)] = "word/" + CABECALHO_COM_NUMERO
        if faltam_ref:
            novo_usado = True
            inserir.append((m, faltam_ref))
    for m, faltam_ref in reversed(inserir):
        sp = m.group(0)
        for t in faltam_ref:
            sp = re.sub(rb'<w:headerReference w:type="%s"[^>]*/>' % t.encode(), b"", sp)
        pos = sp.find(b">") + 1
        refs_novas = b"".join(b'<w:headerReference w:type="%s" r:id="%s"/>'
                              % (t.encode(), id_novo.encode()) for t in faltam_ref)
        doc = doc[:m.start()] + sp[:pos] + refs_novas + sp[pos:] + doc[m.end():]
    return doc, receber, novo_usado, faltavam


CABECALHO_SEM_NUMERO = "header-sem-numero.xml"

CABECALHO_VAZIO = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
    '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    "<w:p/></w:hdr>").encode("utf-8")

TIPO_CABECALHO = ("http://schemas.openxmlformats.org/officeDocument/2006/"
                  "relationships/header")
CT_CABECALHO = ("application/vnd.openxmlformats-officedocument."
                "wordprocessingml.header+xml")


def proximo_rid(rels_xml):
    usados = [int(m.group(1)) for m in
              re.finditer(rb'Id="rId(\d+)"', rels_xml)]
    return "rId%d" % ((max(usados) + 1) if usados else 1)


def com_relacionamento(rels_xml, rid, alvo, tipo=None):
    novo = ('<Relationship Id="%s" Type="%s" Target="%s"/>'
            % (rid, tipo or TIPO_CABECALHO, alvo)).encode("utf-8")
    return rels_xml.replace(b"</Relationships>", novo + b"</Relationships>", 1)


def com_override(ctypes_xml, parte, ct=None):
    novo = ('<Override PartName="%s" ContentType="%s"/>'
            % (parte, ct or CT_CABECALHO)).encode("utf-8")
    return ctypes_xml.replace(b"</Types>", novo + b"</Types>", 1)


def tirar_tampas(doc):
    """Apaga as caixas de texto vazias e brancas que tampam o numero de pagina.

    POR QUE ELAS EXISTEM E POR QUE SAEM

    A NBR 14724 manda contar as folhas pre-textuais e nao numera-las. Sem saber
    fazer isso por secao, ha autor que cobre o numero com um retangulo branco
    sobre fundo branco, um por folha, posto onde o numero fica.

    A gambiarra e fragil por natureza: mudada a margem, a coluna se desloca, o
    numero vai junto e a tampa fica para tras, e os numeros do pre-textual passam
    a aparecer.

    Sai so o que e caixa de texto SEM texto dentro e com preenchimento e contorno
    iguais ao fundo. Caixa com conteudo, imagem e grafico ficam.
    """
    fora = []
    for m in re.finditer(rb"<mc:AlternateContent>.*?</mc:AlternateContent>"
                         rb"|<w:drawing>.*?</w:drawing>", doc, re.S):
        bl = m.group(0)
        # O nome do objeto nao serve de criterio: a tampa pode ter o nome de
        # qualquer forma do Word. O que a define e o que ela e: forma vazia,
        # com preenchimento e contorno iguais ao fundo, isto e, invisivel no
        # papel e util so para cobrir o que esta embaixo.
        if b"<wps:wsp>" not in bl and b"<pic:pic" in bl:
            continue                      # imagem de verdade
        if b"".join(re.findall(rb"<w:t(?:\s[^>]*)?>(.*?)</w:t>", bl, re.S)).strip():
            continue
        if bl.count(b'<a:schemeClr val="bg1"/>') < 2:
            continue
        fora.append((m.start(), m.end()))
    for a, b in sorted(fora, reverse=True):
        doc = doc[:a] + doc[b:]
    return doc, len(fora)


def separar_secao(doc, alvo, id_cabecalho):
    """Faz do pre-textual uma secao com cabecalho proprio, que nao traz numero.

    A capa passa a ser a folha 1 da contagem e o numero so aparece a partir da
    parte textual, que e o que a NBR 14724 pede. O Word continua a contagem de
    uma secao para a outra sozinho: o que muda entre elas e so o cabecalho.

    `alvo` e a fatia (inicio, fim) do ultimo paragrafo do pre-textual, lida do
    documento ja reescrito.
    """
    final = None
    for m in re.finditer(rb"<w:sectPr\b.*?</w:sectPr>", doc, re.S):
        final = m
    if final is None:
        return doc, False

    nova = final.group(0)
    nova = re.sub(rb"<w:headerReference[^>]*/>", b"", nova)
    nova = re.sub(rb"<w:footerReference[^>]*/>", b"", nova)
    nova = re.sub(rb"<w:pgNumType[^>]*/>", b"", nova)
    pos = nova.find(b">") + 1
    nova = (nova[:pos]
            + b'<w:headerReference w:type="default" r:id="%s"/>'
            % id_cabecalho.encode()
            + nova[pos:])
    nova = nova.replace(b"</w:sectPr>", b'<w:pgNumType w:start="%d"/></w:sectPr>'
                        % NORMA["numeracao"]["inicio"])

    # a secao final deixa de reiniciar a contagem: ela continua a da primeira
    doc = (doc[:final.start()]
           + re.sub(rb"<w:pgNumType[^>]*/>", b"", final.group(0))
           + doc[final.end():])

    a, b = alvo
    bloco = doc[a:b]
    p = O.ppr(bloco)
    if b"<w:sectPr" in p:
        return doc, False
    return doc[:a] + escrever_ppr(bloco, p + nova) + doc[b:], True

FOLGA_DE_LINHA = 1.15       # o que o Word acrescenta a uma linha "simples"


def sem_corpo_de_letra(bloco):
    """Tira o corpo de letra escrito a mao, para que o paragrafo siga o estilo.

    E excecao a regra de nao mexer em corrida de texto, e vale onde o que se
    quer e uniformidade: no pre-textual, nos titulos, nas legendas e nas
    citacoes. Corpo escrito na corrida vence o do estilo.
    """
    return re.sub(rb"<w:szCs(?:\s[^>]*)?/>", b"",
                  re.sub(rb"<w:sz(?:\s[^>]*)?/>", b"", bloco))


def sem_espacejamento(bloco):
    """Tira o espacejamento de letras escrito na corrida (`w:spacing w:val`,
    que so existe no rPr: o do paragrafo usa before, after e line). No titulo,
    o Word o leva para a entrada do sumario, e a linha estoura."""
    return re.sub(rb'<w:spacing w:val="-?\d+"/>', b"", bloco)


def altura_de_linha(corpo_pt, entrelinha=1.0):
    return corpo_pt * entrelinha * FOLGA_DE_LINHA * PT


def linhas_de(texto, largura_tw, corpo_pt):
    """Quantas linhas um paragrafo ocupa, estimado pela largura media da letra.

    Meio em de largura media e a regra de bolo das serifadas, e o erro aqui custa
    pouco: ele desloca um bloco da capa em fracao de linha, nao muda o que esta
    escrito. O que NAO se pode fazer e contar um paragrafo por linha, que e o que
    faz o titulo de tres linhas transbordar.
    """
    if not texto:
        return 1
    por_linha = max(1, int(largura_tw / (corpo_pt * 10)))
    return max(1, -(-len(texto) // por_linha))


def blocos_de(fichas, ini, fim):
    """As corridas de paragrafos com texto entre ini e fim, e os vazios entre elas."""
    blocos, atual = [], []
    for k in range(ini, fim + 1):
        if fichas[k]["texto"]:
            atual.append(k)
        elif atual:
            blocos.append(atual)
            atual = []
    if atual:
        blocos.append(atual)
    return blocos


def diagramar_capa(fichas, ini, fim, largura_tw, altura_tw, corpo_pt, entrelinha,
                   fracoes, folga_pe_tw=0):
    """Poe os blocos da capa em fracoes da altura util, e devolve as edicoes.

    POR QUE A CAPA NAO PODE SER FEITA DE LINHAS EM BRANCO

    E comum o autor empurrar cada bloco da capa para o seu lugar com paragrafos
    vazios, contando linhas. Isso amarra a diagramacao ao corpo da letra e a
    altura da pagina que vigoravam quando ele contou: mudada a margem ou o corpo,
    o texto cresce e a ultima linha sai da folha.

    A regra que substitui isso e posicional. Os blocos sao as corridas de
    paragrafos com texto; o primeiro fica no alto, o ultimo termina no pe da
    mancha, e os do meio comecam nas fracoes que a norma fixa. Recalculado a cada
    mudanca de margem ou de corpo, e por isso nao volta a transbordar.
    """
    blocos = blocos_de(fichas, ini, fim)
    if len(blocos) < 2:
        return [], []

    linha_tw = altura_de_linha(corpo_pt, entrelinha)
    alturas = [sum(linhas_de(fichas[k]["texto"], largura_tw, corpo_pt)
                   for k in bl) * linha_tw for bl in blocos]

    # onde cada bloco comeca, em twips do alto da mancha
    alvos = [0.0]
    meio = fracoes[:max(0, len(blocos) - 2)]
    while len(meio) < len(blocos) - 2:
        # bloco do meio a mais do que a norma previu: distribui o que sobra
        ultimo = meio[-1] if meio else 0.25
        meio.append(ultimo + (1.0 - ultimo) / (len(blocos) - len(meio)))
    for f in meio:
        alvos.append(f * altura_tw)
    # O ultimo termina um pouco acima do pe. A altura dos blocos e estimada, e
    # sem folga um erro de meia linha derruba a ultima linha para a folha
    # seguinte.
    alvos.append(altura_tw - alturas[-1] - folga_pe_tw)

    def sem_corpo(bloco):
        return sem_corpo_de_letra(bloco)

    def _sem_corpo_antigo(bloco):
        """Tira o corpo de letra escrito a mao, para que a capa siga o Normal.

        Aqui e excecao a regra de nao mexer em corrida de texto: na capa sao
        poucos paragrafos, e o que se quer e uniformidade."""
        return re.sub(rb"<w:szCs(?:\s[^>]*)?/>", b"",
                      re.sub(rb"<w:sz(?:\s[^>]*)?/>", b"", bloco))

    edicoes, avisos = [], []
    acumulado = 0.0
    for k, bl in enumerate(blocos):
        antes = alvos[k] - acumulado
        if antes < 0:
            avisos.append("o bloco %d nao cabe onde a norma o quer: faltam "
                          "%.1f cm" % (k + 1, -antes / CM))
            antes = 0.0
        primeiro = fichas[bl[0]]
        bloco_1 = sem_corpo(primeiro["bloco"])
        p = tirar(O.ppr(bloco_1), b"w:spacing", b"w:suppressAutoHyphens")
        p += (b'<w:spacing w:before="%d" w:after="0" w:line="%d" '
              b'w:lineRule="auto"/>' % (int(round(antes)),
                                        int(round(240 * entrelinha))))
        p += b"<w:suppressAutoHyphens/>"
        edicoes.append((primeiro["ini"], primeiro["fim"],
                        escrever_ppr(bloco_1, p)))
        # Capa nao se hifeniza: sao nomes e titulo, e a divisao silabica ali le
        # como erro.
        for k2 in bl[1:]:
            b2 = sem_corpo(fichas[k2]["bloco"])
            # As demais linhas do bloco tambem perdem o espacamento escrito a
            # mao: a conta que posiciona os blocos supoe entrelinha simples, e
            # entrelinha escrita em cada linha faz os blocos crescerem alem dela.
            p2 = tirar(O.ppr(b2), b"w:suppressAutoHyphens", b"w:spacing")
            edicoes.append((fichas[k2]["ini"], fichas[k2]["fim"],
                            escrever_ppr(b2, p2 + b"<w:suppressAutoHyphens/>")))
        acumulado = alvos[k] + alturas[k]

    # os vazios da faixa saem: o espaco agora vem do estilo, e nao de linhas
    for k in range(ini, fim + 1):
        if fichas[k]["texto"]:
            continue
        if O.travado(fichas[k]["bloco"]):
            continue
        edicoes.append((fichas[k]["ini"], fichas[k]["fim"], b""))
    return edicoes, avisos


def congelar_espaco(fichas, ini, fim, corpo_velho_pt, quebrar_pagina=False,
                    tirar_corpo=None):
    """Troca corrida de paragrafo vazio por espaco antes, na altura que ela tinha.

    Serve ao pre-textual fora da capa, onde a diagramacao e do autor e nao ha
    norma que diga onde cada coisa deve ficar. Congelar em twips faz o espaco
    parar de depender do corpo da letra: o texto cresce e o vazio nao.
    """
    edicoes, corrida = [], []
    alt = altura_de_linha(corpo_velho_pt, 1.0)
    primeiro = True
    for k in range(ini, fim + 1):
        f = fichas[k]
        bloco = tirar_corpo(f["bloco"]) if tirar_corpo else f["bloco"]
        if not f["texto"]:
            if not O.travado(f["bloco"]):
                corrida.append(k)
            continue
        if corrida or (primeiro and quebrar_pagina):
            p = tirar(O.ppr(bloco), b"w:spacing", b"w:pageBreakBefore")
            if primeiro and quebrar_pagina:
                # Tirados os vazios, a capa deixa de encher a folha e a folha de
                # rosto subiria para junto dela.
                p = b"<w:pageBreakBefore/>" + p
            if corrida:
                p += b'<w:spacing w:before="%d"/>' % int(round(len(corrida) * alt))
            edicoes.append((f["ini"], f["fim"], escrever_ppr(bloco, p)))
            for k2 in corrida:
                edicoes.append((fichas[k2]["ini"], fichas[k2]["fim"], b""))
            corrida = []
        elif tirar_corpo and bloco != f["bloco"]:
            edicoes.append((f["ini"], f["fim"], bloco))
        primeiro = False
    return edicoes

ESTILO_SUBTITULO = "SubtituloCapitulo"


def _textos_da_corrida(corrida):
    return re.findall(rb"<w:t(?:\s[^>]*)?>.*?</w:t>", corrida, re.S)


def _filhos(corpo):
    """Os elementos de primeiro nivel do corpo do paragrafo, com a posicao de cada
    um, contando o aninhamento do mesmo nome. None se o XML nao fechar."""
    saida, pos = [], 0
    while True:
        m = re.compile(rb"<w:([A-Za-z]+)\b[^>]*?(/?)>").search(corpo, pos)
        if not m:
            break
        nome = m.group(1)
        if m.group(2) == b"/":
            fim = m.end()
        else:
            pat = re.compile(rb"<w:" + nome + rb"\b[^>]*?(/?)>|</w:" + nome + rb">")
            prof, k = 1, m.end()
            while prof:
                mm = pat.search(corpo, k)
                if not mm:
                    return None
                if mm.group(0).startswith(b"</"):
                    prof -= 1
                elif mm.group(1) != b"/":
                    prof += 1
                k = mm.end()
            fim = k
        saida.append((m.start(), fim, nome.decode()))
        pos = fim
    return saida


def dividir_no_caractere(bloco, pos):
    """Divide um paragrafo em dois no caractere `pos` do texto concatenado.

    POR QUE NAO SE DIVIDE PELO TEXTO

    O titulo nao e uma corrida so: quando traz uma expressao estrangeira em
    italico, ela esta numa corrida propria. Cortar o texto e reescrever o paragrafo perderia o italico. O corte
    e feito na corrida em que o caractere cai, e o resto passa inteiro para um
    lado ou para o outro.

    O QUE NAO PODE SE PERDER

    Tudo o que o paragrafo tem alem das corridas: marcador, hiperlink, campo,
    comentario. Levar adiante so as corridas faz o resto sumir sem que a
    conferencia de texto perceba, porque o texto nao muda. Cada elemento vai para
    o lado em que esta, e
    quando o corte cairia dentro de um hiperlink, de um campo ou de outro
    elemento que nao se parte, o paragrafo nao e dividido.

    Devolve (antes, depois) ou None se o paragrafo nao puder ser dividido ali.
    """
    m = re.match(rb"<w:p(?:\s[^>]*)?>", bloco)
    if not m or not bloco.endswith(b"</w:p>"):
        return None
    abertura, corpo = m.group(0), bloco[m.end():-len(b"</w:p>")]
    mp = re.match(rb"<w:pPr>.*?</w:pPr>", corpo, re.S)
    ppr, resto = (mp.group(0), corpo[mp.end():]) if mp else (b"", corpo)

    filhos = _filhos(resto)
    if filhos is None:
        return None
    antes, depois, contados, cortou, campo = [], [], 0, False, 0

    def abre_campo(x):
        return (x.count(b'w:fldCharType="begin"') - x.count(b'w:fldCharType="end"'))

    for ini, fim, nome in filhos:
        x = resto[ini:fim]
        if cortou:
            depois.append(x)
            continue
        n = len(O.texto(x))
        if n == 0:
            # Sem texto: marcador, comentario, pedaco de campo. Fica do lado em
            # que esta; no ponto exato do corte, vai para a direita, a menos que
            # feche um campo aberto a esquerda.
            if contados < pos or (contados == pos and campo > 0):
                antes.append(x)
                campo += abre_campo(x)
            else:
                depois.append(x)
            continue
        if contados + n <= pos:
            antes.append(x)
            campo += abre_campo(x)
            contados += n
            continue
        if campo > 0:
            return None             # o corte cairia no meio de um campo
        if contados == pos:
            depois.append(x)
            cortou = True
            continue
        if nome != "r":
            return None             # hiperlink, campo simples, etc.: nao se parte
        # O corte cai dentro desta corrida, que pode ter mais de uma `w:t`: o
        # numero do titulo, a tabulacao e o titulo ficam na mesma corrida.
        dentro = pos - contados
        abre_r = re.match(rb"<w:r(?:\s[^>]*)?>", x).group(0)
        miolo = x[len(abre_r):-len(b"</w:r>")]
        mr_pr = re.match(rb"<w:rPr>.*?</w:rPr>", miolo, re.S)
        rpr = mr_pr.group(0) if mr_pr else b""
        conteudo = miolo[len(rpr):]
        ts = list(re.finditer(rb"<w:t(?:\s[^>]*)?>(.*?)</w:t>", conteudo, re.S))
        acum = 0
        for j, t in enumerate(ts):
            texto = t.group(1).decode("utf-8", "replace")
            if acum + len(texto) > dentro or j == len(ts) - 1:
                break
            acum += len(texto)
        local = dentro - acum

        def wt(y):
            return b'<w:t xml:space="preserve">' + y.encode("utf-8") + b"</w:t>"

        esq = conteudo[:t.start()] + (wt(texto[:local]) if texto[:local] else b"")
        dir_ = (wt(texto[local:]) if texto[local:] else b"") + conteudo[t.end():]
        if esq:
            antes.append(abre_r + rpr + esq + b"</w:r>")
        if dir_:
            depois.append(abre_r + rpr + dir_ + b"</w:r>")
        cortou = True
    if not cortou or not antes or not depois:
        return None
    return (abertura + ppr + b"".join(antes) + b"</w:p>",
            abertura + ppr + b"".join(depois) + b"</w:p>")


def sem_negrito(bloco):
    """Tira o negrito escrito nas corridas, para que o estilo governe.

    O subtitulo herda as corridas do titulo, que sao negrito. Declarar no estilo
    que ele nao e negrito nao basta: direta vence estilo, e o subtitulo sairia
    em negrito com o estilo dizendo o contrario.
    """
    return re.sub(rb"<w:bCs(?:\s[^>]*)?/>", b"",
                  re.sub(rb"<w:b(?:\s[^>]*)?/>", b"", bloco))


def tirar_dois_pontos(bloco):
    """Tira o dois-pontos e os espacos que abrem o subtitulo."""
    m = re.search(rb"<w:t(?:\s[^>]*)?>(.*?)</w:t>", bloco, re.S)
    if not m:
        return bloco
    t = m.group(1).decode("utf-8", "replace").lstrip()
    if t.startswith(":"):
        t = t[1:].lstrip()
    novo = b'<w:t xml:space="preserve">' + t.encode("utf-8") + b"</w:t>"
    return bloco[:m.start()] + novo + bloco[m.end():]

# ------------------------------------------------------------------ plano
NUMERO_CRU = re.compile(r"^(\s*)(\d+(?:\.\d+)*)\.?(\s+)(?=\S)")
ESTILO_DEDICATORIA = "Dedicatoria"
ESTILO_EPIGRAFE = "Epigrafe"
PE_DA_PAGINA = (ESTILO_DEDICATORIA, ESTILO_EPIGRAFE)
TAB = object()      # marca, em `trocar_trecho`, onde vai uma tabulacao


def depois_de_ilustracao(fichas, i):
    """A fonte vem logo abaixo da figura ou da tabela, com no maximo dois vazios
    no meio. A da tabela fica fora da tabela, e a ficha anterior e a ultima
    celula."""
    k, vazios = i - 1, 0
    while k >= 0 and vazios <= 2:
        g = fichas[k]
        if g["figura"] or g["tabela"]:
            return True
        if g["texto"]:
            return False
        vazios += 1
        k -= 1
    return False


def _xml(t):
    return (t.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))


def trocar_trecho(bloco, ini, fim, novo):
    """Troca os caracteres [ini, fim) do texto do paragrafo por `novo`.

    O texto se espalha por varias corridas. O novo entra na `w:t` em que o
    trecho comeca; das seguintes sai o que ainda faltava cortar. As corridas nao
    se fundem, e o italico ou o negrito de cada uma ficam onde estavam. `novo` e
    uma lista de pedacos: texto, ou `TAB` para uma tabulacao.

    As posicoes contam o texto como ele esta escrito no XML, que e o que
    `O.texto` devolve.
    """
    partes, pos, feito, acum = [], 0, False, 0
    for m in re.finditer(rb"(<w:t)((?:\s[^>]*)?)>(.*?)</w:t>", bloco, re.S):
        t = m.group(3).decode("utf-8")
        a, b = acum, acum + len(t)
        acum = b
        if b <= ini or a >= fim and feito:
            continue
        lo, hi = max(ini, a) - a, min(fim, b) - a
        antes, depois = t[:lo], t[hi:]
        attrs = m.group(2)
        pres = attrs if b"xml:space" in attrs else attrs + b' xml:space="preserve"'
        pedacos = []
        if not feito:
            atual = antes
            for x in novo:
                if x is TAB:
                    pedacos.append(m.group(1) + pres + b">" + atual.encode("utf-8")
                                   + b"</w:t><w:tab/>")
                    atual = ""
                else:
                    atual += x
            atual += depois
            pedacos.append(m.group(1) + pres + b">" + atual.encode("utf-8") + b"</w:t>")
            feito = True
        else:
            pedacos.append(m.group(1) + pres + b">" + (antes + depois).encode("utf-8")
                           + b"</w:t>")
        partes.append(bloco[pos:m.start()])
        partes.append(b"".join(pedacos))
        pos = m.end()
        if b >= fim:
            break
    partes.append(bloco[pos:])
    return b"".join(partes)


def com_campo_seq(bloco, rotulo, ini, fim):
    """Poe o numero da ilustracao, em [ini, fim) do texto, dentro de um campo SEQ.

    E o que "Inserir legenda" do Word faz, e e pelo campo que a lista de
    ilustracoes e montada. O resultado guardado e o numero que o autor escreveu,
    e o texto do paragrafo nao muda. So se faz quando o numero cai inteiro numa
    `w:t` so; senao devolve o bloco como estava.
    """
    acum = 0
    for m in re.finditer(rb"<w:r(?:\s[^>]*)?>(?:(?!</w:r>).)*?</w:r>", bloco, re.S):
        corrida = m.group(0)
        ts = list(re.finditer(rb"(<w:t)((?:\s[^>]*)?)>(.*?)</w:t>", corrida, re.S))
        comeco = acum
        acum += sum(len(x.group(3).decode("utf-8")) for x in ts)
        if not (comeco <= ini and fim <= acum):
            continue
        if len(ts) != 1:
            return bloco
        t = ts[0].group(3).decode("utf-8")
        lo, hi = ini - comeco, fim - comeco
        rpr = re.search(rb"<w:rPr>.*?</w:rPr>", corrida, re.S)
        rpr = rpr.group(0) if rpr else b""

        def r(conteudo):
            return b"<w:r>" + rpr + conteudo + b"</w:r>"

        def wt(x):
            return b'<w:t xml:space="preserve">' + x.encode("utf-8") + b"</w:t>"
        nova = ((r(wt(t[:lo])) if t[:lo] else b"")
                + r(b'<w:fldChar w:fldCharType="begin"/>')
                + r(b'<w:instrText xml:space="preserve"> SEQ '
                    + rotulo.encode("utf-8") + rb" \* ARABIC </w:instrText>")
                + r(b'<w:fldChar w:fldCharType="separate"/>')
                + r(wt(t[lo:hi]))
                + r(b'<w:fldChar w:fldCharType="end"/>')
                + (r(wt(t[hi:])) if t[hi:] else b""))
        return bloco[:m.start()] + nova + bloco[m.end():]
    return bloco


def campo_de_lista(rotulo, estilo):
    """A lista de ilustracoes de um rotulo, montada pelo Word a partir dos SEQ."""
    return (b'<w:p><w:pPr><w:pStyle w:val="' + estilo.encode() + b'"/></w:pPr>'
            b'<w:r><w:fldChar w:fldCharType="begin" w:dirty="true"/></w:r>'
            b'<w:r><w:instrText xml:space="preserve"> TOC \\h \\z \\c "'
            + rotulo.encode("utf-8") + b'" </w:instrText></w:r>'
            b'<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
            b"<w:r><w:t>Lista automatica: no Word, clique aqui e tecle F9 para "
            b"atualizar.</w:t></w:r>"
            b'<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>')


PALAVRAS_CHAVE = re.compile(
    r"^\s*(Palavras[- ]chave|Keywords?|Key[- ]words|Palabras[- ]clave|Mots[- ]cl[ée]s)"
    r"\s*:\s*", re.I)
LIGACOES = {"de", "da", "do", "das", "dos", "e", "em", "of", "the", "and", "in",
            "for", "on", "del", "la", "le", "des", "et", "y"}


def palavras_chave_abnt(texto):
    """As palavras-chave como a NBR 6028:2021 (4.1.7) pede.

    Separadas por ponto e virgula, terminadas por ponto, com a inicial em
    minuscula, salvo nome proprio. Nome proprio aqui e a expressao em que toda
    palavra que nao e de ligacao comeca por maiuscula (`Universidade de
    Brasilia`), ou que traz sigla (`IBGE`). E criterio de maquina e erra no nome
    proprio de uma palavra so, como `Brasil`: por isso cada troca sai listada no
    relatorio, para o autor conferir.
    """
    itens = [x.strip() for x in re.split(r"[;.]", texto) if x.strip()]
    saida = []
    for x in itens:
        palavras = [w for w in re.findall(r"[\wÀ-ÿ]+", x) if w.lower() not in LIGACOES]
        sigla = any(len(w) > 1 and w.isupper() for w in palavras)
        proprio = len(palavras) > 1 and all(w[:1].isupper() for w in palavras)
        if not (sigla or proprio) and x[:1].isupper():
            x = x[:1].lower() + x[1:]
        saida.append(x)
    return "; ".join(saida) + "."


def recuo_efetivo(bloco, estilos):
    """O w:left que vale no paragrafo: o dele, ou o do estilo dele."""
    v = O.atributo(O.prop(O.ppr(bloco), b"w:ind") or b"", b"<w:ind", b"w:left")
    if v is None:
        sid = O.estilo_de(bloco)
        x = estilos.get(sid, {}).get("xml", b"") if sid else b""
        v = O.atributo(O.prop(x, b"w:ind") or b"", b"<w:ind", b"w:left")
    try:
        return int(v) if v is not None else 0
    except ValueError:
        return 0


def natureza_no_meio(doc, estilos, meio_tw, limiar_tw):
    """O paragrafo recuado da folha de rosto e da de aprovacao comeca no meio.

    Vale ate o primeiro paragrafo da dedicatoria, da epigrafe ou de titulo do
    pre-textual, e so para quem ja estava recuado alem do limiar: a natureza do
    trabalho e as linhas de assinatura. O `hanging` que o paragrafo tiver fica.
    """
    pars, n = O.paragrafos(doc), 0
    fim = next((k for k, (a, b) in enumerate(pars)
                if O.estilo_de(doc[a:b]) in PE_DA_PAGINA + ("Ttulo0", "Ttulo1")),
               0)
    for k in reversed(range(fim)):
        a, b = pars[k]
        bl = doc[a:b]
        if recuo_efetivo(bl, estilos) < limiar_tw:
            continue
        pp = O.ppr(bl)
        ind = O.prop(pp, b"w:ind")
        # Quem nao tem `hanging` tambem nao leva recuo de primeira linha: o
        # bloco comeca no meio, rente, como a dedicatoria, e nao com o recuo de
        # primeira linha do Normal.
        if ind and b"w:hanging=" in ind:
            novo_ind = re.sub(rb'\s*w:left="[^"]*"', b"", ind).replace(
                b"<w:ind", b'<w:ind w:left="%d"' % meio_tw, 1)
        else:
            novo_ind = b'<w:ind w:left="%d" w:firstLine="0"/>' % meio_tw
        # E em espaco simples, como o resto do pre-textual (NBR 14724, 5.2). O
        # paragrafo guarda o estilo do autor, que herda a entrelinha do Normal.
        # O `before` que a diagramacao da folha de rosto escreveu fica.
        sp = O.prop(pp, b"w:spacing")
        if sp:
            novo_sp = re.sub(rb'\s*w:(?:line|lineRule)="[^"]*"', b"", sp).replace(
                b"<w:spacing", b'<w:spacing w:line="240" w:lineRule="auto"', 1)
        else:
            novo_sp = b'<w:spacing w:line="240" w:lineRule="auto"/>'
        # `w:spacing` e `w:ind` tem lugar certo no pPr: antes de `w:jc` e do
        # `w:rPr` da marca de paragrafo.
        pp = tirar(pp, b"w:ind", b"w:spacing")
        m = re.search(rb"<w:(?:contextualSpacing|mirrorIndents|suppressOverlap|jc"
                      rb"|textDirection|textAlignment|textboxTightWrap|outlineLvl"
                      rb"|divId|cnfStyle|rPr|sectPr|pPrChange)\b", pp)
        k_ins = m.start() if m else len(pp)
        pp = pp[:k_ins] + novo_sp + novo_ind + pp[k_ins:]
        doc = doc[:a] + escrever_ppr(bl, pp) + doc[b:]
        n += 1
    return doc, n


def trocar_comeco(bloco, fim, novo):
    """Troca os `fim` primeiros caracteres do texto do paragrafo por `novo`.

    O texto se espalha por varias corridas, e o numero pode estar numa e o
    espaco noutra. O novo comeco entra na primeira `w:t`; das seguintes sai o
    que ainda faltava cortar. As corridas nao se fundem, e o italico ou o
    negrito de cada uma ficam onde estavam.
    """
    partes, pos, falta, primeiro = [], 0, fim, True
    for m in re.finditer(rb"(<w:t)((?:\s[^>]*)?)>(.*?)</w:t>", bloco, re.S):
        t = m.group(3).decode("utf-8")
        corte = min(falta, len(t))
        falta -= corte
        t = t[corte:]
        if primeiro:
            t = novo + t
            primeiro = False
        attrs = m.group(2)
        if t != t.strip() and b"xml:space" not in attrs:
            attrs += b' xml:space="preserve"'
        partes.append(bloco[pos:m.start()])
        partes.append(m.group(1) + attrs + b">" + t.encode("utf-8") + b"</w:t>")
        pos = m.end()
        if falta == 0:
            break
    partes.append(bloco[pos:])
    return b"".join(partes)


def numerar_titulo(bloco):
    """O numero do titulo com ponto depois, ou sem ele, conforme a norma.

    A Introducao e numerada como qualquer secao primaria: a NBR 14724:2024
    (5.2.3) e a NBR 6024:2012 (4.1, alinea h) nao a poem entre os titulos sem
    indicativo numerico.
    """
    regra = NORMA["numero_de_titulo"]
    cru = O.texto(bloco)
    m = NUMERO_CRU.match(cru)
    if not m:
        return bloco
    lead, num = m.group(1), m.group(2)
    ponto = "." if regra["ponto_depois"] else ""
    if regra["tabulacao"]:
        return trocar_trecho(bloco, 0, m.end(), [lead + num + ponto, TAB])
    novo = lead + num + ponto + " "
    if novo == cru[:m.end()]:
        return bloco
    return trocar_comeco(bloco, m.end(), novo)


def secao_da_dedicatoria(doc, modelo):
    """Poe a dedicatoria numa secao propria, alinhada ao pe da pagina.

    `modelo` e a secao do pre-textual ja escrita por `separar_secao`, com o
    cabecalho vazio. As duas secoes novas usam o mesmo cabecalho, e so a
    primeira secao do documento diz onde a contagem comeca.
    """
    pars = O.paragrafos(doc)
    todos = [k for k, (a, b) in enumerate(pars)
             if O.estilo_de(doc[a:b]) in PE_DA_PAGINA]
    if not todos or todos[0] == 0:
        return doc, "a dedicatoria nao foi achada"
    # Cada corrida de paragrafos seguidos e uma secao: a dedicatoria e uma, a
    # epigrafe e outra. De tras para a frente, para os deslocamentos valerem.
    # A corrida acaba tambem quando o estilo muda: tirado o titulo da epigrafe,
    # ela encosta na dedicatoria, e as duas iam parar na mesma folha.
    corridas = []
    for k in todos:
        if (corridas and k == corridas[-1][-1] + 1
                and O.estilo_de(doc[pars[k][0]:pars[k][1]])
                == O.estilo_de(doc[pars[k - 1][0]:pars[k - 1][1]])):
            corridas[-1].append(k)
        else:
            corridas.append([k])
    avisos = []
    for idx in reversed(corridas):
        doc, aviso = _secao_no_pe(doc, idx, modelo)
        if aviso:
            avisos.append(aviso)

    # So a primeira secao do documento diz onde a contagem comeca.
    todas = list(re.finditer(rb"<w:sectPr\b.*?</w:sectPr>", doc, re.S))
    for mm in reversed(todas[1:]):
        doc = (doc[:mm.start()]
               + re.sub(rb"<w:pgNumType[^>]*/>", b"", mm.group(0))
               + doc[mm.end():])
    return doc, "; ".join(avisos) or None


def _secao_no_pe(doc, idx, modelo):
    pars = O.paragrafos(doc)
    ant = idx[0] - 1
    a0, b0 = pars[ant]
    na_tabela = O.dentro_de_tabela(O.tabelas(doc))
    if na_tabela(a0):
        return doc, "o paragrafo antes da dedicatoria nao pode fechar secao"
    # O paragrafo de antes ja fecha secao quando e o fim de outra corrida no pe,
    # e o ultimo desta ja fecha quando a corrida seguinte, processada antes,
    # abriu a secao dela aqui: dedicatoria e epigrafe em folhas seguidas.
    ant_fecha = b"<w:sectPr" in O.ppr(doc[a0:b0])
    a1, b1 = pars[idx[-1]]

    base = re.sub(rb"<w:pgNumType[^>]*/>", b"", modelo)
    base = re.sub(rb"<w:vAlign[^>]*/>", b"", base)
    def antes_de(xml, seguintes, peca):
        """Poe `peca` na posicao que a ordem do esquema de w:sectPr pede."""
        m = re.search(rb"<w:(?:" + seguintes + rb")\b", xml)
        k = m.start() if m else xml.rindex(b"</w:sectPr>")
        return xml[:k] + peca + xml[k:]

    depois_de_valign = (rb"noEndnote|titlePg|textDirection|bidi|rtlGutter"
                        rb"|docGrid|printerSettings|sectPrChange")
    # A secao pode nao ter `w:cols`, e procurar so por ele deixaria a contagem
    # sem inicio.
    antes = antes_de(base, rb"cols|formProt|vAlign|" + depois_de_valign,
                     b'<w:pgNumType w:start="%d"/>' % NORMA["numeracao"]["inicio"])
    na_base = antes_de(base, depois_de_valign, b'<w:vAlign w:val="bottom"/>')

    # de tras para a frente, para que os deslocamentos continuem valendo
    pp1 = O.ppr(doc[a1:b1])
    ms = re.search(rb"<w:sectPr\b.*?</w:sectPr>", pp1, re.S)
    if ms:
        so = re.sub(rb"<w:vAlign[^>]*/>", b"", ms.group(0))
        so = antes_de(so, depois_de_valign, b'<w:vAlign w:val="bottom"/>')
        pp1 = pp1[:ms.start()] + so + pp1[ms.end():]
    else:
        pp1 = pp1 + na_base
    doc = doc[:a1] + escrever_ppr(doc[a1:b1], pp1) + doc[b1:]
    # A quebra de pagina do primeiro paragrafo era a reserva para o caso de a
    # secao nao poder ser criada. Com a secao, ela sai.
    a, b = O.paragrafos(doc)[idx[0]]
    doc = doc[:a] + escrever_ppr(doc[a:b], tirar(O.ppr(doc[a:b]),
                                                  b"w:pageBreakBefore")) + doc[b:]
    if not ant_fecha:
        a, b = O.paragrafos(doc)[ant]
        doc = doc[:a] + escrever_ppr(doc[a:b], O.ppr(doc[a:b]) + antes) + doc[b:]
    return doc, None


# ------------------------------------------------------ o que veio da Norma antiga
REVISAO_PENDENTE = re.compile(
    rb"<w:(ins|del|moveFrom|moveTo|pPrChange|rPrChange|sectPrChange|tblPrChange"
    rb"|trPrChange|tcPrChange|numberingChange)\b[^>]*w:author=\"([^\"]*)\"")


def alteracoes_pendentes(partes_xml):
    """Quem assina as alteracoes controladas do arquivo, e quantas de cada um."""
    from collections import Counter
    de = Counter()
    for xml in partes_xml:
        for m in REVISAO_PENDENTE.finditer(xml):
            de[m.group(2).decode("utf-8", "replace") or "(sem autor)"] += 1
    return de


def porta_das_alteracoes(de, nome):
    """Para se o arquivo trouxer alteracao controlada por decidir.

    Vem da Norma antiga: nenhum programa aceita ou
    recusa por conta as alteracoes de ninguem. Aqui ha mais uma razao: a copia
    com controle de alteracoes e feita pelo Word comparando o original com a
    copia formatada, e as marcas pendentes do original se misturariam as da
    formatacao, sem que se soubesse de quem e cada uma."""
    if not de:
        return
    total = sum(de.values())
    linhas = ["",
              "%s tem %d alteracao(oes) controlada(s) por decidir, e a formatacao"
              % (nome, total),
              "nao roda assim: as marcas de quem revisou se misturariam as da",
              "formatacao, e aceitar tudo depois apagaria o trabalho dessa pessoa.",
              "", "Quem assina:"]
    for autor, n in de.most_common():
        linhas.append("   %-28s %4d" % (autor[:28], n))
    linhas += ["", "Aceite ou recuse as alteracoes no Word, salve, e rode de novo.",
               "Comentario nao atrapalha e pode ficar."]
    sys.exit("\n".join(linhas))


def no_ppr(pp, elemento, seguintes):
    """Poe `elemento` no pPr antes do primeiro de `seguintes`, que e a ordem do
    esquema."""
    m = re.search(rb"<w:(?:" + seguintes + rb")\b", pp)
    k = m.start() if m else len(pp)
    return pp[:k] + elemento + pp[k:]


DEPOIS_DE_IND = (rb"contextualSpacing|mirrorIndents|suppressOverlap|jc|textDirection"
                 rb"|textAlignment|textboxTightWrap|outlineLvl|divId|cnfStyle|rPr"
                 rb"|sectPr|pPrChange")
DEPOIS_DE_SPACING = rb"ind|" + DEPOIS_DE_IND


def endireitar_separadores(notas_xml):
    """O filete das notas comeca na margem esquerda (NBR 14724, 5.2.1).

    Vem da Norma antiga. O separador e a linha curta que o Word desenha acima
    das notas, e mora em `footnotes.xml`, nas notas de tipo `separator` e
    `continuationSeparator`. Sem recuo proprio, ele herda o de primeira linha do
    Normal e comeca para dentro da margem. Junto costuma vir paragrafo vazio a mais, que abre um
    branco antes das notas, e ele sai."""
    n = [0]

    def um(m):
        bloco = m.group(0)
        pars = list(re.finditer(rb"<w:p(?:\s[^>]*)?/>|<w:p(?:\s[^>]*)?>.*?</w:p>",
                                bloco, re.S))
        for extra in reversed(pars[1:]):
            x = extra.group(0)
            if (not b"".join(re.findall(rb"<w:t(?:\s[^>]*)?>(.*?)</w:t>", x, re.S)).strip()
                    and b"<w:separator" not in x and b"<w:continuationSeparator" not in x):
                bloco = bloco[:extra.start()] + bloco[extra.end():]
                n[0] += 1

        def zera(mp):
            par = mp.group(0)
            pp = tirar(O.ppr(par), b"w:ind")
            pp = no_ppr(pp, b'<w:ind w:left="0" w:firstLine="0"/>', DEPOIS_DE_IND)
            novo = escrever_ppr(par, pp)
            if novo != par:
                n[0] += 1
            return novo
        return re.sub(rb"<w:p(?:\s[^>]*)?/>|<w:p(?:\s[^>]*)?>.*?</w:p>", zera, bloco,
                      flags=re.S)

    saida = re.sub(rb'<w:(footnote|endnote)\s[^>]*w:type="(?:separator|continuationSeparator)"'
                   rb"[^>]*>.*?</w:\1>", um, notas_xml, flags=re.S)
    return saida, n[0]


def juntar_espacos(doc, desde):
    """Espaco repetido vira um so, da parte textual em diante.

    Vem da Norma antiga, e so a partir de `desde`: na capa o alinhamento as
    vezes e feito com espaco repetido. A conferencia compara o texto com os
    espacos normalizados, e por isso isto nao a derruba."""
    n = [0]

    def troca_t(m):
        if m.start() < desde:
            return m.group(0)
        novo = re.sub(rb"  +", b" ", m.group(2))
        if novo == m.group(2):
            return m.group(0)
        n[0] += 1
        return m.group(1) + novo + b"</w:t>"

    return re.sub(rb"(<w:t(?:\s[^>]*)?>)(.*?)</w:t>", troca_t, doc, flags=re.S), n[0]


def parece_titulo(t):
    """Conservador, como na Norma antiga: abre com numeracao de secao, ou e curto
    e todo em maiuscula; e nao termina como frase."""
    if not t or len(t) > 130 or t.rstrip().endswith((".", ";", ",", ":")):
        return False
    letras = [c for c in t if c.isalpha()]
    return bool(O.NUMERO_DE_TITULO.match(t)) or (
        len(letras) >= 4 and t == t.upper())


ESTILO_DE_INDICE = re.compile(r"^(Sumrio|Sumario|TOC|toc|Contents)\d$|^ndicedeilustraes|"
                              r"^Indicedeilustracoes|^TableofFigures|^Tabeladeilustraes")


def faixas_de_indice(doc):
    """As faixas de bytes ocupadas por campo TOC: o sumario automatico e a lista
    de ilustracoes automatica. Os campos se aninham (cada entrada traz um
    PAGEREF), e por isso a varredura leva uma pilha."""
    faixas, pilha = [], []
    for m in re.finditer(rb'<w:fldChar[^>]*w:fldCharType="(begin|end)"|'
                         rb"<w:instrText[^>]*>([^<]*)</w:instrText>", doc):
        if m.group(1) == b"begin":
            pilha.append([m.start(), b""])
        elif m.group(1) == b"end":
            if pilha:
                ini, instr = pilha.pop()
                if instr.strip().startswith(b"TOC"):
                    faixas.append((ini, m.end()))
        elif pilha:
            pilha[-1][1] += m.group(2)
    return faixas


TIPO_NUMERACAO = ("http://schemas.openxmlformats.org/officeDocument/2006/"
                  "relationships/numbering")
CT_NUMERACAO = "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"
SEGUEM_NUMPR = (rb"suppressLineNumbers|pBdr|shd|tabs|suppressAutoHyphens|kinsoku|wordWrap"
                rb"|overflowPunct|topLinePunct|autoSpaceDE|autoSpaceDN|bidi|adjustRightInd"
                rb"|snapToGrid|spacing|ind|contextualSpacing|mirrorIndents|suppressOverlap"
                rb"|jc|textDirection|textAlignment|textboxTightWrap|outlineLvl|divId"
                rb"|cnfStyle|rPr|sectPr|pPrChange")


def sem_numpr(pp):
    return re.sub(rb"<w:numPr>.*?</w:numPr>", b"", pp, flags=re.S)


def com_numpr(pp, ilvl, num_id):
    """O pPr com a lista pedida, no lugar que o esquema quer."""
    return no_ppr(sem_numpr(pp), b'<w:numPr><w:ilvl w:val="%d"/><w:numId w:val="%d"/>'
                  b"</w:numPr>" % (ilvl, num_id), SEGUEM_NUMPR)


def ids_da_numeracao(numbering_xml):
    """Os primeiros identificadores livres de lista abstrata e de lista."""
    if not numbering_xml:
        return 1, 1
    a = [int(x) for x in re.findall(rb'w:abstractNumId="(\d+)"', numbering_xml)]
    n = [int(x) for x in re.findall(rb'<w:num w:numId="(\d+)"', numbering_xml)]
    return (max(a) + 1 if a else 1), (max(n) + 1 if n else 1)


def lista_de_titulos(abstrata, lista, parada_t1_tw, pendentes_tw, t1_no_centro=False):
    """A lista ligada aos estilos Titulo 1 a 4: numeros ate o 3, letra no 4.

    O numero vem sem ponto e seguido de tabulacao (NBR 6024, 4.1). O recuo de
    cada nivel e o do estilo, porque o da lista vence o do estilo: o recuo
    pendente que alinha a segunda linha sob a primeira letra do titulo, ou, com o
    Titulo 1 centralizado, recuo nenhum e a parada do vao.
    """
    textos = ["%1", "%1.%2", "%1.%2.%3", "%4)"]
    formatos = ["decimal", "decimal", "decimal", "lowerLetter"]
    niveis = []
    for k in range(9):
        if k < 4:
            if k == 0 and t1_no_centro:
                ppr = (b'<w:pPr><w:tabs><w:tab w:val="num" w:pos="%d"/></w:tabs>'
                       b'<w:ind w:left="0" w:firstLine="0"/></w:pPr>' % parada_t1_tw)
            else:
                x = pendentes_tw[k + 1]
                ppr = (b'<w:pPr><w:tabs><w:tab w:val="num" w:pos="%d"/></w:tabs>'
                       b'<w:ind w:left="%d" w:hanging="%d"/></w:pPr>' % (x, x, x))
            niveis.append(
                b'<w:lvl w:ilvl="%d"><w:start w:val="1"/><w:numFmt w:val="%s"/>'
                b'<w:pStyle w:val="Ttulo%d"/><w:lvlText w:val="%s"/>'
                b'<w:lvlJc w:val="left"/>%s</w:lvl>'
                % (k, formatos[k].encode(), k + 1, textos[k].encode(), ppr))
        else:
            niveis.append(b'<w:lvl w:ilvl="%d"><w:start w:val="1"/><w:numFmt w:val="none"/>'
                          b'<w:lvlText w:val=""/><w:lvlJc w:val="left"/></w:lvl>' % k)
    return (b'<w:abstractNum w:abstractNumId="%d"><w:multiLevelType w:val="multilevel"/>'
            % abstrata + b"".join(niveis) + b"</w:abstractNum>",
            b'<w:num w:numId="%d"><w:abstractNumId w:val="%d"/></w:num>' % (lista, abstrata))


def com_lista(numbering_xml, abstrata_xml, lista_xml):
    """numbering.xml com a lista nova: a abstrata antes de toda `w:num`, e a
    lista depois delas, como o esquema pede. Cria a parte se ela nao existe."""
    if not numbering_xml:
        return (b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
                b'<w:numbering xmlns:w="http://schemas.openxmlformats.org/'
                b'wordprocessingml/2006/main">' + abstrata_xml + lista_xml
                + b"</w:numbering>")
    i = numbering_xml.find(b"<w:num ")
    if i < 0:
        i = numbering_xml.find(b"<w:numIdMacAtCleanup")
    if i < 0:
        i = numbering_xml.rindex(b"</w:numbering>")
    x = numbering_xml[:i] + abstrata_xml + numbering_xml[i:]
    j = x.find(b"<w:numIdMacAtCleanup")
    if j < 0:
        j = x.rindex(b"</w:numbering>")
    return x[:j] + lista_xml + x[j:]


def planejar(doc, corpo_meio_ponto, linha_twips):
    pars = O.paragrafos(doc)
    tabs = O.tabelas(doc)
    em_tabela = O.dentro_de_tabela(tabs)
    indices = faixas_de_indice(doc)

    fichas = []
    for i, (a, b) in enumerate(pars):
        bloco = doc[a:b]
        t = O.texto(bloco).strip()
        num, nivel = O.numero_de_titulo(t)
        estilo = O.estilo_de(bloco)
        # Entrada de indice automatico nao e titulo, nem corpo: e resultado de
        # campo que o Word regera. Sem isto, as entradas numeradas do sumario
        # seriam lidas como titulos, com quebra de pagina e divisao no
        # dois-pontos.
        indice = bool(estilo and ESTILO_DE_INDICE.match(estilo)) or any(
            x < b and a < y for x, y in indices)
        m_est = re.match(r"(?:Ttulo|Titulo|Heading)(\d)$", estilo or "")
        fichas.append({"i": i, "ini": a, "fim": b, "bloco": bloco, "texto": t,
                       "tabela": em_tabela(a), "pPr": O.ppr(bloco),
                       "num": num, "nivel": 0 if indice else nivel,
                       "nivel_estilo": int(m_est.group(1)) if m_est else 0,
                       "indice": indice, "estilo": estilo})

    # Onde termina o pre-textual: no primeiro titulo de nivel 1, numerado ou em
    # estilo de Titulo 1, que nao seja titulo do pre-textual nem entrada de
    # indice. Ha trabalho que marca os capitulos so por estilo, sem numero.
    inicio_textual = None
    for f in fichas:
        if f["tabela"] or f["indice"] or not f["texto"] or len(f["texto"]) >= 200:
            continue
        if TITULO_PRE.match(f["texto"]):
            continue
        if f["nivel"] == 1 or f["nivel_estilo"] == 1:
            inicio_textual = f["i"]
            break
    if inicio_textual is None:
        sys.exit("nao achei onde comeca a parte textual (nenhum titulo de nivel 1, "
                 "numerado ou em estilo de titulo, fora de tabela e de indice). "
                 "O programa para em vez de adivinhar.")

    # A tabela que guarda o sumario digitado: a que fica no pre-textual e vem
    # depois do paragrafo que diz SUMARIO.
    sumario_par = next((f for f in fichas
                        if f["i"] < inicio_textual and not f["tabela"]
                        and re.match(r"(SUM[ÁA]RIO|CONTENTS|TABLE OF CONTENTS)",
                                     f["texto"].upper())), None)
    tabela_sumario = None
    if sumario_par:
        for a, b in tabs:
            # `>=` e nao `>`: quando nao ha paragrafo vazio entre o titulo
            # SUMARIO e a tabela, ela comeca exatamente onde ele termina, e a
            # comparacao estrita a perderia sem avisar.
            if a >= sumario_par["fim"] and a < pars[inicio_textual][0]:
                tabela_sumario = (a, b)
                break

    # Figura e o paragrafo que so tem desenho. A exigencia de nao ter texto e o
    # que separa a figura do paragrafo que apenas carrega uma ancora, como os da
    # capa, da dedicatoria e dos titulos do pre-textual.
    for f in fichas:
        f["figura"] = bool(re.search(rb"<w:drawing|<w:pict|<w:object", f["bloco"])
                           and not f["texto"])

    # A faixa das referencias: do titulo REFERENCIAS ate o proximo titulo de
    # primeira ordem.
    faixa_ref = None
    for f in fichas:
        if f["i"] < inicio_textual or f["tabela"]:
            continue
        # O titulo pode vir numerado, `5 REFERENCIAS`, como na Norma antiga.
        if (re.match(r"^\s*(?:\d+[.)]?\s+)?(REFER[ÊE]NCIAS?|BIBLIOGRAFIA|REFERENCES"
                     r"|BIBLIOGRAPHY)\b",
                     f["texto"], re.I) and len(f["texto"]) < 60):
            fim = len(fichas)
            for g in fichas[f["i"] + 1:]:
                if g["tabela"]:
                    continue
                if ((g["nivel"] == 1 and len(g["texto"]) < 200)
                        or (ESTRUTURAL_TEXTUAL.match(g["texto"])
                            and len(g["texto"]) < 90)):
                    fim = g["i"]
                    break
            faixa_ref = (f["i"] + 1, fim)
            break

    def recuado(f):
        ind = O.prop(f["pPr"], b"w:ind")
        if not ind:
            return False
        v = O.atributo(ind, b"<w:ind", b"w:left")
        try:
            return v is not None and int(v) >= NORMA["limiar_recuo_tw"]
        except ValueError:
            return False

    # As faixas de lista do pre-textual: do titulo LISTA DE ... ate o proximo
    # titulo do pre-textual.
    LISTA = LISTA_DE
    TIT_PRE = TITULO_PRE
    em_lista = set()
    dentro = False
    for f in fichas:
        if f["i"] >= inicio_textual or f["tabela"]:
            break
        t = f["texto"]
        if t and len(t) < 90 and TIT_PRE.match(t):
            dentro = bool(LISTA.match(t))
            continue
        if dentro and t:
            # subtitulo da lista, como `Graficos` e `Tabelas`, nao e entrada
            if len(t) <= 25 and not re.search(r"\s[–—-]\s", t):
                continue
            em_lista.add(f["i"])

    papeis = {}
    for f in fichas:
        i = f["i"]
        if tabela_sumario and tabela_sumario[0] <= f["ini"] < tabela_sumario[1]:
            papeis[i] = "sumario-velho"
        elif f["indice"]:
            papeis[i] = "indice"
        elif f["tabela"]:
            papeis[i] = "tabela"
        elif i in em_lista:
            papeis[i] = "lista"
        elif i < inicio_textual:
            papeis[i] = "pretextual"
        elif f["figura"]:
            papeis[i] = "figura"
        # Apendice e anexo tem o destaque da secao primaria, centralizados e sem
        # numero (NBR 14724, 4.2.3.3; NBR 6024, 4.1, alinea h), qualquer que seja
        # o estilo que o autor lhes deu.
        elif ESTRUTURAL_SUB.match(f["texto"]) and len(f["texto"]) < 120:
            papeis[i] = "titulo1"
        elif f["nivel"] in (1, 2, 3, 4) and len(f["texto"]) < 200:
            papeis[i] = "titulo%d" % f["nivel"]
        elif f["nivel_estilo"] in (1, 2, 3, 4) and f["texto"] and len(f["texto"]) < 200:
            # Titulo sem numero, marcado so pelo estilo.
            papeis[i] = "titulo%d" % f["nivel_estilo"]
        elif ESTRUTURAL_TEXTUAL.match(f["texto"]) and len(f["texto"]) < 90:
            papeis[i] = "titulo1"
        elif LEGENDA.match(f["texto"]):
            papeis[i] = "legenda"
        elif (FONTE_APOS_FIGURA.match(f["texto"]) and len(f["texto"]) < 400
              and depois_de_ilustracao(fichas, i)):
            papeis[i] = "legenda"
        elif faixa_ref and faixa_ref[0] <= i < faixa_ref[1] and f["texto"]:
            papeis[i] = "referencia"
        elif faixa_ref and faixa_ref[0] <= i < faixa_ref[1] and not f["tabela"]:
            # A linha em branco entre referencias (NBR 6023) vem do espaco
            # depois do estilo, e o vazio que o autor pos para isso sai.
            papeis[i] = "vazio-referencia"
        elif recuado(f):
            papeis[i] = "citacao"
        elif f["texto"]:
            papeis[i] = "corpo"
        else:
            papeis[i] = "vazio"
    return fichas, papeis, inicio_textual, tabela_sumario, tabs


def vazios_antes_de_quebra(doc, fichas, papeis, com_quebra):
    """Os vazios que sobram imediatamente antes de um titulo que quebra pagina.

    POR QUE ISTO E A UNICA EXCECAO A NAO APAGAR VAZIO

    O trabalho separa secao de secao com uma corrida de paragrafos vazios, que
    empurra a proxima secao para a folha seguinte. Posta a quebra de pagina no
    estilo do titulo, essa corrida deixa de empurrar e passa a ocupar uma folha
    inteira sozinha, em branco.

    Apaga-se so o que esta entre o fim do texto anterior e o titulo, so quando o
    titulo carrega quebra de pagina, e so o vazio que nao carrega nada: quebra de
    secao, quebra de pagina, ancora de imagem ou de nota, marcador, comentario,
    campo ou marcacao de revisao ficam onde estao.
    """
    apagar = []
    for f in fichas:
        if papeis[f["i"]] not in com_quebra:
            continue
        k = f["i"] - 1
        while k >= 0:
            g = fichas[k]
            if g["tabela"] or not O.vazio(g["bloco"]):
                break
            # Vazio que carrega ancora, quebra ou marcacao fica onde esta, e a
            # varredura NAO para nele: se parasse, ele protegeria as linhas em
            # branco atras de si, que sobrariam como folhas.
            # A quebra de pagina do vazio repete a do titulo, e nao o protege.
            if not O.travado(g["bloco"], aceitar_quebra=True):
                apagar.append(k)
            k -= 1
    return sorted(set(apagar))

# ------------------------------------------------------------------ campo
def campo_de_sumario():
    return (b'<w:p><w:pPr><w:pStyle w:val="SumarioAutomatico"/></w:pPr>'
            b'<w:r><w:fldChar w:fldCharType="begin" w:dirty="true"/></w:r>'
            b'<w:r><w:instrText xml:space="preserve"> TOC \\o "1-2" \\h \\z \\u '
            b"</w:instrText></w:r>"
            b'<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
            b"<w:r><w:t>Sumario automatico: no Word, clique aqui e tecle F9 para "
            b"atualizar.</w:t></w:r>"
            b'<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>')


# ----------------------------------------------------------------- aplicar
def formatar(origem, destino, corpo_pt=13.0, entrelinha=1.3, fonte="Times New Roman"):
    partes = O.ler(origem)
    with zipfile.ZipFile(origem) as _z:
        _xmls = [_z.read(n) for n in _z.namelist()
                 if re.match(r"word/(document|footnotes|endnotes|header\d*|footer\d*)\.xml$", n)]
    porta_das_alteracoes(alteracoes_pendentes(_xmls), Path(origem).name)
    doc = partes["word/document.xml"]
    est = partes["word/styles.xml"]
    # ---- 0. o estilo de texto colado da web passa ao Normal
    doc, est, n_mso, ids_mso = sem_msonormal(doc, est)
    # Guardado antes de qualquer reescrita: e dele que se le o que um paragrafo
    # herdava, quando se vai tirar dele o estilo que herdava.
    est_original = est
    with zipfile.ZipFile(origem) as _z:
        rels = _z.read("word/_rels/document.xml.rels")

    sz = int(round(corpo_pt * O.MEIO_PONTO))
    linha = int(round(240 * entrelinha))
    fichas, papeis, inicio_textual, tabela_sumario, tabs = planejar(doc, sz, linha)

    # ---- a numeracao dos titulos: digitada vira automatica
    #
    # So quando o trabalho numera os titulos digitando o numero. Quem ja numera
    # pelo Word fica com a numeracao dele.
    regra_num = NORMA["numero_de_titulo"]
    defs_num = O.estilos_definidos(est)

    def numera_pelo_word(f):
        """O titulo ja tem numero de lista, no paragrafo ou no estilo."""
        pp_ = O.ppr(f["bloco"])
        if b"<w:numPr>" in pp_:
            return b'<w:numId w:val="0"/>' not in pp_
        return b"<w:numPr>" in defs_num.get(O.estilo_de(f["bloco"]) or "", {}).get("xml", b"")

    titulos_ = [f for f in fichas if papeis[f["i"]].startswith("titulo")]
    # Basta um titulo numerado pelo Word para o trabalho ficar com a numeracao
    # dele: trocar a lista do autor apagaria os numeros de todos os outros.
    numerar_auto = (regra_num["automatica"]
                    and any(NUMERO_CRU.match(O.texto(f["bloco"])) for f in titulos_)
                    and not any(numera_pelo_word(f) for f in titulos_))
    abs_id, num_id = ids_da_numeracao(partes.get("word/numbering.xml"))
    # A lista renumera a partir de 1, nivel a nivel. Onde o numero digitado nao
    # e o que ela vai dar, as remissoes do texto a ele ficam erradas, e o
    # relatorio aponta.
    avisos_numero = []
    if numerar_auto:
        cont = [0, 0, 0, 0]
        for f in fichas:
            pap = papeis[f["i"]]
            m_n = NUMERO_CRU.match(O.texto(f["bloco"])) if pap.startswith("titulo") else None
            if not m_n:
                continue
            nv = int(pap[len("titulo"):])
            cont[nv - 1] += 1
            for k in range(nv, 4):
                cont[k] = 0
            if nv == 4:
                # O quarto nivel passa de numero a letra: remissao a ele, que o
                # texto faca pelo numero, fica errada sempre.
                dado = "%s)" % chr(ord("a") + (cont[3] - 1) % 26)
            else:
                dado = ".".join(str(x) for x in cont[:nv])
            if dado != m_n.group(2):
                avisos_numero.append((f["texto"][:60], m_n.group(2), dado))

    conta = {k: 0 for k in ("pretextual", "tabela", "titulo1", "titulo2",
                            "legenda", "corpo", "vazio", "sumario-velho")}
    for p in papeis.values():
        conta[p] = conta.get(p, 0) + 1

    # ---- 1. os estilos que blindam quem herda, criados ANTES de mexer no Normal
    #
    # A entrelinha e o corpo que se devolvem sao os que o Normal tem HOJE, lidos
    # dele: o que se promete a quem herda e que nada mude, e para isso e preciso
    # saber o que ele tinha.
    normal_velho = bloco_do_estilo(est_original, "Normal")
    m = (re.search(rb"<w:spacing(?:\s[^>]*)?/>", normal_velho.group(0))
         if normal_velho else None)
    simples = m.group(0) if m else b'<w:spacing w:after="0" w:line="240" w:lineRule="auto"/>'
    sz_velho = (O.atributo(normal_velho.group(0), b"<w:sz", b"w:val")
                if normal_velho else None) or "24"
    herdado = (b'<w:sz w:val="%s"/><w:szCs w:val="%s"/>'
               % (sz_velho.encode(), sz_velho.encode()))

    # O Normal passa a declarar recuo de primeira linha e justificacao, entao
    # quem nao quiser os dois tem de dize-lo. `firstLine="0"` e seguro para quem
    # declara `hanging`: o `hanging` do paragrafo vence o `firstLine` do estilo,
    # e a linha sai no mesmo lugar com ou sem ele.
    sem_recuo = b'<w:ind w:firstLine="0"/>'
    esquerda = b'<w:jc w:val="left"/>'
    centro = b'<w:jc w:val="center"/>'
    sem_hifen = b"<w:suppressAutoHyphens/>"

    leg, cit = NORMA["legenda"], NORMA["citacao"]
    # O pre-textual leva o corpo do texto e entrelinha propria: ali nao ha prosa
    # corrida que peca respiro, ha nome, titulo, lista e resumo.
    ent_pre = NORMA["pretextual"]["entrelinha"]
    # `after=0` nao e detalhe: sem ele o pre-textual herda os 6 pt do Normal, e a
    # conta que posiciona os blocos da capa erra por esse espaco em cada
    # paragrafo.
    entrelinha_pre = (b'<w:spacing w:after="0" w:line="%d" w:lineRule="auto"/>'
                      % int(round(240 * ent_pre)))
    criados = []
    for sid, nome, ppr, rpr in (
        # O pre-textual segue o corpo e a entrelinha do texto, por decisao da
        # norma, e por isso NAO recebe de volta o que o Normal tinha antes. O que
        # ele ainda declara e a ausencia de recuo de primeira linha, que na capa
        # e nas listas seria erro, e o alinhamento a esquerda, que os paragrafos
        # centralizados sobrescrevem por conta propria.
        ("PreTextual", "Pre-textual",
         b"<w:pPr>" + entrelinha_pre + sem_recuo + esquerda + b"</w:pPr>",
         None) if NORMA["pretextual"]["segue_o_normal"] else
        ("PreTextual", "Pre-textual",
         b"<w:pPr>" + simples + sem_recuo + esquerda + b"</w:pPr>",
         b"<w:rPr>" + herdado + b"</w:rPr>"),
        ("TabelaTrabalho", "Tabela do trabalho",
         b"<w:pPr>" + simples + sem_recuo + esquerda + b"</w:pPr>",
         b"<w:rPr>" + herdado + b"</w:rPr>"),
        ("FiguraGrafico", "Figura ou grafico",
         b"<w:pPr><w:keepNext/>"
         + b'<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
         + sem_recuo + centro + sem_hifen + b"</w:pPr>", None),
        # Subtitulo de capitulo: o que vinha depois do dois-pontos. Fica fora do
        # sumario (`outlineLvl` 9) e sem negrito, para nao disputar com o titulo.
        ("SubtituloCapitulo", "Subtitulo de capitulo",
         b"<w:pPr><w:keepNext/><w:keepLines/>"
         + b'<w:spacing w:before="0" w:after="%d" w:line="%d" w:lineRule="auto"/>'
         % (pt(NORMA["titulo1"]["depois_pt"]),
            int(round(240 * NORMA["titulos_entrelinha"])))
         # Na coluna do titulo, quando ele esta a esquerda.
         + (b'<w:ind w:firstLine="0"/><w:jc w:val="center"/>'
            if NORMA["titulo1"]["centralizado"] else
            b'<w:ind w:left="%d" w:firstLine="0"/><w:jc w:val="left"/>'
            % tw(NORMA["numero_de_titulo"]["pendente_cm"][1]))
         + b'<w:suppressAutoHyphens/><w:outlineLvl w:val="9"/></w:pPr>', None),
        # Entrada de lista do pre-textual, das ilustracoes e das siglas. O recuo
        # pendente alinha a continuacao sob o texto da primeira linha, em vez de
        # deixa-la voltar a margem e colar na sigla da entrada seguinte.
        ("ListaPreTextual", "Entrada de lista",
         b"<w:pPr>"
         + b'<w:spacing w:after="%d" w:line="%d" w:lineRule="auto"/>'
         % (pt(NORMA["normal_depois_pt"]),
            int(round(240 * NORMA["pretextual"]["entrelinha"])))
         + b'<w:ind w:left="%d" w:hanging="%d"/>'
         % (tw(NORMA["pretextual"]["recuo_da_lista_cm"]),
            tw(NORMA["pretextual"]["recuo_da_lista_cm"]))
         + b'<w:jc w:val="left"/></w:pPr>', None),
        ("Referencia", "Referencia bibliografica",
         b"<w:pPr>" + sem_recuo + b"</w:pPr>", None),
        # O resumo e um bloco de prosa continua lido como unidade, e a NBR 14724
        # poe em simples a familia a que ele pertence: natureza do trabalho,
        # notas, referencias e legendas.
        ("ResumoTrabalho", "Resumo",
         b"<w:pPr>" + (b'<w:spacing w:after="0" w:line="%d" w:lineRule="auto"/>'
                      % int(round(240 * NORMA["pretextual"]["resumo_entrelinha"])))
         + sem_recuo + b'<w:jc w:val="both"/></w:pPr>', None),
        # O texto corrido do pre-textual (os agradecimentos) e prosa como a do
        # corpo: justificado e com o recuo de primeira linha dele. So a
        # entrelinha e a do pre-textual.
        ("TextoPreTextual", "Texto do pre-textual",
         b"<w:pPr>" + entrelinha_pre
         + b'<w:ind w:firstLine="%d"/><w:jc w:val="both"/></w:pPr>'
         % tw(NORMA["recuo_primeira_cm"]), None),
        ("SumarioAutomatico", "Sumario automatico",
         b"<w:pPr>" + simples + sem_recuo + b"</w:pPr>", None),
    ):
        est, feito = criar_estilo(est, sid, nome, ppr, rpr)
        if feito:
            criados.append(sid)

    # ---- todo estilo que os paragrafos protegidos usam tambem tem de se defender
    #
    # Quando a norma manda o pre-textual seguir o formato geral, o estilo usado
    # SO por ele deixa de ser blindado: ao contrario, ele solta o corpo que
    # declarava, que seguraria a folha de rosto num corpo diferente do resto.
    so_pretextual, em_tabela = set(), set()
    for f in fichas:
        if not f["estilo"]:
            continue
        (em_tabela if papeis[f["i"]] == "tabela" else so_pretextual).add(f["estilo"])
    so_pretextual -= em_tabela
    if not NORMA["pretextual"]["segue_o_normal"]:
        em_tabela |= so_pretextual
        so_pretextual = set()

    soltos = []
    for sid in sorted(so_pretextual):
        m = bloco_do_estilo(est, sid)
        if not m or b"<w:sz " not in m.group(0):
            continue
        limpo = re.sub(rb"<w:szCs(?:\s[^>]*)?/>", b"",
                       re.sub(rb"<w:sz(?:\s[^>]*)?/>", b"", m.group(0)))
        est = est[:m.start()] + limpo + est[m.end():]
        soltos.append(sid)

    blindados = []
    for sid in sorted(em_tabela) + ["Textodenotaderodap", "Rodap",
                                    "Textodecomentrio", "Textodebalo"]:
        est, feito = blindar_estilo(est, sid, simples, herdado, sem_recuo)
        if feito and sid not in blindados:
            blindados.append(sid)

    # O cabecalho corrente vai menor que o corpo, e o corpo tem de vir do estilo:
    # o resultado do campo STYLEREF e escrito pelo Word e segue o estilo do
    # paragrafo, nao o que estiver nas corridas do campo.
    sz_cab = int(NORMA["cabecalho"]["corpo_pt"] * 2)
    # Arquivo cujo cabecalho nunca usou o estilo nao o tem, e reescrever o que
    # nao existe nao faz nada, e o corpo nao chegaria ao cabecalho.
    est, feito = criar_estilo(est, "Cabealho", "header", None, None, embutido=True)
    if feito:
        criados.append("Cabealho")
    est, _ = reescrever_estilo(
        est, "Cabealho",
        ppr=b"<w:pPr>" + simples + sem_recuo + b"</w:pPr>",
        rpr=((b'<w:rPr><w:i/><w:iCs/>' if NORMA["cabecalho"]["italico"] else b"<w:rPr>")
             + b'<w:sz w:val="%d"/><w:szCs w:val="%d"/></w:rPr>' % (sz_cab, sz_cab)))

    # ---- 2. o Normal, que e a norma
    est, achou_normal = reescrever_estilo(
        est, "Normal",
        ppr=(b'<w:pPr><w:spacing w:after="%d" w:line="%d" w:lineRule="auto"/>'
             b'<w:ind w:firstLine="%d"/><w:jc w:val="both"/></w:pPr>'
             % (pt(NORMA["normal_depois_pt"]), linha,
                tw(NORMA["recuo_primeira_cm"]))),
        rpr=(b'<w:rPr><w:rFonts w:ascii="%s" w:eastAsia="%s" w:hAnsi="%s" '
             b'w:cs="%s"/><w:sz w:val="%d"/><w:szCs w:val="%d"/></w:rPr>'
             % (fonte.encode(), fonte.encode(), fonte.encode(), fonte.encode(),
                sz, sz)))
    if not achou_normal:
        sys.exit("este .docx nao tem estilo Normal; o programa para.")

    # ---- 2b. os estilos do sumario
    #
    # Os estilos de sumario costumam vir com forma diferente por nivel, e com
    # `autoRedefine`, que autoriza o Word a reescrever o estilo a partir de
    # qualquer paragrafo que o use.
    #
    # O DESENHO, E O QUE ELE RESOLVE
    #
    # Titulo de capitulo pode ser longo e ocupar varias linhas do sumario. Sem
    # recuo pendente, a segunda
    # linha volta a margem e o pontilhado com o numero fica solto embaixo, de
    # modo que o olho perde a qual entrada o numero pertence. Com recuo pendente,
    # a continuacao alinha sob o texto da primeira linha e o numero fica na
    # coluna da direita, sozinho.
    #
    # O nivel 1 vai em negrito e com espaco antes, que e o que deixa a estrutura
    # de capitulos visivel numa lista de vinte e tres entradas.
    largura_util = 9071
    sec = O.secoes(doc)
    if sec and sec[0].get("larg"):
        x = sec[0]
        largura_util = int(round((x["larg"] - NORMA["margens_cm"]["esq"]
                                  - NORMA["margens_cm"]["dir"]) * O.TWIPS_POR_CM))
    # A dedicatoria: sem recuo de primeira linha, e o texto comecando no meio da
    # mancha, que na folha A4 com margens iguais e o meio da folha.
    est, feito = criar_estilo(
        est, ESTILO_DEDICATORIA, "Dedicatoria",
        b'<w:pPr><w:ind w:left="%d" w:firstLine="0"/></w:pPr>'
        % int(round(largura_util * NORMA["pretextual"]["dedicatoria_recuo_fracao"])),
        None, base=b"PreTextual")
    if feito:
        criados.append(ESTILO_DEDICATORIA)

    sumarios = []
    for nivel in range(1, 4):
        sid = "Sumrio%d" % nivel
        if not bloco_do_estilo(est, sid):
            # Arquivo que nunca gerou sumario nao tem esses estilos. Cria-se,
            # senao o Word os inventa na primeira atualizacao, com a forma dele.
            est, feito = criar_estilo(est, sid, "toc %d" % nivel, None, None)
            if not feito:
                continue
            criados.append(sid)
        d = NORMA["sumario"].get("nivel%d" % nivel, NORMA["sumario"]["nivel2"])
        recuo_tw = tw(d["recuo_cm"])
        pend_tw = tw(d["pendente_cm"])
        ppr = (b'<w:pPr><w:tabs><w:tab w:val="right" w:leader="dot" '
               b'w:pos="%d"/></w:tabs>' % largura_util
               + b'<w:spacing w:before="%d" w:after="%d" w:line="240" '
                 b'w:lineRule="auto"/>' % (pt(d["antes_pt"]), pt(d["depois_pt"]))
               + b'<w:ind w:left="%d" w:hanging="%d"/>'
                 % (recuo_tw + pend_tw, pend_tw)
               + esquerda + sem_hifen + b"</w:pPr>")
        rpr = (b'<w:rPr><w:rFonts w:ascii="%s" w:eastAsia="%s" w:hAnsi="%s" '
               b'w:cs="%s"/>' % (fonte.encode(), fonte.encode(), fonte.encode(),
                                 fonte.encode())
               + (b"<w:b/><w:bCs/>" if d["negrito"] else b"")
               + b'<w:sz w:val="%d"/><w:szCs w:val="%d"/></w:rPr>' % (sz, sz))
        est, _ = reescrever_estilo(est, sid, ppr=ppr, rpr=rpr)
        m = bloco_do_estilo(est, sid)
        est = (est[:m.start()] + m.group(0).replace(b"<w:autoRedefine/>", b"")
               + est[m.end():])
        sumarios.append(sid)

    # ---- 3. os estilos de titulo
    t1, t2 = NORMA["titulo1"], NORMA["titulo2"]
    sz_t1 = int(t1["corpo_pt"] * 2)
    jc_t1 = centro if t1["centralizado"] else esquerda
    linha_tit = int(round(240 * NORMA["titulos_entrelinha"]))

    def ppr_titulo(antes, depois, nivel_topico, jc, com_quebra, tabs=b"",
                   ind=sem_recuo, lista=b""):
        # Na ordem que o esquema pede: keepNext, keepLines, pageBreakBefore,
        # numPr, tabs, suppressAutoHyphens, spacing, ind, jc, outlineLvl.
        return (b"<w:pPr><w:keepNext/><w:keepLines/>"
                + (b"<w:pageBreakBefore/>" if com_quebra else b"")
                + lista + tabs + sem_hifen
                + b'<w:spacing w:before="%d" w:after="%d" w:line="%d" '
                  b'w:lineRule="auto"/>' % (antes, depois, linha_tit)
                + ind + jc
                + b'<w:outlineLvl w:val="%d"/></w:pPr>' % nivel_topico)

    nt = NORMA["numero_de_titulo"]
    # Titulo 1 centralizado: a parada da o vao entre numero e titulo. A esquerda,
    # a forma e a dos outros niveis, e vem de `tabs_pendente` e `ind_pendente`.
    tabs_t1 = (b'<w:tabs><w:tab w:val="left" w:pos="%d"/></w:tabs>'
               % tw(nt["parada_t1_cm"])) if nt["tabulacao"] else b""

    def tabs_pendente(nivel):
        # A parada implicita do recuo pendente nao basta: o paragrafo que traz
        # `numPr` com `numId` 0, resto de uma lista que o autor desligou, perde
        # essa parada, e a tabulacao salta para a parada padrao de 1,25 cm. A
        # parada explicita vale com ou sem lista.
        if not nt["tabulacao"] or nivel not in nt["pendente_cm"]:
            return b""
        return (b'<w:tabs><w:tab w:val="left" w:pos="%d"/></w:tabs>'
                % tw(nt["pendente_cm"][nivel]))

    def ind_pendente(nivel):
        # Titulo alinhado a esquerda: o recuo pendente e a parada implicita da
        # tabulacao, e a segunda linha alinha sob a primeira letra do titulo.
        if not nt["tabulacao"] or nivel not in nt["pendente_cm"]:
            return sem_recuo
        x = tw(nt["pendente_cm"][nivel])
        return b'<w:ind w:left="%d" w:hanging="%d"/>' % (x, x)

    def rpr_titulo(sz_local):
        return (b'<w:rPr><w:rFonts w:ascii="%s" w:eastAsia="%s" w:hAnsi="%s" '
                b'w:cs="%s"/><w:b/><w:bCs/><w:color w:val="auto"/>'
                b'<w:sz w:val="%d"/><w:szCs w:val="%d"/></w:rPr>'
                % (fonte.encode(), fonte.encode(), fonte.encode(),
                   fonte.encode(), sz_local, sz_local))

    def lista_do_nivel(nivel):
        if not numerar_auto:
            return b""
        return (b'<w:numPr><w:ilvl w:val="%d"/><w:numId w:val="%d"/></w:numPr>'
                % (nivel - 1, num_id))

    if t1["centralizado"]:
        ppr_t1 = ppr_titulo(pt(t1["antes_pt"]), pt(t1["depois_pt"]), 0, jc_t1,
                            t1["quebra_pagina"], tabs=tabs_t1, lista=lista_do_nivel(1))
    else:
        ppr_t1 = ppr_titulo(pt(t1["antes_pt"]), pt(t1["depois_pt"]), 0, jc_t1,
                            t1["quebra_pagina"], tabs=tabs_pendente(1),
                            ind=ind_pendente(1), lista=lista_do_nivel(1))
    # O titulo do pre-textual nao tem numero, e fica no centro (NBR 14724, 5.2.3).
    ppr_t0 = ppr_titulo(pt(t1["antes_pt"]), pt(t1["depois_pt"]), 9, centro,
                        t1["quebra_pagina"])

    est, _ = reescrever_estilo(est, "Ttulo1", ppr=ppr_t1, rpr=rpr_titulo(sz_t1))
    est, _ = reescrever_estilo(
        est, "Ttulo2",
        ppr=ppr_titulo(pt(t2["antes_pt"]), pt(t2["depois_pt"]), 1, esquerda, False,
                       tabs=tabs_pendente(2), ind=ind_pendente(2),
                       lista=lista_do_nivel(2)),
        rpr=rpr_titulo(sz))

    # Titulo 0: a forma do Titulo 1, e fora do sumario. `outlineLvl` 9 e o valor
    # que o esquema reserva para corpo de texto, e e o que tira o paragrafo do
    # campo de sumario sem lhe tirar a aparencia.
    est, feito = criar_estilo(est, "Ttulo0", "Titulo 0", ppr_t0, rpr_titulo(sz_t1))
    if feito:
        criados.append("Ttulo0")
    else:
        est, _ = reescrever_estilo(est, "Ttulo0", ppr=ppr_t0, rpr=rpr_titulo(sz_t1))

    # O cabecalho do sumario e o titulo que o Word aplica a palavra SUMARIO.
    est, _ = reescrever_estilo(est, "CabealhodoSumrio", ppr=ppr_t0,
                               rpr=rpr_titulo(sz_t1))

    # O Titulo 3 tem de existir, porque o papel de titulo de terceiro nivel o
    # aplica. Arquivo que nunca o usou nao o tem.
    est, feito = criar_estilo(
        est, "Ttulo3", "heading 3",
        b'<w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="%d" w:after="%d"/>'
        b'<w:outlineLvl w:val="2"/></w:pPr>' % (pt(12), pt(6)),
        b"<w:rPr><w:b/><w:bCs/></w:rPr>", embutido=True)
    if feito:
        criados.append("Ttulo3")
    est, feito = criar_estilo(
        est, "Ttulo4", "heading 4",
        b'<w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="%d" w:after="%d"/>'
        b'<w:outlineLvl w:val="3"/></w:pPr>' % (pt(12), pt(6)),
        b"<w:rPr><w:b/><w:bCs/><w:i/><w:iCs/></w:rPr>", embutido=True)
    if feito:
        criados.append("Ttulo4")

    # Os demais niveis de titulo: sem recuo de primeira linha e sem hifenizacao,
    # que e o que a norma pede de todo titulo.
    for nivel in range(3, 10):
        sid = "Ttulo%d" % nivel
        if not bloco_do_estilo(est, sid):
            continue
        b0 = bloco_do_estilo(est, sid).group(0)
        pp = re.search(rb"<w:pPr>(.*?)</w:pPr>", b0, re.S)
        dentro = pp.group(1) if pp else b""
        jc_o = re.search(rb"<w:jc [^>]*/>", dentro)
        ol = re.search(rb"<w:outlineLvl [^>]*/>", dentro)
        dentro = tirar(dentro, b"w:ind", b"w:jc", b"w:outlineLvl", b"w:suppressAutoHyphens")
        dentro = re.sub(rb"<w:tabs>.*?</w:tabs>", b"", dentro, flags=re.S)
        # A entrelinha dos titulos; o que o estilo tinha antes e depois fica.
        m_sp = re.search(rb"<w:spacing\b[^>]*/>", dentro)
        sp = re.sub(rb'\s+w:(?:line|lineRule)="[^"]*"', b"", m_sp.group(0)) if m_sp \
            else b"<w:spacing/>"
        sp = sp[:-2].rstrip() + b' w:line="%d" w:lineRule="auto"/>' % linha_tit
        dentro = (dentro[:m_sp.start()] + sp + dentro[m_sp.end():]) if m_sp \
            else no_ppr(dentro, sp, DEPOIS_DE_SPACING)
        if numerar_auto and nivel <= 4:
            dentro = com_numpr(dentro, nivel - 1, num_id)
        # Cada peca no lugar que o esquema pede: tabulacao fora de ordem, antes
        # de `keepNext`, o Word ignora.
        dentro = no_ppr(dentro, tabs_pendente(nivel) + sem_hifen,
                        rb"kinsoku|wordWrap|overflowPunct|topLinePunct|autoSpaceDE"
                        rb"|autoSpaceDN|bidi|adjustRightInd|snapToGrid|" + DEPOIS_DE_SPACING
                        + rb"|spacing")
        dentro = no_ppr(dentro, ind_pendente(nivel), DEPOIS_DE_IND)
        dentro = no_ppr(dentro, (jc_o.group(0) if jc_o else b"")
                        + (ol.group(0) if ol else b""),
                        rb"textDirection|textAlignment|textboxTightWrap|divId|cnfStyle"
                        rb"|rPr|sectPr|pPrChange")
        est, _ = reescrever_estilo(est, sid, ppr=b"<w:pPr>" + dentro + b"</w:pPr>")

    # ---- 3a. a legenda, as notas, as referencias e a lista de ilustracoes
    #
    # A legenda e o estilo que o Word chama `caption`, cujo styleId muda com a
    # lingua (`Legenda` em portugues). Ele e reescrito por inteiro: o do Word traz
    # italico e cor, e a norma quer o texto em preto (NBR 14724, 5.1).
    id_legenda = next((sid for sid, x in O.estilos_definidos(est).items()
                       if x["nome"] == "caption"), None)
    if id_legenda is None:
        id_legenda = "Legenda"
        est, feito = criar_estilo(est, id_legenda, "caption", None, None,
                                  embutido=True)
        if feito:
            criados.append(id_legenda)
    sz_leg = int(leg["corpo_pt"] * 2)
    est, _ = reescrever_estilo(
        est, id_legenda,
        ppr=(b'<w:pPr><w:spacing w:before="%d" w:after="%d" w:line="240" '
             b'w:lineRule="auto"/>' % (pt(leg["antes_pt"]), pt(leg["depois_pt"]))
             + sem_recuo + centro + b"</w:pPr>"),
        rpr=((b"<w:rPr><w:i/><w:iCs/>" if leg["italico"]
              else b'<w:rPr><w:i w:val="0"/><w:iCs w:val="0"/>')
             + b'<w:color w:val="auto"/>'
             b'<w:sz w:val="%d"/><w:szCs w:val="%d"/></w:rPr>' % (sz_leg, sz_leg)))

    sz_nota = int(NORMA["notas"]["corpo_pt"] * 2)
    m_nota = bloco_do_estilo(est, "Textodenotaderodap")
    if m_nota:
        bn = m_nota.group(0)
        bn = re.sub(rb"<w:szCs(?:\s[^>]*)?/>", b"", re.sub(rb"<w:sz(?:\s[^>]*)?/>", b"", bn))
        if b"<w:rPr>" in bn:
            bn = bn.replace(b"<w:rPr>", b'<w:rPr><w:sz w:val="%d"/><w:szCs w:val="%d"/>'
                            % (sz_nota, sz_nota), 1)
        else:
            bn = bn.replace(b"</w:style>", b'<w:rPr><w:sz w:val="%d"/><w:szCs w:val="%d"/>'
                            b"</w:rPr></w:style>" % (sz_nota, sz_nota))
        est = est[:m_nota.start()] + bn + est[m_nota.end():]

    ref = NORMA["referencia"]
    est, _ = reescrever_estilo(
        est, "Referencia",
        ppr=(b'<w:pPr><w:spacing w:after="%d" w:line="%d" w:lineRule="auto"/>'
             % (pt(ref["depois_pt"]), int(round(240 * ref["entrelinha"])))
             + sem_recuo + esquerda + b"</w:pPr>"))

    id_lista_il = "ndicedeilustraes"
    li = NORMA["lista_de_ilustracoes"]
    ppr_li = (b'<w:pPr><w:tabs><w:tab w:val="right" w:leader="dot" w:pos="%d"/>'
              b'</w:tabs>' % largura_util
              + sem_hifen
              + b'<w:spacing w:before="0" w:after="%d" w:line="240" w:lineRule="auto"/>'
              % pt(3)
              + b'<w:ind w:left="%d" w:hanging="%d"/>'
              % (tw(li["pendente_cm"]), tw(li["pendente_cm"]))
              + esquerda + b"</w:pPr>")
    if not bloco_do_estilo(est, id_lista_il):
        est, feito = criar_estilo(est, id_lista_il, "table of figures", ppr_li, None,
                                  embutido=True)
        if feito:
            criados.append(id_lista_il)
    else:
        est, _ = reescrever_estilo(est, id_lista_il, ppr=ppr_li)

    # A epigrafe tem a forma da dedicatoria.
    est, feito = criar_estilo(est, ESTILO_EPIGRAFE, "Epigrafe", None, None,
                              base=ESTILO_DEDICATORIA.encode())
    if feito:
        criados.append(ESTILO_EPIGRAFE)

    # ---- 3b. a citacao, que e o paragrafo recuado
    #
    # `Citao` e o styleId que o Word da a "Quote" num documento em portugues.
    # Num arquivo que nunca usou o estilo ele nao existe, e ai se cria.
    ppr_cit = (b'<w:pPr><w:spacing w:before="%d" w:after="%d" w:line="240" '
               b'w:lineRule="auto"/><w:ind w:left="%d" w:firstLine="0"/>'
               b'<w:jc w:val="both"/></w:pPr>'
               % (pt(cit["antes_pt"]), pt(cit["depois_pt"]),
                  tw(cit["recuo_esq_cm"])))
    rpr_cit = (b'<w:rPr><w:sz w:val="%d"/><w:szCs w:val="%d"/></w:rPr>'
               % (sz - 2 * cit["menor_pt"], sz - 2 * cit["menor_pt"]))
    if not bloco_do_estilo(est, "Citao"):
        est, feito = criar_estilo(est, "Citao", "Quote", ppr_cit, rpr_cit)
        if feito:
            criados.append("Citao")
    est, _ = reescrever_estilo(
        est, "Citao", ppr=ppr_cit, rpr=rpr_cit)

    # ---- 4. as edicoes no documento, de tras para a frente
    edicoes, divisoes, renumerados, trocas_chave = [], [], [], []
    estilos_antes = O.estilos_definidos(est)
    mudou = {k: 0 for k in ("pretextual", "tabela", "titulo", "legenda",
                            "corpo", "capa-destitulada", "figura", "citacao",
                            "referencia", "lista", "aprovacao", "dedicatoria",
                            "numero", "palavras-chave", "legenda-seq",
                            "lista-automatica", "natureza", "epigrafe",
                            "vazios-ilustracao", "vazios-referencia",
                            "vazios-corpo-apagados",
                            "espacos", "separadores")}
    # O numero das ilustracoes so vai para campo SEQ se o autor as numerou na
    # ordem em que aparecem. O campo renumera ao ser atualizado, e numeracao fora
    # de ordem viraria outra numeracao, com as remissoes do texto apontando para
    # a ilustracao errada.
    seq_ok, vistos = {}, {}
    for f in fichas:
        if papeis[f["i"]] != "legenda":
            continue
        m = ROTULO_DE_ILUSTRACAO.match(f["texto"])
        if not m:
            continue
        rot = rotulo_seq(m.group(1))
        vistos[rot] = vistos.get(rot, 0) + 1
        if int(m.group(2)) != vistos[rot]:
            seq_ok[rot] = False
        seq_ok.setdefault(rot, True)
    avisos_legenda = ["%s fora de ordem: numeros mantidos a mao" % r
                      for r, ok in seq_ok.items() if not ok]
    ESTILO_DE_PAPEL = {"tabela": "TabelaTrabalho",
                       "titulo1": "Ttulo1", "titulo2": "Ttulo2", "titulo3": "Ttulo3",
                       "titulo4": "Ttulo4",
                       "legenda": id_legenda, "figura": "FiguraGrafico",
                       "lista": "ListaPreTextual",
                       "citacao": "Citao", "referencia": "Referencia"}
    # Titulo do pre-textual: a norma manda Titulo 0, que tem a forma do Titulo 1
    # e fica fora do sumario.
    TITULO_PRETEXTUAL = TITULO_PRE

    # O corpo do resumo e do abstract vai do titulo ate o proximo titulo do
    # pre-textual, e leva entrelinha propria.
    dentro_do_resumo = set()
    marcando = False
    for f in fichas:
        if f["i"] >= inicio_textual or f["tabela"]:
            break
        t = f["texto"]
        eh_tit = bool(t) and len(t) < 90 and bool(TITULO_PRETEXTUAL.match(t))
        if eh_tit:
            marcando = bool(re.match(r"^(RESUMO|ABSTRACT|RESUMEN)", t, re.I))
            continue
        if marcando and t:
            dentro_do_resumo.add(f["i"])

    # O texto dos agradecimentos, do titulo ate o proximo titulo do pre-textual.
    dos_agradecimentos = set()
    marcando = False
    for f in fichas:
        if f["i"] >= inicio_textual or f["tabela"]:
            break
        t = f["texto"]
        if bool(t) and len(t) < 90 and bool(TITULO_PRETEXTUAL.match(t)):
            marcando = bool(re.match(r"^(AGRADECIMENTOS?|ACKNOWLEDG)", t, re.I))
            continue
        if marcando and t:
            dos_agradecimentos.add(f["i"])

    for f in fichas:
        papel = papeis[f["i"]]
        bloco, p = f["bloco"], f["pPr"]

        if papel == "sumario-velho":
            continue                      # a tabela inteira sai de uma vez
        if papel == "indice":
            continue                      # resultado de campo, que o Word regera

        if papel in ("pretextual", "vazio") and f["i"] < inicio_textual:
            t = f["texto"]
            eh_titulo = bool(t) and len(t) < 90 and bool(TITULO_PRETEXTUAL.match(t))
            sid = ("Ttulo0" if eh_titulo
                   else "ResumoTrabalho" if f["i"] in dentro_do_resumo
                   else "TextoPreTextual" if f["i"] in dos_agradecimentos
                   else "PreTextual")
            eh_estilo_de_titulo = bool(f["estilo"]) and bool(
                re.match(r"(Ttulo|T[íi]tulo|Heading)\d", f["estilo"]))

            if f["i"] in dentro_do_resumo:
                # O resumo inteiro no corpo do estilo, palavras-chave inclusive:
                # o corpo escrito em cada corrida venceria o do estilo.
                sem = sem_corpo_de_letra(bloco)
                if sem != bloco:
                    mudou["palavras-chave"] += 1
                bloco = sem
                cru = O.texto(bloco)
                mk = PALAVRAS_CHAVE.match(cru)
                if mk and NORMA["palavras_chave_abnt"]:
                    import html
                    resto = html.unescape(cru[mk.end():]).strip()
                    novo_t = palavras_chave_abnt(resto)
                    if novo_t != resto:
                        nb = trocar_trecho(bloco, mk.end(), len(cru), [_xml(novo_t)])
                        renumerados.append((f["texto"], O.texto(nb).strip()))
                        trocas_chave.append((resto, novo_t))
                        bloco = nb
                p = O.ppr(bloco)

            # Paragrafo que ja tem estilo proprio fica com ele, a menos que seja
            # um titulo do pre-textual, que a norma manda pôr em Titulo 0.
            # Trocar o estilo leva junto tudo o que ele dava: na folha de rosto,
            # `Recuodecorpodetexto` carregava o recuo do bloco de qualificacao e
            # das linhas de assinatura, e substitui-lo jogou as assinaturas para
            # fora da margem.
            if (f["estilo"] and not eh_estilo_de_titulo and not eh_titulo
                    and f["i"] not in dentro_do_resumo
                    and f["i"] not in dos_agradecimentos):
                continue

            if eh_estilo_de_titulo and not eh_titulo:
                # Tirar o estilo leva junto o que ele dava. Congela primeiro o
                # alinhamento, o recuo, o espacamento e o negrito que vinham de
                # la, e so depois troca: sem isto, o que estava centralizado e em
                # negrito na capa perderia as duas coisas.
                bloco = congelar(
                    bloco,
                    do_estilo(est_original, f["estilo"], b"w:pPr",
                              [b"w:jc", b"w:ind", b"w:spacing"]),
                    do_estilo(est_original, f["estilo"], b"w:rPr",
                              [b"w:b", b"w:bCs", b"w:i", b"w:sz", b"w:szCs"]))
                p = O.ppr(bloco)
                mudou["capa-destitulada"] += 1

            if eh_titulo:
                # A forma passa a morar no Titulo 0, entao a direta sai.
                p = tirar(p, b"w:jc", b"w:ind", b"w:spacing", b"w:outlineLvl")
                mudou["titulo"] += 1
            elif f["i"] in dos_agradecimentos:
                # A forma mora no estilo: o recuo e o alinhamento escritos a mao
                # sairiam diferentes de um trabalho para outro.
                p = tirar(p, b"w:jc", b"w:ind", b"w:spacing")
            else:
                # O mesmo vale para o resto do pre-textual: a entrelinha vem do
                # estilo, e a entrelinha escrita no paragrafo venceria a dele.
                p = tirar(p, b"w:spacing")

            novo = escrever_ppr(bloco, com_pstyle(p, sid))
            if novo != f["bloco"]:
                edicoes.append((f["ini"], f["fim"], novo))
                mudou["pretextual"] += 1
            continue

        if papel == "tabela":
            if f["estilo"]:               # ja tem estilo proprio: nao se troca
                continue
            novo = escrever_ppr(bloco, com_pstyle(p, "TabelaTrabalho"))
            if novo != bloco:
                edicoes.append((f["ini"], f["fim"], novo))
                mudou["tabela"] += 1
            continue

        texto_atual = f["texto"]
        if papel.startswith("titulo"):
            # Com a numeracao automatica, o numero digitado sai; sem ela, passa a
            # ser seguido de tabulacao. E texto do autor que muda, e por isso cada
            # troca vai registrada para a conferencia.
            numerado = False
            if numerar_auto:
                m_n = NUMERO_CRU.match(O.texto(bloco))
                nb = trocar_trecho(bloco, 0, m_n.end(), [""]) if m_n else bloco
                numerado = bool(m_n)
            else:
                nb = numerar_titulo(bloco)
            if nb != bloco:
                renumerados.append((f["texto"], O.texto(nb).strip()))
                bloco = nb
                texto_atual = O.texto(nb).strip()
                mudou["numero"] += 1
            # O corpo do titulo vem do estilo: o escrito em cada corrida o
            # venceria. O espacejamento de letras tambem sai.
            bloco = sem_espacejamento(sem_corpo_de_letra(bloco))
            # A lista desligada (`numId` 0) sai do titulo. Ela so diz "sem
            # numeracao", e os estilos de titulo reescritos aqui ja nao numeram;
            # mas com ela o Word desfaz o recuo pendente do estilo, e a segunda
            # linha do titulo volta a margem.
            p = re.sub(rb"<w:numPr>(?:(?!</w:numPr>).)*?"
                       rb'<w:numId w:val="0"/>(?:(?!</w:numPr>).)*?</w:numPr>',
                       b"", O.ppr(bloco), flags=re.S)
            if numerar_auto:
                # O titulo numerado leva a lista no proprio paragrafo, alem do
                # estilo; o que nao tinha numero leva a lista desligada, para que
                # o estilo nao o numere.
                nv = int(papel[len("titulo"):]) if papel[len("titulo"):].isdigit() else 1
                p = com_numpr(p, nv - 1, num_id) if (numerado and nv <= 4) \
                    else com_numpr(p, 0, 0)

        if (papel == "titulo1" and NORMA["cabecalho"]["dividir_no_dois_pontos"]
                and ":" in texto_atual):
            # O titulo de capitulo passa a ser o que vem antes do dois-pontos, e
            # o que explica desce para um paragrafo proprio, fora do sumario.
            # Sem isso o cabecalho corrente traria 153 caracteres, que sao tres
            # linhas, e a alternativa seria guardar uma copia curta do titulo,
            # que envelhece em silencio quando o autor edita o original.
            limpo = tirar(p, b"w:jc", b"w:ind", b"w:spacing", b"w:outlineLvl")
            com_estilo = escrever_ppr(bloco, com_pstyle(limpo, "Ttulo1"))
            # O deslocamento tem de ser contado no texto CRU, e nao no texto
            # aparado de `f["texto"]`: um espaco a mais no comeco do paragrafo
            # desloca o corte, e foi o que separou o S de HISTORICOS do resto.
            duas = dividir_no_caractere(com_estilo, O.texto(bloco).index(":"))
            if duas:
                a_bl, b_bl = duas
                b_bl = sem_negrito(tirar_dois_pontos(b_bl))
                # O subtitulo nasce com o pPr do titulo, lista inclusive, e sairia
                # numerado.
                b_bl = escrever_ppr(
                    b_bl, com_pstyle(sem_numpr(tirar(O.ppr(b_bl), b"w:spacing")),
                                     ESTILO_SUBTITULO))
                edicoes.append((f["ini"], f["fim"], a_bl + b_bl))
                divisoes.append((texto_atual, O.texto(a_bl).strip(),
                                 O.texto(b_bl).strip()))
                mudou["titulo"] += 1
                continue

        if papel == "legenda":
            # O corpo vem do estilo: o escrito em cada corrida o venceria.
            bloco = sem_corpo_de_letra(bloco)
            p = O.ppr(bloco)
            m = ROTULO_DE_ILUSTRACAO.match(texto_atual)
            if m:
                rot = rotulo_seq(m.group(1))
                cru = O.texto(bloco)
                desl = len(cru) - len(cru.lstrip())
                if seq_ok.get(rot):
                    nb = com_campo_seq(bloco, rot, desl + m.start(2), desl + m.end(2))
                    if nb != bloco:
                        bloco = nb
                        mudou["legenda-seq"] += 1
            if LEGENDA.match(texto_atual):
                # Sem o negrito que o autor pos a mao, por decisao da
                # coordenacao: a legenda e italico 11 e o estilo governa. O Word
                # copia o negrito direto do titulo para a lista de ilustracoes.
                bloco = sem_negrito(bloco)
                # o titulo fica com a figura ou a tabela que vem abaixo dele
                p = O.ppr(bloco)
                if b"<w:keepNext/>" not in p:
                    p = b"<w:keepNext/>" + p

        if papel == "citacao":
            # O corpo vem do estilo: o escrito em cada corrida o venceria.
            bloco = sem_corpo_de_letra(bloco)
            p = O.ppr(bloco)

        if papel in ESTILO_DE_PAPEL:
            # A forma passa a morar no estilo, entao a direta sai.
            limpo = tirar(p, b"w:jc", b"w:ind", b"w:spacing", b"w:outlineLvl")
            if (papel == "titulo1" and not NORMA["titulo1"]["centralizado"]
                    and not NUMERO_CRU.match(O.texto(f["bloco"]))
                    and not numera_pelo_word(f)):
                # Sem numero, no centro e sem o recuo pendente (NBR 14724, 5.2.3).
                limpo = no_ppr(limpo, b'<w:ind w:left="0" w:firstLine="0"/>',
                               DEPOIS_DE_IND)
                limpo = no_ppr(limpo, b'<w:jc w:val="center"/>',
                               rb"textDirection|textAlignment|textboxTightWrap|outlineLvl"
                               rb"|divId|cnfStyle|rPr|sectPr|pPrChange")
            if papel == "legenda" and LEGENDA.match(texto_atual):
                # o titulo da ilustracao fica perto dela
                limpo += b'<w:spacing w:after="%d"/>' % pt(leg["titulo_depois_pt"])
            novo = escrever_ppr(bloco, com_pstyle(limpo, ESTILO_DE_PAPEL[papel]))
            if novo != f["bloco"]:
                edicoes.append((f["ini"], f["fim"], novo))
                mudou[{"titulo1": "titulo", "titulo2": "titulo",
                       "titulo3": "titulo", "titulo4": "titulo"}.get(papel, papel)] += 1
            continue

        if bloco != f["bloco"]:
            # titulo de nivel sem estilo na norma, que so trocou o numero
            edicoes.append((f["ini"], f["fim"], bloco))
            continue

        if papel == "corpo":
            # A forma do corpo passa a morar no Normal, entao a direta sai
            # inteira: reescreve-la em cada paragrafo seria o defeito que o
            # programa existe para desfazer.
            pc = tirar(p, b"w:spacing", b"w:jc", b"w:ind")
            # A lista desligada (`numId` 0) tambem sai, se o estilo nao numera:
            # com ela o Word zera o recuo, e o da primeira linha, que vinha
            # escrito no paragrafo e acabou de sair, nao volta do Normal.
            if (b'<w:numId w:val="0"/>' in pc
                    and b"<w:numPr>" not in estilos_antes.get(O.estilo_de(bloco) or "Normal",
                                                              {}).get("xml", b"")):
                pc = sem_numpr(pc)
            novo = escrever_ppr(bloco, pc)
            if novo != bloco:
                edicoes.append((f["ini"], f["fim"], novo))
                mudou["corpo"] += 1

    # ---- 4c. a capa e a folha de rosto, diagramadas por espaco e nao por linhas
    avisos_capa, capa_fim, rosto_fim = [], None, None
    pre = NORMA["pretextual"]
    if pre["capa_por_espacamento"]:
        ANO = re.compile(r"^\d{4}$")
        for f in fichas[:80]:
            if f["i"] >= inicio_textual or f["tabela"]:
                break
            if ANO.match(f["texto"]):
                capa_fim = f["i"]
                break
        if capa_fim is not None:
            for f in fichas[capa_fim + 1:]:
                if f["i"] >= inicio_textual:
                    break
                if (f["texto"] and len(f["texto"]) < 90
                        and TITULO_PRETEXTUAL.match(f["texto"])):
                    rosto_fim = f["i"] - 1
                    break

    if capa_fim is not None:
        sec0 = O.secoes(doc)[0] if O.secoes(doc) else {}
        larg_tw = tw(NORMA["margens_cm"] and
                     (sec0.get("larg", 21.0) - NORMA["margens_cm"]["esq"]
                      - NORMA["margens_cm"]["dir"]))
        alt_tw = tw(sec0.get("alt", 29.7) - NORMA["margens_cm"]["sup"]
                    - NORMA["margens_cm"]["inf"])
        novas, avisos_capa = diagramar_capa(
            fichas, 0, capa_fim, larg_tw, alt_tw, corpo_pt, pre["entrelinha"],
            list(pre["capa_fracoes"]), tw(pre["capa_folga_pe_cm"]))
        if rosto_fim is not None and rosto_fim > capa_fim:
            novas += congelar_espaco(
                fichas, capa_fim + 1, rosto_fim, float(sz_velho) / 2,
                quebrar_pagina=True,
                tirar_corpo=(sem_corpo_de_letra
                             if pre["segue_o_normal"] else None))
        faixas_capa = {(a, b) for a, b, _ in novas}
        edicoes = [e for e in edicoes if (e[0], e[1]) not in faixas_capa]
        # o estilo que cada paragrafo ja recebeu volta para dentro da edicao nova
        arrumadas = []
        for a, b, dados in novas:
            if not dados:
                arrumadas.append((a, b, dados))
                continue
            f = next(g for g in fichas if g["ini"] == a)
            eh_tit = bool(f["estilo"]) and bool(
                re.match(r"(Ttulo|T[íi]tulo|Heading)\d", f["estilo"]))
            if eh_tit:
                # Na capa nao ha titulo: ha nome de autor e nome de trabalho.
                # Congela o que o estilo dava antes de tira-lo, como no resto do
                # pre-textual.
                dados = congelar(
                    dados,
                    do_estilo(est_original, f["estilo"], b"w:pPr", [b"w:jc"]),
                    do_estilo(est_original, f["estilo"], b"w:rPr",
                              [b"w:b", b"w:bCs", b"w:i"]))
                mudou["capa-destitulada"] += 1
            sid = "PreTextual" if eh_tit else (f["estilo"] or "PreTextual")
            arrumadas.append((a, b, escrever_ppr(dados,
                                                 com_pstyle(O.ppr(dados), sid))))
        edicoes += arrumadas

    # ---- 4d. a folha de aprovacao vai para folha propria
    #
    # A NBR 14724 a poe depois da folha de rosto, em folha propria. Aqui ela esta
    # dentro da folha de rosto, nos blocos de assinatura. O programa so abre a
    # folha: nome, titulo, natureza e data de aprovacao, que a norma tambem exige
    # ali, sao texto que o autor tem de escrever.
    if pre["aprovacao_em_folha_propria"]:
        for f in fichas:
            if f["i"] >= inicio_textual or f["tabela"] or not f["texto"]:
                continue
            if not re.match(r"^_{5,}", f["texto"]):
                continue
            bl = f["bloco"]
            pp = O.ppr(bl)
            if b"<w:pageBreakBefore/>" not in pp:
                novo_bl = escrever_ppr(bl, b"<w:pageBreakBefore/>"
                                       + tirar(pp, b"w:pageBreakBefore"))
                edicoes = [e for e in edicoes if (e[0], e[1]) != (f["ini"], f["fim"])]
                edicoes.append((f["ini"], f["fim"], novo_bl))
                mudou["aprovacao"] += 1
            break

    # ---- 4e. a dedicatoria perde o titulo, que a NBR nao quer
    #
    # Tirado o titulo, a quebra de pagina que vinha dele passa para o primeiro
    # paragrafo do texto da dedicatoria, senao ela sobe para a folha anterior.
    removidos = []
    if pre["dedicatoria_sem_titulo"]:
        for k, f in enumerate(fichas):
            if f["i"] >= inicio_textual or f["tabela"]:
                continue
            mt = re.match(r"^(Dedicat[óo]ria|Ep[íi]grafe)\.?$", f["texto"].strip(), re.I)
            if not mt:
                continue
            estilo_pe = (ESTILO_EPIGRAFE if mt.group(1).lower().startswith("ep")
                         else ESTILO_DEDICATORIA)
            edicoes = [e for e in edicoes if (e[0], e[1]) != (f["ini"], f["fim"])]
            edicoes.append((f["ini"], f["fim"], b""))
            removidos.append(f["texto"])
            mudou["epigrafe" if estilo_pe == ESTILO_EPIGRAFE else "dedicatoria"] += 1
            # O texto da dedicatoria vai do titulo tirado ate o proximo titulo do
            # pre-textual. Os vazios dali saem: na secao alinhada ao pe, vazio
            # depois do texto o empurraria para cima.
            primeiro = True
            for g in fichas[k + 1:]:
                if g["i"] >= inicio_textual or g["tabela"]:
                    break
                t = g["texto"]
                if t and len(t) < 90 and TITULO_PRETEXTUAL.match(t):
                    break
                edicoes = [e for e in edicoes
                           if (e[0], e[1]) != (g["ini"], g["fim"])]
                if not t:
                    if not O.travado(g["bloco"]):
                        edicoes.append((g["ini"], g["fim"], b""))
                    continue
                pp = tirar(O.ppr(g["bloco"]), b"w:pageBreakBefore", b"w:ind",
                           b"w:jc", b"w:spacing")
                if pre["dedicatoria_na_base"]:
                    pp = com_pstyle(pp, estilo_pe)
                if primeiro:
                    # reserva: se a secao do pe nao puder ser criada, a quebra
                    # de pagina ainda tira a dedicatoria da folha anterior
                    pp = b"<w:pageBreakBefore/>" + pp
                    primeiro = False
                edicoes.append((g["ini"], g["fim"], escrever_ppr(g["bloco"], pp)))

    # ---- 4f. a lista de ilustracoes, montada pelo Word
    #
    # A lista digitada sai, com os subtitulos dela, e entra um campo por rotulo
    # numerado por SEQ. Com numero de pagina, o pontilhado passa a levar a algum
    # lugar (NBR 14724, 4.2.1.9). Tabela nao entra: por decisao da coordenacao,
    # nao ha lista de tabelas, e o " E TABELAS" sai do titulo.
    li = NORMA["lista_de_ilustracoes"]
    listas_trocadas = set()
    rotulos = [r for r, ok in seq_ok.items() if ok]
    if li["automatica"] and rotulos:
        LISTA_IL = re.compile(r"^(LISTA DE|[ÍI]NDICE DE) (GR[ÁA]FICOS|ILUSTRA|FIGURAS|QUADROS)"
                              r"|^LIST OF (FIGURES|CHARTS|GRAPHS|ILLUSTRATIONS)", re.I)
        for k, f in enumerate(fichas):
            if f["i"] >= inicio_textual:
                break
            if not (f["texto"] and LISTA_IL.match(f["texto"])):
                continue
            faixa = []
            for g in fichas[k + 1:]:
                if g["i"] >= inicio_textual or g["tabela"]:
                    break
                t = g["texto"]
                if t and len(t) < 90 and TITULO_PRETEXTUAL.match(t):
                    break
                faixa.append(g)
            if not any(g["texto"] for g in faixa):
                continue
            # Lista que ja e campo do Word fica como esta, e a busca segue para a
            # proxima.
            if any(g["indice"] for g in faixa):
                continue
            primeiro = True
            for g in faixa:
                if O.travado(g["bloco"]) and not g["texto"]:
                    continue
                listas_trocadas.add((g["ini"], g["fim"]))
                edicoes = [e for e in edicoes if (e[0], e[1]) != (g["ini"], g["fim"])]
                if primeiro:
                    novo = b"".join(campo_de_lista(r, id_lista_il) for r in rotulos)
                    primeiro = False
                else:
                    novo = b""
                edicoes.append((g["ini"], g["fim"], novo))
                if g["texto"]:
                    removidos.append(g["texto"])
            mudou["lista-automatica"] += 1
            if li["sem_tabelas"] and re.search(r"\s+E\s+TABELAS\s*$", f["texto"], re.I):
                atual = next((e for e in edicoes if (e[0], e[1]) == (f["ini"], f["fim"])),
                             (f["ini"], f["fim"], f["bloco"]))
                edicoes = [e for e in edicoes if (e[0], e[1]) != (f["ini"], f["fim"])]
                cru = O.texto(atual[2])
                mt = re.search(r"\s+E\s+TABELAS\s*$", cru, re.I)
                nb = trocar_trecho(atual[2], mt.start(), len(cru), [""])
                edicoes.append((f["ini"], f["fim"], nb))
                renumerados.append((f["texto"], O.texto(nb).strip()))
            break

    # ---- 4g. os vazios em volta das ilustracoes e entre as referencias
    #
    # O autor separa a ilustracao do texto com paragrafos vazios, e o espaco
    # agora vem do estilo da legenda. O vazio entre o titulo e a figura, alem
    # disso, quebra a cadeia do `keepNext`, e o titulo pode ficar no pe de uma
    # folha com a figura na seguinte.
    def apagar(g):
        nonlocal edicoes
        if g["texto"] or g["tabela"] or g["figura"] or O.travado(g["bloco"]):
            return False
        edicoes = [e for e in edicoes if (e[0], e[1]) != (g["ini"], g["fim"])]
        edicoes.append((g["ini"], g["fim"], b""))
        return True

    n_vazios_ilustracao = n_vazios_ref = 0
    for k, f in enumerate(fichas):
        papel = papeis[f["i"]]
        if papel == "vazio-referencia":
            n_vazios_ref += apagar(f)
            continue
        if papel not in ("legenda", "figura"):
            continue
        j = k - 1
        while j >= 0 and not fichas[j]["texto"] and papeis[fichas[j]["i"]] != "figura":
            if not apagar(fichas[j]):
                break
            n_vazios_ilustracao += 1
            j -= 1
        if papel == "legenda" and LEGENDA.match(f["texto"]):
            continue        # o que vem depois do titulo e a figura, vista acima
        j = k + 1
        while j < len(fichas) and not fichas[j]["texto"] and papeis[fichas[j]["i"]] != "figura":
            if not apagar(fichas[j]):
                break
            n_vazios_ilustracao += 1
            j += 1
    mudou["vazios-ilustracao"] = n_vazios_ilustracao
    mudou["vazios-referencia"] = n_vazios_ref

    # ---- 4h. os vazios que sobram no corpo
    #
    # Vazio encostado em titulo, citacao ou ilustracao sai, porque o estilo ja
    # da o espaco. Vazio depois de paragrafo de corpo tambem sai, por decisao da
    # coordenacao: o corpo abre com o recuo de primeira linha do Normal (o
    # escrito a mao saiu no passo 4), e o recuo ja marca o paragrafo novo.
    ENCOSTA = ("titulo1", "titulo2", "titulo3", "titulo4", "citacao", "legenda", "figura")
    k = 0
    while k < len(fichas):
        f = fichas[k]
        if (f["i"] < inicio_textual or f["texto"] or f["tabela"] or f["figura"]
                or papeis[f["i"]] != "vazio"):
            k += 1
            continue
        j = k
        while (j < len(fichas) and not fichas[j]["texto"] and not fichas[j]["tabela"]
               and not fichas[j]["figura"] and papeis[fichas[j]["i"]] == "vazio"):
            j += 1
        corrida = [g for g in fichas[k:j] if not O.travado(g["bloco"])]
        antes = fichas[k - 1] if k > 0 else None
        depois = fichas[j] if j < len(fichas) else None
        pa = papeis[antes["i"]] if antes else None
        pd = papeis[depois["i"]] if depois else None
        if corrida and pd == "titulo1":
            pass        # antes de titulo que quebra pagina: e a varredura 4b
        elif corrida and (pa in ENCOSTA or pd in ENCOSTA or pa == "corpo"):
            # Conta so o que esta varredura apaga: a das ilustracoes (4g) ja
            # pode ter apagado parte, e contar de novo inflaria o relatorio.
            for g in corrida:
                faixa = (g["ini"], g["fim"])
                if any((e[0], e[1]) == faixa and e[2] == b"" for e in edicoes):
                    continue
                edicoes = [e for e in edicoes if (e[0], e[1]) != faixa]
                edicoes.append((g["ini"], g["fim"], b""))
                mudou["vazios-corpo-apagados"] += 1
        k = j

    # ---- 8 e 9. o que so se aponta: titulo sem estilo e espaco posto a mao
    titulos_por_marcar = [
        (f["i"], f["texto"][:80]) for f in fichas
        if f["i"] >= inicio_textual and not f["tabela"]
        and not papeis[f["i"]].startswith("titulo")
        and papeis[f["i"]] not in ("legenda", "figura", "lista", "referencia", "indice")
        and parece_titulo(f["texto"])]
    empurroes = []
    for f in fichas:
        if f["i"] < inicio_textual or f["tabela"] or not f["texto"]:
            continue
        v = max([int(x) for x in re.findall(
            rb'w:(?:before|after)="(\d+)"', O.prop(f["pPr"], b"w:spacing") or b"")] or [0])
        if v >= 720:
            empurroes.append((f["i"], f["texto"][:60], v / 20.0, papeis[f["i"]]))

    # ---- 4b. os vazios que a quebra de pagina tornou folha em branco
    papel_com_quebra = {"titulo1"}
    apagados = []
    if NORMA["titulo1"]["quebra_pagina"]:
        alvos = set(vazios_antes_de_quebra(doc, fichas, papeis, papel_com_quebra))
        # o titulo do pre-textual tambem quebra pagina, e ele nao tem papel
        # proprio: reconhece-se pelo estilo que acabou de receber.
        for f in fichas:
            if f["i"] >= inicio_textual or papeis[f["i"]] == "sumario-velho":
                continue
            t = f["texto"]
            if not (t and len(t) < 90 and TITULO_PRETEXTUAL.match(t)):
                continue
            k = f["i"] - 1
            while k >= 0:
                g = fichas[k]
                if g["tabela"] or not O.vazio(g["bloco"]):
                    break
                if not O.travado(g["bloco"], aceitar_quebra=True):
                    alvos.add(k)
                k -= 1
        # O mesmo paragrafo ja pode ter recebido uma edicao de estilo la atras.
        # Duas edicoes para a mesma faixa de bytes se aplicam uma sobre a outra
        # e corrompem o arquivo, que o Word recusa abrir. A exclusao vence, e a edicao de estilo daquele paragrafo sai.
        faixas = {(fichas[k]["ini"], fichas[k]["fim"]) for k in alvos}
        edicoes = [e for e in edicoes if (e[0], e[1]) not in faixas]
        for k in sorted(alvos):
            edicoes.append((fichas[k]["ini"], fichas[k]["fim"], b""))
            apagados.append(k)

    # ---- 5. o sumario digitado sai, o campo entra
    trocou_sumario = False
    if tabela_sumario:
        edicoes.append((tabela_sumario[0], tabela_sumario[1], campo_de_sumario()))
        trocou_sumario = True

    # O que as edicoes tiram de proposito (paragrafo apagado, sumario digitado,
    # lista digitada) sai da conta da estrutura; o resto tem de ficar.
    intencionais = {(tabela_sumario[0], tabela_sumario[1])} if tabela_sumario else set()
    intencionais |= listas_trocadas
    estrutura_antes = contar_estrutura(doc)
    tirado = {k: 0 for k in ESTRUTURA}
    for ini, fim, novo in edicoes:
        if novo == b"" or (ini, fim) in intencionais:
            for k, v in contar_estrutura(doc[ini:fim]).items():
                tirado[k] += v - contar_estrutura(novo)[k]
    for ini, fim, novo in sorted(edicoes, key=lambda e: e[0], reverse=True):
        doc = doc[:ini] + novo + doc[fim:]

    # ---- 6. a pagina: margens, hifenizacao e o numero de pagina
    doc, n_margens = escrever_margens(doc, NORMA["margens_cm"])

    # ---- 6d. o paragrafo em estilo de titulo que nao e titulo numerado
    #
    # Com a lista ligada aos estilos, todo paragrafo em Titulo 1 a 4 sairia
    # numerado. O que nao traz lista propria leva a lista desligada.
    if numerar_auto:
        pars_n = O.paragrafos(doc)
        for a_, b_ in reversed(pars_n):
            bl_ = doc[a_:b_]
            if not re.fullmatch(r"Ttulo[1-4]", O.estilo_de(bl_) or ""):
                continue
            pp_ = O.ppr(bl_)
            if b"<w:numPr>" in pp_:
                continue
            doc = doc[:a_] + escrever_ppr(bl_, com_numpr(pp_, 0, 0)) + doc[b_:]

    # ---- 6b. a natureza do trabalho, do meio da mancha a margem direita
    n_natureza = 0
    if NORMA["pretextual"]["natureza_no_meio"]:
        doc, n_natureza = natureza_no_meio(
            doc, O.estilos_definidos(est),
            int(round(largura_util * NORMA["pretextual"]["dedicatoria_recuo_fracao"])),
            NORMA["limiar_recuo_tw"])
    mudou["natureza"] = n_natureza

    # ---- 6c. espaco repetido, da parte textual em diante
    pars_f = O.paragrafos(doc)
    desde = next((a for a, b in pars_f if O.estilo_de(doc[a:b]) == "Ttulo1"), len(doc))
    doc, mudou["espacos"] = juntar_espacos(doc, desde)

    # ---- 7. a numeracao: pre-textual contado e nao numerado
    #
    # As tampas saem primeiro, porque sao a gambiarra que isto substitui.
    n_tampas, separou, aviso_ded = 0, False, None
    novo_hdr = None
    if NORMA["numeracao"]["mostra_do_textual"]:
        doc, n_tampas = tirar_tampas(doc)
        # O limite e reachado no documento JA reescrito: os deslocamentos de
        # `fichas` sao do arquivo de entrada e nao valem mais aqui.
        pars2 = O.paragrafos(doc)
        limite = None
        for k, (a, b) in enumerate(pars2):
            if O.estilo_de(doc[a:b]) == "Ttulo1":
                limite = k
                break
        if limite and limite > 0:
            alvo = pars2[limite - 1]
            novo_hdr = proximo_rid(rels)
            doc, separou = separar_secao(doc, alvo, novo_hdr)
            if separou and NORMA["pretextual"]["dedicatoria_na_base"]:
                pars3 = O.paragrafos(doc)
                a, b = pars3[limite - 1]
                modelo = re.search(rb"<w:sectPr\b.*?</w:sectPr>", doc[a:b], re.S)
                doc, aviso_ded = secao_da_dedicatoria(doc, modelo.group(0))

    # ---- 7b. o numero de pagina, quando a parte textual nao o tem
    receber_numero, hdr_numero_usado, n_faltava, aviso_numero_pag = set(), False, 0, None
    hdr_numero = None
    if NORMA["numeracao"]["inserir_se_faltar"]:
        # `O.ler` nao traz cabecalho nem rodape; sem eles, todo trabalho
        # pareceria sem numero, e o que ja o tem ganharia um segundo.
        with zipfile.ZipFile(origem) as z_:
            cab_rod = {n: z_.read(n) for n in z_.namelist()
                       if re.match(r"word/(header|footer)\d*\.xml$", n)}
        if separou:
            hdr_numero = "rId%d" % (int(novo_hdr[3:]) + 1)
            doc, receber_numero, hdr_numero_usado, n_faltava = numero_onde_falta(
                doc, dict(partes, **cab_rod), rels, novo_hdr, hdr_numero)
        elif not any(b"PAGE" in dados_ for dados_ in cab_rod.values()):
            aviso_numero_pag = ("o trabalho nao tem numero de pagina, e o pre-textual nao "
                                "ficou em secao propria: pondo o numero, ele sairia tambem "
                                "na capa; nao pus")

    # O programa escreve `r:id` nas referencias de cabecalho, e o prefixo tem de
    # estar declarado. Todo .docx do Word o declara, mas um arquivo gerado por
    # outro programa pode nao declarar, e o XML nao fecharia.
    raiz = re.search(rb"<w:document\b[^>]*>", doc)
    if raiz and b"xmlns:r=" not in raiz.group(0) and b"r:id=" in doc:
        doc = (doc[:raiz.start()] + raiz.group(0).replace(
            b"<w:document", b'<w:document xmlns:r="http://schemas.openxmlformats.org/'
            b'officeDocument/2006/relationships"', 1) + doc[raiz.end():])

    # ---- grava
    shutil.copyfile(origem, destino)
    with zipfile.ZipFile(destino) as z:
        itens = [(i, z.read(i.filename)) for i in z.infolist()]
    n_hifen, n_pagina, n_titulo = False, 0, 0
    hifen_desligada = False
    with zipfile.ZipFile(destino, "w", zipfile.ZIP_DEFLATED) as z:
        for info, dados in itens:
            nome = info.filename
            if nome == "word/document.xml":
                dados = doc
            elif nome == "word/styles.xml":
                dados = est
            elif nome in ("word/footnotes.xml", "word/endnotes.xml"):
                dados, k_sep = endireitar_separadores(dados)
                mudou["separadores"] += k_sep
            elif nome == "word/settings.xml" and NORMA["hifenizar"]:
                dados, n_hifen = ligar_hifenizacao(dados)
            elif nome == "word/settings.xml":
                dados, hifen_desligada = desligar_hifenizacao(dados)
            elif nome == "word/numbering.xml" and numerar_auto:
                dados = com_lista(dados, *lista_de_titulos(
                    abs_id, num_id, tw(regra_num["parada_t1_cm"]),
                    {k: tw(v) for k, v in regra_num["pendente_cm"].items()},
                NORMA["titulo1"]["centralizado"]))
            elif nome == "word/_rels/document.xml.rels":
                if separou:
                    dados = com_relacionamento(dados, novo_hdr, CABECALHO_SEM_NUMERO)
                if hdr_numero_usado:
                    dados = com_relacionamento(dados, hdr_numero, CABECALHO_COM_NUMERO)
                if numerar_auto and "word/numbering.xml" not in partes:
                    dados = com_relacionamento(dados, proximo_rid(dados), "numbering.xml",
                                               TIPO_NUMERACAO)
            elif nome == "[Content_Types].xml":
                if separou:
                    dados = com_override(dados, "/word/" + CABECALHO_SEM_NUMERO)
                if hdr_numero_usado:
                    dados = com_override(dados, "/word/" + CABECALHO_COM_NUMERO)
                if numerar_auto and "word/numbering.xml" not in partes:
                    dados = com_override(dados, "/word/numbering.xml", CT_NUMERACAO)
            elif re.match(r"word/header\d*\.xml$", nome):
                if nome in receber_numero:
                    dados = com_numero_de_pagina(dados)
                dados, k = numero_de_pagina_a_direita(dados)
                n_pagina += k
                if NORMA["cabecalho"]["titulo_no_cabecalho"]:
                    dados, t = titulo_no_cabecalho(dados, largura_util,
                                                   com_numero=numerar_auto)
                    n_titulo += int(t)
            z.writestr(info, dados)
        if separou:
            z.writestr("word/" + CABECALHO_SEM_NUMERO, CABECALHO_VAZIO)
        if hdr_numero_usado:
            z.writestr("word/" + CABECALHO_COM_NUMERO, CABECALHO_NUMERADO)
        if numerar_auto and "word/numbering.xml" not in partes:
            z.writestr("word/numbering.xml", com_lista(None, *lista_de_titulos(
                abs_id, num_id, tw(regra_num["parada_t1_cm"]),
                {k: tw(v) for k, v in regra_num["pendente_cm"].items()},
                NORMA["titulo1"]["centralizado"])))

    return {"conta": conta, "mudou": mudou, "criados": criados,
            "tampas": n_tampas, "secao_pretextual": separou,
            "avisos_capa": avisos_capa,
            "blindados": blindados, "soltos": soltos, "sumarios": sumarios,
            "apagados": apagados, "divisoes": divisoes, "removidos": removidos, "renumerados": renumerados, "numeracao_auto": numerar_auto, "num_id": num_id, "avisos_numero": avisos_numero, "estrutura_antes": estrutura_antes, "estrutura_tirada": tirado, "titulos_por_marcar": titulos_por_marcar, "empurroes": empurroes, "trocas_chave": trocas_chave, "avisos_legenda": avisos_legenda, "aviso_dedicatoria": aviso_ded, "titulo_cab": n_titulo, "numero_faltava": n_faltava, "numero_posto": sorted(receber_numero) + (["novo cabecalho"] if hdr_numero_usado else []), "aviso_numero_pag": aviso_numero_pag, "margens": n_margens, "hifen": n_hifen, "hifen_desligada": hifen_desligada, "n_mso": n_mso, "ids_mso": ids_mso, "pagina": n_pagina, "trocou_sumario": trocou_sumario,
            "inicio_textual": inicio_textual, "sz": sz, "linha": linha,
            "doc_novo": doc, "doc_velho": partes["word/document.xml"],
            "tabela_sumario": tabela_sumario}


# ------------------------------------------------------------- conferencia
def conferir(res, destino):
    """O texto tem de ser o mesmo, fora o sumario digitado que saiu."""
    velho, novo = res["doc_velho"], res["doc_novo"]
    a, b = res["tabela_sumario"] or (0, 0)
    velho_sem = velho[:a] + velho[b:] if res["tabela_sumario"] else velho

    # So os paragrafos COM texto entram na comparacao: o programa apaga vazio
    # antes de titulo que quebra pagina, e comparar a lista inteira acusaria
    # isso como perda de conteudo. O numero de vazios apagados e relatado a
    # parte, e cada um deles passou pelas travas de O.travado.
    tv = [t for t in (O.texto(velho_sem[x:y]).strip()
                      for x, y in O.paragrafos(velho_sem)) if t]
    tn = [t for t in (O.texto(novo[x:y]).strip()
                      for x, y in O.paragrafos(novo)) if t]
    tn_sem_campo = [t for t in tn if not t.startswith(("Sumario automatico:",
                                                       "Lista automatica:"))]
    # Titulo dividido em titulo e subtitulo: o esperado passa a ser a dupla, no
    # lugar do texto antigo. O dois-pontos e o espaco que o seguia saem, porque a
    # quebra de paragrafo passou a fazer o que eles faziam.
    # Titulo cujo numero mudou: o esperado e o numero novo.
    ren = dict(res.get("renumerados", []))
    tv = [ren.get(t, t) for t in tv]
    mapa = {velho_t: (a, b) for velho_t, a, b in res.get("divisoes", [])}
    tv = [x for t in tv for x in (mapa[t] if t in mapa else (t,))]
    # E do titulo que a norma manda tirar, como o da dedicatoria.
    for t in res.get("removidos", []):
        if t in tv:
            tv.remove(t)

    # Espaco repetido que virou um so nao e texto que mudou.
    tv = [" ".join(t.split()) for t in tv]
    tn_sem_campo = [" ".join(t.split()) for t in tn_sem_campo]
    falhas = []
    if len(tv) != len(tn_sem_campo):
        falhas.append("o numero de paragrafos com texto mudou: %d contra %d"
                      % (len(tv), len(tn_sem_campo)))
    else:
        difs = [i for i, (x, y) in enumerate(zip(tv, tn_sem_campo)) if x != y]
        if difs:
            falhas.append("o texto mudou em %d paragrafos, o primeiro em %d"
                          % (len(difs), difs[0]))

    # A estrutura: o que o arquivo tinha, menos o que se tirou de proposito,
    # tem de ser o que ele tem. A conferencia de texto nao ve marcador nem
    # hiperlink perdidos, porque o texto nao muda.
    if "estrutura_antes" in res:
        agora = contar_estrutura(novo)
        for k, v in res["estrutura_antes"].items():
            esperado = v - res["estrutura_tirada"].get(k, 0)
            if agora[k] != esperado:
                falhas.append("%s: o arquivo tinha %d, tirou-se de proposito %d, e "
                              "ficaram %d" % (k, v, res["estrutura_tirada"].get(k, 0),
                                              agora[k]))

    import xml.etree.ElementTree as ET
    with zipfile.ZipFile(destino) as z:
        for nome in ("word/document.xml", "word/styles.xml"):
            try:
                ET.fromstring(z.read(nome))
            except ET.ParseError as e:
                falhas.append("%s nao e XML bem formado: %s" % (nome, e))
    return falhas


def abre_no_word(caminho):
    """Pede ao Word que abra o arquivo. Devolve None, ou o que ele reclamou.

    POR QUE ISTO ESTA NA CONFERENCIA

    Um arquivo pode ter o XML bem formado e o texto igual paragrafo a paragrafo,
    e ainda assim o Word se recusar a abri-lo: basta uma faixa de bytes com duas
    edicoes, uma aplicada sobre a outra, que duplica conteudo sem quebrar a
    arvore. XML valido nao e documento valido, e so quem le o formato sabe dizer
    a diferenca.

    Abre com ReadOnly e fecha sem salvar, entao nao muda o arquivo.
    """
    try:
        import win32com.client.dynamic as _d
    except ImportError:
        return "pywin32 nao esta instalado; esta conferencia nao rodou"
    try:
        import win32com.client as _wc
        # Word proprio, e nao o do usuario. `Dispatch` se anexa ao Word que estiver
        # aberto, e o programa esconderia e fecharia a janela e os documentos de
        # quem esta trabalhando. `DispatchEx` abre uma instancia nova, e so ela e
        # fechada.
        w = _d.Dispatch(_wc.DispatchEx("Word.Application")._oleobj_)
    except Exception as e:
        return "nao consegui abrir o Word: %s" % e
    w.Visible = False
    w.DisplayAlerts = False
    try:
        d = w.Documents.Open(str(Path(caminho).resolve()), ReadOnly=True,
                             AddToRecentFiles=False)
        paginas = int(d.ComputeStatistics(2))
        d.Close(0)
        if paginas < 1:
            return "o Word abriu e contou zero paginas"
        return None
    except Exception as e:
        return "o Word recusou abrir: %s" % e
    finally:
        try:
            w.Quit()
        except Exception:
            pass


def main():
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("trabalho")
    ap.add_argument("--saida", required=True)
    ap.add_argument("--corpo", type=float, default=13.0)
    ap.add_argument("--entrelinha", type=float, default=1.3)
    ap.add_argument("--fonte", default="Times New Roman")
    ap.add_argument("--impressao", action="store_true",
                    help="versao final de impressao: liga a hifenizacao")
    ap.add_argument("--sem-word", action="store_true",
                    help="nao pede ao Word que abra o resultado (nao recomendado)")
    a = ap.parse_args()

    origem, destino = Path(a.trabalho), Path(a.saida)
    if not origem.exists():
        sys.exit("nao encontrei %s" % origem)
    if origem.resolve() == destino.resolve():
        sys.exit("a saida nao pode ser o arquivo de entrada.")
    destino.parent.mkdir(parents=True, exist_ok=True)

    if a.impressao:
        NORMA["hifenizar"] = True
    res = formatar(origem, destino, a.corpo, a.entrelinha, a.fonte)

    # O relatorio vai tambem para um anexo ao lado do arquivo, como na Norma
    # antiga: quem recebe o arquivo formatado precisa saber o que mudou nele, e o
    # que o programa diz so aparecia no terminal de quem o rodou.
    import io
    anexo = destino.with_name(destino.stem + "-relatorio.md")
    buf, tela = io.StringIO(), sys.stdout

    class _Tee:
        def write(self, s):
            tela.write(s)
            buf.write(s)

        def flush(self):
            tela.flush()

    sys.stdout = _Tee()
    try:
        _relatar(res, a, origem, destino)
    finally:
        sys.stdout = tela
        anexo.write_text(
            "# O que a formatacao mudou em %s\n\n" % origem.name
            + "Gerado por `formatar.py`. Este anexo acompanha o arquivo formatado,\n"
            + "para que quem o recebe saiba o que mudou nele.\n\n```\n"
            + buf.getvalue() + "```\n", encoding="utf-8")
        print("  anexo do relatorio em %s" % anexo)
        print()


def _relatar(res, a, origem, destino):
    print()
    print("FORMATACAO DE %s" % origem.name)
    print("  norma: %s, %g pt, entrelinha %g" % (a.fonte, a.corpo, a.entrelinha))
    print("  o pre-textual vai do paragrafo 0 ao %d" % (res["inicio_textual"] - 1))
    print()
    print("  ESTILOS CRIADOS   %s" % (", ".join(res["criados"]) or "nenhum"))
    print("  ESTILOS BLINDADOS %s" % (", ".join(res["blindados"]) or "nenhum"))
    print("  ESTILOS SOLTOS   %s (deixam de fixar corpo, e seguem o Normal)"
          % (", ".join(res["soltos"]) or "nenhum"))
    print("  ESTILOS DO SUMARIO %s" % (", ".join(res["sumarios"]) or "nenhum"))
    print()
    print("  PARAGRAFOS TOCADOS")
    for k, rotulo in (("pretextual", "pre-textual, que passa a ter estilo proprio"),
                      ("tabela", "dentro de tabela, idem"),
                      ("titulo", "titulos, que passam a usar estilo de titulo"),
                      ("legenda", "legendas"),
                      ("figura", "figuras, que passam a ter estilo proprio"),
                      ("citacao", "paragrafos recuados, que viram citacao"),
                      ("referencia", "entradas da bibliografia"),
                      ("lista", "entradas das listas do pre-textual"),
                      ("aprovacao", "folha de aprovacao aberta em folha propria"),
                      ("dedicatoria", "titulo da dedicatoria, que a NBR nao quer"),
                      ("numero", "numeros de titulo digitados, que saem ou ganham tabulacao"),
                      ("legenda-seq", "titulos de ilustracao com numero em campo SEQ"),
                      ("lista-automatica", "listas de ilustracoes trocadas por campo"),
                      ("natureza", "paragrafos recuados da folha de rosto, ao meio"),
                      ("epigrafe", "titulo da epigrafe, que a NBR nao quer"),
                      ("vazios-ilustracao", "vazios em volta das ilustracoes, apagados"),
                      ("vazios-referencia", "vazios entre as referencias, apagados"),
                      ("vazios-corpo-apagados", "vazios do corpo encostados em titulo, citacao ou ilustracao"),
                      ("espacos", "trechos com espaco repetido, que virou um so"),
                      ("separadores", "ajustes no separador das notas de rodape"),
                      ("palavras-chave", "paragrafos do resumo com corpo escrito a mao"),
                      ("corpo", "corpo do texto, cuja forma passa ao Normal")):
        print("    %5d  %s" % (res["mudou"][k], rotulo))
    print("    %5d  paragrafos da capa que usavam estilo de titulo sem ser titulo"
          % res["mudou"]["capa-destitulada"])
    print()
    print()
    print("  A PAGINA")
    print("    margens: %s, escritas em %d secao(oes)"
          % (", ".join("%s %.1f cm" % (k, v)
                       for k, v in NORMA["margens_cm"].items()), res["margens"]))
    print("    hifenizacao: %s" % (
        "ligada, para a versao de impressao" if res["hifen"]
        else "desligada: o arquivo a tinha ligada, e ela so vai na versao de impressao"
        if res["hifen_desligada"] else "desligada; so a versao de impressao hifeniza"))
    if res["ids_mso"]:
        print("    estilo de texto colado da web (%s): %d paragrafo(s) passam ao "
              "Normal, e o estilo sai" % (", ".join(res["ids_mso"]), res["n_mso"]))
    print("    tampas de numero de pagina apagadas: %d" % res["tampas"])
    print("    pre-textual em secao propria, contada e nao numerada: %s"
          % ("sim" if res["secao_pretextual"] else "nao"))
    print("    dedicatoria no pe da folha, em secao propria: %s"
          % (res["aviso_dedicatoria"] or "sim"))
    print("    numeracao dos titulos: %s"
          % ("automatica, pela lista do Word ligada aos estilos Titulo 1 a 4"
             if res["numeracao_auto"] else "a que o trabalho ja tinha"))
    if res["trocas_chave"]:
        print()
        print("  PALAVRAS-CHAVE, como a NBR 6028 pede (conferir nomes proprios)")
        for antes_k, depois_k in res["trocas_chave"]:
            print("    antes:  %s" % antes_k)
            print("    depois: %s" % depois_k)
    if res["numero_faltava"]:
        print("    numero de pagina: faltava em %d secao(oes) da parte textual, e foi "
              "posto no cabecalho (%s)" % (res["numero_faltava"], ", ".join(res["numero_posto"])))
    print("    numero de pagina: %s"
          % ("alinhado a direita em %d paragrafo(s) de cabecalho" % res["pagina"]
             if res["pagina"] else "ja estava a direita, nada a fazer"))
    print()
    print("    %d paragrafo(s) vazio(s) apagado(s) antes de titulo que quebra"
          % len(res["apagados"]))
    print("    pagina; nenhum deles carregava quebra, ancora, campo ou marcacao.")
    # Cada secao sai mesmo vazia, com o zero: secao ausente nao distingue
    # "nada encontrado" de "nada conferido".
    avisos = (["na numeracao: '%s' era %s e passa a %s; as remissoes do texto a ele "
               "ficam erradas" % x for x in res["avisos_numero"]]
              + (["no numero de pagina: %s" % res["aviso_numero_pag"]]
                 if res["aviso_numero_pag"] else [])
              + ["nas legendas: %s" % x for x in res["avisos_legenda"]]
              + ["na capa: %s" % x for x in res["avisos_capa"]])
    print()
    print("  AVISOS: %d" % len(avisos))
    for x in avisos:
        print("    %s" % x)
    print()
    # A lista digitada sai inteira e a automatica a refaz a partir das legendas;
    # o que nao tem legenda de ilustracao nao volta, e e isso que se aponta.
    nao_volta = [t for t in res["removidos"] if re.match(r"(?i)\s*(tabela|table)\b", t)
                 or re.match(r"(?i)\s*(dedicat|ep[ií]grafe)", t)]
    print("  TEXTO DO AUTOR QUE SAIU E NAO VOLTA: %d" % len(nao_volta))
    print("  (entrada de tabela na lista, que a norma nao tem; titulo de dedicatoria ou")
    print("  epigrafe, que a norma nao poe. Conferir na copia com controle de alteracoes.)")
    for t in nao_volta[:20]:
        print("    %s" % t[:80])
    print()
    print("  PARECEM TITULO E NAO RECEBERAM ESTILO DE TITULO: %d"
          % len(res["titulos_por_marcar"]))
    print("  (o programa nao os converte; quem decide e quem escreveu)")
    for i, t in res["titulos_por_marcar"][:20]:
        print("    par. %4d  %s" % (i, t))
    print()
    print("  ESPACO POSTO A MAO, de 36 pt ou mais: %d" % len(res["empurroes"]))
    print("  (se era para abrir pagina, o certo e quebra de pagina antes)")
    for i, t, v, papel in res["empurroes"][:20]:
        print("    par. %4d  %4.0f pt  %-10s %s" % (i, v, papel, t))
    print()
    print("  SUMARIO: %s"
          % ("o digitado saiu e entrou um campo automatico"
             if res["trocou_sumario"] else "nao achei sumario digitado; nada feito"))
    print()
    falhas = conferir(res, destino)
    if falhas:
        print("  CONFERENCIA REPROVADA")
        for f in falhas:
            print("    %s" % f)
        sys.exit(1)
    print("  CONFERENCIA")
    print("    o texto e o mesmo, paragrafo a paragrafo")
    print("    o XML das partes alteradas e bem formado")
    if a.sem_word:
        print("    o Word NAO foi consultado, e sem isso as duas linhas acima")
        print("    nao bastam: XML valido nao e documento valido.")
    else:
        erro = abre_no_word(destino)
        if erro:
            print("    O WORD NAO ABRIU: %s" % erro)
            sys.exit(1)
        print("    o Word abre o arquivo e conta paginas")
    print()
    print("  gravado em %s" % destino)
    print()


if __name__ == "__main__":
    main()

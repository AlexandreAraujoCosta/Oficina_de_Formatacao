# Oficina de Formatação

Ferramentas sobre a **forma** de um trabalho acadêmico em `.docx`. É repositório
separado da Oficina de Orientação por causa do risco, e não do assunto: a
oficina de orientação lê e relata, e esta escreve arquivos.

Nenhuma ferramenta daqui altera o arquivo de quem escreveu o trabalho. O
formatador grava uma cópia, e a cópia com controle de alterações é feita pelo
próprio Word, para que o autor aceite ou recuse cada mudança.

## A Norma

A Norma formata um trabalho na norma do PMPD com um comando só, e entrega, ao lado
do arquivo, a cópia formatada, a mesma cópia com controle de alterações, o
relatório do que mudou e o PDF:

    python scripts/norma.py trabalho.docx

A hifenização fica desligada. Ela é a última providência, e só a versão final de
impressão a liga: `python scripts/norma.py trabalho.docx --impressao`.

Ela roda, em ordem, os programas abaixo, e para na primeira etapa que reprovar.
Para usá-la por um agente (Claude Code, ou VS Code com Copilot em modo agente),
o `AGENTS.md` traz as instruções que o agente segue, e a página da Norma nas
Oficinas Acadêmicas traz o pedido a colar.

## As ferramentas

| | o que faz |
|---|---|
| `norma.py` | o comando único: formata, faz a cópia com controle, exporta o PDF e confere a página |
| `formatar.py` | grava uma cópia do `.docx` na norma do PMPD, que está em `NORMA-PMPD.md` |
| `revisao.py` | faz, a partir do original e da cópia formatada, a cópia com controle de alterações |
| `mostrador.py` | mostra as formas que o `.docx` usa, e não escolhe nenhuma; não escreve nada |
| `exportar_pdf.py` | exporta `.docx` para PDF pelo Word, com a marcação desligada |
| `medir_pagina.py` | mede a geometria das páginas de um PDF: margem, mancha, linha, densidade |
| `ooxml.py` | a leitura de OOXML que os outros usam; não escreve nada |

E os controles, que provam que as conferências reprovam o que têm de reprovar:

| | o que prova |
|---|---|
| `controle_formatar.py` | desenha um `.docx` sintético, formata e confere cada regra da norma |
| `controle_norma.py` | desenha PDFs com folha em branco e referência quebrada, e exige que a Norma acuse |
| `controle_medir_pagina.py` | desenha páginas de geometria conhecida e exige que o medidor as leia |
| `controle_capa.py` | desmonta de propósito a capa de um trabalho e exige que o medidor acuse |

## O que é preciso

Python 3, PyMuPDF e pywin32, e o Word instalado no Windows. O formatador em si é
Python puro; o Word entra para conferir que o arquivo abre, para fazer a cópia
com controle de alterações, para atualizar sumário e listas, e para exportar o
PDF.

Os programas abrem um Word próprio, separado do que estiver aberto, e fecham só
esse. Pode-se trabalhar no Word enquanto eles rodam.

## Como usar cada etapa em separado

    python scripts/formatar.py trabalho.docx --saida trabalho-norma.docx
    python scripts/revisao.py trabalho.docx trabalho-norma.docx --saida trabalho-norma-com-controle.docx
    python scripts/exportar_pdf.py --saida DIR trabalho-norma.docx
    python scripts/medir_pagina.py DIR/trabalho-norma.pdf

O `formatar.py` grava, ao lado da cópia, um anexo com o relatório do que mudou
(`trabalho-norma-relatorio.md`), para quem recebe o arquivo. Ele recusa arquivo
com alteração controlada pendente e diz quem a assina: as alterações têm de ser
aceitas ou recusadas no Word antes.

Antes de confiar numa mudança do programa:

    python scripts/controle_formatar.py
    python scripts/controle_norma.py
    python scripts/controle_medir_pagina.py

## O que as conferências garantem, e o que não

O `formatar.py` confere, antes de terminar, que o texto é o mesmo parágrafo a
parágrafo, que marcadores, hiperlinks, chamadas de nota e comentários
continuam lá, que o XML fecha e que o Word abre o arquivo. O `revisao.py`
confere que aceitar todas as alterações devolve exatamente a cópia formatada.

Nenhuma delas vê a página. Capa fora do lugar, folha em branco e título partido
só aparecem no PDF, e por isso o resultado se confere exportando e olhando, ou
medindo com o `medir_pagina.py`.

A exportação tem de ser feita com a marcação desligada: com uma única alteração
controlada no arquivo, o Word reserva a coluna de balão em todas as páginas, e o
que se mede no PDF é o balão.

## Sobre material de terceiro

Nenhum trabalho de outra pessoa entra neste repositório, nem o que dele se
extrai, nem o registro de como os programas foram testados. O `.gitignore`
recusa `.docx`, `.pdf`, as pastas de trabalho, o registro e os prompts de
sessão.

# Norma: instruções para o agente

Você vai formatar um trabalho acadêmico em `.docx` na norma do PMPD com a Norma,
o programa `scripts/norma.py` deste repositório, e entregar o resultado a quem
pediu. A norma está em `NORMA-PMPD.md`.

## Antes de rodar

- **A Norma só roda em Windows com o Word instalado.** Se a máquina não for
  Windows, ou não tiver Word, diga isso e pare.
- Instale o que falta: `pip install pymupdf pywin32`.
- O trabalho tem de ser `.docx`. Se for `.doc`, peça que seja salvo como `.docx`
  no Word.
- Se a cópia formatada de uma rodada anterior estiver aberta no Word, a gravação
  falha: peça que seja fechada. O original pode ficar aberto.

## Rodar

Na pasta do trabalho:

    python Oficina_de_Formatacao/scripts/norma.py "NOME DO ARQUIVO.docx"

Ela grava, ao lado do trabalho, `…-norma.docx`, `…-norma-com-controle-de-alteracoes.docx`,
`…-norma-relatorio.md` e `…-norma.pdf`, e para na primeira etapa que reprovar.

**Se ela recusar por alteração controlada pendente**, explique que as alterações
têm de ser aceitas ou recusadas no Word, por quem tem o arquivo, e pare. Não
aceite nem recuse alteração de ninguém.

**Se ela parar em outra etapa**, relate a mensagem dela e a etapa. Não altere os
programas deste repositório para fazê-la passar, e não entregue o que ficou
gravado: arquivo de etapa reprovada não é entrega.

## Olhar as páginas

A Norma confere o arquivo e procura no PDF folha em branco e referência de
sumário quebrada. O resto só se vê na página. Renderize com o PyMuPDF e olhe,
pelo menos: a capa, a folha de rosto, o sumário, a primeira página de cada
capítulo e as páginas com ilustração. Procure bloco de capa fora do lugar ou
caído para a folha seguinte, título que não abre página, sumário desalinhado,
legenda separada da figura.

Diga o que viu, com o número da página. Página que você não abriu não entra no
relato como conferida.

## Relatar

Liste os quatro arquivos e diga para que serve cada um: a revisão se faz pela
cópia com controle de alterações, aceitando ou recusando no Word. Do relatório
(`…-relatorio.md`), traga o que pede decisão de quem escreveu: as trocas nas
palavras-chave, os parágrafos que parecem título e não receberam estilo, o
espaço posto à mão, as pausas trocadas por espaço, e os avisos.

## O que não fazer

- Não altere o arquivo do trabalho. A Norma escreve cópias.
- Não rode git além do clone. Não publique nada.
- Não envie o trabalho, nem trecho dele, a serviço externo.

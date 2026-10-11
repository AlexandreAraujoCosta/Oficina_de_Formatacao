# Norma de formatação do PMPD

Esta é a especificação que o `scripts/formatar.py` aplica. Ela existe porque o
programa precisa de uma norma para ler, e porque um formatador sem norma escrita
inventa uma que ninguém pediu.

**Estado:** decidida pelo orientador, e ainda não submetida à coordenação do
programa. Cada linha abaixo é
decisão de quem coordena, e não do programa. O que o programa garante é aplicar
o que está escrito aqui e dizer quando não conseguiu.

A norma foi cotejada com a NBR 14724:2024 (quarta edição, versão
corrigida de 01/04/2025), a NBR 6024:2012, a NBR 6027:2012 e a NBR 6028:2021, e
incorporou o que não comprometia o desenho. Onde uma regra vem de lá, a seção da
NBR vai citada; onde a norma do PMPD diverge, a divergência vai dita, na seção
própria e na lista do fim.

A norma mora em `NORMA`, no alto do `scripts/formatar.py`, num dicionário só.
Mudar um número ali muda o resultado; mudar aqui sem mudar lá não muda nada.

## A página

| | |
|---|---|
| papel | A4 (não alterado pelo programa) |
| margem esquerda | 3,3 cm |
| margem direita | 3,3 cm |
| margem superior | 3,3 cm |
| margem inferior | 2,5 cm |
| posição vertical do cabeçalho | no meio da margem superior (cerca de 1,45 cm da borda) |
| cabeçalho | só o número de página, à direita, em 10 pt e itálico |
| contagem | começa na folha de rosto; a capa é a folha 0 (NBR 14724, 5.3) |
| exibição | a partir da primeira folha da parte textual |
| hifenização | desligada; só a versão final de impressão (`--impressao`) a liga, com no máximo 2 hífens seguidos |

## Os estilos

### Normal, que é o corpo do texto

Times New Roman 13, entrelinha 1,3, justificado, recuo de primeira linha de
1,25 cm, 3 pt de espaço depois. Hifenizado.

O recuo e o espaço marcam a mesma coisa, que é o começo de parágrafo novo. O
recuo fica, porque é a marca clássica do livro; o espaço encolhe a 3 pt, o
bastante para o olho respirar sem marcar duas vezes.

O corpo do texto não escreve nada disso em si: a forma mora no `Normal`, e o
parágrafo apenas o segue. É essa a diferença entre um trabalho formatado por
estilo e um formatado à mão, e é o que permite mudar a norma inteira num lugar
só.

### O corpo menor

Citação e nota de rodapé vão em 12 pt; legenda e fonte da ilustração, em 11 pt e
em itálico. A NBR 14724 (5.1) quer esses elementos em tamanho menor e uniforme.
O número de página vai em 10 pt.
Onde o autor escreveu o corpo à mão nas corridas, ele sai, para que o estilo
governe: corpo escrito na corrida vence o do estilo.

### Título 1, das seções primárias

Alinhado à esquerda, 16 pt, negrito, entrelinha simples, 0 pt antes e 36 pt
depois, quebra de página antes. Recuo pendente de 1,25 cm, o mesmo do recuo de
primeira linha do corpo: o número fica na margem e o título começa onde começa o
parágrafo recuado. Sem hifenização. Entra no sumário. O 0 antes põe o título no
alto da mancha (NBR 14724, 5.2.2).

O Título 1 **sem número** (REFERÊNCIAS, APÊNDICES, ANEXOS) fica centralizado e
sem recuo (NBR 14724, 5.2.3; NBR 6024, 4.1, alínea h).

O título longo fica inteiro: encurtá-lo é decisão de quem escreve. (O programa
sabe dividi-lo no dois-pontos, com o complemento num parágrafo de subtítulo, e a
opção está desligada.)

### Título 0, dos títulos do pré-textual

A forma do Título 1, mas centralizado e sem recuo, porque não tem número, e
fora do sumário. É o que recebe RESUMO, ABSTRACT, as listas, os agradecimentos e
o próprio SUMÁRIO.

Fora do sumário quer dizer `outlineLvl` 9, que é o valor que o esquema do OOXML
reserva para corpo de texto. O parágrafo fica com a aparência de título e o campo
de sumário não o enxerga.

### Cabeçalho do Sumário

A forma do Título 0. É o estilo que o Word aplica à palavra SUMÁRIO quando ele
próprio insere o campo, e sem isto ele sai em Cambria azul.

### Título 2, das seções secundárias

Alinhado à esquerda, 13 pt, negrito, entrelinha simples, 18 pt antes e 12 pt
depois. Recuo pendente de 1,25 cm, com parada de tabulação no mesmo ponto: o
título começa na coluna do parágrafo recuado e a segunda linha alinha sob a
primeira letra dele (NBR 6024, 4.1, alínea i). Sem hifenização. Entra no
sumário.

A parada explícita é necessária. Parágrafo que traz `numPr` com `numId` 0, resto
de uma lista que o autor desligou, perde a parada implícita do recuo pendente, e o
Word também desfaz o recuo, e a segunda linha volta à margem; o programa tira
esse resto do título.

**Diverge da ABNT:** a NBR 14724 (5.2.2) quer o título de subseção separado do
texto por um espaço de 1,5 antes e depois. O espaço menor depois prende o título
ao texto que ele abre.

### O número do título

O número vem seguido de **tabulação**, sem ponto: `2⇥TÍTULO DO CAPÍTULO`,
`2.1⇥Título da seção`. Ponto que o autor tenha posto sai. A NBR 6024:2012
(4.1, alínea d) não admite sinal entre o indicativo e o título.

A tabulação, e não o espaço, é o que alinha os títulos numa coluna, no próprio
título e no sumário (NBR 6027, 5.2). Cada estilo numerado declara a sua parada:
sem ela, o Word usa as paradas padrão.

**A Introdução é numerada**, como as demais seções primárias. A NBR 14724:2024
(5.2.3) e a NBR 6024:2012 (4.1, alínea h) listam os títulos sem indicativo
numérico, e a Introdução não está entre eles; a numeração das seções primárias
começa em 1 (NBR 6024, 4.1, alínea f).

**A numeração é automática.** O número digitado pelo autor sai, e o Word passa a
numerar os títulos por uma lista ligada aos estilos: acrescentar, tirar ou mover
um título renumera os outros. A numeração progressiva vai até o Título 3
(`2.1.3`); o Título 4 leva só uma letra, `a)`, que recomeça a cada Título 3, na
forma da alínea (NBR 6024, 4.2). Cada número tirado vai
registrado, para que a conferência de texto saiba o que esperar, e o relatório
avisa quando a sequência digitada não é a que a lista dá, porque aí as remissões
do texto ("ver seção 4.4") ficam erradas; o de quarto nível, que passa de número
a letra, é sempre avisado. O título sem número recebe a lista
desligada, para que o estilo não o numere.

O trabalho que já numera algum título pela lista do Word fica com a numeração
dele: trocar a lista do autor apagaria os números dos títulos que ela numera. Aí
o número digitado que houver passa a ser seguido de tabulação.

### Apêndice e anexo

Forma de seção primária, centralizados e sem número, qualquer que seja o estilo
que o autor lhes deu (NBR 14724, 4.2.3.3 e 4.2.3.4; NBR 6024, 4.1, alínea h).

### Os demais níveis de título

Entrelinha simples e sem hifenização. Os Títulos 3 e 4 levam recuo pendente e
parada a 1,25 cm, como o Título 2. O resto fica como o arquivo trouxer.

### Citação, que é o parágrafo recuado

12 pt. Recuo esquerdo de 2,5 cm, sem recuo de primeira linha, entrelinha
simples, 12 pt antes e 12 pt depois.

Reconhece-se pelo recuo esquerdo declarado no parágrafo, acima de 1.000 twips.
É heurística, e por isso o programa diz quantos achou.

### Figura ou gráfico

Centralizado, sem recuo, sem espaço antes ou depois, e preso ao parágrafo
seguinte, que é a fonte. Reconhece-se pelo parágrafo que tem desenho e não tem
texto. Parágrafo que tem desenho **e** texto não é figura: é texto com âncora,
como costuma acontecer na capa e na dedicatória.

### Legenda

É o estilo de legenda do próprio Word, o `caption` que "Inserir legenda" aplica,
e não um estilo criado pelo programa. Centralizada, 11 pt, em itálico e em
preto, 12 pt antes e 12 pt depois. O título e a fonte dividem o estilo, e os dois
saem assim; o negrito que o autor pôs à mão no título fica. Recebe:

- o título da ilustração ou da tabela: o parágrafo que começa por Gráfico,
  Tabela, Figura, Quadro ou Imagem seguido de número. O separador que a NBR pede
  pode faltar (`Gráfico 1 Negativas...`), e então o que o distingue do texto
  corrido é a maiúscula logo depois do número;
- a fonte ou a nota: o parágrafo que começa por Fonte, Nota, Elaboração ou
  equivalente, logo abaixo de uma figura ou de uma tabela, com no máximo dois
  vazios no meio.

O parágrafo seguinte à fonte não entra, mesmo que seja curto.

**O bloco da ilustração fica junto.** O título leva 12 pt depois, escritos no
próprio parágrafo, e fica preso à figura; o bloco se separa do texto por 12 pt.
Os parágrafos vazios que o autor pôs em volta da ilustração saem, porque o
espaço agora vem do estilo, e porque um vazio entre o título e a figura quebra a
cadeia que os prende: o título pode ficar no pé de uma folha e a figura ir
para a seguinte.

**O número da ilustração vai num campo `SEQ`**, como o Word faz ao inserir
legenda, e o texto não muda. É pelo campo que a lista de ilustrações é montada.
Isso só se faz quando o autor numerou as ilustrações na ordem em que aparecem:
o campo renumera ao ser atualizado, e numeração fora de ordem viraria outra, com
as remissões do texto apontando para a ilustração errada. Nesse caso o programa
avisa e deixa o número como está.

### Referência bibliográfica

Espaço simples, 12 pt depois, sem recuo de primeira linha, alinhada à esquerda.
Os 12 pt fazem a linha em branco entre referências que a NBR 14724 (5.2) pede, e o
parágrafo vazio que o autor pôs para isso sai.

A faixa vai do título REFERÊNCIAS, numerado ou não (`5 REFERÊNCIAS`), até o
próximo título de primeira ordem.

### Lista de ilustrações

Gerada pelo Word, por um campo que lê os `SEQ` das legendas, com número de
página e pontilhado (NBR 14724, 4.2.1.9). A lista digitada sai. Recuo pendente
de 2,2 cm, que é a largura de `Gráfico 5 – ` em Times 13 medida pela métrica da
fonte: a continuação de uma entrada longa alinha sob o título dela.

Não há lista de tabelas, por decisão da coordenação, e o ` E TABELAS` sai do
título da lista.

### Entrada de lista

Da lista de siglas. Recuo pendente de 1,2 cm, para que a continuação de uma
entrada longa alinhe sob o texto da primeira linha em vez de voltar à margem e
colar na sigla seguinte. Entrelinha simples, alinhada à esquerda, sem número de
página, que a lista de siglas não tem.

### Resumo

Entrelinha simples, sem recuo de primeira linha, justificado. A NBR 14724 põe em
simples a família a que o resumo pertence, que é a do bloco de prosa contínua
lido como unidade: natureza do trabalho, notas, referências e legendas. O
ABSTRACT acompanha.

**As palavras-chave seguem a NBR 6028:2021 (4.1.7):** separadas por ponto e
vírgula, terminadas por ponto, com a inicial minúscula, salvo nome próprio. Nome
próprio aqui é a expressão em que toda palavra que não é de ligação começa por
maiúscula (`Universidade de Brasília`), ou que traz sigla. É critério de máquina
e erra no nome próprio de uma palavra só, como `Brasil`; por isso o programa
lista cada troca no relatório, para o autor conferir. O corpo escrito à mão sai,
e o negrito do rótulo `Palavras-chave:` fica.

### Nota de rodapé

12 pt, entrelinha simples. O filete que separa as notas do texto começa na margem
esquerda (NBR 14724, 5.2.1): sem recuo próprio, ele herda o de primeira linha
do Normal e começa para dentro da margem. O parágrafo vazio a mais que costuma vir
com ele sai.

### Pré-textual

Segue o corpo do texto, e vai em **entrelinha simples**: ali não há prosa corrida
que peça respiro, há nome, título, lista e resumo. Sem espaço depois, porque o
espaço entre blocos vem da posição deles.

Estilo usado só pelo pré-textual solta o corpo que declarava, para seguir o
`Normal`. O que ele declara de próprio, como um recuo, fica.

O texto dos agradecimentos é prosa corrida, e vai no estilo `TextoPreTextual`:
justificado e com o recuo de primeira linha do corpo, na entrelinha simples do
pré-textual. O recuo e o alinhamento escritos à mão saem.

### Tabela do trabalho

Não é forma: é defesa. Devolve a quem herda o corpo e a entrelinha que o `Normal`
tinha antes da norma. Sem ela, mudar o `Normal` alcança todas as células de
tabela. O corpo menor não chega à tabela, que a NBR 14724
(5.9) remete às normas do IBGE.

## A capa e a folha de rosto

**A capa não é feita de linhas em branco.** É comum o autor empurrar cada bloco
para o seu lugar com parágrafos vazios, contando linhas. Isso amarra a diagramação ao corpo da letra e à altura de página que
vigoravam quando ele contou, e qualquer mudança de margem a desmonta.

A regra é posicional. Os blocos são as corridas de parágrafos com texto, e a
norma diz onde cada um começa, em fração da altura útil da mancha:

| bloco | onde começa |
|---|---|
| instituição e programa | no alto |
| autor | 25% |
| título e subtítulo | 45% |
| local e ano | termina 0,5 cm acima do pé da mancha |

A folga de 0,5 cm no pé não é enfeite. A altura dos blocos é estimada, e sem ela
um erro de meia linha derruba a última linha para a folha seguinte.

O local e o ano ficam em **duas linhas**, e não em `Brasília – 2026`. São dois
elementos distintos na norma, e a forma de uma linha obriga a escolher entre
hífen, travessão e vírgula, escolha que a norma não faz e que varia de estudante
para estudante.

O programa recalcula a cada mudança de margem ou de corpo, e por isso a capa não
volta a transbordar. A capa não se hifeniza, porque ali só há nomes e título.

**A folha de rosto não é reinventada.** Ali a diagramação é do autor. O que o
programa faz é trocar cada corrida de parágrafos vazios por espaço antes, na
altura que ela tinha: o texto cresce e o vazio fica onde estava.

**A natureza do trabalho começa no meio da mancha**, rente, sem recuo de primeira
linha e em espaço simples, na folha de rosto e na de aprovação (NBR 14724, 5.2).
O espaço simples precisa ser escrito no parágrafo: ele guarda o estilo do autor,
que herda a entrelinha do Normal. Vale para o
parágrafo que o autor já tinha recuado além de 1.000 twips, o que inclui as
linhas de assinatura, que guardam o recuo pendente que tinham. É o mesmo eixo da
dedicatória: um só alinhamento vertical para o pré-textual.

**Cada elemento pré-textual começa em folha nova**, pelo Título 0.

**A folha de aprovação vai em folha própria**, depois da folha de rosto, como a
NBR 14724 pede. O programa abre a folha, e nada mais: nome, título, natureza e
data de aprovação, que a norma também exige ali, são texto que o autor escreve.

**A dedicatória e a epígrafe não levam título** (NBR 14724, 5.2.4).

**A dedicatória e a epígrafe ficam no pé da folha**, sem recuo de primeira linha
e com recuo esquerdo de metade da mancha (NBR 14724, 5.2.4). Na folha A4 com
margens de 3,3 cm, o texto começa em 10,5 cm da borda, que é o meio da folha. O pé
vem de uma seção própria com alinhamento vertical inferior, e não de espaço
calculado: se o autor reescrever o texto, ele continua no pé. Cada uma tem a sua
seção, mesmo quando uma vem logo depois da outra. As seções novas usam o
cabeçalho vazio do pré-textual e não reiniciam a contagem.

## O cabeçalho

Só o número de página, à direita, em 10 pt e itálico, por decisão da coordenação.
A linha fica no meio da margem superior: a distância da borda é a metade do que
sobra da margem depois da altura da linha. O corpo e o itálico vêm do estilo
`Cabealho`, que o programa cria quando o arquivo não o tem e aplica ao parágrafo
do número; o corpo escrito nas corridas sai. O número fica na margem direita do
texto, e não a 2 cm da borda como a NBR 14724 (5.3) pede.

O programa sabe pôr o título do capítulo no cabeçalho, pelo campo `STYLEREF 1`,
com o número dele quando a numeração é automática; a opção está desligada.

O pré-textual não tem cabeçalho, porque está em outra seção.

## A numeração

A NBR 14724 (5.3) manda contar as folhas a partir da folha de rosto e não
numerar as pré-textuais. A capa é a folha 0, o pré-textual vira uma seção com
cabeçalho próprio, que não traz o campo de página, e a seção textual mantém o
cabeçalho com o número. O Word continua a contagem de uma para a outra sozinho.

**O estilo `msonormal`**, que o Word cria quando se cola texto de página da web
ou de e-mail, passa ao `Normal`: o parágrafo que o usa perde o estilo, e o estilo
sai do arquivo.

**Trabalho sem número de página ganha o número.** Seção da parte textual que não
mostra número, nem no cabeçalho nem no rodapé, recebe o campo de página: no
cabeçalho dela, se ela tem um, e num cabeçalho novo, se ela herdaria o cabeçalho
vazio do pré-textual. Número posto pelo autor no rodapé fica onde está. Se o
pré-textual não ficou em seção própria, o número sairia também na capa, e o
programa não o põe: o relatório avisa.

Há trabalho que resolve o mesmo problema com retângulos brancos sobre fundo
branco, um por folha, cobrindo o número. O programa apaga a
caixa de texto que esteja vazia e tenha preenchimento e contorno iguais ao fundo,
e deixa em paz caixa com conteúdo, imagem e gráfico.

## O que o programa reconhece, e onde

**A parte textual começa no primeiro título de nível 1**, numerado (`1 INTRODUÇÃO`)
ou marcado só pelo estilo de Título 1 (`INTRODUÇÃO`), desde que não seja título do
pré-textual nem entrada de índice. Se não houver nenhum, o programa para em vez
de adivinhar.

**Título** é o parágrafo numerado de nível 1 a 3, ou o que usa estilo de Título 1
a 3. O de terceiro nível recebe o Título 3.

**Índice automático fica intocado**: o sumário e a lista de ilustrações que o
Word gera por campo, reconhecidos pelo campo TOC e pelos estilos de sumário e de
índice de ilustrações. Suas entradas numeradas não são títulos: sem isto, cada
entrada do sumário seria lida como título de capítulo. A lista de ilustrações que
já é campo não é trocada.

**O vocabulário vale em português e em inglês**: References, Bibliography,
Appendix, Annex, Abstract, Contents, List of, Acknowledgements, Dedication,
Epigraph; Figure, Chart, Graph, Table nas legendas; Source e Note nas fontes.

**O corpo escrito à mão sai também dos títulos e de todas as linhas da capa.**
Corpo escrito em cada corrida vence o do estilo, e entrelinha escrita em cada
linha da capa faz os blocos crescerem além da conta que os posiciona.

**Dividir um título no dois-pontos não perde nada do parágrafo.** Marcador,
comentário e pedaço de campo vão para o lado em que estão. Quando o corte cairia
dentro de um hiperlink ou de um campo, o título não é dividido.

**A conferência conta a estrutura**, além do texto: marcadores, hiperlinks,
chamadas de nota e comentários. O que o arquivo tinha, menos o que se tirou de
propósito, tem de ser o que ele tem.

## O que veio da Norma antiga

A Norma era o programa anterior desta oficina. Ele alinhava cada papel à forma que já predominava no
arquivo, e isso a norma do PMPD substituiu. Algumas peças dele valiam por si e
foram trazidas:

- **A porta das alterações controladas.** Arquivo com alteração controlada por
  decidir não é formatado: o programa para e diz quem assina cada uma. As marcas
  de quem revisou se misturariam às da formatação na cópia com controle de
  alterações, e aceitar tudo depois apagaria o trabalho dessa pessoa.
- **As travas completas.** Parágrafo vazio com quebra de página antes, quebra de
  coluna, fim de marcador, referência de comentário ou texto apagado sob revisão
  não é espaço em branco, e não se apaga. A exceção é o vazio logo antes de um
  título que já quebra página: ali a quebra dele repete a do título.
- **Espaço repetido** vira um só, da parte textual em diante. Na capa ele às
  vezes faz alinhamento, e fica.
- **Os vazios do corpo.** O encostado em título, citação ou ilustração sai,
  porque o estilo já dá o espaço. O que vem depois de parágrafo de corpo também
  sai: o corpo abre com recuo de primeira linha, e o recuo já marca o parágrafo
  novo.
- **O anexo do relatório**, gravado ao lado do arquivo formatado
  (`...-relatorio.md`): quem recebe o arquivo precisa saber o que mudou nele.
- **O que só se aponta.** Parágrafo que parece título e não recebeu estilo de
  título, e espaço de 36 pt ou mais posto à mão, que costuma ser a gambiarra de
  empurrar conteúdo para a folha seguinte. O programa não converte nem corrige:
  diz onde está, e quem decide é quem escreveu.

## Onde a norma do PMPD diverge da ABNT

Todas contra um "deve" da norma, e mantidas por decisão da coordenação:

- **Entrelinha de 1,3**, contra 1,5 (NBR 14724, 5.2).
- **Margens de 3,3 cm e 2,5 cm**, contra 3 cm e 2 cm (NBR 14724, 5.1). As da
  NBR são assimétricas por causa da encadernação.
- **Título 2 com 18 pt antes e 12 pt depois**, contra 1,5 antes e depois
  (NBR 14724, 5.2.2).
- **Títulos em entrelinha simples**, contra 1,5 (NBR 14724, 5.2).
- **Número de página na margem do texto e no meio da margem superior**, contra
  2 cm da borda direita e 2 cm da borda superior (NBR 14724, 5.3).
- **Três corpos menores em vez de um**: 12 pt para citação e notas, 11 pt para
  legenda e fonte, 10 pt para cabeçalho e número de página. A NBR 14724 (5.1)
  quer esses elementos num tamanho só.

## O que a norma não decide, e por isso o programa não faz

**Parágrafo vazio travado.** O vazio que carrega quebra de seção, quebra de
página ou de coluna, âncora de imagem ou de nota, marcador, comentário, campo ou
marcação de revisão fica onde está.

**Fonte e corpo dentro de corrida de texto**, onde a aparência muda palavra a
palavra. As exceções são a capa, o resumo, as legendas e as citações, onde o
corpo escrito à mão sai para que o estilo governe.

**A forma das referências**, que é a NBR 6023 e o estilo `pmpd.csl` do Zotero, e
não matéria de estilo de parágrafo.

**A remissão do texto a capítulos e seções.** O texto cita "capítulo 3" pelo
número digitado; o programa não reescreve prosa.

## O que ainda não está decidido

- **A legenda como o Word a insere pela figura.** O orientador quer a legenda que
  o Word cria ao clicar com o botão direito na figura, como parte dela. O que o
  Word grava nesse caso ainda não foi examinado: um parágrafo à parte, como o
  programa faz hoje, ou algo preso à própria imagem. A largura da legenda
  (NBR 14724, 5.8) depende disso.
- A fonte. Times New Roman está em toda máquina Windows e **não é de uso livre**.
  Um modelo distribuído aos estudantes precisa de fonte livre e incorporada no
  arquivo, porque `.docx` não carrega fonte: sem ela instalada, o Word substitui
  em silêncio e a norma inteira se desfaz. Candidatas conferidas no CTAN:
  Libertinus (SIL OFL, com Serif, Sans, Mono e Math desenhados juntos),
  EB Garamond (LPPL e OFL) e ETbb (LPPL 1.3 e MIT), que é a do tema do memorial.

---
key: right-of-way-monitoring
template: solution
title: Levantamento de Ocupações na Zona de Protecção | AfriScan
description: Registo das construções a menos de 50 e 100 m do seu gasoduto, linha de energia, estrada ou ferrovia em Moçambique, com distância ao eixo e troços prioritários.
h1: Levantamento das construções a menos de 50 m da sua infra-estrutura
crumb: Levantamento de ocupações
eyebrow: Proteger corredores e locais
lead: Envie o traçado. Devolvemos a lista das construções por faixa, a distância de cada uma ao eixo e os troços de 500 m com mais ocupação, prontos para a equipa de servidões ou para o plano de regularização da faixa. Depois, se quiser, voltamos a levantar o traçado com a periodicidade acordada e dizemos-lhe o que mudou. Uma pessoa revê cada resultado antes da entrega.
used_in: [oil-gas, power-utilities, rail-roads]
buttons:
  - {label: Pedir proposta, intent: proposal}
  - {label: Ver o exemplo do gasoduto, key: results}
related: [change-detection, mz-protection-zone, oil-gas]
service:
  name: Levantamento de ocupações na zona de protecção parcial
  type: Levantamento de ocupações em faixas de servidão e levantamentos periódicos
  description: Registo das construções dentro das distâncias que se aplicam a um gasoduto, oleoduto, linha de energia, estrada ou ferrovia em Moçambique, com a distância ao eixo, as coordenadas e uma classificação de densidade de ocupação por troço de 500 m, seguido de levantamentos periódicos e de um aviso das alterações. Revisto por uma pessoa e entregue em PDF e ficheiros SIG.
og:
  headline: Levantamento de ocupações na zona de protecção
  subline: Construções a menos de 50 e 100 m, medidas até ao eixo e revistas por uma pessoa
cta:
  title: Envie o traçado e as larguras que se aplicam
  text: KML, KMZ, GeoJSON, Shapefile, GPX ou GeoPackage, ou desenhamos a linha consigo. Respondemos com o âmbito, um plano de imagens e uma proposta escrita.
  button: Pedir proposta
faq:
  - q: Que infra-estruturas podem levantar?
    a: Qualquer infra-estrutura linear com um traçado conhecido, como gasodutos, oleodutos, adutoras, linhas de transporte e distribuição de energia, estradas e ferrovias, e também áreas como concessões, instalações ou locais de projecto, dentro do limite, numa faixa à sua volta ou em ambos.
  - q: Com que larguras trabalham?
    a: Por defeito, 50 e 100 metros. Acrescentamos as larguras da lei que se aplicam, como os 200 metros do corredor Pande–Temane, as do contrato de concessão ou a norma da sua empresa, até seis distâncias por levantamento.
  - q: O registo substitui a fiscalização no terreno?
    a: Não. Diz à sua equipa onde ir primeiro e o que vai encontrar. A verificação no terreno, o contacto com as comunidades e as decisões sobre cada construção ficam com a sua equipa e com as autoridades.
  - q: E se a nossa linha atravessar zonas com muita vegetação?
    a: Construções debaixo de árvores ou com coberturas que se confundem com o solo podem não se ver. Nesses troços, o revisor marca à mão o que a imagem permite confirmar, o relatório assinala o que fica por verificar e, se for preciso, um levantamento por drone do troço dá mais detalhe, sujeito às licenças e autorizações que cada trabalho exige.
  - q: O resultado é um levantamento topográfico ou cadastral?
    a: Não. É um registo de construções feito a partir de imagens, para gerir a faixa. A demarcação de direitos sobre a terra é feita pelos Serviços de Cadastro ou por um agrimensor ajuramentado.
---

::::section{id="o-que-e" eyebrow="O que é" title="Um registo datado do que está dentro da sua faixa"}
:::::columns{split="1-1"}
::::col
Um levantamento de ocupações responde a três perguntas: que construções estão dentro das faixas que se aplicam à sua infra-estrutura, a que distância do eixo está cada uma, e em que troços se concentram. É o ponto de partida para gerir uma faixa de servidão: para a equipa de servidões priorizar as visitas, para o plano de regularização, para o diálogo com as comunidades e, numa linha nova, para o plano de reassentamento.

Cada levantamento fica associado às imagens usadas e à respectiva data, o que permite comparar levantamentos e responder, mais tarde, à pergunta que a Lei de Electricidade torna central: o que já estava lá, e o que apareceu depois.
::::
::::col
### As faixas da lei, em resumo

| Infra-estrutura | Faixa | Fonte |
|---|---|---|
| Condutas de petróleo, gás e água; linhas de electricidade e telecomunicações | 50 m de cada lado | Lei n.º 19/97, art. 8, al. g) |
| Infra-estruturas petrolíferas | 50 m | Lei n.º 8/2026, art. 75, n.º 3 |
| Linhas de energia (servidão) | até 50 m do eixo | Lei n.º 12/2022, art. 43, n.º 4 |
| Corredor Pande–Temane | 50 m e 200 m | Decreto n.º 36/2001, arts. 1 e 2 |
| Linhas férreas | 50 m de cada lado do eixo | Lei n.º 19/97, art. 8, al. f) |
| Estradas primárias; secundárias e terciárias | 30 m; 15 m | Lei n.º 19/97, art. 8 |

Situação a 26 de Setembro de 2026; resumo informativo, não constitui aconselhamento jurídico. [O guia completo da zona de 50 metros](/mz/pt/zona-de-proteccao-parcial-50-metros).
::::
:::::
::::

::::section{id="entrada" tone="alt" eyebrow="O que nos envia" title="Do seu SIG para o nosso, sem visita ao local"}
:::checklist
- **O traçado ou o limite**, em KML, KMZ, GeoJSON, Shapefile, GPX ou GeoPackage, tal como está no seu SIG. Sem ficheiro, desenhamos o traçado consigo e enviamo-lo para confirmação.
- **As larguras a medir**: as da lei, as do contrato de concessão ou as da norma da empresa.
- **A data que o registo deve reflectir**, se houver uma, porque a escolha das imagens depende dela.
- **As imagens que já tem**, como ortofotomapas de drone ou cenas de satélite georreferenciadas.
- **A língua do relatório** e o prazo.
:::
::::

::::section{id="saida" eyebrow="O que recebe" title="Uma lista de construções, e uma lista de troços a visitar primeiro"}
:::::columns{split="1-1" align="center"}
::::col
:::figure{src="diagrams/corridor-pt" alt="Esquema de um traçado de 2,5 km com as faixas de 50 m e 100 m, as construções coloridas por faixa e uma barra por baixo com a densidade de ocupação de cada troço de 500 m: baixa, média, alta, média, baixa" caption="Como se lê um levantamento de corredor" credit="Esquema desenhado pela AfriScan para ilustração; não é um local real." size="half"}
<span class="band band--a">Até 50 m</span> <span class="band band--b">50–100 m</span> <span class="band band--c">Além de 100 m</span> Cada troço de 500 m é classificado pelas construções dentro da faixa mais larga.
:::
::::
::::col
- **O registo**: cada construção com referência, distância ao eixo, faixa, distância ao longo do traçado e coordenadas em WGS84 e UTM.
- **A densidade de ocupação por troço**: alta com mais de cinco construções dentro da faixa mais larga, média com uma a cinco, baixa com nenhuma. É uma regra de contagem para priorizar visitas, não uma avaliação de segurança.
- **O relatório PDF**, em português ou em inglês, com o mapa geral, a tabela de troços, uma fotografia de cada construção e o registo de coordenadas.
- **As camadas SIG** (GeoPackage, GeoJSON, KMZ e Shapefile) e um **mapa interactivo em ficheiro** para as equipas no terreno.

[Ver o exemplo do gasoduto](/mz/pt/resultados-de-exemplo)
::::
:::::
::::

::::section{id="como" tone="alt" eyebrow="Como funciona" title="Primeiro o levantamento de base, depois os levantamentos periódicos"}
:::steps
:::step{title="Âmbito"}
Acordamos o traçado, as larguras, as imagens e os produtos a entregar, por escrito, na proposta.
:::
:::step{title="Rastreio"}
Bases de dados abertas de edifícios e modelos de segmentação aplicados às imagens do trabalho propõem as construções.
:::
:::step{title="Revisão"}
Um revisor confirma, corrige e acrescenta construções, e lista para verificação no terreno o que a imagem não permite decidir.
:::
:::step{title="Periodicidade"}
Se quiser, voltamos a levantar o traçado com a periodicidade acordada consigo e enviamos um aviso a dizer o que mudou e onde.
:::
:::
::::

::::section{id="proposta" eyebrow="A proposta" title="O que define o âmbito de cada proposta"}
Cada proposta é preparada para o seu traçado ou terreno. O âmbito depende da extensão da linha ou da área, das larguras a medir, das imagens de que o trabalho precisa (as suas, cenas de catálogo, uma nova captação ou um levantamento por drone), dos produtos a entregar, da língua do relatório, do prazo e, se houver, da periodicidade dos levantamentos seguintes. Tudo isto fica escrito antes de qualquer trabalho começar.
::::

::::section{id="servicos" tone="alt" eyebrow="Serviços" title="Os serviços por trás do levantamento"}
:::catalogue{services="S01,S08,S09,S10,S18,S22,S13,S38,S43,S44,S45"}
:::
::::

::::section{id="limites" eyebrow="Limites" title="O que um levantamento de ocupações não é"}
:::::columns{split="1-1"}
::::col
### O que mostra
:::checklist
- As construções que a revisão confirma dentro das faixas, na data das imagens
- A distância de cada uma ao eixo e os troços onde se concentram
- O que ficou por verificar no terreno, e porquê
:::
::::
::::col
### O que não mostra
:::checklist{tone="no"}
- Construções que as árvores ou as coberturas escondem por completo
- Quem vive, a quem pertence ou para que serve cada construção
- Se uma construção está autorizada ou tem direito a compensação
- Limites cadastrais ou direitos sobre a terra
:::
::::
:::::

:::callout{tone="note" title="Construções perto do limite de uma faixa"}
As imagens têm pequenos desvios de posição, por isso uma construção a poucos metros do limite de uma faixa pode cair de um lado ou do outro. O relatório assinala esses casos para verificação no terreno.
:::
::::

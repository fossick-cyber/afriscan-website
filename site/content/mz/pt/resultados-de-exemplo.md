---
key: results
icon: map
title: Exemplo de Levantamento de Ocupações num Gasoduto | AfriScan
description: "O que um levantamento AfriScan entrega: construções ao longo do traçado, distância ao eixo, registo das faixas de 50 e 100 m, densidade por troço, PDF e SIG."
h1: Exemplo de um levantamento de ocupações
crumb: Exemplo de resultados
section: resources
nav_group: resources
nav_order: 10
nav_label: Exemplo de resultados
nav_blurb: Uma revisão feita num gasoduto de alta pressão em Moçambique
eyebrow: Recursos
lead: Resultados reais de uma revisão feita no traçado de um gasoduto de alta pressão em Moçambique, mostrados com a autorização do proprietário do traçado. Todas as marcações são manuais, feitas por um revisor, e aparecem exactamente como foram registadas; a versão pública omite as coordenadas.
buttons:
  - {label: Pedir o exemplo anonimizado, intent: sample-report}
  - {label: Como trabalhamos, key: how-we-work}
og:
  headline: Exemplo de um levantamento de ocupações
  subline: Um gasoduto de alta pressão em Moçambique · faixas de 50 e 100 m · marcações do revisor
cta:
  title: Quer ver o seu próprio corredor?
  text: Envie o ficheiro do traçado (KML, GeoJSON ou Shapefile) e as distâncias que contam para a sua empresa, e preparamos o âmbito de um levantamento da sua linha ou do seu terreno.
  button: Enviar o traçado
  intent: proposal
faq:
  - q: Este exemplo é detecção automática?
    a: Não. Cada marcação deste exemplo foi colocada por um revisor; não se mostra nenhum resultado de detecção automática. Num levantamento para um cliente, as bases de dados abertas de edifícios e os modelos de segmentação propõem primeiro as construções, e um revisor confirma, corrige e acrescenta.
  - q: Porque é que as vistas de perto usam imagens Google, e porque não têm data?
    a: O revisor marcou este exemplo sobre imagens de satélite Google, na nossa ferramenta de revisão, por isso as vistas de perto mostram as marcações sobre essas imagens, com o crédito «Imagens © Google». A Google não indica quando essas imagens foram captadas, por isso mostram onde fica cada marcação, não quando uma construção apareceu. Os levantamentos que precisam de data usam imagens que se podem datar e entregar, como um levantamento por drone, uma cena de satélite adquirida ou imagens georreferenciadas do cliente. A vista da classificação por troços, mais abaixo, usa uma cena Copernicus Sentinel-2 datada.
  - q: Porque é que alguns círculos ficam ao lado de uma cobertura e não em cima dela?
    a: Cada círculo está centrado no ponto que o revisor colocou, e a revisão foi feita com uma ampliação menor do que a destas vistas de perto, por isso um círculo pode ficar ao lado da cobertura a que se refere. As distâncias do registo são medidas a partir desses pontos, exactamente como foram registados.
  - q: Porque é que o registo aparece numa faixa recta?
    a: A vista linear endireita o traçado. Cada construção fica na sua distância ao longo do traçado e na sua distância à linha, com o lado norte da linha em cima. As distâncias na transversal estão desenhadas com o dobro da escala das distâncias ao longo do traçado, para que as faixas de 50 m e 100 m se leiam bem. É um gráfico do registo, não um mapa.
  - q: Pode faltar alguma construção no exemplo?
    a: Sim, e as vistas de perto mostram algumas. A revisão registou as construções que o revisor marcou na altura, dentro da área de pesquisa à volta do traçado, e as coberturas sem círculo dentro das faixas não foram marcadas nela. Publicamos o exemplo tal como foi registado, sem acrescentos. Construções sob árvores, coberturas que se confundem com o solo e o que fica fora da área de pesquisa também escapam a partir do ar. Um levantamento entregue é revisto sobre as imagens que indica, e o que não se resolve a partir do ar fica listado para verificação no terreno.
  - q: As instalações e os poços junto do traçado contam como ocupação?
    a: Não. As instalações perto das duas pontas deste traçado pertencem ao operador do gasoduto. As instalações do próprio operador fazem parte do activo, não são ocupação, e não entram no registo.
related: [oil-gas, right-of-way-monitoring, change-detection]
---

::::section{id="resumo" eyebrow="O exemplo em resumo" title="Um gasoduto de alta pressão em Moçambique"}
:::facts{cols="4"}
- Extensão do traçado: 10,78 km
- Faixas: 50 m e 100 m
- Método: marcação do revisor (manual)
- Construções marcadas: 59
- A menos de 50 m do eixo: 11
- A menos de 100 m do eixo: 36
- Troços de 500 m: 22
- Com densidade alta: 4
:::

As contagens são cumulativas: «a menos de 100 m» inclui as 11 construções a menos de 50 m. As distâncias são medidas até ao traçado tal como foi fornecido, na respectiva zona UTM (36S). As outras 23 marcações ficam entre os 100 m e o limite da área de pesquisa.
::::

::::section{id="imagens" tone="alt" eyebrow="Nas imagens de satélite" title="As marcações do revisor sobre imagens de satélite Google" lead="Todo o traçado e, depois, vistas de perto de cerca de 600 por 400 m dos troços onde o revisor colocou marcações, cada uma com o traçado, as faixas de 50 m e 100 m e um círculo em cada marcação do revisor. As letras na vista geral mostram onde fica cada vista de perto."}
:::sample-gallery{data="sample-pipeline-google" priority="true"}
:::
::::

::::section{id="mapa" eyebrow="Vista do registo" title="O troço mais denso, construção a construção" lead="Do km 5,0 ao km 6,5, onde a linha acompanha uma picada existente entre machambas e habitações. Cada marcação fica na sua distância ao longo do traçado e na sua distância à linha, e leva a referência usada na tabela abaixo."}
:::figure{src="samples/sample-pipeline-register-km5-6-pt" alt="Vista linear do traçado do gasoduto entre o km 5,0 e o km 6,5: o traçado como uma linha recta cor de laranja, com as faixas de 50 m a vermelho e de 100 m a âmbar dos dois lados, e vinte marcações do revisor, R20 a R39, colocadas pela distância ao longo do traçado e pela distância à linha; em baixo, os três troços de 500 m classificados com densidade média (4), alta (16) e alta (7)" caption="O gasoduto do exemplo, km 5,0 a 6,5: marcações do revisor pela distância ao longo do traçado e pela distância à linha, com a classificação de cada troço de 500 m" badge="Revisto · marcação manual" size="wide" credit="Desenhado pela AfriScan a partir do registo do exemplo. Sem imagens; as distâncias na transversal estão desenhadas com o dobro da escala longitudinal."}
<span class="band band--a">Até 50 m</span> <span class="band band--b">50–100 m</span> <span class="band band--c">Além de 100 m</span> O lado norte da linha fica em cima.
:::

### Excerto do registo deste troço

A distância ao longo do traçado conta-se a partir do início da linha. Este exemplo público omite as coordenadas. O registo entregue a um cliente dá cada construção em WGS84 e em UTM, e segue apenas para os contactos que o cliente indicar.

:::register{data="sample-pipeline"}
:::
::::

::::section{id="trocos" tone="alt" eyebrow="Densidade de ocupação" title="Cada troço de 500 m, classificado" lead="A classificação é uma regra de contagem que indica onde enviar primeiro as equipas. Não é uma avaliação de segurança nem de integridade do gasoduto."}
:::figure{src="samples/sample-pipeline-route-ratings" alt="O traçado do gasoduto sobre uma cena de satélite Sentinel-2, com cerca de 10 km desde uma instalação de gás a oeste, passando por uma povoação, até uma zona húmida a leste; o traçado está colorido pela classificação: vermelho nos troços de densidade alta entre o km 4 e o km 6,5, âmbar nos de densidade média e cinzento nos de densidade baixa" caption="Todo o traçado, com cada troço de 500 m colorido pela classificação: vermelho alta, âmbar média, cinzento baixa" size="wide" credit="Traçado sobre uma cena Copernicus Sentinel-2 de 2 de Agosto de 2026 (contém dados Copernicus Sentinel modificados, 2026), mostrada para localização. Com píxeis de 10 m, a cena não mostra construções individuais; as classificações vêm do registo revisto, não desta cena."}
:::

:::segments{data="sample-pipeline"}
Cada troço é classificado a partir das construções dentro da faixa mais larga, neste traçado a de 100 m, medida a partir de qualquer ponto do troço. Uma construção perto do limite entre dois troços conta nos dois, por isso a soma dos troços é maior do que o total do traçado. Nas três vistas abaixo, a área contornada é a que a classificação conta; as marcações fora dela aparecem esbatidas.
:::

:::cards{cols="3"}
:::figure{src="samples/sample-pipeline-register-high-pt" alt="Vista linear do km 5,5 ao km 6,0, com densidade alta: dezasseis marcações do revisor dentro da área a menos de 100 m do troço, duas delas a menos de 50 m da linha; duas marcações fora dessa área aparecem esbatidas" caption="Alta · km 5,5 a 6,0" size="third" credit="Vista do registo; mesma escala nas três"}
16 construções a menos de 100 m do troço, onde o traçado passa entre habitações dos dois lados. Três delas ficam logo depois das pontas do troço e contam também para o troço vizinho.
:::
:::figure{src="samples/sample-pipeline-register-medium-pt" alt="Vista linear do km 9,0 ao km 9,5, com densidade média: duas marcações do revisor dentro da faixa de 50 m, do lado norte da linha, logo depois do km 9,0" caption="Média · km 9,0 a 9,5" size="third" credit="Vista do registo; mesma escala nas três"}
2 construções, ambas dentro da faixa de 50 m: poucas, mas perto da linha. Ficam logo depois do km 9,0, por isso o troço do km 8,5 ao 9,0 também tem densidade média.
:::
:::figure{src="samples/sample-pipeline-register-low-pt" alt="Vista linear do km 7,5 ao km 8,0, com densidade baixa: as faixas de 50 m e 100 m sem marcações do revisor" caption="Baixa · km 7,5 a 8,0" size="third" credit="Vista do registo; mesma escala nas três"}
Nenhuma construção a menos de 100 m: a linha atravessa mato e capim queimado.
:::
:::
::::

::::section{id="entrega" eyebrow="O que uma entrega contém" title="Os mesmos resultados, para o seu traçado"}
:::cards{cols="3"}
:::card{title="Relatório PDF" icon="file-text"}
Uma capa com os números principais; os dados do levantamento (método, fonte e data das imagens, zona UTM); um mapa geral com as faixas, a quilometragem e a escala; a tabela de troços; uma fotografia de cada construção pela ordem do traçado; e o registo completo de coordenadas. Em português ou em inglês.
:::
:::card{title="Camadas SIG" icon="layers"}
GeoPackage, GeoJSON, KMZ e Shapefile. Cada construção leva a distância à linha e a faixa; o traçado vem dividido em troços classificados. Abrem no QGIS, no ArcGIS e no Google Earth.
:::
:::card{title="Mapa interactivo em ficheiro" icon="map"}
Um mapa autónomo dos resultados que abre num navegador. As camadas de resultados funcionam sem ligação; as imagens de fundo precisam de ligação, salvo se as imagens do levantamento forem incluídas.
:::
:::

:::cta{title="Quer abrir os ficheiros antes de uma proposta?" text="Peça o exemplo anonimizado: o relatório e as camadas SIG no formato da entrega, com o traçado generalizado, sem coordenadas e sem imagens de mapas de base, para que as equipas de SIG e de terras vejam como os ficheiros estão organizados." button="Pedir o exemplo anonimizado" intent="sample-report"}
:::
::::

::::section{id="sobre" tone="alt" eyebrow="Sobre este exemplo" title="O que este exemplo mostra, e o que não mostra"}
:::::columns{split="1-1"}
::::col
:::checklist
- Onde ficam as construções em relação à linha, por faixa e por troço do traçado
- Como a classificação de densidade transforma um registo numa lista de troços a visitar primeiro
- O que a equipa de SIG recebe e como os ficheiros estão organizados
:::
::::
::::col
:::checklist{tone="no"}
- Quem vive numa construção, a quem pertence ou para que serve
- Se uma construção está autorizada: isso cabe ao proprietário do traçado e às autoridades
- Qualquer coisa sobre o estado ou a integridade do gasoduto
:::
::::
:::::

:::callout{tone="scope" title="Mostrado com autorização"}
Nunca publicamos o traçado, as imagens ou os resultados de um cliente sem autorização escrita. O traçado é mostrado aqui com a autorização do respectivo proprietário. As instalações perto das duas pontas do traçado são do próprio operador: fazem parte do activo, não são ocupação, e não entram no registo.
:::
::::

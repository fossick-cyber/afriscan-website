---
key: results
icon: map
title: Exemplo de Levantamento de Ocupações (T-9) | AfriScan
description: "O que um levantamento AfriScan entrega: construções ao longo do traçado, distância ao eixo, registo das faixas de 50 e 100 m, densidade por troço, PDF e SIG."
h1: Exemplo de um levantamento de ocupações
crumb: Exemplo de resultados
section: resources
nav_group: resources
nav_order: 10
nav_label: Exemplo de resultados
nav_blurb: Uma revisão do traçado do gasoduto T-9, em Inhambane
eyebrow: Recursos
lead: Resultados reais de uma revisão feita no traçado do gasoduto de substituição T-9, na província de Inhambane, mostrados com a autorização do proprietário do traçado. As marcações aparecem exactamente como o revisor as registou.
buttons:
  - {label: Pedir um relatório de exemplo, intent: sample-report}
  - {label: Como trabalhamos, key: how-we-work}
og:
  headline: Exemplo de um levantamento de ocupações
  subline: Gasoduto de substituição T-9, Inhambane · faixas de 50 e 100 m · marcações do revisor
cta:
  title: Quer ver o seu próprio corredor?
  text: Envie o ficheiro do traçado (KML, GeoJSON ou Shapefile) e as distâncias que contam para a sua empresa, e preparamos o âmbito de um levantamento da sua linha ou do seu terreno.
  button: Enviar o traçado
  intent: proposal
faq:
  - q: Este exemplo é detecção automática?
    a: Não. Cada marcação deste exemplo foi colocada por um revisor sobre imagens de satélite; não se mostra nenhum resultado de detecção automática. Num levantamento para um cliente, as bases de dados abertas de edifícios e os modelos de segmentação propõem primeiro as construções, e um revisor confirma, corrige e acrescenta.
  - q: Porque é que as imagens não têm data?
    a: Esta revisão foi feita sobre um mapa de base de satélite da Google, que não indica quando as imagens foram captadas. Serve para um exemplo e para um rastreio interno, mas não para um registo que tenha de reflectir uma data. Os levantamentos que precisam de data usam imagens datadas, como um levantamento por drone, uma cena de satélite adquirida ou imagens georreferenciadas do cliente.
  - q: Há construções na imagem sem marcação. Porquê?
    a: A revisão registou as construções que o revisor confirmou na altura, dentro da área de pesquisa à volta do traçado. Construções sob árvores, coberturas que se confundem com o solo e o que fica fora da área de pesquisa não estão marcados. Um levantamento entregue é revisto sobre as imagens que indica, e o que não se resolve a partir do ar fica listado para verificação no terreno.
  - q: Porque é que alguns quadrados ficam ligeiramente fora dos telhados?
    a: As marcações do revisor são pontos. Os quadrados são desenhados à volta de cada ponto para se verem a esta escala, e um ponto colocado na beira de um telhado deixa o quadrado em parte sobre o terreno ao lado. As distâncias são medidas a partir do ponto.
  - q: As instalações e os poços junto do traçado contam como ocupação?
    a: Não. As instalações perto das duas pontas deste traçado pertencem ao operador do gasoduto. As instalações do próprio operador fazem parte do activo, não são ocupação, e não entram no registo.
---

::::section{id="resumo" eyebrow="O exemplo em resumo" title="Gasoduto de substituição T-9, Inhambane"}
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

::::section{id="mapa" tone="alt" eyebrow="Vista do mapa" title="Como o registo se vê no terreno" lead="O troço mais denso do traçado, onde a linha acompanha uma picada existente entre machambas e habitações. Cada marcação leva a referência usada na tabela abaixo."}
:::figure{src="samples/t9-km5-6-pt" alt="Vista de satélite do traçado T-9 entre o km 5,0 e o km 6,3, com o traçado a laranja, a faixa de 50 m a vermelho, a faixa de 100 m a âmbar e as marcações do revisor R21 a R40 em habitações dos dois lados da linha" caption="Gasoduto de substituição T-9, km 5,0 a 6,3: traçado, faixas de 50 m e 100 m e marcações do revisor por faixa" badge="Revisto · marcação manual" size="wide" priority="true" credit="Imagens © Google. O fundo é um mapa de base de satélite da Google, mostrado para ilustração: não tem data de captação e não é imagem entregue num levantamento. Traçado, faixas, marcações e referências desenhados pela AfriScan a partir dos dados da revisão."}
As marcações vermelhas estão a menos de 50 m da linha, as âmbar entre 50 e 100 m e as verde-azuladas além de 100 m.
:::

### Excerto do registo deste troço

A distância ao longo do traçado conta-se a partir do início da linha. O exemplo público omite as coordenadas; o registo entregue dá cada construção em WGS84 e em UTM.

:::register{data="t9"}
:::
::::

::::section{id="trocos" eyebrow="Densidade de ocupação" title="Cada troço de 500 m, classificado" lead="A classificação é uma regra de contagem que indica onde enviar primeiro as equipas. Não é uma avaliação de segurança nem de integridade do gasoduto."}
:::segments{data="t9"}
Cada troço é classificado a partir das construções dentro da faixa mais larga, neste traçado a de 100 m. Uma construção perto do limite entre dois troços conta nos dois, por isso a soma dos troços é maior do que o total do traçado.
:::

:::cards{cols="3"}
:::figure{src="samples/t9-rating-high-pt" alt="Vista aproximada do km 5,5 ao km 6,0: habitações dos dois lados do traçado, dentro das faixas de 50 m e 100 m" caption="Alta · km 5,5 a 6,0" size="third" credit="Imagens © Google"}
16 construções a menos de 100 m, onde o traçado passa entre habitações dos dois lados.
:::
:::figure{src="samples/t9-rating-medium-pt" alt="Vista aproximada à volta do km 9,0: duas marcações do revisor em pequenas parcelas dentro da faixa de 50 m, logo a norte do traçado" caption="Média · à volta do km 9,0" size="third" credit="Imagens © Google"}
2 construções, ambas dentro da faixa de 50 m: poucas, mas perto da linha. Ficam no limite entre dois troços, por isso os troços do km 8,5 ao 9,0 e do km 9,0 ao 9,5 têm ambos densidade média.
:::
:::figure{src="samples/t9-rating-low-pt" alt="Vista aproximada à volta do km 8,0: o traçado atravessa mato e capim queimado, sem construções dentro de qualquer das faixas" caption="Baixa · km 7,5 a 8,5" size="third" credit="Imagens © Google"}
Nenhuma construção a menos de 100 m em qualquer dos dois troços: uma faixa desimpedida entre mato e capim queimado.
:::
:::
::::

::::section{id="entrega" tone="alt" eyebrow="O que uma entrega contém" title="Os mesmos resultados, para o seu traçado"}
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

:::cta{title="Quer o relatório de exemplo completo?" text="Podemos enviar o PDF e os ficheiros SIG deste exemplo para que as equipas de SIG e de terras os abram." button="Pedir um relatório de exemplo" intent="sample-report"}
:::
::::

::::section{id="sobre" eyebrow="Sobre este exemplo" title="O que este exemplo mostra, e o que não mostra"}
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
Nunca publicamos o traçado, as imagens ou os resultados de um cliente sem autorização escrita. O traçado T-9 é mostrado aqui com a autorização do proprietário do traçado. As instalações perto das duas pontas do traçado são do próprio operador: fazem parte do activo, não são ocupação, e não entram no registo.
:::
::::

---
key: route-site-selection
template: solution
title: "Route & Site Selection: Structures Affected | AfriScan"
description: Compare structures along alternative alignments or candidate sites, with slope, drainage and flood screening, before your route or site is fixed.
h1: Compare routes and sites by the structures they affect
eyebrow: Baselines and records
lead: The cheapest structure to deal with is the one your route never crosses. Send us the alignments or candidate sites you are weighing, and we compare how many structures each puts inside the widths that matter, where they cluster, and which stretches are steep, erosion-prone or flood-exposed. Your engineers and land team then choose with the land impact on the table.
used_in: [power-utilities, oil-gas, project-finance-esia, rail-roads, renewables, telecom-fibre]
buttons:
  - {label: Request a proposal, intent: proposal}
  - {label: How we measure structures, key: methodology, anchor: "#distances"}
related: [resettlement-cut-off-baselines, terrain-flood-post-event, right-of-way-monitoring]
service:
  name: Route and site option comparison
  type: Route alternatives and site selection screening
  description: Side-by-side structure counts along alternative pipeline, power-line, road or rail alignments or across candidate sites, with slope, drainage-crossing and flood-exposure screening, estimated households and a land-use baseline, as an input to route and site selection.
og:
  headline: Compare routes and sites by the structures they affect
  subline: Structure counts, terrain and flood screening for each option, before the route is fixed
cta:
  title: Send us your route options.
  text: Two alignments or twenty, as KML, KMZ, GeoJSON, Shapefile, GPX or GeoPackage. Tell us the widths and the decision date, and we reply with a written proposal.
  button: Request a proposal
faq:
  - q: Is this a full alternatives assessment?
    a: No. It is one input to route and site selection, the land-impact part, measured the same way for each option. Your alternatives assessment also weighs engineering, cost, environmental and social factors that stay with your engineers and ESIA team.
  - q: Will choosing the option with the fewest structures avoid resettlement?
    a: Not necessarily. It shows which options affect fewer structures on the imagery date, which is often a large part of the land question. Whether resettlement or compensation is needed is settled by your land and resettlement process, on the ground.
  - q: How many options can you compare?
    a: As many as you send. Each option is measured with the same widths, the same imagery plan and the same review, so the counts are comparable. Short deviations around a single village can be added as their own options.
  - q: Is the terrain screening a geotechnical study?
    a: No. Slope, erosion and drainage screening from elevation data highlights stretches for your engineers to look at. It is not a geotechnical, hydrological or hydraulic study, and not engineering design. Where a stretch needs detail, a drone elevation model can be added, subject to the permits the flight requires.
  - q: What happens once the route is fixed?
    a: The same register becomes the baseline for the chosen route. We can then add a dated cut-off-date record for resettlement, the reports your land team needs for landowner engagement and servitude acquisition, and re-surveys during construction.
---

::::section{id="what" eyebrow="What it is" title="Land impact, weighed before the line is drawn"}
:::::columns{split="2-1"}
::::col
Every new pipeline, transmission line, road or railway starts as a set of options on a map. Engineering constraints narrow them, and one of the biggest remaining questions is the land: how many homesteads, outbuildings and other structures each option would put inside its servitude or protection zone, and how many more sit just beyond it.

Answered late, that question becomes resettlement cost, programme delay and difficult community meetings on a route that can no longer move. Answered early, it is one more column in the options table.

We measure each alignment or candidate site the same way: the structures within the distances you set, rated every 500 m for encroachment density, with a land-use baseline and estimated households alongside. Terrain, drainage and flood screening highlight the stretches your engineers will want to check. The result is a like-for-like comparison, not a recommendation.
::::
::::col
:::callout{tone="note" title="South Africa's transmission build-out"}
Under the [power-line Standard (GN 2313 of 2022)](https://www.dffe.gov.za/sites/default/files/legislations/nema_powerlinessubstationsdevelopmet_g47095gon2313.pdf), compliant lines and substations inside the strategic transmission corridors are excluded from needing environmental authorisation. That brings land baselines forward, to the point where routes are still being chosen.
:::
::::
:::::
::::

::::section{id="how" tone="alt" eyebrow="How it works" title="From options on a map to a comparison table"}
:::steps
:::step{title="Options and widths" icon="route"}
You send each alignment or candidate site, or we draw them with you from your design drawings. We agree the servitude or protection-zone width, a wider watch band, and any sites to avoid.
:::
:::step{title="Structures per option" icon="houses"}
We map the structures along every option from dated imagery and open building datasets. A reviewer checks each option to the same standard, so a difference in the counts is a difference on the ground, not in the method.
:::
:::step{title="Screening" icon="layers"}
Slope, erosion-prone ground and drainage crossings from elevation data; historically flooded, low-lying and river-crossing stretches; a land-use baseline of built-up land, cropland, bare ground, water and tree cover; and imagery history where a stretch has changed.
:::
:::step{title="Comparison and handover" icon="file-text"}
A comparison table and maps for your options workshop. Once a route is chosen, the register for that route becomes its baseline.
:::
:::
::::

::::section{id="deliverables" eyebrow="What you receive" title="A comparison your options workshop can use"}
The comparison table puts every option in the same columns. We fill it from the survey; the headings below are what each column means.

| Column | What it shows |
|---|---|
| Length or area | The option's length, or the site's area, as supplied |
| Structures within each width | Cumulative counts: within 50 m includes those within 30 m, and so on, for the widths you set |
| High-density stretches | How many 500 m segments carry more than 5 structures inside the widest band |
| Estimated households | An estimate from the structure count and stated persons-per-household assumptions |
| Land use | Built-up land, cropland, tree cover and water crossed, from the land-use baseline |
| Terrain and water | Steep, erosion-prone and drainage-crossing stretches, and historically flooded or low-lying ground |
| Change | Where imagery history shows recent building or clearing along the option |

:::checklist
- **The comparison table** in the PDF report, with a short note on each option
- **Maps of each option** with its buffers, rated segments and structures
- **GIS layers** for every option in GeoPackage, GeoJSON, KMZ and Shapefile
- **For the chosen route:** a dated register for landowner engagement, servitude acquisition and, where people are affected, a resettlement cut-off-date record
:::

### The services behind it

Every proposal names the services it includes. These are the ones this solution draws on.

:::catalogue{services="S04,S27,S28,S33,S35,S36,S01,S31,S37"}
:::
::::

::::section{id="limits" tone="alt" eyebrow="Limits" title="What route comparison is not"}
:::::columns{split="1-1"}
::::col
### What it gives you
:::checklist
- A like-for-like count of the structures each option affects on the imagery date
- The stretches of each option to look at first, for land, terrain and water
- A baseline that carries forward once the route is fixed
:::
::::
::::col
### What it is not
:::checklist{tone="no"}
- A full alternatives assessment, or a recommendation of a route
- A geotechnical, hydraulic or flood study, or engineering design
- A census: household figures are estimates, with the assumptions stated
- A cadastral or statutory survey of the servitude
:::
::::
:::::
::::

::::section{id="who" eyebrow="Who uses it" title="Teams choosing where a new line or site will go"}
:::cards{cols="3"}
:::card{title="Power & utilities" icon="power" key="power-utilities#new-lines" cta="New lines and servitudes"}
Transmission and distribution planners, servitude-acquisition teams and EPC contractors comparing line routes and substation sites.
:::
:::card{title="Oil & gas" icon="pipeline" key="oil-gas#new-pipelines" cta="New pipelines and loops"}
Route engineers and land teams on new pipelines, loops and gathering lines, before the route and its protection zone are fixed.
:::
:::card{title="Project finance & ESIA" icon="clipboard" key="project-finance-esia#esia-rap-baselines" cta="ESIA and RAP baselines"}
ESIA consultancies that need structure counts for the alternatives chapter, delivered as GIS layers for their own maps.
:::
:::card{title="Rail & roads" icon="rail" key="rail-roads#new-alignments" cta="New alignments"}
Highway and railway project teams comparing alignments and bypasses through settled land.
:::
:::card{title="Renewables" icon="sun" key="renewables#site-selection" cta="Site selection"}
Solar and wind developers comparing candidate sites and the routes of their connection lines.
:::
:::

Also used for [cross-border interconnectors](key:power-utilities#interconnectors) and [cross-border pipelines](key:oil-gas#cross-border-lines), by telecom operators planning [new fibre routes](key:telecom-fibre#new-routes), and by mining companies weighing haul-road and conveyor routes, and dump or tailings sites.
::::

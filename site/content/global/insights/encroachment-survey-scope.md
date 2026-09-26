---
key: insight-survey-scope
template: article
published: 2026-09-26
parent: resources
nav_group: resources
nav_langs: [en]
nav_order: 45
nav_blurb: An RFQ checklist and questions for bidders
title: How to Scope a Right-of-Way Encroachment Survey | AfriScan
description: A checklist for RFQs and scopes of work for pipeline, power-line and concession encroachment surveys, and the questions that separate strong bids.
h1: How to scope a right-of-way encroachment survey
crumb: Scoping an encroachment survey
summary: A checklist for writing an RFQ or scope of work for a right-of-way, servitude or concession encroachment survey, and the questions to ask every bidder.
icon: clipboard
eyebrow: Guide · Procurement
lead: "An encroachment survey is only as useful as its scope. Leave the reference line, the distances or the imagery date vague, and you get a register nobody can rely on or compare with the next one. This guide is for integrity, land and procurement teams writing an RFQ or scope of work for a pipeline, power-line, road, rail or concession survey, whoever they buy it from."
buttons:
  - {label: Send tender documents, intent: tender}
  - {label: How we work, key: how-we-work}
related: [right-of-way-monitoring, encroachment-surveys, change-detection]
about: [home, right-of-way-monitoring, encroachment-surveys, change-detection, oil-gas, power-utilities, rail-roads, za-water-utilities, za-land-invasion, mz-procurement, za-procurement, ng-procurement]
og:
  headline: How to scope a right-of-way encroachment survey
  subline: A checklist for RFQs and scopes of work, and the questions to ask every bidder
faq:
  - q: Should we ask bidders for an accuracy figure?
    a: Ask how quality is controlled and how you can check it, rather than for a single percentage. Detection quality depends on the imagery, the season, tree cover and building styles on your route, so a figure measured somewhere else says little about yours. A bid that names its imagery, has every result reviewed by a person, lists uncertain structures for a ground check and agrees a sample check against your field data gives you something you can actually test.
  - q: Should the first survey cover the whole route?
    a: Not necessarily. Starting with one stretch, one site or one concession boundary lets your team check the register against what it knows before the rest is surveyed, and settles the reference line, the bands and the report format early. The scope can then extend to the whole route on the same terms.
  - q: Who should own the imagery?
    a: Say in the scope what you need to do with it. For purchased satellite scenes, the licence decides who may use and share the imagery; for drone surveys, national law may also control who may receive and publish it, as Mozambique's Lei n.º 6/2024 does. Ask each bidder to state the licence terms and any authorisations for the imagery they will deliver.
  - q: How often should a right of way be re-surveyed?
    a: It depends on how fast the land around it changes and on the imagery supply, not on a fixed rule. Fast-growing peri-urban stretches may justify frequent re-surveys; remote stretches much less often. Ask for a cadence that matches the imagery that can realistically be obtained, and for each re-survey to use the same method and bands as the baseline.
cta:
  title: Writing a scope or a tender?
  text: Send your draft scope or the tender documents. We reply with a written proposal that answers each item, or with the questions the scope still needs to settle.
  button: Request a proposal
  intent: tender
---

## Start with what the register is for {#purpose}

Every other line of the scope follows from one question: what decision or record is the survey for? Screening a long network to plan patrols, preparing a notice programme, sizing a census, comparing route options and building a record that may be relied on in a dispute all need different imagery, review depth and deliverables. Write the purpose into the scope in one sentence, and ask each bidder to say how their method serves it.

## Define the asset and the reference line {#reference-line}

Most disagreements about encroachment registers are really disagreements about which line distances were measured from.

- **Supply the line.** A route file in a common GIS format (KML, KMZ, GeoJSON, Shapefile, GPX or GeoPackage), with its coordinate reference system stated.
- **Say what the line represents.** The pipe centreline, the power line's axis, the edge of a registered servitude, a road reserve boundary or a concession limit. Statutory strips are defined differently: Mozambique's Land Law describes a strip of 50 m on each side of a conduit, while its Electricity Law counts a servitude from the line's axis. See [the 50 m protection zone](/mz/50m-protection-zone).
- **Include the polygons you have.** Servitude or wayleave polygons, road and rail reserves, concession boundaries and facility perimeters let the register report inside and outside them directly.
- **Name the facilities that are yours.** Plants, well pads, camps and stations belonging to the operator should be listed so they are recorded as what they are and never counted as encroachment.

## Set the distances and bands {#bands}

Specify the distances the register must report, in metres, and whether bands are cumulative (within 100 m including within 50 m) or exclusive (50 to 100 m). Tie them to something your organisation already uses: the statutory strip, the servitude width, a company standard or the widths in a previous survey you want to compare with.

| Country | Widths commonly written into scopes | Source |
|---|---|---|
| Mozambique | 50 m on each side of oil, gas, water, electricity and telecommunications lines; 50 m each side of a railway's axis; 30 m and 15 m for roads | [Lei n.º 19/97, art. 8](https://www.pdul.gov.mz/content/download/486/2635/file/Lei%20de%20Terras.pdf) |
| South Africa | The registered servitude width for each line or pipeline; 60 m from a national road boundary outside urban areas | Servitude diagrams; [SANRAL Act, s48](https://www.sagc.org.za/pdf/legislation/S%20A%20National%20Roads%20Agency%20Act%207%20of%201998.pdf) |
| Nigeria | 50 m, 30 m and 11 m rights of way for 330 kV, 132 kV and 33 kV lines; restrictions up to 100 feet from the land in a pipeline licence, by order | [NESIS Regulations 2015](https://nemsa.gov.ng/wp-content/uploads/2019/11/NESIS-Regulations-2015-1-1.pdf); [Oil Pipelines Act, s.12](https://lawsofnigeria.placng.org/view2.php?sn=425) |

Ask for a search margin beyond the widest band, so the register also shows what sits just outside your strip, and for the distance of each structure, not only its band, so structures near a band edge can be checked first.

## Specify the imagery, not just "satellite" {#imagery}

"Satellite imagery" in a scope can mean anything from a dated half-metre scene to a web-map layer with no date at all. State what you need:

:::checklist
- **Named imagery with its capture date** in the report, for every stretch
- **A maximum age** for the imagery, or a date window, tied to the purpose
- **A minimum resolution** fine enough to see individual buildings on your route
- **Cloud and season**: how cloudy scenes and rainy-season gaps will be handled
- **No web-map basemaps as delivered imagery**, and no dated record built on them
- **Licence terms** for any imagery delivered to you
- **For drone work**, the permits and authorisations the flight needs, who holds them, and any rules on receiving and publishing the imagery
:::

Where one date matters, such as a cut-off date or a handover, say so, and ask how close to it each bidder's imagery plan can realistically get. The trade-offs are set out in [satellite or drone?](/insights/satellite-or-drone-corridor-surveys).

## Ask for review, not just detection {#review}

Automatic detection is fast, but generic detectors miss thatch, mud-brick and zinc roofs and homesteads under trees, and propose bushes, rocks and shadows as buildings. Different automatic sources can disagree widely on the same stretch. Write into the scope that:

- a person reviews every automatic result before delivery, and the register shows which structures were confirmed, which were added by hand and which were removed;
- anything the imagery cannot settle goes on a ground-check list with its location and a photo crop, rather than being guessed;
- change between surveys is confirmed by a reviewer on both images before it is reported.

## Say what the register must contain {#register}

| Field | Why it matters |
|---|---|
| Structure ID, stable across re-surveys | Lets field teams, the census, the land team and later surveys refer to the same structure |
| Coordinates in WGS84 and the local UTM zone | Opens in any GIS and on a field phone |
| Distance to the reference line and band | The core of the register |
| Chainage along the route | How integrity and line teams find things |
| Photo crop and imagery date | Lets anyone check the entry without opening the imagery |
| Review status | Confirmed, added by hand, or to verify on the ground |
| Category, where needed | Main building, outbuilding, enclosure, as far as the imagery shows |

If you want a density measure to prioritise work, specify it: for example, structures inside the widest band per 500 m of route. Ask for it to be called what it is, a density or pressure measure, not a safety or integrity rating.

## Deliverables and formats {#deliverables}

Name the formats your teams will open: a PDF report with maps, the segment table, the register and photo crops; GIS layers in GeoPackage, GeoJSON, KMZ or Shapefile with distance and band on each feature; and, if field teams need it, a map file that opens in a web browser. State the language (English or Portuguese), the coordinate reference system, and that the deliverables are files your organisation keeps.

## Repeat surveys and change {#repeat}

If the survey is the first of a series, say so now. Ask for the same method, bands and imagery type at each re-survey, new and removed structures between dated surveys confirmed by a reviewer, and a short notice after each survey saying what changed and where. Web-map layers and footprint datasets cannot support change detection, because they have no reliable date.

## Data, confidentiality and permissions {#data}

- An NDA before route files are shared, if your organisation needs one.
- Who may publish or reuse the route, imagery or results: normally nobody, without your written permission.
- Where client data is stored and processed, and for how long, stated in the proposal.
- Data-protection duties in the country: POPIA in South Africa, the Nigeria Data Protection Act, and the data and publication rules for aerial surveys in Mozambique's Lei n.º 6/2024.

## Questions to ask every bidder {#bidders}

:::steps{style="list"}
:::step{title="Which imagery, from which dates, for each stretch?"}
A strong bid names the imagery type, the realistic dates and what happens where the archive is too old or too cloudy.
:::
:::step{title="Who reviews the automatic results, and how can we see it?"}
Look for a review of every result, a ground-check list and a register that shows review status.
:::
:::step{title="What exactly is measured, from which line?"}
The answer should restate your reference line and bands, in metres on the ground.
:::
:::step{title="What can the register not tell us?"}
A good bidder lists limits without being asked: occupancy, ownership and eligibility; anything under canopy or cloud; positional error near band edges.
:::
:::step{title="For drone work, which permits does this job need, and who holds them?"}
The answer should come from the country's rules, not from a general statement. See [drone rules by country](/drone-regulations).
:::
:::step{title="Can we start with one stretch and check it?"}
A pilot stretch with a sample check against your own field data is the cheapest way to see whether a register is fit for your purpose.
:::
:::

:::callout{tone="warn" title="Claims to test before you accept them"}
Be careful with bids that promise more than imagery allows. No optical satellite or drone sees through cloud or at night. No satellite can be booked for a guaranteed date. A single accuracy figure with no method behind it tells you nothing about your route. A land survey should not identify, count or follow people, and a structure register should not call anyone an intruder: whether a structure is authorised is for you and the authorities to decide.
:::

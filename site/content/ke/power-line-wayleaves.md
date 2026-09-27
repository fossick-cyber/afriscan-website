---
key: power-utilities
template: industry
title: Power-Line Wayleaves & Interconnectors in Kenya | AfriScan
description: Structures and vegetation in Kenya's transmission and distribution wayleaves, cut-off-date checks for resettlement plans, and route comparison for new lines.
h1: Structures and vegetation in Kenya's power-line wayleaves
crumb: Power lines
eyebrow: Power & utilities · Kenya
lead: Kenya's transmission grid is growing, with new 400 kV and 132 kV lines, interconnectors to Ethiopia, Tanzania and Uganda, and lines planned under the Transmission Master Plan 2025–2044. Wayleave compensation is paid for land use, crops, trees and structures, so knowing what stood in the corridor on a given date matters to the budget as much as to safety. We map it from dated imagery, and a person checks each result.
buttons:
  - {label: Request a proposal, intent: proposal}
  - {label: Cut-off-date checks, href: "#cut-off"}
service:
  name: Power-line wayleave register and cut-off-date checks, Kenya
  type: Wayleave structure register, resettlement cut-off-date check, vegetation change and route comparison
  description: Structures, vegetation change, new tracks, cleared ground and excavations inside transmission and distribution wayleaves in Kenya, measured to the centreline, checked against the resettlement cut-off date that applies, compared between dated scans, checked by a person and delivered as a PDF report with GeoPackage, GeoJSON, KMZ and Shapefile layers.
og:
  headline: Structures and vegetation in Kenya's power-line wayleaves
  subline: Cut-off-date checks, wayleave registers and route comparison, checked by a person
related: [resettlement-cut-off-baselines, route-site-selection, vegetation-land-cover-fire]
faq:
  - q: Which wayleave widths do you use?
    a: The width in your resettlement action plan, easement or gazette order. KETRACO's published plans use 40 m (20 m each side of the centreline) for a 220 kV line and 30 m (15 m each side) for a 132 kV line. We found no published figure for 400 kV lines or distribution lines in the sources we checked, so send the width that applies, and each structure is reported with its distance to the centreline and its band.
  - q: Which cut-off date do you check against?
    a: The one you tell us applies. Each published plan defines its cut-off date as the completion of the census and asset inventory, and provides for re-ratifying it by Gazette notice where the project is delayed by two years or more. Ask your RAP consultant whether the original date or a re-ratified one governs, and we compare the imagery closest to that date with today's.
  - q: Can you tell us who is eligible for compensation?
    a: No. We show which structures appear on imagery from the cut-off date and which appeared later. The census, the asset inventory and every eligibility decision stay with you, your RAP consultant and, where land is acquired, the National Land Commission.
  - q: Can the imagery show vandalism or theft near towers?
    a: No. Imagery cannot see vandalism or theft. It shows land change, such as new tracks, cleared ground, excavations and structures near towers and substations between dated images, which can help your teams decide where to patrol.
  - q: Can you fly drones along the line?
    a: Only with KCAA's specific permission. Operating in or around high-tension cables without it is treated as negligent or reckless operation under the Civil Aviation (Unmanned Aircraft Systems) Regulations 2025 (reg 45(2)(c)), so any drone flight near a line needs that permission and is flown by a Kenyan company with a KCAA RPAS Operator Certificate, subject to the approvals and security clearances the job requires.
  - q: Do you map distribution lines as well as transmission lines?
    a: Yes. The method is the same at any width, with the line as a route file, the width you give us, and a register of the structures and vegetation change inside it, ranked by 500 m segment.
cta:
  title: Check your corridor against its cut-off date
  text: "Send the line route or the RAP corridor as a file, the wayleave width, the counties it crosses and the cut-off date that applies. We reply with the dated imagery that exists for it and a written proposal."
  button: Request a proposal
  secondary: Wayleave law guide
  secondary_href: /ke/wayleave-law-guide
---

::::section{id="cut-off" eyebrow="Cut-off-date checks" title="What was there at the census, and what appeared since" lead="Many resettlement action plans for new lines fixed their cut-off date long before construction, and some are re-ratified later. We compare dated imagery from the cut-off date that applies with today's, so you can see which structures were there at the census and which appeared since. The census and eligibility stay with you and your RAP consultant."}
:::::columns{split="1-1" align="center"}
::::col
:::figure{src="diagrams/ke-rap-wayleaves" alt="Plan-view schematic of two overhead lines drawn to one scale: a 220 kV line with a 40 m red band, 20 m each side, and a 132 kV line with a 30 m band, 15 m each side, each centred on the line with tower symbols along it; square markers stand for structures, red inside the band and teal outside" caption="The wayleave widths in published resettlement action plans for a 220 kV and a 132 kV line, drawn to one scale" credit="Schematic drawn by AfriScan for illustration; the structures are invented." size="half"}
Red: the wayleave each plan sets, divided equally either side of the centreline. The squares show how a register bands structures: <span class="band band--a">inside the wayleave</span> or <span class="band band--c">outside it</span>.
:::
::::
::::col
KETRACO's published resettlement action plans set a **40 m wayleave** (20 m each side) for a [220 kV line](https://www.ketraco.co.ke/sites/default/files/reports/Resettlement%20Action%20Plan%20for%20the%20proposed%20220kV%20Malindi-Weru%20Transmission%20Line.pdf) (final report April 2022) and a **30 m wayleave** (15 m each side) for a [132 kV line](https://www.ketraco.co.ke/sites/default/files/reports/Proposed%20RAP%20for%20the%20132kV%20Narok-Bomet.pdf) (final report November 2022). Their cut-off dates fall in December 2021, and in one county in November 2022, for lines not yet built.

Each plan defines the cut-off date as the completion of the census and asset inventory, and provides that a delay of two years or more would see it "ratified by the gazette notice". So the first question on any old corridor is which date applies. At least one KETRACO plan, for another 132 kV line, already lists "use of satellite imagery" in its methodology.

Documentation matters to the payments too. In May 2026 the Energy Cabinet Secretary told Parliament that payments to 163 of 836 people affected on three transmission lines were still pending, citing "budget constraints, documentation gaps and other administrative challenges" ([Capital FM, 6 May 2026](https://capitalfm.africa/govt-pays-sh2-23bn-in-wayleave-compensation/)).
::::
:::::

:::cards{cols="3"}
:::card{title="The cut-off record" icon="calendar" eyebrow="On the date"}
The structures inside the wayleave on the dated imagery closest to the cut-off date, each with an ID, coordinates, band and distance to the centreline.
:::
:::card{title="What appeared since" icon="compare" eyebrow="Change"}
New and removed structures between the cut-off imagery and today's, flagged automatically and confirmed by a reviewer, with before-and-after views.
:::
:::card{title="Where to look first" icon="search" eyebrow="For the field team"}
The 500 m segments with the most change, ranked, so the census update or verification visit starts where the corridor has moved most.
:::
:::
::::

::::section{id="operating-lines" tone="alt" eyebrow="Operating lines" title="Registers and vegetation along lines already built"}
:::::columns{split="2-1"}
::::col
On operating lines, the question moves from the cut-off date to the wayleave itself: which structures stand inside it, how close each is to the centreline, and what has changed since the last scan. We list each structure the review confirms, rate each 500 m segment by the number of structures it holds, and flag what is new at each repeat scan, so wayleave officers visit the busiest stretches first.

**Vegetation.** Corridors are cleared at periodic maintenance, and KETRACO's land FAQ says trees are paid for once, at construction, not at later maintenance clearing. We map where vegetation has regrown or been cleared along the wayleave between dated images, and where tall vegetation stands inside it from drone elevation models or open canopy-height data.

**The law behind access.** Energy infrastructure may be developed "on, through, over or under any public, community or private land" (Energy Act 2019, s.170), with entry to operate, inspect, repair or remove it (s.176) and lines across streets, roads, railways, forests, national parks and reserves (s.178). Occupiers claim for loss or damage within three months after the development (s.173(1)(b)), which is one more reason to keep a dated record from before the works.
::::
::::col
### What you receive

:::checklist
- A structure register with IDs, bands, distances and coordinates
- A rating for each 500 m segment, busiest first
- Vegetation regrowth and clearing between dated images
- Tall vegetation inside the wayleave, where elevation data exist
- New and removed structures at each repeat scan
- A PDF report in English and GIS layers
:::

[Vegetation, land cover and fire](/solutions/vegetation-land-cover-fire)
::::
:::::
::::

::::section{id="new-lines" eyebrow="New lines and interconnectors" title="Compare routes by the structures they would affect"}
:::::columns{split="1-1"}
::::col
Twenty-nine transmission lines are due for completion by 2028 ([KBC, 4 September 2025](https://www.kbc.co.ke/ketraco-to-complete-29-transmission-line-project-by-2028/)). Regional links include a 500 kV direct-current interconnector to Ethiopia, a 400 kV line towards Tanzania and a link to Uganda, and further lines are planned under the Transmission Master Plan 2025–2044.

For a new line, we score alternative alignments by the structures within 15 m and 20 m of each centreline, or within whatever half-width the design uses, so the land, compensation and resettlement each option implies can be weighed before a route is fixed. Once it is chosen, a dated record from before the preliminary survey shows what stood on the land, because the Land Act compensates damage from that survey work too (s.148(3)).
::::
::::col
:::callout{tone="note" title="Routes through forests"}
A 2026 amendment to the forest law allows easements and wayleaves in forests with the Kenya Forest Service's approval, and it has been challenged in the Environment and Land Court ([Capital FM, 2 July 2026](https://capitalfm.africa/green-belt-movement-seeks-to-nullify-new-forest-act/)). Where a route option crosses a forest, a land-cover baseline along each option shows how much tree cover it would cross.
:::

[Route and site selection](/solutions/route-site-selection)
::::
:::::
::::

::::section{id="towers" tone="alt" eyebrow="Near towers and substations" title="Land change that helps direct patrols"}
:::::columns{split="2-1"}
::::col
Imagery cannot see vandalism or theft; it shows land change. Between dated images we map new tracks to towers, cleared ground, fresh excavations and new structures near towers and substations, so your teams can decide where to patrol and what to check first. We map the ground, never people or vehicles, and we never publish imagery of a client's installations without written permission.
::::
::::col
:::callout{tone="legal" title="Drones near the line"}
Flying in or around high-tension cables without permission from the Kenya Civil Aviation Authority is treated as negligent or reckless operation under the Civil Aviation (Unmanned Aircraft Systems) Regulations 2025 (regulation 45). Any drone flight near a line needs that specific permission and is flown by a Kenyan company with a KCAA RPAS Operator Certificate.

[Drone law in Kenya](/ke/drone-regulations#restricted)
:::
::::
:::::
::::

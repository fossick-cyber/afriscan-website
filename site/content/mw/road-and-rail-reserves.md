---
key: rail-roads
template: industry
slug: road-and-rail-reserves
title: Road & Rail Reserve Encroachment in Malawi | AfriScan
description: Structures, cleared ground and excavations inside road and rail reserves in Malawi, measured against the reserve width, with dated change for RAPs.
h1: What is inside Malawi's road and rail reserves, mapped against the legal width
crumb: Road and rail reserves
eyebrow: Rail & roads · Malawi
lead: Malawi's road reserves are 60 m for main roads and 36 m for most others. Building, planting or new cultivation in them needs the highway authority's written consent, and the authority can direct that an encroachment is removed. AfriScan maps the structures and land change inside the reserve width you give us, segment by segment, and what changed between dated surveys. A person reviews every result.
buttons:
  - {label: Send the road or rail chainage file, intent: proposal}
  - {label: See sample outputs, key: results}
service:
  name: Road and rail reserve encroachment survey, Malawi
  type: Reserve encroachment survey, alignment baselines and change detection
  description: Structures, tracks and excavations inside road reserves, measured against the width for the road class, inside rail reserves and along buried cable routes in Malawi, with larger patches of new clearing flagged, compared between dated surveys, reviewed by a person and delivered as PDF reports and GIS layers.
og:
  headline: Road and rail reserves in Malawi, mapped against the legal width
  subline: Structures, cleared ground and excavations, and what changed between dated surveys
related: [route-site-selection, resettlement-cut-off-baselines, power-utilities]
faq:
  - q: Do you measure from the centre line or the reserve edge?
    a: >-
      From the centre line of the carriageway, unless you send the reserve boundary. Under s.10(2) of the
      consolidated Public Roads Act the reserve's centre line lies down the centre line of the carriageway
      unless the Minister directs otherwise by notice in the Gazette, so a 36 m reserve normally runs 18 m
      either side. Where the Minister has directed otherwise, a carriageway has been realigned, or a reserve
      was surveyed differently, send the reserve polygons and we measure from those.
  - q: Which width applies to our road?
    a: >-
      The one for its class, as the consolidated Act lists them: 60 m for a main road; 36 m for a secondary,
      tertiary or district road; 18 m for a branch or estate road. A newly built road can take a reserve
      of up to 60 m in total. Send the class of each section, or the widths you work to, and up to six
      widths can be reported in one survey.
  - q: Can you map farming inside the reserve?
    a: >-
      In part. The register lists structures. Alongside it, cropland along the road is shown from generalised
      land-cover maps, and larger patches of new clearing are flagged between dated surveys for a reviewer to
      check. Small plots, and gardens under tree cover, can be missed. Whether a plot had the highway
      authority's consent, or was already cultivated when the land became road reserve, is for the authority
      to establish.
  - q: Do you decide which structures must be removed?
    a: >-
      No. The highway authority may direct, in writing, anyone who encroaches on a road reserve to remove the
      encroachment (s.36(4)). We record what stands where, and on which imagery date; the authority, the
      owner and, where it comes to that, the courts decide what happens next.
  - q: What rail-reserve width do you use?
    a: >-
      We do not assume one. Send the rail-reserve polygons or the width the railway works to, and we list
      what stands inside the reserve and in a band beyond it, with the chainage along the line.
  - q: Can you see stolen rails or cut cables?
    a: >-
      No. Satellites cannot see theft or vandalism. We flag structures, new tracks, worn crossing paths and
      fresh excavations beside the railway or along a cable route between surveys, so your teams know where
      to check.
cta:
  title: Send the road or rail chainage file
  text: "The centre line or the reserve polygons (KML, KMZ, GeoJSON, Shapefile, GPX or GeoPackage), the road class or reserve width for each section, the districts and the chainage you use. For a road project, add the alignment options or the cut-off date. We reply with a scope, an imagery plan and a written proposal."
  button: Request a proposal
  secondary: Drone law in Malawi
  secondary_href: /mw/drone-regulations
---

::::section{id="problem" eyebrow="The problem, in your words" title="Reserves are wide on paper and narrow on the ground"}
:::cards{cols="2"}
:::card{title="“Gardens go right up to the carriageway.”" icon="leaf"}
The Roads Authority is running nationwide sensitisation against farming in road reserves, working with district councils and traditional leaders (June 2026). New cultivation in a reserve needs the highway authority's written consent.
:::
:::card{title="“Stalls and houses keep moving closer.”" icon="building"}
At trading centres and on the edges of towns, the reserve is where people trade and build. Each new structure is easier to discuss before it is finished than after the road is widened.
:::
:::card{title="“The upgrade is waiting on compensation.”" icon="calendar"}
A lender-funded upgrade of a road on the Nacala corridor was reported in July 2026 to be more than a year behind schedule, partly because of compensation issues along the construction path. A dated record of what stood on the land gives those questions a dated starting point.
:::
:::card{title="“Road works cut our fibre.”" icon="excavation"}
A telecom operator in Malawi has reported fibre outages caused by vandalism and by road construction works. Earthworks beside a buried route are visible from the air before a cable is hit.
:::
:::

### Who this is for

:::chips
- Road agency and district council engineers
- Design and supervision consultants
- Resettlement and RAP teams
- Railway reserve and property managers
- Fibre and cable route owners
- Lenders' environmental and social supervisors
:::
::::

::::section{id="widths" tone="alt" eyebrow="The Public Roads Act" title="The reserve widths, and what needs consent inside them"}
:::::columns{split="1-1"}
::::col
:::figure{src="diagrams/mw-road-reserves" alt="Plan-view schematic of three roads drawn to one scale, each with its reserve shaded red and centred on the carriageway: 60 m, 36 m and 18 m; square markers stand for structures, red inside the reserve and teal outside it" caption="Road reserves by road class, drawn to one scale" credit="Schematic drawn by AfriScan for illustration; the carriageways and structures are invented." size="half"}
From the top: a main road (60 m); a secondary, tertiary or district road (36 m); a branch or estate road (18 m), each centred on the carriageway, as the consolidated Public Roads Act sets it unless the Minister directs otherwise by Gazette notice (s.10(2)). Red squares stand inside the reserve, teal ones outside it.
:::
::::
::::col
The consolidated [Public Roads Act](https://malawilii.org/akn/mw/act/1962/11/eng@2017-12-31) (Cap. 69:02) lists road reserves of "60 metres" for a main road, "36 metres" for secondary, tertiary and district roads and "18 metres" for branch and estate roads, and under s.10(2) the reserve's centre line "shall in every case lie down the centre line of the carriageway of the road unless the Minister shall in any case otherwise direct by notice published in the Gazette".

- **Consent (s.10(6)).** Without "the consent in writing of the highway authority", no one may "erect or alter any structure", "plant any tree or bush" or prepare for cultivation land that was not prepared for cultivation when it became road reserve. Even with consent, no compensation is due for what was done if the land is later needed for the road.
- **Notice before works.** The highway authority gives one month's notice before works likely to damage a structure, three months for a building, with compensation for the damage.
- **New roads (s.28).** The reserve of a newly built road is up to "a total width of 60 metres", and "Compensation shall be payable".
- **Encroachment.** No one may encroach on a road or road reserve by making or altering a structure, ditch or other obstacle, or by planting trees. The highway authority may, by notice in writing, direct the person to remove it or fill it in (s.36(4)), and may do the work itself and recover the cost.
::::
:::::

:::::columns{split="2-1"}
::::col
:::callout{tone="legal" title="Read with care"}
In the consolidated text on MalawiLII the list of widths stands apart from s.10(1), which still speaks of a width "not exceeding sixty metres"; the list appears to be the 2017 amendment's substitution. Confirm the operative text with your counsel before relying on a width in a dispute. This page quotes no penalties. It summarises public rules for information, as read on 27 September 2026, and is not legal advice.
:::
::::
::::col
:::sources{law="mw" ids="public-roads-act,public-roads-amendment-2017"}
:::
::::
:::::
::::

::::section{id="registers" eyebrow="Road-reserve registers" title="A register of the reserve, measured from the carriageway"}
:::::columns{split="2-1"}
::::col
We buffer the road centre line you supply at half the reserve width either side, or use your reserve polygons, and list each structure the review confirms with its distance to the centre line, whether it stands inside the reserve, its chainage and its coordinates. Each 500 m segment is rated high, medium or low for **encroachment density** by a count rule, so your inspection teams know which stretches to walk first. The rating is not a road-safety rating.

New tracks and access routes onto the road are mapped between dates, and larger patches of new clearing along the reserve are flagged for a reviewer to check; cropland along the road comes from generalised land-cover maps. After the baseline, the road is re-surveyed on a schedule agreed with you: new and removed structures are flagged automatically and confirmed by a reviewer, with before-and-after views, and your team receives a notice after each survey.
::::
::::col
### What you receive

:::checklist
- A reserve register: ID, distance to the centre line, inside or outside the reserve, chainage, coordinates
- An encroachment-density rating for each 500 m segment
- Larger patches of new clearing along the reserve, and generalised cropland maps
- New tracks and access routes onto the road
- New and removed structures between surveys, confirmed by a reviewer
- A PDF report and GIS layers your engineers can load straight away
:::
::::
:::::

:::solutions{keys="right-of-way-monitoring,change-detection,route-site-selection" cols="3"}
:::
::::

::::section{id="resettlement" tone="alt" eyebrow="Upgrades and new alignments" title="Count the structures first, then date the record"}
:::::columns{split="2-1"}
::::col
Land for a public road is acquired under the Public Roads Act rather than the Land Act's public-utility procedure, and compensation is payable for the reserve of a newly built road (s.28). Lender-funded road projects add a resettlement action plan with a cut-off date. A road-corridor plan disclosed in September 2026 applied a cut-off date of 30 April 2023; its census was done in 2023, and valuation and verification in 2025.

Before the alignment is fixed, we compare the structures along each option within the same reserve width. At the cut-off date, a reviewed register on the closest dated imagery the archive holds gives the census team a starting layer and gives everyone the same record of what stood where. During implementation, a comparison shows what is new since the cut-off date, so questions about late arrivals are answered from the imagery rather than from memory.
::::
::::col
:::callout{tone="scope" title="For a road project"}
- Structure counts along each alignment option
- A dated baseline register at the cut-off date
- New structures since the cut-off date, confirmed by a reviewer
- Borrow pits and earthworks between dated drone surveys, where flown
- It supports the census and asset inventory; it does not replace them

[Cut-off-date baselines](key:resettlement-cut-off-baselines)
:::
::::
:::::
::::

::::section{id="rail" eyebrow="Rail reserves" title="What stands beside the railway"}
:::::columns{split="1-1"}
::::col
Malawi's railway runs under a concession and connects to the line to the port of Nacala in Mozambique. We do not assume a rail-reserve width: we work from the reserve polygons or the width the railway gives us.

Inside the reserve and in a band beyond it we list structures, new tracks, informal crossings that show as worn paths, and fresh excavations near the track, compared between dated surveys and confirmed by a reviewer. Paths under tree cover and very narrow ones can be missed.
::::
::::col
:::callout{tone="note" title="Along the track"}
- Structures inside the rail reserve and in a band beyond it, by chainage
- Worn paths across the line, where they show in the imagery
- Fresh excavation, spoil and cleared ground near the track
- Satellites cannot see theft of rails or fittings; the register shows your teams where the ground has changed
:::
::::
:::::
::::

::::section{id="fibre" tone="alt" eyebrow="Fibre and cable routes" title="Find the works before they find the cable"}
:::::columns{split="2-1"}
::::col
Road construction works are among the causes of fibre outages that a telecom operator in Malawi has reported, alongside vandalism. Along a buried route we map new excavations, trenches, spoil heaps, earthworks and construction between dated surveys, the third-party works that damage cables, so your patrols know where to check.

Only works visible at the surface are flagged, and each flag is confirmed by a reviewer. We cannot see a cable cut or cable theft, and the register does not replace route patrols.
::::
::::col
:::solutions{keys="excavation-mapping" cols="1"}
:::
::::
:::::
::::

::::section{id="nacala" eyebrow="The Nacala corridor" title="Routes that continue into Mozambique"}
:::::columns{split="1-1"}
::::col
Road and rail on the Nacala corridor continue into Mozambique to the port of Nacala. For a route that crosses the border, the proposal covers both sides, with the widths that apply in each country. The Mozambican rules, including the 50 m partial protection zone, are on our Mozambique pages.

[The Mozambique site](/mz/) · [The 50 m partial protection zone](/mz/50m-protection-zone)
::::
::::col
:::callout{tone="scope" title="Field and drone scope on the corridor"}
We send no field or drone teams to the areas of Mozambique where the UK Foreign, Commonwealth & Development Office advises against all or all but essential travel (advice of 20 August 2026): Cabo Delgado province, Memba and Eráti in Nampula, and Mecula and Marrupa in Niassa. In Malawi, drone flights stay away from border posts and borderlands.
:::
::::
:::::
::::

::::section{id="how" tone="alt" eyebrow="How it works" title="From chainage file to reviewed reserve register"}
:::steps
:::step{title="Agree the widths"}
You send the centre line or the reserve polygons, the road class or reserve width for each section and the chainage you use. The written proposal fixes the widths, the imagery and the deliverables.
:::
:::step{title="Screen the whole route"}
Structures and cleared ground are mapped along the full length from dated satellite imagery, open building datasets or imagery you already hold.
:::
:::step{title="Look closer where needed"}
Junctions, river crossings and borrow pits can be flown by drone, subject to the approvals and security clearances each job requires.
:::
:::step{title="Review and deliver"}
A person checks every result; each structure is measured, each segment rated, and the report and GIS files go to the contacts you name.
:::
:::

More on [how it works](/features), the [methodology](/methodology) and [imagery and data sources](/imagery), and on [rail and road work across Africa](/industries/rail-roads).
::::

::::section{id="scope" eyebrow="Honest scope" title="What we do, and what we don't"}
:::::columns{split="1-1"}
::::col
### What we do

:::checklist
- Map structures, tracks and excavations inside the reserve, and larger patches of clearing along it
- Measure each structure from the centre line, against the width for the road class or the reserve you supply
- Flag change between dated surveys, confirmed by a reviewer
- Deliver dated records that support consent, compensation, engagement and legal processes
:::
::::
::::col
### What we don't do

:::checklist{tone="no"}
- Decide whether a structure had consent or must be removed
- Detect theft of rails, fittings or cables, or identify who did anything
- Identify, count or follow people
- Replace a land surveyor, or prepare records for evictions
:::
::::
:::::
::::

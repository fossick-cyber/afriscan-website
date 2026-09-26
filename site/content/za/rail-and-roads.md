---
key: rail-roads
template: industry
slug: rail-and-roads
title: Rail & Road Reserve Encroachment Mapping SA | AfriScan
description: Structures inside rail reserves and SANRAL's building restriction area in South Africa, measured from the reserve boundary and compared between dated surveys.
h1: Rail reserves and road building restriction areas, mapped from the air
crumb: Rail and roads
eyebrow: Rail & roads · South Africa
lead: A dated register of the structures inside your rail reserve or within SANRAL's building restriction area, each measured from the reserve boundary you supply, with new structures, tracks and informal crossings flagged between surveys. For reserve managers, road agencies and the engineers planning new alignments. A person reviews every result.
buttons:
  - {label: Send us your reserve file, intent: proposal}
  - {label: See sample outputs, key: results}
service:
  name: Rail and road reserve encroachment survey, South Africa
  type: Reserve encroachment survey, building restriction area register and alignment baselines
  description: Structures inside rail reserves and national road building restriction areas in South Africa, measured from the reserve boundary, with informal crossings, tracks and change between dated surveys, reviewed by a person and delivered as PDF reports and GIS layers.
og:
  headline: Rail and road reserves in South Africa, mapped from the air
  subline: Structures measured from the reserve boundary, and what changed between surveys
related: [route-site-selection, change-detection, terrain-flood-post-event]
faq:
  - q: Do you measure from the centreline or the reserve boundary?
    a: Whichever the rule uses. SANRAL's building restriction area is measured from the boundary of the national road, so we need the road reserve boundary, not only the centreline; the 500 m around a point of intersection is measured from that point. For rail reserves, send the reserve polygons and we list what stands inside them and in a band beyond.
  - q: The building restriction area excludes urban areas. How do you handle that?
    a: We mask the areas you mark as urban, using your own mapping or the municipal boundaries you tell us to use, and report them separately so nothing is lost. Which land counts as urban for the Act is for SANRAL and your legal team to confirm.
  - q: Do you decide which structures must move?
    a: No. We record what stands where and when it first appears. Whether a structure has SANRAL's permission, predates the road or has to move is for the road or rail authority, the owner and the courts to decide, and relocation from a rail reserve is a housing matter as much as an operational one.
  - q: Can you map informal crossings of the railway?
    a: Yes, where they show as worn paths or tracks across the reserve in the imagery. Paths under tree cover or very narrow ones can be missed, so a reviewer confirms each one.
  - q: Can your borrow-pit volumes stand as a statutory survey?
    a: No. Volumes from drone elevation models support your registered surveyors and quantity surveyors; they are not a statutory survey or a valuation. Drone surveys are subject to the permits and authorisations each job requires.
cta:
  title: Send us your reserve file
  text: "The road reserve boundary or rail reserve polygons (Shapefile, GeoPackage, KML or GeoJSON), the points of intersection if you have them, and the areas you treat as urban. We reply with a scope, an imagery plan and a written proposal."
  button: Request a proposal
  secondary: Route and servitude baselines
  secondary_href: /za/transmission-route-baseline
---

::::section{id="problem" eyebrow="The problem, in your words" title="Reserves are long, narrow and easy to build in"}
:::cards{cols="2"}
:::card{title="“The rail reserve has filled up.”" icon="rail"}
Commuter lines through the metros run beside some of the densest settlement in the country. In March 2024 the Housing Development Agency told Parliament that 3,941 households were still living on the Central Line between Philippi and Khayelitsha ([GroundUp, 14 March 2024](https://groundup.org.za/article/thousands-of-shack-dwellers-are-still-living-on-prasa-rail-lines-in-langa-philippi-and-khayelitsha/)). Planning anything on a reserve like that starts with a current, located picture.
:::
:::card{title="“Structures keep going up beside the national road.”" icon="building"}
Outside urban areas, any structure within 60 m of a national road's boundary needs SANRAL's written permission. Filling stations, stalls, walls and houses still appear in the restriction area, and they are cheapest to deal with before they are finished.
:::
:::card{title="“We need the count before the land take.”" icon="clipboard"}
A widening, a new interchange or a new alignment needs to know what stands on the land before negotiations start, and needs that record dated.
:::
:::card{title="“Nobody recorded the borrow pits.”" icon="excavation"}
Construction leaves borrow pits, haul roads and spoil behind. Dated imagery of each stage settles questions about rehabilitation and volumes later.
:::
:::

### Who this is for

:::chips
- Rail reserve and property managers
- Road agencies' land-acquisition and building-control teams
- Commuter-rail and metro planning teams
- Housing and human-settlements programmes working on reserves
- EPC contractors and supervising engineers
- Environmental assessment practitioners on new alignments
:::
::::

::::section{id="roads" tone="alt" eyebrow="National roads" title="SANRAL's building restriction area, measured the way the Act draws it"}
:::::columns{split="1-1" align="center"}
::::col
Section 48 of the [SANRAL Act 7 of 1998](https://www.sagc.org.za/pdf/legislation/S%20A%20National%20Roads%20Agency%20Act%207%20of%201998.pdf) requires SANRAL's written permission for any structure on or over a national road, or in its building restriction area. The building restriction area is land **outside urban areas** that lies:

- within **60 m of the boundary** of the national road; or
- within **500 m of a point of intersection**, where the national road crosses or links with another road.

So the register is measured from the road reserve boundary, not from the centreline, and around each point of intersection. We list each structure inside either zone with its distance from the boundary or from the intersection point, mask the areas you mark as urban and report them separately, and flag new structures between surveys.
::::
::::col
:::figure{src="diagrams/za-road-restriction" alt="Schematic of a national road with its road reserve, a 60 m building restriction strip along each side of the reserve boundary, and a 500 m circle around a point of intersection with a side road; an urban area on the right is excluded from the restriction area" caption="Schematic: the building restriction area along a national road" size="half" credit="Schematic drawn by AfriScan for illustration, not to scale. It is not a real road."}
Orange: the 60 m strip beyond each road reserve boundary. Dashed circle: 500 m around a point of intersection. Grey: an urban area, outside the building restriction area. Squares are structures, red inside the restriction area and teal outside it.
:::
::::
:::::

Provincial and municipal roads have their own reserve widths and building lines under provincial legislation and each authority's standards. Send the widths and we measure to them in the same way.
::::

::::section{id="rail" eyebrow="Rail reserves" title="What stands inside the reserve, and what crosses it"}
:::::columns{split="2-1"}
::::col
The Transnet Rail Infrastructure Manager manages, operates and maintains the national rail network ([TRIM](https://www.transnet.net/TRIM)), after the separation of Transnet Freight Rail announced in October 2024; commuter lines in the metros run on their own reserves. Reserve widths differ from line to line, so we work inside the reserve polygons you supply and in a band beyond them.

The register lists each structure with its distance to the track centreline and to the reserve edge, and rates each 500 m stretch for encroachment density. Worn paths and informal crossings over the line are mapped as tracks, and new ones between surveys are flagged and confirmed by a reviewer.

Where a reserve is already densely settled, the register serves planning rather than enforcement: a dated structure layer, and on request estimated households with the assumptions stated, for the municipality, the housing agency and the rail operator to work from the same map. It is not a census and it does not decide who moves.
::::
::::col
### What you receive

:::checklist
- Structures inside the reserve and in a band beyond it, with distances to the track and the reserve edge
- A density rating for each 500 m stretch
- Informal crossings and tracks across the reserve
- New and removed structures between surveys, with before-and-after views
- Estimated households with stated assumptions, on request
- A PDF report and GeoPackage, GeoJSON, KMZ and Shapefile layers
:::
::::
:::::
::::

::::section{id="new-alignments" tone="alt" eyebrow="New alignments and widening" title="Count the structures before the land take"}
:::::columns{split="1-1"}
::::col
For a new alignment, a widening or an interchange, we compare the structures along each option, screen steep, erosion-prone and flood-prone stretches and drainage crossings, and fix a dated baseline once the alignment is chosen. Where the project is lender-financed, the dated register supports the census and asset inventory and the IFC Performance Standard 5 cut-off-date records; it does not replace them.

Expropriation for public roads still runs under the Expropriation Act 63 of 1975, because the 2024 Act faces constitutional challenges ([PMG, 19 December 2025](https://pmg.org.za/committee-question/35164/)). The survey that defines what is taken is a professional land surveyor's.
::::
::::col
### During construction {#construction}

Dated drone orthophotos and elevation models of the works, compared side by side, show where ground has been cut or filled between visits, with stockpile and borrow-pit volumes for your surveyors to check. Drone surveys are subject to the permits and authorisations each job requires; flights over or within 50 m of a public road need specific SACAA approval, and flights over a road need it closed ([drone law](/za/drone-regulations#fifty-metres)).
::::
:::::
::::

::::section{id="how" eyebrow="How it works" title="From reserve file to reviewed register"}
:::steps
:::step{title="Scope"}
You send the reserve boundary or polygons, the points of intersection, the areas you treat as urban and what the register is for. We agree the zones, imagery and deliverables in a written proposal.
:::
:::step{title="Screen from satellite"}
Structures and tracks are mapped along the whole reserve from dated satellite imagery, open building datasets or the imagery you hold.
:::
:::step{title="Check up close"}
Drone surveys capture the stretches that need detail, subject to the permits and authorisations each job requires.
:::
:::step{title="Review and deliver"}
A person checks every result; each structure is measured to the boundary or intersection and the report and GIS layers go to the contacts you name.
:::
:::
::::

::::section{id="scope" tone="alt" eyebrow="Honest scope" title="What we do, and what we don't"}
:::::columns{split="1-1"}
::::col
### What we do

:::checklist
- Map structures, tracks, crossings and cleared ground inside and beside the reserve
- Measure each structure the way the rule is drawn: from the boundary, the centreline or the intersection
- Flag change between dated surveys, confirmed by a reviewer
- Deliver dated, credited records that support planning, building control and engagement
:::
::::
::::col
### What we don't do

:::checklist{tone="no"}
- Decide whether a structure has permission, or who must move
- Identify, count or follow people or vehicles
- Certify volumes or replace a professional land surveyor
- Take any part in removals: decisions about structures rest with the authorities and the courts
:::
::::
:::::
::::

---
key: power-utilities
template: industry
slug: power-lines
title: Power-Line Servitude Encroachment, South Africa | AfriScan
description: Structures, fresh excavations and vegetation change inside transmission and distribution servitudes in South Africa, measured to the line and reviewed.
h1: Servitude encroachment on power lines, mapped from the air
crumb: Power lines
eyebrow: Power & utilities · South Africa
lead: A dated register of the structures inside your transmission and distribution servitudes, each measured to the line, with fresh digging near towers, vegetation change and anything new since the last survey flagged for your servitude officers. Satellite first, your own drone orthophotos where you already fly, and a person reviews every result.
buttons:
  - {label: Send us your line route, intent: proposal}
  - {label: See sample outputs, key: results}
service:
  name: Power-line servitude encroachment survey, South Africa
  type: Servitude encroachment survey, change detection and route baselines
  description: Structures, fresh excavations and vegetation change inside the servitudes of transmission and distribution lines in South Africa, each measured to the line and grouped by 500 m stretch or by span, compared between dated surveys and reviewed by a person, with PDF reports and GIS layers.
og:
  headline: Power-line servitudes in South Africa, mapped from the air
  subline: Structures by distance to the line, change between surveys, reviewed by a person
related: [route-site-selection, vegetation-land-cover-fire, za-land-invasion]
faq:
  - q: Can you measure clearance to conductors?
    a: No. We show where structures stand in the servitude and how far each is from the line, and where tall vegetation stands, from drone elevation models where flown and open canopy-height data for wider context. None of that is a measured clearance to a conductor, which stays with your line engineers.
  - q: Which servitude widths do you use?
    a: The widths registered for each line, which vary by voltage and by servitude. Send the widths per line or section, or the servitude polygons from your GIS, together with any building line or company standard you work to. Up to six widths can be reported in one survey, each structure with its distance to the line and its band.
  - q: Can you report by span rather than by distance along the line?
    a: Yes. Send the tower positions with the route and the register groups structures by span between towers, as well as by 500 m stretch, so the list lines up with how your maintenance teams already work.
  - q: Do you detect cable theft or tower vandalism?
    a: No. We flag fresh excavation, spoil and works visible at the surface near towers and along the servitude between surveys, so your teams know where to check. What happened there, and why, is for your field teams to establish.
  - q: We already fly drones along our lines. Can you work with that imagery?
    a: Yes. Send the orthophotos or GeoTIFFs and we run the same structure, change and vegetation analysis on them, with a reviewer confirming every result. Between your flights, dated satellite scenes show what changed where they exist for the line.
  - q: How do your fire notices relate to AFIS?
    a: South Africa already has a free satellite fire information service in CSIR's AFIS. Our notices are an add-on for the lines you name, with hotspots from NASA FIRMS filtered to your servitude buffers, checked by a reviewer and followed by a burnt-area map. They are not an emergency or early-warning service, and small, short-lived or cloud-covered fires can be missed.
  - q: Can the register be used to register a servitude?
    a: It can show your land surveyor what stands on the proposed strip, but a servitude diagram is the work of a professional land surveyor and needs Surveyor-General approval. Our registers are monitoring and planning records built from imagery, not cadastral surveys.
cta:
  title: Send us your line route
  text: "KML, KMZ, GeoJSON, Shapefile, GPX or GeoPackage, with tower positions if you have them, the servitude widths and the province. Tell us if your team already flies the line. We reply with a scope, an imagery plan and a written proposal."
  button: Request a proposal
  secondary: Drone law in South Africa
  secondary_href: /za/drone-regulations
---

::::section{id="problem" eyebrow="The problem, in your words" title="What goes up in the servitude between inspections"}
:::cards{cols="2"}
:::card{title="“Houses keep going up under our lines.”" icon="houses"}
Eskom has publicly urged people not to build under high-voltage lines ([Engineering News, 19 February 2024](https://www.engineeringnews.co.za/article/eskom-urges-public-not-to-build-structures-below-high-voltage-powerlines-2024-02-19)). Structures still appear, and each one is a safety concern, a maintenance-access problem and a harder conversation once it is occupied.
:::
:::card{title="“We only see a span when the patrol gets there.”" icon="route"}
Lines run for hundreds of kilometres across farmland, townships and open veld, and inspections cover them on a cycle. What changed in between is found late, sometimes by a maintenance crew that cannot reach a tower.
:::
:::card{title="“Wayleave applications keep crossing our servitudes.”" icon="clipboard"}
Roads, pipelines, fibre, housing projects and solar plants all want to cross or share the strip. Each application needs a view of what already stands there, and wayleaves can end up in court: in February 2026 the Gauteng High Court gave judgment on a refused wayleave for a solar plant crossing an Eskom servitude ([Eskom statement](https://www.eskom.co.za/eskom-notes-the-high-court-judgment-in-a-matter-brought-by-sibanye-stillwater-and-others-for-a-wayleave-application/)).
:::
:::card{title="“Servitude acquisition is holding up the new line.”" icon="calendar"}
The national Transmission Development Plan names servitude acquisition among its risks, and records that the Kusile–Lulamisa 400 kV line "was delayed due to servitude acquisition challenges" ([TDP 2024](https://www.ntcsa.co.za/wp-content/uploads/2024/12/TDP-2024-Public-Report_Rev1.pdf)).
:::
:::

### Who this is for

:::chips
- Servitude and wayleave officers
- Land and rights managers
- Line maintenance and vegetation managers
- Asset and GIS teams
- Land-acquisition teams on new lines
- Transmission bidders and EPC contractors
- IPP grid-connection teams
- Municipal electricity departments
:::
::::

::::section{id="servitudes" tone="alt" eyebrow="Transmission and distribution servitudes" title="A register of the structures inside the servitude, then what changed"}
:::::columns{split="2-1"}
::::col
We buffer the line route you supply in its UTM zone, at the widths registered for each servitude, and list each structure the review confirms with its distance to the line, its band, its chainage and its coordinates. Each 500 m stretch is rated for **encroachment density**, high, medium or low, by a count rule; with tower positions, the register also groups structures by span. The rating tells your servitude officers where to go first. It is not a safety or clearance rating.

Around substations and depots the line becomes a boundary: we list what stands inside the site and in a ring around it.

After the baseline, lines are re-surveyed on a schedule you agree with us. New and removed structures between dated surveys are flagged automatically and confirmed by a reviewer, with before-and-after views, and your team receives a notice after each survey saying what changed and where. Fresh digging and spoil near towers, new tracks across the servitude and cleared ground that often comes before building are flagged the same way, so field teams know where to check. They do not replace inspections.
::::
::::col
### What you receive

:::checklist
- A servitude register: ID, distance to the line, band, span, chainage and coordinates
- A density rating for each 500 m stretch, and a ranked list of stretches
- Structures inside and around substations and depots
- New and removed structures between surveys, with before-and-after views
- A change notice after each survey
- Fresh excavation, new tracks and cleared ground near the line, flagged for checking
- A PDF report and GeoPackage, GeoJSON, KMZ and Shapefile layers
:::
::::
:::::

:::figure{src="diagrams/corridor" alt="Schematic of a line route with a 50 m band and a 100 m band, structures coloured by distance to the line, and a strip below rating each 500 m stretch low, medium, high, medium and low from the structures inside the widest band" caption="Schematic: how a servitude register is read" size="wide" credit="Schematic drawn by AfriScan for illustration. It is not a real line, and the widths shown are examples: your register uses the widths in your servitudes."}
Structures are coloured by band: <span class="band band--a">Within 50 m</span> <span class="band band--b">50 to 100 m</span> <span class="band band--c">Beyond 100 m</span>. The strip below rates each 500 m stretch from the structures inside the widest band: high above five, medium one to five, low none.
:::

:::solutions{keys="right-of-way-monitoring,change-detection,excavation-mapping" cols="3"}
:::
::::

::::section{id="wayleaves" eyebrow="Wayleaves" title="What stands where another service wants to cross"}
:::::columns{split="1-1"}
::::col
A wayleave is permission to place a service in, or work in, another party's servitude or road reserve. When an application to cross your servitude arrives, a dated picture of the strip around the crossing shows what the applicant's drawings may leave out: structures, tracks, cleared ground and recent excavation, each located and measured to your line.

The same works the other way. When your own line needs a wayleave through a municipal road reserve or a pipeline servitude, a dated baseline of what stands there before construction starts protects both sides later.
::::
::::col
:::callout{tone="note" title="For each crossing"}
- A register of structures within the widths you set around the crossing point
- Tracks, cleared ground and excavation visible in the latest imagery
- The imagery source and date, and earlier scenes where the history matters
- GIS layers your wayleave team can overlay on the applicant's drawings
:::
::::
:::::
::::

::::section{id="new-lines" tone="alt" eyebrow="New lines and servitude acquisition" title="Count the structures before the route is fixed"}
:::::columns{split="2-1"}
::::col
The [Transmission Development Plan 2024](https://www.ntcsa.co.za/wp-content/uploads/2024/12/TDP-2024-Public-Report_Rev1.pdf) plans about 14,494 km of new lines and 210 transformers for 2025 to 2034. The first Independent Transmission Projects prequalified seven bidders on 15 December 2025, with the final request for proposals expected in the 2026/27 financial year ([DBSA](https://www.dbsa.org/sites/default/files/media/documents/2025-12/15%20Dec%202025%20ITP%20PQBs%20and%20REIPPPP%20BW7.pdf)). Every new line needs a route, a servitude and a record of what stood on the land when the servitude was negotiated.

The DFFE's [power-line standard (GN 2313, 2022)](https://www.dffe.gov.za/sites/default/files/legislations/nema_powerlinessubstationsdevelopmet_g47095gon2313.pdf) excludes compliant lines inside the strategic transmission corridors from environmental authorisation. That removes one of the moments when land is described in detail, so a land baseline is needed earlier in the programme.

We compare the structures along alternative alignments, screen slope, drainage crossings and flood exposure, and pull the imagery history of contested parcels. Once the route is fixed, a dated register with reviewer categories supports landowner engagement and, where lenders apply IFC Performance Standard 5, the cut-off-date records. Expropriation, where it is used, still runs under the Expropriation Act 63 of 1975: the Minister's reply to Parliament of 19 December 2025 says the 2024 Act faces constitutional challenges ([PMG](https://pmg.org.za/committee-question/35164/)).
::::
::::col
:::callout{tone="scope" title="Route baselines for new transmission lines"}
Alignment options compared by the structures they affect, then a dated baseline before servitude acquisition starts, for utilities, ITP bidders and IPPs.

[Route and servitude baselines](/za/transmission-route-baseline)
:::
::::
:::::
::::

::::section{id="vegetation-fire" eyebrow="Vegetation and fire" title="Where the servitude has regrown, and where it has burnt"}
:::cards{cols="3"}
:::card{title="Cleared and regrown vegetation" icon="leaf"}
Maps of where vegetation in the servitude has been cleared or has grown back between dates, from Copernicus Sentinel data. It shows larger patches of change, not individual trees.
:::
:::card{title="Tall vegetation in the servitude" icon="tree"}
Where tall vegetation stands, from drone elevation models where flown and open canopy-height data for wider context. The open data is older in places, so a drone survey gives the current picture. It is not a measured clearance to conductors.
:::
:::card{title="Fire near the line" icon="alert"}
Satellite-detected fire hotspots from NASA FIRMS, filtered to your servitude buffers and checked by a reviewer, then burnt-area maps. CSIR's AFIS already covers South Africa, so these notices complement it for the lines you name. Not an emergency or early-warning service.
:::
:::

[Vegetation, land cover and fire, in detail](/solutions/vegetation-land-cover-fire)
::::

::::section{id="your-imagery" tone="alt" eyebrow="Your drone programme" title="Get more from the drone imagery you already fly"}
:::::columns{split="2-1"}
::::col
Several South African utilities and metros run their own drone programmes under their own UASOC ([SACAA operator list](https://www.caa.co.za/industry-information/flight-operations/)). The flying is rarely the bottleneck; turning each campaign's orthophotos into a list of what is new, where, and how close it is to the line is. Send the orthophotos or GeoTIFFs and we return the register, the change layers and the vegetation maps, with a person reviewing every result.

Where you want flagged spans flown, drone surveys are subject to the permits and authorisations each job requires. Along a line those include specific approvals for flights within 50 m of structures, people or public roads, the landowner's signed permission for every flight, and SACAA notice on form CA 101-20 before flying near a substation that is a national key point.
::::
::::col
:::callout{tone="legal" title="Drone law in South Africa"}
The UASOC and Air Service Licence, the 50 m rules, landowner permission and key points, with a 12-point checklist for clients, as of 26 September 2026.

[Read the guide](/za/drone-regulations)
:::
::::
:::::
::::

::::section{id="how" eyebrow="How it works" title="From line route to reviewed register"}
:::steps
:::step{title="Scope"}
You send the line route, tower positions if you have them, the servitude widths and what the register is for: maintenance, wayleaves, a new route or a baseline before acquisition. We agree the bands, imagery and deliverables in a written proposal.
:::
:::step{title="Screen from satellite"}
Structures are mapped along the whole line from dated satellite imagery, open building datasets or the orthophotos you hold.
:::
:::step{title="Check up close"}
Your own flights or drone surveys cover the spans that need detail, subject to the permits and authorisations each job requires.
:::
:::step{title="Review and deliver"}
A person checks every result; each structure is measured, banded and rated, and the report and GIS layers go to the contacts you name.
:::
:::

More on [how it works](/features), the [methodology](/methodology) and [imagery and data sources](/imagery).
::::

::::section{id="scope" tone="alt" eyebrow="Honest scope" title="What we do, and what we don't"}
:::::columns{split="1-1"}
::::col
### What we do

:::checklist
- Map structures, cleared ground, fresh excavations and tracks visible from the air
- Measure each structure to the line, at your servitude widths
- Flag change between dated surveys, confirmed by a reviewer
- Deliver dated, credited records that support your wayleave, engagement and legal processes
:::
::::
::::col
### What we don't do

:::checklist{tone="no"}
- Measure conductor clearances or inspect towers and hardware
- Detect theft or vandalism, or identify who did anything
- Identify, count or follow people
- Replace a professional land surveyor, or decide whether a structure is lawful
:::
::::
:::::
::::

---
key: power-utilities
template: industry
slug: power-lines
title: Power-Line Servitude Encroachment in Botswana | AfriScan
description: Structures, ploughed fields and cleared ground inside 33 kV to 400 kV power-line servitudes in Botswana, measured to the line, dated and reviewed by a person.
h1: What stands and grows inside power-line servitudes in Botswana
crumb: Power lines
eyebrow: Power transmission · Botswana
lead: A dated register of what stands and grows inside your line servitudes, from the 400 kV backbone down to 66 kV sub-transmission and 33 kV distribution. Each structure is measured to the line, ploughed and cleared ground inside the servitude is mapped, each 500 m stretch is rated, and anything new since the last survey is flagged. For new lines and substations, a baseline before the route is fixed. Satellite first, and a person reviews every result.
buttons:
  - {label: Send us your line route, intent: proposal}
  - {label: See sample outputs, key: results}
service:
  name: Power-line servitude survey, Botswana
  type: Servitude encroachment survey, route baseline and change detection
  description: Structures, ploughed fields, cleared ground and vegetation change inside the servitudes of transmission, sub-transmission and distribution lines in Botswana, each structure measured to the line and grouped by 500 m stretch, with route baselines for new lines and substations, compared between dated surveys and reviewed by a person, delivered as PDF reports and GIS layers.
og:
  headline: Power-line servitudes in Botswana, mapped and dated
  subline: Structures, fields and vegetation inside the servitude, reviewed by a person
related: [route-site-selection, resettlement-cut-off-baselines, vegetation-land-cover-fire]
faq:
  - q: Which servitude widths do you measure for lines in Botswana?
    a: The widths that apply to each of your lines. The only widths we found published are the typical designs for BPC's World Bank-financed 66 kV and 33 kV lines, with permanent rights of way of 30 m and 15 m, which "will be further defined" as the project is implemented ([ESMF, Table 2-1](https://documents.worldbank.org/curated/en/099052224012584077/pdf/P18122113057240211a4a61be136f2f526c.pdf)). For 132 kV, 220 kV and 400 kV lines, send the width of each line or the servitude polygons from your GIS. Up to six widths can be reported in one survey.
  - q: Can a new 66 kV or 33 kV line be built entirely inside a road reserve?
    a: Often not, on the framework's own reading. Main-road servitudes are typically 60 m wide and the Roads Department wants lines as close to their edge as possible, so about half the line's right of way is likely to extend onto the land beside the road, and the framework calls the assumption that such lines can be built within road reserves without affecting surrounding land uses "optimistic". That strip beside the road is what we map before the alignment is fixed.
  - q: We plan to build near a power line. Can you help with our request to BPC?
    a: We can map what already stands in and around the servitude on a dated image, with distances to the line, as map layouts to go with your letter of request to BPC's technical-services desk. Whether the development can go ahead, and on what terms, is for BPC to decide.
  - q: Can our register go with a wayleave application to the Land Authority?
    a: It can show your land surveyor what stands on the strip, but it is not a cadastral survey. The framework asks consultants for survey reports and cadastral survey drawings that meet the Land Authority's submission standard, and cadastral surveys are regulated and approved by the Department of Surveys and Mapping under the Land Survey Act (Cap. 33:01).
  - q: Is a house inside the servitude on tribal land unlawful?
    a: That is not for us to say. Tribal land, allocated by Land Boards, is the most common form of tenure in Botswana, and a structure may stand on a plot that was granted before the line, or after it. We record what stands where, and when it first appears in the imagery; the Land Board, the Land Tribunal and the courts decide what it means.
  - q: Do you detect cable theft or vandalism on the network?
    a: No. Satellites cannot see theft, vandalism or people. We map land change, such as fresh digging, spoil, new tracks, cleared ground and new structures along a line or cable route, so that your patrols know where to look first.
  - q: Can you measure clearance to the conductors?
    a: No. We show where structures and tall vegetation stand in the servitude and how far each structure is from the line. A measured clearance to a conductor stays with your line engineers.
cta:
  title: Send us the line route
  text: "KML, KMZ, GeoJSON, Shapefile, GPX or GeoPackage, with pole or tower positions if you have them, the voltage and servitude width of each line, and the district. For a new line, send the alignment options. We reply with a scope, an imagery plan and a written proposal."
  button: Request a proposal
  secondary: Drone law in Botswana
  secondary_href: /bw/drone-regulations
---

::::section{id="problem" eyebrow="The problem, in your words" title="What happens in a servitude between patrols"}
:::cards{cols="2"}
:::card{title="“Fields are ploughed right up to the poles.”" icon="leaf"}
On the typical designs for the 66 kV and 33 kV lines of BPC's World Bank-financed project, the permitted use of the permanent servitude is "Grazing. No cultivation or built infrastructure permitted." ([ESMF, May 2024, Table 2-1](https://documents.worldbank.org/curated/en/099052224012584077/pdf/P18122113057240211a4a61be136f2f526c.pdf)). Ploughing fields, livestock enclosures and new houses beside a line are where that rule meets the land.
:::
:::card{title="“The new line follows the road, and still needs land.”" icon="route"}
The same framework notes that main-road servitudes are typically 60 m wide and that the Roads Department wants lines placed as close to their edge as possible, so half a line's right of way is likely to fall on the land beside the road: 15 m for a 66 kV line and 7.5 m for a 33 kV line.
:::
:::card{title="“Someone wants to build next to our line.”" icon="building"}
BPC's technical-services desk takes requests about "New Developments that Encroache on BPC Servitude", with a letter of request and map layouts, alongside requests to relocate BPC infrastructure and to manage vegetation on it ([BPC technical services](https://www.bpc.bw/technical-services/)). Each one is easier to settle with a dated picture of what already stands there.
:::
:::card{title="“The bush grows back before the next clearing.”" icon="tree"}
On those typical designs the right of way is "Cleared of bush and trees" and "maintained to prevent trees from damaging conductors". Knowing which stretches have regrown, or burnt, tells the clearing crews where to start.
:::
:::

### Who this is for

:::chips
- Servitude and wayleave officers
- Transmission and distribution line managers
- Vegetation and bush-clearing supervisors
- GIS and asset-records teams
- Project teams on new lines and substations
- Resettlement and permissions teams
- Engineering consultants preparing wayleave applications
- Developers planning to build near a line
:::
::::

::::section{id="servitudes" tone="alt" eyebrow="Existing lines" title="A register of what stands inside each servitude, then what changed"}
:::::columns{split="2-1"}
::::col
We buffer the line route you supply in its UTM zone, at the servitude width of each line, and list each structure the review confirms with its distance to the line, its band, its chainage and its coordinates. Where the imagery shows them, ploughed fields and cleared ground inside the servitude are mapped too: on a line whose servitude is meant for grazing, a field matters as much as a building. Each 500 m stretch is rated high, medium or low for encroachment density by a count rule, so your servitude officers know which stretches to visit first. The rating is not a safety or clearance rating.

Botswana's national grid comprises 1,246 km of 400 kV, 2,200 km of 220 kV and 2,162 km of 132 kV transmission lines and 54 substations, with sub-transmission at 66 kV and distribution at 33 kV and 11 kV ([BPC, The Transmission Grid](https://www.bpc.bw/portfolio-items/gaborone-power-project/)). Satellite screening covers whole lines at once, and a line is then re-surveyed on the schedule you agree with us. New and removed structures between dated surveys are flagged automatically and confirmed by a reviewer, with before-and-after views, and your team receives a notice after each survey saying what changed and where.

Around substations the line becomes a boundary: we list what stands inside the site and in a ring around it.
::::
::::col
### What you receive

:::checklist
- A servitude register: ID, distance to the line, band, chainage and coordinates
- Ploughed and cleared ground inside the servitude, where the imagery shows it
- A density rating for each 500 m stretch, and a ranked list of stretches
- Structures inside and around substations
- New and removed structures between surveys, with before-and-after views
- New tracks and fresh digging near the line, flagged for checking
- A PDF report and GeoPackage, GeoJSON, KMZ and Shapefile layers
:::
::::
:::::

:::solutions{keys="right-of-way-monitoring,change-detection,excavation-mapping" cols="3"}
:::
::::

::::section{id="road-reserves" eyebrow="Lines along roads" title="When the line follows the road, half its servitude is on the land beside it"}
New lines are often routed along main roads. The Environmental and Social Management Framework (ESMF) for BPC's World Bank-financed 66 kV and 33 kV lines, whose conceptual routes share the servitudes of existing main roads, explains why that does not keep them off other people's land. The Roads Department, which approves joint use of road servitudes, generally wants lines "as close to the edge of the servitude as possible", so half of the right of way is likely to extend "into adjacent land use". The framework concludes that the assumption that the lines can be built within existing road reserves without affecting surrounding land uses "is optimistic", and that the right of way "will extend beyond the boundaries of existing road servitudes in places".

:::figure{src="diagrams/bw-road-reserve" alt="Plan-view schematic drawn to one scale: a main road in a 60 m road reserve, a 66 kV line on poles along the reserve's upper edge, and the line's 30 m right of way shaded red, half inside the reserve and 15 m on the land beside it; a ploughed field runs into the right of way, square markers inside it are red and those outside are teal" caption="Schematic: a 66 kV line along the edge of a road reserve" credit="Schematic drawn by AfriScan for illustration, after the typical designs in the ESMF of May 2024. The road, field and structures are invented." size="wide"}
Structures are coloured by where they stand: <span class="band band--a">Inside the right of way</span> <span class="band band--c">Outside it</span>. The hatched block is a ploughed field. The widths are the framework's: a 60 m road servitude, a 30 m right of way, 15 m of it beyond the reserve.
:::

:::::columns{split="1-1"}
::::col
Placing power, water or communication services inside a road reserve also needs a Department of Roads "Access Control to Road Reserve Space (Wayleave)" permit, and the A1 drawings sent with the application must show the utilities affected or close by ([gov.bw](https://www.gov.bw/transport/access-control-road-reserve-space-wayleave)).
::::
::::col
We map what stands along both edges of the reserve before the alignment is fixed: structures, ploughed fields, cleared ground and tracks, each measured to the proposed line and to the reserve boundary you give us.
::::
:::::
::::

::::section{id="new-lines" tone="alt" eyebrow="New lines, substations and interconnectors" title="Count the structures before the route is fixed"}
:::::columns{split="2-1"}
::::col
BPC's published project list includes grid extensions, new 132 kV substations and two planned 400 kV interconnectors with neighbouring countries ([BPC](https://www.bpc.bw/portfolio-items/gaborone-power-project/)), and the Minister of Minerals and Energy told Parliament in March 2026 that transmission is to operate as a Transmission System Operator, allowing Independent Transmission Providers to take part ([Committee of Supply speech, 4 March 2026](https://www.parliament.gov.bw/documents/BUDGET-2026-27-MARCH---Ministry-of-Minerals-and-Energy-Committee-of-Supply-Final-Speech_12_59_20_04_08_2026.pdf)). Every new line needs a route, a servitude and a record of what stood on the land when the route was agreed.

We compare alternative alignments by the structures, fields and cleared ground they affect, and once the route is fixed we produce a dated register with reviewer categories for the land and resettlement teams. Under the [Botswana Power Corporation Act (Cap. 74:01)](https://www.bpc.bw/wp-content/uploads/2025/04/BPC-Act.pdf), BPC's functions count as public purposes for the law on compulsory acquisition (s.24), the terms of any resettlement of people living on communally owned land are subject to the agreement of the Government and of the local authority of the area (s.25), and BPC must do as little damage as possible and make full compensation, with arbitration where the amount is disputed (s.26).

Where lenders' standards apply, a resettlement action plan comes before the works: the framework for BPC's World Bank-financed lines requires that "all compensation of Project affected persons (PAPs) must be completed before construction of the subcomponent starts". IFC Performance Standard 5 lets a project decline compensation for people who move into the project area after a cut-off date that "has been clearly established and made public" ([PS5, para. 23](https://www.ifc.org/content/dam/ifc/doc/2010/2012-ifc-performance-standard-5-en.pdf)). A dated register made on that date is the record the census and asset inventory go back to. It supports them; it does not replace them.
::::
::::col
:::callout{tone="scope" title="Route and cut-off baselines"}
- Structure counts along each alignment option
- A dated register on the cut-off date, with imagery, coordinates and IDs
- What has appeared since, confirmed by a reviewer

[Route and site selection](/solutions/route-site-selection) · [Resettlement cut-off-date baselines](/solutions/resettlement-cut-off-baselines)
:::

:::callout{tone="note" title="Tribal land, state land, freehold"}
Tribal land is the most common form of tenure in Botswana, followed by state land (ESMF), and customary land rights, once granted, are held in perpetuity ([gov.bw](https://www.gov.bw/land-management/application-customary-law-land-grant)). Whether a structure in a servitude is allowed is for the Land Board, the Land Tribunal and the courts; our register gives them a dated, located list to work from.
:::
::::
:::::
::::

::::section{id="vegetation-fire" eyebrow="Vegetation and fire" title="Where the servitude has regrown, and where it has burnt"}
:::cards{cols="3"}
:::card{title="Cleared and regrown bush" icon="leaf"}
Maps of where vegetation in the servitude has been cleared or has grown back between dates, from Copernicus Sentinel data, so clearing crews start with the stretches that have changed most. They show larger patches of change, not individual trees.
:::
:::card{title="Tall vegetation in the servitude" icon="tree"}
Where tall vegetation stands, from drone elevation models where flown and open canopy-height data for the wider line. The open data is older in places, so a drone survey gives the current picture. It is not a measured clearance to the conductors.
:::
:::card{title="Fire near the line" icon="alert"}
Satellite-detected fire hotspots from NASA FIRMS, filtered to the servitude buffers you set and checked by a reviewer, followed by a burnt-area map. Not an emergency or early-warning service: small, short-lived or cloud-covered fires can be missed.
:::
:::

[Vegetation, land cover and fire, in detail](/solutions/vegetation-land-cover-fire)
::::

::::section{id="drones" tone="alt" eyebrow="Why we start from satellite" title="Drones near power lines in Botswana"}
:::::columns{split="2-1"}
::::col
The Civil Aviation (Remotely Piloted Aircraft) Regulations, 2024 keep a drone 50 m from any structure not under the control of the person in charge of it, unless the Civil Aviation Authority of Botswana (CAAB) approves other distances (S.I. No. 71 of 2024, reg. 24). CAAB's 2016 bye-law on remotely operated aircraft, still published on CAAB's website, also lists "Within a lateral distance of 200m from any Power Line" among its restricted areas, together with major public roads and built-up areas. Whether that list still applies after 2024 is a question to put to CAAB in writing.

That is why a line survey starts from satellite imagery, which needs no drone permit and covers the whole route. Drone checks of flagged spans are planned with the line's owner and CAAB, and are subject to the approvals and security clearances each job requires. Where your own teams already fly the line, we run the same analysis on your orthophotos.
::::
::::col
:::callout{tone="legal" title="Drone law in Botswana"}
The operator certificate, security vetting, the authorisation each operation needs, the flight limits along a line and the older bye-law restrictions, with a checklist for clients, last reviewed 27 September 2026.

[Read the guide](/bw/drone-regulations)
:::
::::
:::::
::::

::::section{id="how" eyebrow="How it works" title="From line route to reviewed register"}
:::steps
:::step{title="Scope"}
You send the line route, pole or tower positions if you have them, the servitude widths and what the register is for: maintenance, a development request, a new route or a cut-off-date baseline. We agree the widths, imagery and deliverables in a written proposal.
:::
:::step{title="Screen from satellite"}
Structures, fields and cleared ground are mapped along the whole line from dated satellite imagery, open building datasets or the orthophotos you hold.
:::
:::step{title="Review"}
A person checks every result, marks what the models missed and lists what the imagery cannot settle for a ground check.
:::
:::step{title="Deliver and repeat"}
The report and GIS layers go to the contacts you name, and re-surveys follow the schedule you agree with us.
:::
:::

More on [how it works](/features), the [methodology](/methodology), [imagery and data sources](/imagery) and our global [power and water utilities page](/industries/power-utilities).
::::

::::section{id="scope" tone="alt" eyebrow="Honest scope" title="What we do, and what we don't"}
:::::columns{split="1-1"}
::::col
### What we do

:::checklist
- Map structures, ploughed and cleared ground, fresh digging and tracks visible from the air
- Measure each structure to the line, at the widths you give us
- Flag change between dated surveys, confirmed by a reviewer
- Deliver dated, credited records that support wayleave, compensation and community processes
:::
::::
::::col
### What we don't do

:::checklist{tone="no"}
- Measure conductor clearances or inspect poles, towers and hardware
- Detect theft or vandalism, or identify who did anything
- Identify, count or follow people
- Replace a cadastral survey, or decide whether a structure or field is lawful
:::
::::
:::::
::::

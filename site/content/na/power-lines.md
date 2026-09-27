---
key: power-utilities
template: industry
slug: power-lines
title: Power-Line Servitude & Wayleave Mapping, Namibia | AfriScan
description: A dated register of structures, cleared ground and bush inside power-line servitudes in Namibia, measured to your centre line and wayleave edge, and reviewed.
h1: Structures and bush inside your power-line servitudes in Namibia, mapped and dated
crumb: Power lines
eyebrow: Power & utilities · Namibia
lead: In Namibia the width of a power-line servitude, and what may stand inside it, are set in each wayleave agreement, and the Electricity Safety Code sets how close live conductors may come to any structure. We measure each structure's distance to your centre line and to the servitude edge you give us, rate every 500 m of the line, and show which structures and cleared areas are new since the last survey. A person reviews every result.
buttons:
  - {label: Send us your line route, intent: proposal}
  - {label: What the wayleaves say, href: "#wayleaves"}
service:
  name: Power-line servitude and wayleave mapping, Namibia
  type: Servitude encroachment survey, change detection, vegetation change and route baselines
  description: Structures, cleared ground and vegetation change inside the servitudes of transmission and distribution lines in Namibia, each structure measured to the centre line and the servitude edge in the wayleave agreement, grouped into 500 m stretches, compared between dated surveys and reviewed by a person, with route-option counts and cut-off-date baselines for new lines, delivered as PDF reports and GIS layers.
og:
  headline: Power-line servitudes in Namibia, mapped and dated
  subline: Structures measured to your centre line and wayleave edge, reviewed by a person
related: [route-site-selection, vegetation-land-cover-fire, change-detection]
faq:
  - q: How wide is a power-line servitude in Namibia?
    a: There is no single national width. Widths are agreed in each wayleave. For one new 400 kV line, the published resettlement framework describes an 80 m servitude, 40 m either side of the line, and 25 m either side in densely populated areas, with a 12 m strip cleared of vegetation. Send the widths in your agreements, per line or per section, or the servitude polygons from your GIS, and we measure against those.
  - q: Does the register check the Electricity Safety Code clearances?
    a: No. Table 1 of the Code sets the minimum distance between live conductors and a structure by voltage; whether a structure meets it depends on the conductor's height and sag at that point, which we don't measure. We map where each structure stands in plan and how far it is from the centre line and the servitude edge. Checking clearance to the conductors stays with your line engineers, and the register tells them where to look first.
  - q: Our line runs over farms on a wayleave agreement, not a registered servitude. Can you still map it?
    a: Yes. The survey measures against the widths you give us, whether they come from a wayleave agreement, a registered servitude or your own standard. Up to six widths can be reported in one survey, so a stretch with 40 m either side and another with 25 m either side sit in the same register.
  - q: Do you detect copper theft or vandalism?
    a: No. Satellites cannot see theft, vandalism or people. We map land change along your lines, such as new structures, fresh digging, cleared ground and new tracks, so your field teams know where to look first.
  - q: Can the register be used in wayleave and compensation negotiations?
    a: It gives your wayleave officers and the landowner one dated map of what stands on the strip, which supports land, compensation and community processes. It does not decide what is owed or whether a structure is lawful, and it is not a cadastral survey. General plans and diagrams need approval and a land surveyor's signature under the Land Survey Act 1993.
  - q: Can a drone fly our line?
    a: Subject to the approvals and security clearances each job requires. A commercial flight needs an NCAA Letter of Approval, flights beyond visual line of sight need an RPAS Operator Certificate that only Namibian persons and companies can hold, and a drone near structures must stay under the operator's direct control. That is why a line survey in Namibia starts from satellite imagery. See [drone law in Namibia](/na/drone-regulations).
cta:
  title: Send us your line route
  text: "KML, KMZ, GeoJSON, Shapefile, GPX or GeoPackage, with tower positions if you have them, the voltage, the regions the line crosses and the servitude widths in your wayleave agreements. Tell us if your team already flies the line. We reply with a scope, an imagery plan and a written proposal."
  button: Request a proposal
  secondary: Drone law in Namibia
  secondary_href: /na/drone-regulations
---

::::section{id="problem" eyebrow="The problem, in your words" title="Long lines, farm by farm, and a strip that should stay clear" lead="A Namibian transmission line crosses commercial farms, government-leased farms, communal land and the edges of towns. The wayleave framework published for one new 400 kV line allows grazing and cropping inside the servitude but no permanent structures other than the towers, and the only way to know what stands there is to look."}
:::cards{cols="2"}
:::card{title="“Towers only, inside the servitude.”" icon="power"}
Under that framework the landowner keeps farming the strip, but nothing permanent may be built in it except the towers. A house, a kraal wall or a store built inside it is a safety question first and a cost later.
:::
:::card{title="“We don't know who is on that farm now.”" icon="houses"}
Land-risk reports on Namibian transmission projects describe government-leased farms whose occupants change often and are not always formally registered. A dated map of what stands on each farm gives the wayleave team a starting point before the first visit.
:::
:::card{title="“The bush has grown back.”" icon="leaf"}
Bush regrows along lines and service roads, and clearing crews need to know which spans to go to first. Dated imagery shows where vegetation has grown or been cleared, span by span.
:::
:::card{title="“Something has been dug near a tower.”" icon="excavation"}
Fresh excavation, spoil and new tracks near tower bases are visible from above. We flag them between surveys so your line teams know which towers to check. What happened there is for your field teams to establish.
:::
:::

### Who this is for

:::chips
- Wayleave and land officers
- Line maintenance and bush-clearing teams
- Transmission project and route teams
- Distribution network teams
- Environmental and social teams on ECC and lender work
- GIS and asset-records teams
:::
::::

::::section{id="wayleaves" tone="alt" eyebrow="What the Code and your wayleaves say" title="The widths come from your agreements, the clearances from the Code" lead="A summary of the public rules as last reviewed on 27 September 2026, for orientation. It is not legal advice; the sources are linked."}
:::::columns{split="1-1" align="center"}
::::col
:::figure{src="diagrams/na-wayleave" alt="Plan-view schematic of a 400 kV line with towers along it: an 80 m red servitude band, 40 m either side of the line, narrowing to a 50 m band through a settled stretch, with a 12 m grey strip along the line; square markers stand for structures, red inside the servitude and teal outside" caption="An example servitude, widths drawn to scale" credit="Schematic drawn by AfriScan for illustration; the structures are invented." size="half"}
Red: an 80 m servitude, 40 m either side of the line, narrowing to 25 m either side through a densely populated stretch; grey: the 12 m strip cleared of vegetation for a service road. The widths are the ones one published framework for a new 400 kV line describes. Your register uses the widths in your own agreements.
:::
::::
::::col
**Wayleaves, not a national width.** Namibian power-line servitude widths are set in wayleave agreements, not by statute. For a new 400 kV line, the utility's published resettlement framework (November 2023) says "The servitude will be 80 m wide for the entire line", with about 12 m "totally cleared of vegetation and obstacles to create a service road", and 25 m either side of the line in densely populated areas. Where the new line runs beside an existing 220 kV line, the combined strip is 111 m ([NamPower, Land Risks and Impacts, 2023](https://www.nampower.com.na/Media/Document/2029f04a-3d58-4a66-8302-31ea9d6ef0c7.pdf)).

**What may stand inside.** The same framework says: "For safety and technical reasons, no permanent structures other than the towers are allowed within the servitude. Grazing and cultivation of fields with associated farming activities may be accommodated within this area, except for the 12 m strip." The wayleave "does not expropriate the property from the landowner, but provides NamPower with the right to build, operate and maintain a transmission line", and the framework notes that "A formal servitude is not registered over the farm, there is only an agreement".
::::
:::::

### Clearances to structures in the Electricity Safety Code

The [Namibian Electricity Safety Code](https://www.ecb.org.na/wp-content/uploads/2022/07/Namibia-Electricity-Safety-Code.pdf) (GN 200 of 2011, in operation since 31 October 2012) sets, in Table 1, "the minimum distance between live conductors and such structures" for structures that are not part of the power line:

| Line voltage | Minimum clearance |
|---|---|
| Up to 33 kV | 3.0 m |
| 66 kV | 3.2 m |
| 88 kV | 3.4 m |
| 132 kV | 3.8 m |
| 220 kV | 4.2 m |
| 275 kV | 4.7 m |
| 330 kV | 5.3 m |
| 400 kV | 5.6 m |

These are clearances measured from the conductor, which our registers do not measure (see the [scope](#scope) below). What the register gives your engineers is the list of structures standing close enough to the line to need that check.

**Entry and obstacles.** Where a licensee's poles, cables or foundations stand on someone's land, the [Electricity Act 4 of 2007](https://www.lac.org.na/laws/annoSTAT/Electricity%20Act%204%20of%202007.pdf) lets it enter, after reasonable notice except in an emergency, to inspect, maintain, remove, replace or renew them, and it "may require from the owner or occupier of the premises to remove any tree, shrub or growth or any fence or other obstacle" (s.38). We found no statutory vegetation-clearance width for line corridors: the widths come from your wayleaves and your own standards.
::::

::::section{id="servitudes" eyebrow="Existing lines" title="A register of what stands inside each servitude, then what changed"}
:::::columns{split="2-1"}
::::col
We buffer the centre line you supply in its UTM zone and report each structure the review confirms with its distance to the centre line and to the servitude edge, the band it falls in, its chainage along the line and its coordinates. Bands follow your agreements: 40 m either side on one stretch and 25 m on another, plus outer bands of your choosing so you can see where pressure is building beside the strip.

Each 500 m stretch is rated for **encroachment density**: high where more than five structures stand inside the widest band, medium for one to five, low for none. It is a count rule that tells your wayleave officers which spans and farms to visit first, not a safety rating.

On a repeat survey, new and removed structures are flagged automatically and confirmed by a reviewer, with before-and-after views, and your team receives a notice after each survey. Fresh digging, cleared ground and new tracks inside the servitude are flagged the same way.
::::
::::col
### What you receive

:::checklist
- A structure register per line: ID, distance to the centre line and to the servitude edge, band, chainage and coordinates
- An encroachment-density rating for each 500 m stretch
- New and removed structures between surveys, with before-and-after views
- Fresh digging and cleared ground near towers, flagged for checking
- A PDF report with GeoPackage, GeoJSON, KMZ and Shapefile layers
- A change notice after each survey
:::
::::
:::::

:::figure{src="diagrams/corridor" alt="Schematic of a line route with a 50 m band and a 100 m band, structures coloured by distance to the line, and a strip below rating each 500 m stretch low, medium, high, medium and low from the structures inside the widest band" caption="How a servitude register is read" size="wide" credit="Schematic drawn by AfriScan for illustration. It is not a real line, and the widths shown are examples: your register uses the widths in your wayleaves."}
:::

:::solutions{keys="right-of-way-monitoring,change-detection,evidence-packs" cols="3"}
:::
::::

::::section{id="new-lines" tone="alt" eyebrow="New lines and interconnectors" title="Count what each route option would cross, before the first farm visit"}
:::::columns{split="1-1"}
::::col
New 400 kV lines and a planned interconnector with Angola each start with a route, and a wayleave for every farm and plot the route crosses. For each alternative alignment we count the structures within the same bands, so the route team can weigh the options on the land they would affect. Slope, erosion-prone ground, drainage crossings and ground that has flooded in past satellite records are screened along each option.

Once the route is fixed, a dated baseline records the structures in the servitude with their IDs, coordinates and imagery. It gives the wayleave team, the landowners and the environmental consultant one shared map, and on lender-financed lines it supports the cut-off-date records the lender's standard asks for. In Namibia "resettlement" usually means the land-reform programme; here we mean [project land acquisition](key:resettlement-cut-off-baselines).
::::
::::col
:::callout{tone="legal" title="When agreement fails, and before construction"}
- **Expropriation.** A licensee may expropriate land or a right over it only with Cabinet's approval, after the Electricity Control Board has held a public hearing with at least 14 days' written notice, and only if the licensee "has been unable to acquire the land or right concerned on reasonable terms, other than terms relating to compensation, by agreement with the owner". Compensation not agreed is set under the Expropriation Ordinance 13 of 1978 ([Electricity Act, s.35](https://www.lac.org.na/laws/annoSTAT/Electricity%20Act%204%20of%202007.pdf)).
- **Communal land.** A land board may grant occupational land rights for projects of a State-owned enterprise ([Communal Land Reform Act, s.36A](https://www.lac.org.na/laws/annoSTAT/Communal%20Land%20Reform%20Act%205%20of%202002.pdf)).
- **Environmental clearance.** "The transmission and supply of electricity" is a listed activity that needs an Environmental Clearance Certificate ([GN 29 of 2012](https://www.lac.org.na/laws/2012/4878.pdf)).
:::

:::solutions{keys="route-site-selection,resettlement-cut-off-baselines" cols="1"}
:::
::::
:::::
::::

::::section{id="vegetation" eyebrow="Vegetation and bush" title="Where the servitude has regrown, span by span"}
:::::columns{split="1-1"}
::::col
Cleared and regrown vegetation along the servitude and the service road is compared between dates from Copernicus Sentinel imagery, and tall vegetation inside the servitude is mapped from drone elevation models where flown, with open canopy-height data for wider context. The results are grouped by span or 500 m stretch, so bush-clearing crews can be sent where the growth is.

In the dry season we can add notices of satellite-detected fire hotspots near your lines and substations, drawn from NASA FIRMS and checked by our team, followed by maps of the burnt area.
::::
::::col
:::callout{tone="note" title="The limits"}
Vegetation mapping shows larger patches of change, not individual trees, and it is not a measured clearance to conductors. Fire notices are not an emergency or early-warning service: small, short-lived or cloud-covered fires can be missed. Copernicus and NASA FIRMS data are credited as their licences require.
:::

:::solutions{keys="vegetation-land-cover-fire" cols="1"}
:::
::::
:::::
::::

::::section{id="your-imagery" tone="alt" eyebrow="Your drone programme" title="Already flying your lines? Send us the orthophotos"}
:::::columns{split="2-1"}
::::col
Many utilities fly drones for inspection and substation mapping, and the flying is rarely the bottleneck. Turning each campaign into a list of what is new, where, and how close it is to the line is.

Send the orthophotos or GeoTIFFs you already hold. We run the same structure, change and corridor analysis on them, add the servitude bands and the 500 m ratings, and a person reviews every result. Between your flights, dated satellite scenes, where they exist for the line, show what has changed since the last one.
::::
::::col
:::callout{tone="note" title="Analysis of your own imagery"}
Results depend on your imagery's resolution and quality, so we check a sample before we commit. The analysis runs on imagery you already hold and adds no flight of our own.

[Imagery: yours, archive or new capture](/solutions/imagery)
:::
::::
:::::
::::

::::section{id="drones" eyebrow="Drones along the line" title="Why a line survey in Namibia starts from satellite"}
:::::columns{split="1-1"}
::::col
Under NAMCAR Part 101, a drone may not fly beyond the pilot's direct unaided sight, or more than 300 m from the point of operation, without a specific approval, and flights beyond sight need an RPAS Operator Certificate that the NCAA issues only to Namibian persons and companies. Near any person, vehicle or structure the drone must be under the operator's direct control, and Part 101 bars using a drone to keep someone's property under watch without the owner's consent.

So we screen the whole line from satellite first and propose drone checks only for the spans that need them, subject to the approvals and security clearances each job requires. Afridrone is working towards the approvals Namibia requires; every drone proposal names the company that will fly and its approvals.
::::
::::col
:::callout{tone="legal" title="The Namibia drone-law guide"}
The Letter of Approval, the Operator Certificate, remote pilot certificates, height and airfield limits, restricted areas and a 12-point checklist for anyone commissioning a drone survey, last reviewed on 27 September 2026.

[Drone law in Namibia](/na/drone-regulations)
:::
::::
:::::
::::

::::section{id="how" tone="alt" eyebrow="How it works" title="From line route to reviewed register"}
:::steps
:::step{title="Scope"}
You send the line route or tower coordinates, the voltage, the regions it crosses and the widths in your wayleaves. We agree the bands, the imagery and the deliverables in a written proposal.
:::
:::step{title="Screen from satellite"}
Structures, cleared ground, tracks and vegetation are mapped along the whole line from dated satellite imagery, open building datasets or your own orthophotos.
:::
:::step{title="Check up close"}
Your own flights, or drone surveys where approvals and landowners allow, cover the spans that need more detail.
:::
:::step{title="Review and deliver"}
A person checks every result; each structure is measured, banded and rated, and the report and GIS layers go to the contacts you name.
:::
:::

More on [how it works](/features), the [methodology](/methodology) and [power and utilities work in other countries](/industries/power-utilities).
::::

::::section{id="scope" eyebrow="Honest scope" title="What we do, and what we don't"}
:::::columns{split="1-1"}
::::col
### What we do

:::checklist
- Map structures, cleared ground, fresh digging, tracks and vegetation change visible from the air
- Measure each structure to the centre line and the servitude edge in your wayleave
- Flag change between dated surveys, confirmed by a reviewer
- Count the structures along alternative routes for new lines
:::
::::
::::col
### What we don't do

:::checklist{tone="no"}
- Inspect conductors, insulators or towers, or measure clearance to conductors
- Detect theft or vandalism, or say who was responsible
- Identify, count or follow people or vehicles
- Decide whether a structure is lawful, or what compensation is owed
:::
::::
:::::
::::

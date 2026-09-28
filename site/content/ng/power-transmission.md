---
key: power-utilities
template: industry
title: Transmission Line Right-of-Way Mapping, Nigeria | AfriScan
description: Structures inside 330 kV, 132 kV and 33 kV rights of way under NESIS, measured to the line, plus household estimates for mini-grid sites, reviewed by a person.
h1: Secure your transmission rights of way in Nigeria from the air
crumb: Power lines
eyebrow: Power & utilities · Nigeria
lead: "NERC's standard is plain: no structures under the right of way of an overhead line. Structures still go up. We measure each one against the NESIS width for its voltage, rate every 500 m of the line, and flag disturbed ground around towers and burning in the right of way. For mini-grid programmes, the same imagery gives structure layers and estimated households."
buttons:
  - {label: Send us your line route, intent: proposal}
  - {label: Mini-grid site estimates, href: "#energy-access"}
service:
  name: Transmission and distribution right-of-way surveys, Nigeria
  type: Right-of-way encroachment survey, change detection and household estimates
  description: Structures inside the NESIS rights of way of 330 kV, 132 kV, 33 kV and 11 kV lines in Nigeria, each measured to the centre line, with 500 m encroachment-density ratings, disturbed ground near towers, vegetation and fire in the right of way, and estimated households for mini-grid candidate communities, reviewed by a person and delivered as a PDF report with GIS layers.
og:
  headline: Transmission rights of way in Nigeria, mapped from the air
  subline: Structures inside NESIS widths, change between surveys and mini-grid household estimates
related: [vegetation-land-cover-fire, drone-surveys, terrain-flood-post-event]
faq:
  - q: Which widths do you use for a Nigerian line?
    a: By default the NESIS right of way for the voltage, measured from the centre line you supply and split equally either side (25 m each side of a 330 kV line, 15 m of a 132 kV line, 5.5 m of a 33 kV or 11 kV line), plus outer bands of 50 m and 100 m so you can see where pressure is building. You can add your own standard or a lender's width, up to six widths in one survey.
  - q: Do you detect vandalism or cable theft?
    a: No. Imagery cannot show who damaged a tower or took steel from it. We flag fresh excavation, disturbed ground, new tracks and burning visible at the surface around tower bases and along the right of way, so your line teams know where to check.
  - q: Can a drone survey our 330 kV line?
    a: Only with a special NCAA authorisation. Part 21 lists "areas of high RF transmission/interference (e.g. radar sites, high tension wires)" among the operations that need one, requested at least 30 days ahead, on top of the operator's own certificate and ONSA clearance. That is why transmission work in Nigeria starts from satellite imagery.
  - q: Can you measure clearance to the conductors?
    a: No. We show where structures and tall vegetation stand in the right of way, using drone elevation models where flown and open canopy-height data for wider context. It is not a measured clearance to conductors, and open canopy-height data is older in places.
  - q: Can your household estimates be used in a DARES or other grant application?
    a: They can inform your site selection and sizing, but we make no claim that they meet any programme's data requirements. Households are estimated from reviewed structure counts using persons-per-household assumptions that the report states, so your team and the programme can judge them. It is not a census.
  - q: We already fly drones along our lines. Can you use that imagery?
    a: Yes. We run the same structure, change and vegetation analysis on your own orthophotos and elevation models, and use dated satellite scenes to keep the record moving between your flights. Results depend on the imagery's quality and resolution.
  - q: Do you work for distribution companies as well as TCN?
    a: Yes. The NESIS widths cover 33 kV and 11 kV lines as well as the transmission grid, so the same register works for distribution companies, independent power producers building connection lines and state electrification projects.
cta:
  title: Send us your line route
  text: "KML, KMZ, GeoJSON, Shapefile, GPX or GeoPackage, or tower coordinates we can draw the line from. Tell us the voltage, the states the line crosses and what the record is for. For mini-grid work, send the candidate sites or communities. We reply with a scope, an imagery plan and a written proposal."
  button: Request a proposal
  secondary: Drone law in Nigeria
  secondary_href: /ng/drone-regulations
---

::::section{id="problem" eyebrow="The problem, in your words" title="Under the line, and around the towers" lead="A transmission line crosses hundreds of kilometres of farmland, bush and growing towns. The right of way beside it changes between line patrols, and every change is a safety question first and a cost later."}
:::cards{cols="2"}
:::card{title="“People are building under the line.”" icon="houses"}
Homes, shops and extensions go up inside the right of way, most often where a line passes the edge of a town or runs beside a road. NESIS says no structures shall be built there; knowing where they stand is the first step to dealing with them.
:::
:::card{title="“There has been digging around a tower.”" icon="excavation"}
Excavation and disturbed ground near tower bases are visible from above before a line patrol reaches them. We flag them between surveys so line teams know which towers to check first.
:::
:::card{title="“The right of way is overgrown, and it burns.”" icon="leaf"}
Regrowth and dry-season bush burning in the right of way threaten lines and towers. Dated imagery shows where vegetation has grown back and where fires have burnt.
:::
:::card{title="“We need to know how many households are there.”" icon="layers"}
Mini-grid developers and electrification programmes need a structure count and a household estimate for each candidate community before anyone is sent to survey it.
:::
:::

TCN's own statements show the pressure on the grid. In June 2026 it reported damage to six towers, T125 to T130, on its Apir–Lafia 330 kV lines, and condemned "the continued vandalism of power transmission infrastructure" ([TCN, 2 June 2026](https://www.tcn.org.ng/blog_post_sidebar262.php)). Imagery cannot prevent that or say who was responsible. It can show, reviewed and dated, where the ground around the towers and inside the right of way has changed.

### Who this is for

:::chips
- Right-of-way and wayleave officers
- Line maintenance and vegetation managers
- Transmission project and land-acquisition teams
- Distribution companies and independent power producers
- Mini-grid developers' site-selection teams
- Energy-access programme and M&E teams
- GIS and asset-records teams
:::
::::

::::section{id="row-widths" tone="alt" eyebrow="The widths in the standard" title="NESIS rights of way, drawn to scale" lead="A summary of the public rules as of 26 September 2026, for orientation. It is not legal advice; the sources are linked."}
| Voltage | Right of way (NESIS Table 3.1) | Each side of the centre line |
|---|---|---|
| 330 kV | 50 m | 25 m |
| 132 kV | 30 m | 15 m |
| 33 kV | 11 m | 5.5 m |
| 11 kV | 11 m | 5.5 m |

:::::columns{split="1-1" align="center"}
::::col
:::figure{src="diagrams/ng-nesis-row" alt="Plan-view schematic of three overhead lines drawn to one scale: a 330 kV line with a 50 m red band, a 132 kV line with a 30 m band and a 33 kV line with an 11 m band, each centred on the line with tower symbols along it; square markers stand for structures, red inside the band and teal outside" caption="The NESIS rights of way, drawn to one scale" credit="Schematic drawn by AfriScan for illustration; the structures are illustrative." size="half"}
Red: the right of way for each voltage in NESIS Regulations 2015, Table 3.1, divided equally either side of the centre line. The squares show how a register bands structures: red inside the right of way, teal outside it.
:::
::::
::::col
The [NESIS Regulations 2015](https://nemsa.gov.ng/wp-content/uploads/2019/11/NESIS-Regulations-2015-1-1.pdf) (NERC/Reg/1/2015, s.3.1) say "no structures shall be built under the Overhead line Right of Way", and that where structures are built after the line, the licensee "shall not be liable for any mishap caused by contact with the line". That makes the date a structure appeared worth recording.

NESIS was made under the Electric Power Sector Reform Act 2005, which the [Electricity Act 2023](https://nemsa.gov.ng/wp-content/uploads/2024/05/ELECTRICITY-ACT-2023.pdf) repeals. NEMSA still lists NESIS among its [technical standards](https://nemsa.gov.ng/technical-standard/); whether it continues under the 2023 Act's savings provisions is a point for your counsel.
::::
:::::
::::

::::section{id="servitudes" eyebrow="What we map along the line" title="A register of what stands in the right of way, then what changed"}
:::::columns{split="2-1"}
::::col
We buffer the line route you supply in its UTM zone and report each structure the review confirms with its distance to the centre line: inside the NESIS right of way for the voltage, and in outer bands of 50 m and 100 m. The line is cut into 500 m stretches, each rated for **encroachment density**: high where more than five structures stand inside the widest band, medium for one to five, low for none. It is a count rule that tells your right-of-way officers where to go first, not a safety rating.

Between surveys, new and removed structures are flagged automatically and confirmed by a reviewer, with before-and-after views, and your team receives a notice after each survey. Fresh excavation and disturbed ground around tower bases and new tracks to the line are flagged the same way, so patrols know which towers to check. They do not replace patrols.

A dated record matters for the NESIS liability rule too. Where archive imagery exists from before and after the date a line was built or a structure went up, we compare the dated scenes and say roughly when each structure first appears.
::::
::::col
### What you receive

:::checklist
- A structure register for each line: ID, distance to the centre line, band, chainage and coordinates
- An encroachment-density rating for each 500 m stretch
- New and removed structures between surveys, with before-and-after views
- Disturbed ground and fresh excavation near towers, flagged for checking
- Imagery history for structures whose date matters
- A change notice after each survey
:::
::::
:::::

:::solutions{keys="right-of-way-monitoring,change-detection,imagery-history-due-diligence" cols="3"}
:::
::::

::::section{id="vegetation-fire" tone="alt" eyebrow="Vegetation and bush burning" title="Regrowth and fire in the right of way"}
:::::columns{split="1-1"}
::::col
Cleared and regrown vegetation along the right of way is compared between dates from Copernicus Sentinel imagery, and tall vegetation is mapped from drone elevation models where flown, with open canopy-height data for wider context. In the dry season we send notices of satellite-detected fire hotspots within the buffer around your lines and substations, drawn from NASA FIRMS and checked by our team, followed by maps of the burnt area.
::::
::::col
:::callout{tone="note" title="The limits"}
Vegetation mapping shows larger patches of change, not individual trees, and it is not a measured clearance to conductors. Fire notices are not an emergency or early-warning service: small, short-lived or cloud-covered fires can be missed. NASA FIRMS and Copernicus data are credited as their licences require.
:::
::::
:::::
::::

::::section{id="new-lines" eyebrow="New lines" title="Route options and a baseline before the right of way is acquired"}
:::::columns{split="1-1"}
::::col
For a new transmission or connection line, we count the structures along each alternative alignment within the same bands, so the route team can weigh options on the land they would affect. Slope, drainage crossings and ground that has flooded in past satellite records are screened along each option.

Once the route is fixed, a dated baseline records the right of way with reviewer categories and estimated households, ready for compensation enumeration under the Land Use Act and, on World Bank or other lender-financed lines, for the resettlement standard the lender applies.
::::
::::col
:::solutions{keys="route-site-selection,resettlement-cut-off-baselines" cols="1"}
:::
::::
:::::
::::

::::section{id="energy-access" tone="alt" eyebrow="Mini-grids and energy access" title="Structure counts and estimated households for candidate communities" lead="Choosing and sizing mini-grid sites starts with a simple question: how many buildings are there, and how many households do they suggest? Dated imagery answers it for many communities before a field team visits one."}
:::::columns{split="2-1"}
::::col
The World Bank-financed Distributed Access through Renewable Energy Scale-up (DARES) project is implemented by the Rural Electrification Agency (REA). Approved by the World Bank's Board on 14 December 2023, it supports privately owned and operated solar hybrid mini-grids in unserved and underserved areas, through a minimum subsidy tender and performance-based grants, alongside standalone solar systems ([World Bank, P179687](https://projects.worldbank.org/en/projects-operations/project-detail/P179687)).

For developers and programme teams we provide:

- **Structure layers** across the candidate communities, reviewed by a person, as GIS files your team can load;
- **Estimated households**, derived from structure counts with the persons-per-household assumptions stated in the report;
- **Growth trends** showing where built-up land has expanded year by year;
- **Land cover and history** of the generation site, and the structures inside its boundary.

Open settlement data such as GRID3 Nigeria are a good first screen, and we credit them where used. What a survey adds is a reviewed structure layer on dated imagery, with the assumptions written down.
::::
::::col
:::callout{tone="scope" title="Estimates, stated as estimates"}
Households are always estimated from structure counts, never counted, and the report sets out every assumption. It is not a census. We make no claim that the estimates meet any programme's or grant's data requirements; your team and the programme decide how to use them.

[Send us your candidate sites](/ng/contact?intent=proposal&country=ng)
:::

:::solutions{keys="household-estimates" cols="1"}
:::
::::
:::::
::::

::::section{id="drones" eyebrow="Drones near high-tension lines" title="Why transmission work in Nigeria starts from satellite"}
:::::columns{split="1-1"}
::::col
Part 21 of the Nigeria Civil Aviation Regulations lists "areas of high RF transmission/interference (e.g. radar sites, high tension wires)" among the operations that need a special authorisation, requested at least 30 days ahead (21.9.6.21). That comes on top of the operator's RPAS Operator Certificate, ONSA security clearance and the rules that keep drones at least 30 m from people not involved in the operation and away from populated areas.

So we screen the whole line from satellite first, and propose drone checks only for the spans that need them, subject to the NCAA and ONSA authorisations each job requires. Afridrone is working towards the authorisations Nigeria requires; every drone proposal lists the approvals its flights need and when each must be in place.
::::
::::col
:::callout{tone="legal" title="The Nigeria drone-law guide"}
The approvals, the corridor rules and a 15-point checklist for anyone commissioning a drone survey, with sources, as of 26 September 2026.

[Drone law in Nigeria](/ng/drone-regulations)
:::
::::
:::::
::::

::::section{id="buyers" tone="alt" eyebrow="Who buys in Nigeria" title="TCN, distribution companies and project developers"}
The Transmission Company of Nigeria runs the 330 kV and 132 kV grid; distribution companies run the 33 kV and 11 kV networks; NERC regulates the sector and NEMSA enforces technical standards. Federal bodies such as TCN procure under the Public Procurement Act 2007 through the Bureau of Public Procurement's contractor registration, and publish tenders on [TCN's notices page](https://www.tcn.org.ng/page_notices.php). Independent power producers, mini-grid developers and state electrification projects buy directly. Our [procurement notes](/ng/nigerian-content) set out the registrations that apply.
::::

::::section{id="how" eyebrow="How it works" title="From line route to reviewed register"}
:::steps
:::step{title="Scope"}
You send the line route or tower coordinates, the voltage, the states it crosses and what the record is for: encroachment, a new line, a baseline or mini-grid sites. We agree the widths, the imagery and the deliverables in a written proposal.
:::
:::step{title="Screen from satellite"}
Structures, ground change and vegetation are mapped along the whole line from dated satellite imagery, open building datasets or the imagery you hold.
:::
:::step{title="Check up close"}
Drone surveys capture the spans that need detail, subject to the NCAA special authorisation and the other approvals each job requires.
:::
:::step{title="Review and deliver"}
A person checks every result; each structure is measured, banded and rated, and the report goes to the contacts you name with the GIS layers.
:::
:::

More on [how it works](/features), the [methodology](/methodology) and [imagery and data sources](/imagery).
::::

::::section{id="scope" tone="alt" eyebrow="Honest scope" title="What we do, and what we don't"}
:::::columns{split="1-1"}
::::col
### What we do

:::checklist
- Map structures, cleared ground, fresh excavations, tracks and vegetation change visible from the air
- Measure each structure to the centre line against the NESIS width for the voltage
- Flag change between dated surveys, confirmed by a reviewer
- Estimate households from structure counts, with the assumptions stated
:::
::::
::::col
### What we don't do

:::checklist{tone="no"}
- Detect vandalism, theft or tampering, or say who was responsible
- Measure clearance to conductors
- Identify, count or follow people or vehicles
- Decide whether a structure is authorised, or count households in a census
:::
::::
:::::
::::

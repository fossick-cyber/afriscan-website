---
key: power-utilities
template: industry
title: Power-Line Servitude Mapping in Mozambique | AfriScan
description: Structures inside transmission and distribution line servitudes in Mozambique, dated and measured to the axis, for existing lines, new routes and registration.
h1: Structures in power-line servitudes in Mozambique, mapped from the air
crumb: Power lines
eyebrow: Power & utilities · Mozambique
lead: A dated register of the structures inside the servitude of your transmission and distribution lines, each measured to the line's axis, for maintenance, regularisation, new routes and the servitude you are about to register. Reviewed by a person and reported in Portuguese or English.
buttons:
  - {label: Send us your line route, intent: proposal}
  - {label: See the sample register, key: results}
service:
  name: Power-line servitude encroachment survey, Mozambique
  type: Servitude encroachment survey, route comparison and scheduled re-surveys
  description: Structures inside the administrative servitudes of transmission and distribution lines in Mozambique, each measured to the line's axis with the imagery date, for existing lines, new routes and servitude registration, reviewed by a person and reported in Portuguese or English with GIS layers.
og:
  headline: Power-line servitudes in Mozambique, mapped from the air
  subline: Structures by distance to the axis, dated, for existing and new lines
related: [route-site-selection, vegetation-land-cover-fire, terrain-flood-post-event]
faq:
  - q: How wide is the servitude you measure?
    a: The Electricity Law sets an administrative servitude of up to 50 m from the line's axis, with the width depending on the voltage and on whether the setting is rural or urban (Lei n.º 12/2022, art. 43(4)–(5)). We measure to the width recorded in your concession or servitude, report a band outside it where pressure builds, and can add the safety zone inside the servitude if you give us its width.
  - q: Why does the date of each structure matter?
    a: Because the Electricity Law owes no compensation to people who acquired their rights after the electrical infrastructure was built (Lei n.º 12/2022, art. 43(10)). A register tied to dated imagery, and a comparison with imagery from before construction where it exists, shows what stood in the servitude and when. Whether a particular right predates the line is for your land team and the authorities to decide.
  - q: Can you tell which structures are on the servitude of a line not yet built?
    a: Yes. For a planned line we count the structures along each route option within the same bands, so the choice can weigh the land affected. Once the route is fixed, a dated register of the corridor is the baseline for the servitude registration and for the resettlement plan.
  - q: Do you detect vandalism or energy theft?
    a: No. We map structures, cleared ground, vegetation change, fire and tracks visible from the air. Vandalism, theft and illegal connections are not things imagery can show, and we do not claim to detect them.
  - q: Can a drone fly along a live line?
    a: Only within the rules and the authorisations for that flight. IACM's directive keeps drones at least 5 m laterally and 20 ft vertically from live high-tension wires, needs the structure owner's written permission within 50 m of structures, and allows no flights beyond the pilot's sight. Drone surveys are subject to the approvals and security clearances each job requires.
  - q: Can the report go to the district and the land services in Portuguese?
    a: Yes. The PDF report is available in Portuguese or English, so the servitude team, the district government and the Serviços de Cadastro read the same register.
cta:
  title: Send us your line route
  text: "KML, KMZ, GeoJSON, Shapefile, GPX or GeoPackage, or the tower coordinates. Tell us the voltage, the servitude width and what the record is for, and we reply with a scope, an imagery plan and a written proposal."
  button: Request a proposal
  secondary: See the sample register
  secondary_href: /results
---

::::section{id="problem" eyebrow="The problem, in your words" title="The servitude fills up after the line is built" lead="Land along a transmission line looks clear on the day the line is energised. Over the following seasons, houses, machambas and stalls move in beside it, and each one becomes a clearance question for maintenance and a cost for the next project."}
:::::columns{split="1-1"}
::::col
The chair of EDM, the national electricity utility, has described the *ocupação desenfreada das zonas de protecção parcial*, the unchecked occupation of the protection zones along its lines, as a serious constraint that drives up compensation and resettlement costs ([AIM, 29 September 2023](https://aimnews.org/2023/09/29/construcao-de-infra-estruturas-junto-as-linhas-de-transporte-de-energia-preocupa-edm/)). That cost falls on the next line, the next upgrade and the next time a crew needs access to a tower.

A register of what stands in the servitude, measured to the axis and dated, lets the servitude team work from facts: which spans are clear, which are filling up, and which structures were there before the line.
::::
::::col
:::cards{cols="1"}
:::card{title="“Houses are going up under the line.”" icon="houses"}
Structures inside the servitude, listed by distance to the axis and by chainage, with the stretches that are filling fastest shown first.
:::
:::card{title="“Who was there before we built?”" icon="calendar"}
A dated record, and a comparison with earlier imagery where it exists, for the question the Electricity Law makes decisive.
:::
:::card{title="“Which route affects fewest homes?”" icon="route"}
Structure counts for each option of a new line, within the same bands, before the route is fixed.
:::
:::
::::
:::::

### Who this is for

:::chips
- Servitude and land teams at the utility and transmission companies
- Line maintenance and asset managers
- Project teams planning new lines and interconnectors
- Resettlement and social-safeguards teams on lender-funded lines
- ESIA and RAP consultancies
- GIS teams who keep the line records
:::
::::

::::section{id="servitudes" tone="alt" eyebrow="What the law sets" title="The electricity servitude, article by article" lead="The Electricity Law in force since 2022 sets out how a line's servitude is created, registered and compensated. A summary as of 26 September 2026, for orientation; it is not legal advice."}
| Lei n.º 12/2022, de 11 de Julho | What it says |
|---|---|
| [Art. 43(4)](https://arene.org.mz/wp-content/uploads/2022/08/Lei-de-Electricidade-2022.pdf) | An administrative servitude of up to 50 m, counted from the line's axis, recorded on the concession |
| Art. 43(5) | The width depends on the voltage and on whether the setting is rural or urban |
| Art. 43(6) | A safety zone lies inside the servitude |
| Art. 43(7) | The servitude is registered in the Cadastro de Terras and the Conservatória do Registo Predial |
| Art. 43(8) | The resettlement and compensation rules apply |
| Art. 43(10) | No compensation is owed where the holders or owners acquired their rights after the electrical infrastructure was built |
| Art. 44 | Expropriation, with fair compensation |

The Land Law adds its own strip: electricity and telecommunications lines, with 50 m on each side, are a partial protection zone where no DUAT can be acquired, only special licences ([Lei n.º 19/97, arts. 8(g) and 9](https://www.pdul.gov.mz/content/download/486/2635/file/Lei%20de%20Terras.pdf)). Electricity is regulated by ARENE ([Lei n.º 11/2017](https://arene.org.mz/wp-content/uploads/2021/05/Lei-que-cria-a-arene.pdf)).
::::

::::section{id="existing-lines" eyebrow="Existing lines" title="An occupancy baseline, then the spans that change"}
:::::columns{split="2-1"}
::::col
We buffer the line route in its UTM zone, list each structure inside the servitude with its distance to the axis, its band and its position along the line, and rate each 500 m for **encroachment density**: high where more than five structures stand inside the widest band, medium for one to five, low for none. It is a count rule that tells the servitude team where to go first, not a clearance or safety calculation.

Re-surveys follow on a schedule agreed with you. New and removed structures between dated surveys are flagged automatically and confirmed by a reviewer, and a notice after each survey says which stretches changed. Along the line we also map where vegetation has been cleared or has regrown, where tall vegetation stands inside the servitude, and where *queimadas* have burnt close to towers in the dry season.
::::
::::col
### What you receive

:::checklist
- A structure register by distance to the axis, band and position along the line
- A density rating for each 500 m
- New and removed structures between surveys, confirmed by a reviewer
- A change notice after each survey
- Vegetation change and tall vegetation inside the servitude
- Fire-hotspot notices and burnt-area maps near the line
:::
::::
:::::

:::solutions{keys="right-of-way-monitoring,change-detection,vegetation-land-cover-fire" cols="3"}
:::
::::

::::section{id="new-lines" tone="alt" eyebrow="New lines and interconnectors" title="Weigh the homes on each route before it is fixed"}
:::::columns{split="1-1"}
::::col
Mozambique is building a large part of its high-voltage grid now, with lender-funded projects that come with resettlement plans. A new 563 km, 400 kV line in the south is part-financed by the World Bank ([project P160427](https://projects.worldbank.org/en/projects-operations/project-detail/P160427)). A 400 kV line to the northern coast combines about 633 km of 400 kV double circuit with about 183.5 km at 220 kV, and its resettlement plan covered about 1,600 dwellings. A planned hydropower project includes a high-voltage line of about 1,300 km to Maputo.

On lines of that length, the structures a route avoids are the cheapest ones to deal with. We count the structures along each option within the same bands, screen each for steep, erosion-prone and flood-exposed stretches, and estimate households for consultation planning.
::::
::::col
### Before the servitude is registered {#registration}

The Electricity Law asks for the servitude to be registered in the land cadastre and the property register (art. 43(7)). A dated register of the corridor at that point, with each structure's coordinates, distance to the axis and imagery date, is the record your land team and the resettlement plan start from, and the baseline every later survey is compared with.

:::solutions{keys="route-site-selection,resettlement-cut-off-baselines" cols="1"}
:::
::::
:::::
::::

::::section{id="sample" eyebrow="What a register looks like" title="The same register, on a line"}
:::::columns{split="1-1" align="center"}
::::col
:::figure{src="samples/sample-pipeline-register-medium" alt="Strip view of a 500 m stretch of a gas pipeline route rated medium: two reviewer marks inside the 50 m band on the north side of the line" caption="Sample register detail: two structures inside the 50 m band" badge="Reviewed · manual marks" size="half" credit="A high-pressure gas pipeline in Mozambique, shown with the route owner's permission. Drawn by AfriScan from the sample register; no imagery."}
:::
::::
::::col
Our public sample is a gas pipeline route, but the register on a power line is the same: each structure with its distance to the line, its band and its position, and each 500 m rated. Two structures inside the band are enough to rate a stretch medium even where the land around is empty, and the register then shows how close each one stands to the line.

For a line, the route can be the line as built or a set of tower positions, and the bands are the servitude width, the safety zone and any outer band you choose.

[See the full sample](/results)
::::
:::::
::::

::::section{id="after-events" tone="alt" eyebrow="Cyclones and floods" title="Before-and-after checks along the line"}
Cyclones and river floods reach the lines that cross the central and southern provinces. After an event, radar imagery maps the flooded area even under cloud, and a reviewer compares structures and assets along your line before and after on optical or drone imagery captured once the weather clears. Before the season, we highlight the spans that cross low-lying, flood-exposed or erosion-prone ground. It covers visible change only: it is not a structural assessment of towers or conductors.

:::solutions{keys="terrain-flood-post-event" cols="1"}
:::
::::

::::section{id="drones" eyebrow="Drones along lines" title="Close checks, within the rules for live wires"}
:::::columns{split="1-1"}
::::col
A drone survey of flagged spans captures orthophotos and elevation models where satellite detail is not enough, subject to the approvals and security clearances each job requires. IACM's RPAS directive keeps drones at least **5 m laterally and 20 ft vertically** from live high-tension wires, requires the written permission of a structure's owner and a cordon within 50 m of structures, keeps flights at or below 400 ft and within the pilot's sight, and allows none beyond it. Lei n.º 6/2024 adds the Defence authorisation for the survey and the authorisation to release its data.

Afridrone is working towards the operator approvals Mozambique requires; each drone proposal lists the authorisations its flights need.
::::
::::col
:::callout{tone="legal" title="Drone law in Mozambique"}
IACM approvals, the Lei n.º 6/2024 authorisations and data rules, the flight limits and a checklist for anyone commissioning a drone survey, with sources.

[Read the guide](/mz/drone-regulations)
:::
::::
:::::
::::

::::section{id="scope" tone="alt" eyebrow="Honest scope" title="What we do, and what we don't"}
:::::columns{split="1-1"}
::::col
### What we do

:::checklist
- Map structures, cleared ground, vegetation change, fire and tracks along the line
- Measure each structure to the axis, within the widths you set
- Flag change between dated surveys, confirmed by a reviewer
- Deliver dated, credited records in Portuguese or English
:::
::::
::::col
### What we don't do

:::checklist{tone="no"}
- Detect vandalism, energy theft or illegal connections
- Calculate conductor clearances or assess towers
- Identify, count or follow people or vehicles
- Register servitudes, demarcate DUATs or decide who is owed compensation
:::
::::
:::::
::::

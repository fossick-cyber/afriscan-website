---
key: oil-gas
template: industry
slug: pipelines
title: Pipeline Servitude Encroachment Mapping in SA | AfriScan
description: Structures and fresh digging inside fuel and gas pipeline servitudes in South Africa, measured to the line, compared between surveys and reviewed by a person.
h1: Pipeline servitudes and gas sites, mapped from the air
crumb: Pipelines
eyebrow: Oil & gas · South Africa
lead: Fuel and gas pipelines in South Africa cross farms, townships and fast-growing metros. A house, a spaza shop or a wall can appear on a servitude between two patrols, and so can a trench beside the line. We map the whole servitude, flag what is new since the last survey and give your wayleave and servitude teams a ranked list of stretches to visit. A person reviews every result.
buttons:
  - {label: Send us your route file, intent: proposal}
  - {label: See sample outputs, key: results}
service:
  name: Pipeline servitude encroachment survey, South Africa
  type: Servitude encroachment survey, third-party works flagging and scheduled re-surveys
  description: Structures, fresh excavations, tracks and cleared ground inside the servitudes of fuel and gas pipelines in South Africa, each measured to the line, with 500 m encroachment-density ratings and change between dated surveys, reviewed by a person and delivered as PDF reports with GIS layers.
og:
  headline: Pipeline servitudes in South Africa, mapped from the air
  subline: Structures and fresh digging by distance to the line, change between surveys
related: [route-site-selection, excavation-mapping, hazard-zone-registers]
faq:
  - q: Can you detect leaks, illegal taps or fuel theft?
    a: No. We map what is visible at the surface, such as structures, fresh digging, spoil heaps, tracks and cleared ground, and flag it so your field teams know where to look. The pipe, its condition and anything underground are outside what imagery can show, and we never identify who was there.
  - q: Does this replace our patrols?
    a: No. It tells patrols where to look between visits. A ranked list of the stretches where something new has appeared lets your teams spend their time where the change is, and the register gives them the coordinates.
  - q: Which widths do you measure?
    a: The width of each servitude as registered, which differs between lines and sometimes between sections of one line, plus any wider band your standard uses for third-party works. Send the widths per section, or the servitude polygons, and we report each structure's distance to the line and its band.
  - q: Can you support our class-location or population-density reviews?
    a: Yes, with structure counts within the corridor widths and unit lengths your engineers set, and reviewer categories that separate main buildings from outbuildings. The class study and its conclusions stay with your engineers.
  - q: Can a drone fly near our pump or compressor stations?
    a: If a station is a national key point or strategic installation, the operator notifies SACAA on form CA 101-20 before the flight, with the controlling authority's written permission. Satellite work needs no such notice, and our deliverables leave security measures at your stations out of every image either way.
  - q: Your published sample is from Mozambique. Why?
    a: Because the route owner gave permission to show it. We never publish a client's route, imagery or results without written permission, so the sample on this site is the one we are allowed to show. It is a high-pressure gas pipeline in Mozambique, and the register format is the same in South Africa.
cta:
  title: Send us your route file
  text: "KML, KMZ, GeoJSON, Shapefile, GPX or GeoPackage, or the servitude polygons from your GIS. Tell us the widths, the province and what the register is for, and we reply with a scope, an imagery plan and a written proposal."
  button: Request a proposal
  secondary: See sample outputs
  secondary_href: /results
---

::::section{id="problem" eyebrow="The problem, in your words" title="Between two patrols, the servitude changes"}
:::cards{cols="2"}
:::card{title="“Something new is standing on the servitude.”" icon="houses"}
Homes, spaza shops, walls, carports and sheds go up on pipeline servitudes that run through townships and peri-urban land. Each one is easier to talk about in the week it appears than in the year after.
:::
:::card{title="“Someone has been digging near the line.”" icon="excavation"}
Trenches, foundations, borrow pits and spoil heaps near a buried pipeline are third-party works your teams need to know about early. Patrols on a cycle find them late.
:::
:::card{title="“The metro is growing along our route.”" icon="route"}
Lines laid through open land now run beside new estates, industrial parks and informal settlements. The stretches that need attention move year by year, and a fixed patrol plan does not.
:::
:::card{title="“We need a baseline before the new line goes in.”" icon="calendar"}
New gas import and transmission infrastructure needs routes, servitudes and a dated record of what stood on the land before construction, before settlement follows the new access roads.
:::
:::

### Who this is for

:::chips
- Servitude and wayleave teams
- Pipeline integrity engineers, as users of structure data
- Asset-protection managers, as recipients of flagged areas
- HSE managers who own station safety zones
- GIS and asset-records teams
- Route engineers and EPC contractors on new lines and loops
- Environmental assessment practitioners working for the operator
:::
::::

::::section{id="context" tone="alt" eyebrow="The South African picture" title="Fuel lines, gas lines and the infrastructure still to come"}
:::::columns{split="1-1"}
::::col
NERSA issues the licences for petroleum and gas pipelines under the Petroleum Pipelines Act 60 of 2003 and the Gas Act 48 of 2001. The network spans very different ground: Transnet's multi-product fuel trunk line from Durban inland ([Transnet Pipelines](https://www.transnet.net/SubsiteRender.aspx?id=6794475)); Transnet's inland gas lines to Durban, about 153 km and 420 km ([NERSA](https://www.nersa.org.za/files/files/2024/07/RFD-Transnet-SOC-Ltds-Application-for-Piped-Gas-Tariff-for-2023-to-2026.pdf)); and a long cross-border gas pipeline from Mozambique.

More is coming. Supplies from Mozambique's onshore gas fields, which have provided most of South Africa's gas for two decades, are expected to begin falling after 2028 ([The Conversation, 21 July 2026](https://theconversation.com/a-sharp-fall-in-gas-supplies-in-2028-threatens-south-africas-economy-how-to-manage-the-fallout-286861)), which brings new import and transmission infrastructure, and with it new routes and servitudes. The [Gas Bill B6-2026](https://www.parliament.gov.za/bill/2327140), introduced in Parliament on 5 March 2026, would repeal the Gas Act.
::::
::::col
| Instrument | What it means for a pipeline |
|---|---|
| Petroleum Pipelines Act 60 of 2003; Gas Act 48 of 2001 | NERSA licenses construction and operation of petroleum and gas pipelines |
| NEMA EIA Regulations 2014, as amended | New gas transmission pipelines trigger listed activities in Listing Notice 1 (activity 60) and Listing Notice 2 (activity 7), so they need environmental authorisation |
| [GN 834 of 31 July 2020](https://www.dffe.gov.za/sites/default/files/legislations/nema_genericEMPr_gaspipeline_g43571gon834.pdf) | Consultation on a draft generic environmental management programme for gas transmission pipelines; check whether a final version has been adopted |
| Your servitude agreements | The width, the restrictions on building and the access rights along each section |

A summary for orientation as of 26 September 2026, not legal advice.
::::
:::::
::::

::::section{id="rights-of-way" eyebrow="What we map along the line" title="A register of the servitude, then what changed"}
:::::columns{split="2-1"}
::::col
We buffer the route you supply in its UTM zone at the servitude width for each section, and at any wider band your standard uses for third-party works, and list each structure the review confirms with its distance to the line, its band, its chainage and its coordinates. Each 500 m stretch is rated for **encroachment density**, high, medium or low, by a count rule that tells your servitude team where to go first. It is not an integrity, leak or safety rating.

After the baseline, the route is re-surveyed on a schedule you agree with us. New and removed structures between dated surveys are flagged automatically and confirmed by a reviewer, with before-and-after views, and your team receives a notice after each survey saying what changed and where. Fresh digging, spoil heaps, trenches and new tracks on or near the servitude are flagged the same way, so patrols know where to look. They do not replace patrols.

For class-location and population-density reviews, we supply structure counts within the widths and unit lengths your engineers set, with reviewer categories that separate main buildings from outbuildings. The class study stays with your engineers.
::::
::::col
### What you receive

:::checklist
- A structure register: ID, distance to the line, band, chainage and coordinates
- A density rating for each 500 m stretch, and a ranked list of stretches
- Fresh excavation, spoil, trenches and new tracks near the line, flagged for checking
- New and removed structures between surveys, with before-and-after views
- A change notice after each survey
- Structure counts per unit length for class-location reviews, where scoped
- A PDF report and GeoPackage, GeoJSON, KMZ and Shapefile layers
:::
::::
:::::

:::solutions{keys="right-of-way-monitoring,change-detection,excavation-mapping" cols="3"}
:::
::::

::::section{id="sample" tone="alt" eyebrow="What a register looks like" title="Our published sample: a high-pressure gas pipeline in Mozambique" lead="A reviewed sample from the pipeline's route, shown with the route owner's permission. The chart below is the encroachment-density strip from that register: each bar is a 500 m stretch, its height the number of structures within 100 m of the line, the widest band on that route."}
:::segments{data="sample-pipeline"}
:::

The densest stretches are where the route runs beside an existing track through farmland and homesteads; the quietest cross bush and burnt grassland. The operator's own facilities near both ends of the route are part of the asset, not encroachment, and are not in the register. The format, the rating rule and the deliverables are the same for a South African servitude.

[The full sample, with the register excerpt and example stretches](/results)
::::

::::section{id="stations" eyebrow="Stations and gas sites" title="Around pump stations, compressor stations and plants"}
:::::columns{split="1-1"}
::::col
Around a pump station, compressor station or terminal, the line becomes a boundary and the question becomes a ring: what stands inside the site, what stands in a band around it, and what sits inside the safety zones your engineers define. We map all three and show how settlement around the station and its access roads changes between surveys. The zones come from your engineers; we do not model releases or blast radii.

Your own stations, valve sites and yards are part of the asset. We record them as your installations, never as encroachment, and keep them out of the encroachment register unless you ask for them to be listed separately.
::::
::::col
:::callout{tone="legal" title="Key points and security measures"}
Some stations are national key points: one cross-border gas pipeline operator states that a compressor station near the border has been declared one. Drone flights adjacent to or above a key point need SACAA notice on form CA 101-20 with the controlling authority's permission. Our deliverables leave the security measures at your stations out of every image and report, and we never publish imagery of a client's installations without the client's written permission.

[Drone law in South Africa](/za/drone-regulations#sensitive-sites)
:::
::::
:::::
::::

::::section{id="new-lines" tone="alt" eyebrow="New pipelines and loops" title="Weigh the land before the route is fixed"}
For new gas lines, loops and repurposed sections, we compare the structures along alternative alignments, screen steep, erosion-prone and flood-prone stretches and drainage crossings, and fix a dated baseline once the route is chosen, so that what stood on the land before construction is on record before new access roads bring new building. The comparison supports your environmental assessment practitioner's work and your landowner engagement; it is not the alternatives assessment itself.

[Route and servitude baselines](/za/transmission-route-baseline) · [Route and site selection, in detail](/solutions/route-site-selection)
::::

::::section{id="imagery" eyebrow="Satellite first, drone where it adds something" title="The right picture for each stretch"}
:::cards{cols="3"}
:::card{title="Satellite screening" icon="satellite"}
Dated very-high-resolution satellite scenes and open building datasets cover the whole servitude with no site visit and no flight, however long the line.
:::
:::card{title="Your own imagery" icon="layers"}
If your teams or contractors already fly the line, we run the same analysis on your orthophotos, and dated satellite scenes fill the time between flights where they exist.
:::
:::card{title="Drone checks" icon="drone"}
Drone surveys of the flagged stretches, subject to the approvals and security clearances each job requires. In South Africa that includes a UASOC operator with an Air Service Licence, landowner permission for each flight, and approvals for flights within 50 m of structures, people or roads.
:::
:::
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
- Detect leaks, taps, theft or anything about the pipe's integrity
- Identify, count or follow people or vehicles
- Replace patrols, a professional land surveyor or your integrity programme
- Decide whether a structure is lawful, or who must move
:::
::::
:::::
::::

---
key: oil-gas
template: industry
title: Pipeline Wayleave Structure Registers in Kenya | AfriScan
description: Structures, kiosks and new works inside petroleum pipeline wayleaves in Kenya, mapped by band and segment from dated imagery, with repeat scans.
h1: Pipeline wayleave structure registers for Kenya
crumb: Pipelines
eyebrow: Oil & gas · Kenya
lead: Kenya's petroleum pipelines run from Mombasa through Nairobi to Eldoret and Kisumu, often through dense markets and housing. Safety and compensation processes need a dated, neutral record of what stands where, how far it is from the line, and what was there on the date that matters. AfriScan produces that record from satellite imagery, checked by a person.
buttons:
  - {label: Send us your route file, intent: proposal}
  - {label: See sample outputs, key: results}
service:
  name: Pipeline wayleave structure register, Kenya
  type: Wayleave structure register, excavation flagging and repeat scans
  description: Registers of the structures, new excavations, tracks and cleared ground inside petroleum pipeline wayleaves and station buffers in Kenya, each measured to the centreline, ranked by 500 m segment, compared between dated scans, checked by a person and delivered as a PDF report with GeoPackage, GeoJSON, KMZ and Shapefile layers.
og:
  headline: Pipeline wayleave structure registers for Kenya
  subline: Structures and new works by band and 500 m segment, compared between dated scans
related: [route-site-selection, excavation-mapping, change-detection]
faq:
  - q: What width is a pipeline wayleave in Kenya?
    a: It depends on the line and the section, and no standard operator width is published in the sources we checked. Send the width in your wayleave agreement, gazette order or company standard for each section, or the wayleave itself as a polygon, plus any wider safety buffer you use around depots and stations. Each structure is reported against each band.
  - q: Can you detect leaks, illegal taps or fuel theft?
    a: No. Imagery cannot see a tap, a leak, a theft or the condition of the pipe. We map what is visible at the surface, such as structures, fresh digging, spoil heaps, tracks and cleared ground, and flag it so your field teams know where to look.
  - q: We already have fibre sensing on the pipeline. What does this add?
    a: A different view. Sensing on the pipe can warn of digging as it happens at the pipe. A register covers the whole width of the wayleave and its buffers, shows what stands there and what has changed between two dated images, and gives your land and wayleave teams a located, dated list to work from. The two are complementary; neither replaces patrols.
  - q: Can the register support a notice or compensation process?
    a: Yes, as a record. It shows what stood where, how far from the line and on which imagery date, so notices under the Land Act, compensation assessments and any tribunal or court process can work from the same facts. Whether a structure is authorised, and what happens to it, is for the operator, the National Land Commission and the courts under the procedures the Land Act sets. We identify no one.
  - q: Can a drone fly near our depots and pump stations?
    a: Ask KCAA whether they count as strategic installations, because operating in or around one without KCAA's permission is treated as negligent or reckless operation (UAS Regulations 2025, reg 45(2)(c)), and photographing a prohibited place without the authority of the officer in charge is an offence under the Official Secrets Act. Satellite registers need no flight, and we never publish imagery of a client's installations without written permission.
  - q: Your published sample is from Mozambique. Why?
    a: Because its owner gave permission to show it. We never publish a client's route, imagery or results without written permission, so the sample on this site is the one we are allowed to show. It is a high-pressure gas pipeline in Mozambique, and the register format is the same for a Kenyan wayleave.
cta:
  title: Send us your route file
  text: "KML, KMZ, GeoJSON, Shapefile, GPX or GeoPackage, or the wayleave polygons from your GIS. Tell us the widths per section, the counties, any station buffers and what the register is for, and we reply with a scope, an imagery plan and a written proposal."
  button: Request a proposal
  secondary: See sample outputs
  secondary_href: /results
---

::::section{id="problem" eyebrow="The problem, in your words" title="Where the line meets the town"}
Where pipelines pass through towns, wayleaves and safety buffers fill with kiosks, stalls, sheds and houses. Each one is a safety risk for the people nearby and for the line, and each needs to be handled through a lawful, documented process.

:::cards{cols="2"}
:::card{title="“There are stalls on the wayleave again.”" icon="houses"}
Kiosks, market stalls, mabati sheds, jua kali workshops and houses go up on wayleaves that run through towns and estates. A structure is easier to talk about in the month it appears than in the year after.
:::
:::card{title="“Someone has been digging beside the line.”" icon="excavation"}
Foundations, trenches, pits and spoil heaps near a buried pipeline are third-party works your teams need to see early. Patrols on a fixed cycle find them late.
:::
:::card{title="“The town has grown along our route.”" icon="route"}
Lines laid through open land now run beside new estates, markets and roads. The stretches that need attention move from year to year, and a fixed patrol plan does not.
:::
:::card{title="“We need a record before the new line goes in.”" icon="calendar"}
New lines need routes, wayleaves and a dated record of what stood on the land before the preliminary survey work began, because the Land Act compensates damage from that work too.
:::
:::

### Who this is for

:::chips
- Wayleave and land teams
- Pipeline integrity engineers, as users of structure data
- HSE managers who own depot and station buffers
- GIS and asset-records teams
- Route engineers and EPC contractors on new lines
- ESIA and RAP consultants working for the operator
:::
::::

::::section{id="law" tone="alt" eyebrow="The law, once" title="What Kenyan law says about pipeline wayleaves"}
:::::columns{split="1-1"}
::::col
| Provision | What it says |
|---|---|
| Land Act 2012, s.143 | A wayleave for a public authority or a corporate body runs with the servient land and binds every owner and occupier from time to time; the holder's staff, agents and contractors may enter to build, maintain and pass |
| Land Act, s.144(4) | Before it is created, the applicant serves notice on all occupiers, including customary pastoral rights holders and everyone in actual occupation of urban and peri-urban land |
| Land Act, s.148 | Compensation for the use of the land and for damage to trees, crops and buildings, including damage from preliminary survey work, paid promptly by the applicant |
| Energy Act 2019, s.170 | Energy infrastructure may be developed "on, through, over or under any public, community or private land" |
| Energy Act, s.171 | The owner's prior consent to enter land to survey it; where the owner cannot be traced, 15 days' notice in at least two national newspapers and on local radio for two weeks |
| Petroleum Act 2019, s.99(1)(h) | Trespassing or encroaching on a petroleum pipeline wayleave or installation is an offence carrying a minimum fine or a minimum term of imprisonment |
::::
::::col
The same facts run through all of these: which structures stand inside the wayleave, how far each is from the line, and whether it was there on the date of a notice, a Gazette order or an earlier scan. Occupiers can claim for loss or damage within three months after the development (Energy Act, s.173(1)(b)), so a record made before work starts is part of the project file.

A summary for orientation as of 27 September 2026, not legal advice. The [wayleave law guide](/ke/wayleave-law-guide) sets out the procedure, the compensation rules and the offences in full, with sources.
::::
:::::
::::

::::section{id="what-we-map" eyebrow="What we map along the line" title="A register of the wayleave, ranked by segment"}
:::::columns{split="2-1"}
::::col
We buffer the route you supply in its UTM zone at the wayleave width for each section, and at any wider band your standard uses for safety or third-party works, and list each structure the review confirms with its distance to the centreline, its band, its chainage and its coordinates. Each 500 m segment is rated high, medium or low by the number of structures it holds, and the busiest segments are ranked first, so your wayleave team knows which stretches to walk first. It is a count rule, not an integrity, leak or safety rating.

Fresh excavations, spoil heaps, trenches, new tracks and cleared ground on or near the wayleave are flagged the same way. They point field teams to the ground that has changed; they do not replace patrols.

**Around depots and pump stations**, the line becomes a boundary and the question becomes a ring: what stands inside the site, what stands in a band around it, and what sits inside the safety buffers your engineers define. We map all three. The buffers come from your engineers; we do not model releases or blast radii. Your own depots, stations and yards are part of the asset, never encroachment, and stay out of the structure register unless you ask for them to be listed separately.
::::
::::col
### What you receive

:::checklist
- A structure register: ID, distance to the centreline, band, chainage and coordinates
- A rating for each 500 m segment, and a ranked list of segments
- Fresh excavation, spoil, trenches and new tracks near the line, flagged for checking
- Rings and buffers around depots and stations
- The imagery date and source for each feature
- A PDF report in English and GeoPackage, GeoJSON, KMZ and Shapefile layers
:::
::::
:::::

:::solutions{keys="right-of-way-monitoring,change-detection,excavation-mapping" cols="3"}
:::
::::

::::section{id="repeat-scans" tone="alt" eyebrow="After the baseline" title="Repeat scans that show what is new"}
:::::columns{split="1-1"}
::::col
Repeat scans show where new structures or works have appeared in the corridor since the last dated image, so the operator can follow up through its own lawful process. The route is re-scanned on a schedule agreed with you. New and removed structures between dated scans are flagged automatically and confirmed by a reviewer, with before-and-after views, and your team receives a notice after each scan saying what changed and where.

After a notice exercise, a relocation or a compensation round, the next scan shows whether the wayleave has stayed as the record left it, stretch by stretch.
::::
::::col
:::callout{tone="note" title="What decides the schedule"}
- How fast the busiest segments change
- The dates your own process needs: a notice, a Gazette order, a follow-up visit
- Archive and new-capture availability for the stretch
- Cloud in the rainy seasons: optical imagery cannot see through it, so radar comparisons can cover larger changes in between, with flagged areas followed up on optical imagery or with a drone check

[Change detection and repeat scans](/solutions/change-detection)
:::
::::
:::::
::::

::::section{id="sample" eyebrow="What a register looks like" title="Our published sample: a high-pressure gas pipeline in Mozambique" lead="A reviewed sample, shown with its owner's permission. Each bar below is a 500 m segment of that route; its height is the number of structures within 100 m of the line, the widest band on that job."}
:::segments{data="sample-pipeline"}
:::

The busiest segments are where the route runs beside an existing track through farmland and homesteads; the quietest cross bush and burnt grassland. The operator's own facilities near both ends of the route are part of the asset, not encroachment, and are not in the register. For a Kenyan wayleave the bands follow your widths, and the format, the rating rule and the deliverables stay the same.

[The full sample, with the register excerpt and example segments](/results) · [Oil and gas pipelines across Africa](/industries/oil-gas)
::::

::::section{id="new-lines" tone="alt" eyebrow="New lines and extensions" title="Weigh the land before a route is fixed"}
:::::columns{split="2-1"}
::::col
A new eastern pipeline from Mombasa to Nairobi and a cross-border line from Eldoret towards Uganda are planned. For new lines, loops and extensions, we compare the structures along alternative alignments, so the land and compensation each option would affect can be weighed before the route is fixed. Once a route is chosen, a dated imagery record from before the preliminary survey work shows what stood on the land before anyone entered it, which matters because the Land Act compensates damage from that work (section 148(3)).

The comparison supports your ESIA consultant's work and your engagement with landowners and counties; it is not the alternatives assessment itself.
::::
::::col
:::callout{tone="scope" title="What a register is not"}
We don't detect leaks, illegal taps or theft, and imagery is not real-time. Fibre sensing on the pipe can warn of digging as it happens; our registers show what stands and what has changed across the whole corridor between dates. We map structures, not people, and whether a structure is authorised is for the operator and the authorities to decide.
:::

[Route and site selection](/solutions/route-site-selection) · [Imagery history and land due diligence](/solutions/imagery-history-due-diligence)
::::
:::::
::::

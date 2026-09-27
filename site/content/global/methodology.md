---
key: methodology
title: How We Detect, Review & Measure Structures | AfriScan
description: Inputs, dated imagery, detection sources, reviewer checks, distance and buffer rules, the 500 m density rating, and what results can and cannot tell you.
h1: How we detect, review and measure structures
crumb: Methodology & review
section: how
nav_group: how
nav_order: 20
nav_label: Methodology & review
nav_blurb: Sources, reviewer checks, distance rules and limits
eyebrow: Methodology
lead: The method behind every AfriScan register, written for the integrity engineers, GIS leads, land teams and ESIA authors who have to rely on it and explain it to others. It sets out what is automatic, what a person checks, how distances and ratings are worked out, and where the limits are.
buttons:
  - {label: Request a proposal, intent: proposal}
  - {label: See a reviewed sample, key: results}
og:
  headline: How we detect, review and measure structures
  subline: Sources, reviewer checks, distance rules and limits
faq:
  - q: Do you publish an accuracy figure?
    a: No. Detection quality depends on the imagery, the season, tree cover and local building styles, so a single figure would mislead you about your own route. Instead, a person reviews every result, each report states the imagery used, and structures the imagery cannot settle are listed for a ground check. On a large project we can agree a check of a sample of structures against your own field data before the full register is delivered.
  - q: Which structures do you count?
    a: Buildings and other roofed structures visible on the imagery, including thatch, mud-brick and metal-roofed houses, outbuildings and larger sheds. We do not count people, vehicles, crops or trees. On request, reviewers can tag structure categories the imagery shows, such as main building or outbuilding. Cleared ground, excavations and tracks are mapped as separate layers where the scope includes them.
  - q: How are structures near a buffer line handled?
    a: Each structure is placed in a band by its measured distance. Georeferencing of any imagery has some error, so structures within a few metres of a band edge can fall on either side; the register gives each distance so your team can see which ones are close to the edge and check them first.
  - q: Can you use our own survey data or imagery?
    a: Yes. Send georeferenced drone orthophotos or satellite scenes and we run the same analysis on them. Results depend on the imagery's quality and resolution, which we check before we scope the work. Your own survey control or field data can also be used to check positions.
  - q: What happens to low-confidence detections?
    a: They are not silently dropped or silently kept. A reviewer decides each one on the imagery; anything the imagery cannot settle goes to a "verify on the ground" list in the report, with its location and a photo crop.
  - q: Is the register a cadastral or legal survey?
    a: No. It records what stands on the ground, where, and how far from your line, on the imagery date. Boundaries, servitude diagrams and title are the work of professional land surveyors and the land authorities in each country. Our registers are designed to sit alongside that work, not to replace it.
related: [right-of-way-monitoring, change-detection, insight-satellite-or-drone]
cta:
  title: Want the method applied to your ground?
  text: Send a route or boundary and the distances that matter. The proposal states the imagery, the review steps and the limits for your area before any work starts.
  button: Request a proposal
---

::::section{id="principles" eyebrow="In one paragraph" title="Automatic sources propose, a person decides"}
:::::columns{split="2-1"}
::::col
AfriScan combines open building-footprint datasets and open-source segmentation models, run on imagery chosen for your job, to propose where structures stand. A reviewer then goes through the whole route or site on the imagery, confirms what is real, removes what is not, adds what every automatic source missed and lists anything uncertain for a ground check. Distances are measured to the line or boundary you supply, in metres on the ground, and each structure is placed in the bands you choose. The register names its imagery and, where the source gives one, its capture date. Nothing automatic reaches you without that review.
::::
::::col
:::callout{tone="scope" title="What this page is for"}
Use it to judge whether our registers are fit for your purpose, to brief a colleague, or to answer a tender's methodology question. The same rules apply in every country we work in, and to satellite, drone and client imagery.
:::
::::
:::::
::::

::::section{id="inputs" tone="alt" eyebrow="1 · Inputs" title="What the method starts from"}
:::cards{cols="2"}
:::card{title="A route or a boundary" icon="route"}
KML, KMZ, GeoJSON, Shapefile, GPX or GeoPackage. If you have no file, we draw the line or polygon with you from drawings or known points, and you confirm it before any measurement. Distances are only as good as this line, so we ask what it represents: the pipe centreline, the line's axis, the edge of the servitude, or a concession limit.
:::
:::card{title="The distances that matter" icon="ruler"}
Up to six buffer distances, set by you: 50 m and 100 m by default, or the widths your concession, a statutory strip, a servitude or your company standard sets. For areas we report inside the boundary, in a ring around it, or both, with an optional setback inside the boundary.
:::
:::card{title="The date the record must reflect" icon="calendar"}
A resettlement cut-off date, the start of a construction season, the date of a previous survey or a handover date. The date decides which imagery can serve the job, so we ask for it first.
:::
:::card{title="What the record is for" icon="clipboard"}
Screening a long network, planning patrols, a census, a route comparison or a record that may be relied on later. The purpose sets the imagery plan, the review depth and the deliverables, and it is written into the proposal.
:::
:::
::::

::::section{id="imagery" eyebrow="2 · Imagery and dates" title="Every report names its imagery"}
:::callout{tone="note" title="Imagery statement"}
Every report names its imagery and, where the source provides it, the capture date. Where the date matters (cut-off dates, change detection, evidence packs) we use dated imagery: a drone survey, a purchased satellite scene or your own georeferenced imagery.
:::

| Imagery | Capture date | What it is good for | What to know |
|---|---|---|---|
| Very-high-resolution satellite scene | Stated with the scene | Corridor and area registers, change between dated surveys | Archive coverage, dates and cloud vary by place; we check what exists before scoping |
| New satellite capture | Stated once captured | Areas where the archive is too old | The date depends on satellite availability and weather, and is never guaranteed |
| Drone orthophoto | Recorded at capture | Detail on flagged stretches, dense villages, a fixed date | Subject to the approvals and security clearances each job requires |
| Your own imagery | As supplied | Programmes that already fly or buy imagery | Results depend on its resolution, georeferencing and date |
| Open building footprints | Varies; often years old | A first set of candidate structures | A starting point only, never a current count |

**Resolution.** The open segmentation models we use are built for imagery of about 50 cm per pixel or finer. Coarser open satellite data, such as Copernicus Sentinel, is used only for larger patterns (vegetation, land-cover change, burnt areas), never to count individual buildings. On drone orthophotos, reviewers work at the full resolution of the image.

**Basemaps.** The satellite layers behind common web maps carry no capture date and their terms do not make them survey imagery. We use them to screen and plan, and Google's, credited, to show our [pipeline sample](/results), never as delivered imagery and never for a record that depends on a date. The imagery sources and every open-data credit are listed on [imagery and data sources](/imagery) and [data sources and credits](/data-sources).
::::

::::section{id="detection" tone="alt" eyebrow="3 · Detection, then review" title="Several sources propose; a person decides"}
:::steps
:::step{title="Open building datasets" icon="layers"}
Google Open Buildings, Microsoft Building Footprints and OpenStreetMap give a first set of mapped footprints, each credited as its licence requires. They reflect the imagery those projects used, which may be years older than your survey, so they are a starting list. Footprint layers are filtered to building-sized objects; the largest industrial sheds can drop out, and reviewers add them by hand.
:::
:::step{title="Segmentation on the survey imagery" icon="search"}
Open-source building-segmentation models from the humanitarian mapping community, RAMP and HOT fAIr, run on the survey imagery itself to find structures the datasets miss or that are newer than them. They are tuned to find more rather than less, so they also propose false hits, which is what the review is for.
:::
:::step{title="Merge and record sources" icon="grid"}
Where sources overlap, their results are merged so each structure is counted once. Each structure keeps a record of which sources found it and how confident each was, so a reviewer can see at a glance what the sources agree on and what they do not.
:::
:::step{title="Reviewer check" icon="user-check"}
A reviewer works along the whole route or site on the imagery, in chainage order. Real structures are confirmed; bushes, rocks, shadows and bare patches are removed; structures every source missed are marked by hand; anything the imagery cannot settle goes to a "verify on the ground" list. On long routes several reviewers can share the work, and the register records, for each structure it lists, whether it was proposed automatically or added by a reviewer.
:::
:::

:::::columns{split="1-1"}
::::col
### Why a person reviews every result
- **Local building styles.** Generic detectors miss thatch, mud-brick and zinc roofs, small outbuildings and homesteads under trees.
- **Sources disagree.** Different automatic sources can give very different counts for the same stretch of route. A register has to settle on one answer, and a reviewer is the one who can.
- **False hits look like buildings.** Bushes, termite mounds, rocks and shadows can be proposed as structures, especially by sensitive models.
- **Image edges.** A structure cut by the edge of an image tile can be proposed twice. The merge step and the reviewer remove the duplicate.
::::
::::col
### What the reviewer records
:::checklist
- Confirmed structures, with any category the scope asks for
- Structures added by hand, marked as reviewer additions
- Proposals removed as false, kept in the working record
- Structures to verify on the ground, each with a photo crop
- Who reviewed which stretch
:::
Where automatic detection is not enough, for example in dense villages or under mixed tree cover, reviewers mark the whole stretch by hand on the imagery. The [pipeline sample](/results) was built that way and is labelled accordingly.
::::
:::::
::::

::::section{id="distances" eyebrow="4 · Distances and bands" title="How each structure is measured"}
:::figure{src="diagrams/corridor" alt="Schematic of a line route with a 50 m band and a 100 m band, structures coloured by distance to the line, kilometre marks along the route, and a strip below rating each 500 m stretch low, medium, high, medium and low from the structures inside the widest band" caption="Schematic: how distances, bands and the 500 m rating fit together" size="wide" credit="Schematic drawn by AfriScan for illustration. It is not a real route, and the widths shown are the defaults: your register uses the widths you set."}
<span class="band band--a">Within 50 m</span> <span class="band band--b">50 m to 100 m</span> <span class="band band--c">Beyond 100 m</span>
:::

- **Measured to your line.** Distances run from each structure to the route or boundary **as you supplied it**, in the local UTM zone of the route, so they are ground distances in metres, not map-screen distances.
- **Which line is which.** A statutory strip may be counted from a pipe centreline, from a line's axis or from the edge of a servitude. Tell us what your file represents, or send the servitude polygon as well, and we set the bands so the register reads the way your strip is defined.
- **Cumulative bands.** "Within 100 m" includes all the structures within 50 m. A "beyond" band reaches to the edge of the search area, which extends 100 m past the widest buffer, so you can see what sits just outside your strip.
- **The register.** Each structure has an ID, its distance to the line, its band, its chainage (distance along the route), its coordinates in WGS84 and UTM, and a photo crop from the imagery.
- **Edges of bands.** Georeferencing error in any imagery, and GPS-only positioning of drone surveys without ground control, mean a structure within a few metres of a band edge can fall on either side. Those distances are shown so they can be checked first.
::::

::::section{id="rating" tone="alt" eyebrow="5 · Encroachment density" title="The 500 m rating, and what it is not"}
The route is divided into 500 m segments. Each segment is rated from the number of structures inside the widest buffer (100 m by default):

:::facts{cols="3"}
- High: More than 5 structures
- Medium: 1 to 5 structures
- Low: No structures
:::

A structure close to a segment boundary counts in each segment it touches, so segment counts can add up to more than the route total. The report's cover gives the route an overall rating too: high when more than two segments, or at least half of them, are high.

The rating is a density measure that helps you decide where to send people first. **It is not a safety, hazard or integrity assessment**, and it says nothing about whether any structure is authorised. Where your engineers have their own risk classes, we report structures against their zones and leave the risk judgement with them.
::::

::::section{id="areas" eyebrow="6 · Areas and sites" title="Concessions, plant perimeters and project footprints"}
:::::columns{split="1-2"}
::::col
:::figure{src="diagrams/area-ring" alt="Schematic of a site boundary with an inner setback, an outer ring, a facility with an engineer-defined zone, and structures coloured by whether they sit inside the boundary, in the ring or inside the zone" caption="Schematic: how an area survey is read" size="half" credit="Schematic drawn by AfriScan for illustration. It is not a real site."}
<span class="band band--a">Inside the boundary</span> <span class="band band--b">In the ring</span> <span class="band band--c">Beyond the ring</span>
:::
::::
::::col
For a concession, lease, plant perimeter, resettlement site or project footprint, the same method runs on a polygon instead of a line. Structures are reported **inside the boundary**, **in a ring around it** of the width you choose, or both, with an optional setback inside the boundary for a fence line or a buffer strip.

The report gives the area in km², the count in each zone and each structure's distance to the boundary. Where your engineers supply a zone of their own (a blast radius, a flood line, the area downstream of a tailings facility) we list the structures inside it with their distance to the source. Zones come from your engineers; we do not model them.

Operator-owned facilities inside a boundary, such as your own plant, camp or well pad, are recorded as what they are and are never counted as encroachment.
::::
:::::
::::

::::section{id="change" tone="alt" eyebrow="7 · Change between dates" title="New and removed structures"}
Change detection compares structures between two or more surveys with known dates: drone surveys, dated satellite scenes or your own dated imagery. Footprint datasets are never used for change detection, because they describe one moment, not two.

:::cards{cols="3"}
:::card{title="Match" icon="compare"}
Structures on the later date are matched to the earlier one when their footprints overlap or their centres lie within a few metres. Unmatched structures become candidate new or removed structures.
:::
:::card{title="Review" icon="user-check"}
A reviewer checks each candidate on side-by-side and swipe views of the two dates, and on three or four dates where they exist. Only confirmed changes enter the register, each with the date it was first seen.
:::
:::card{title="Compare like with like" icon="history"}
A re-survey is compared with earlier surveys made with the same method on the same kind of imagery. Where the imagery changed, from drone to satellite for example, the report says so.
:::
:::

Two images of the same place are never perfectly aligned, and seasons change how roofs, crops and bare ground look. An offset of a few metres between dates can make one structure look like a removed one and a new one. That is why an automatic flag is a prompt for review, not a finding, and why no change count reaches you before a person has confirmed it.
::::

::::section{id="quality" eyebrow="8 · Quality" title="How you can check our work"}
:::::columns{split="1-1"}
::::col
We do not publish an accuracy percentage. Quality on your route depends on your imagery, the season, tree cover and local building styles, and a figure measured somewhere else would mislead you. What we offer instead can be checked:

:::checklist
- The imagery named in every report, with its capture date where the source gives one
- A reviewer's decision on every automatic proposal
- The sources that found each structure, kept in the working record
- A "verify on the ground" list rather than a guess
- The same method and review rules on every re-survey
:::
::::
::::col
:::callout{tone="scope" title="Check a sample against your own data"}
On a large project we can agree a check before the full register is delivered: your field team visits a sample of structures, including some from the ground-check list, and we compare. Many projects start with one stretch, one site or one concession boundary for exactly this reason. See [how we work](/how-we-work).
:::
::::
:::::
::::

::::section{id="limits" tone="alt" eyebrow="9 · Limits" title="What results can and cannot tell you"}
:::::columns{split="1-1"}
::::col
### What a register shows
:::checklist
- Structures visible from above, their footprint and their distance on the imagery date
- Change between dated surveys, confirmed by a reviewer
- Where structures cluster along a route or around a site
- Cleared ground, excavations and tracks, where the scope includes them
:::
::::
::::col
### What it cannot show
:::checklist{tone="no"}
- Occupancy, ownership, use, value or eligibility for compensation
- Anything under dense canopy, inside buildings or underground
- The condition of a pipeline or line
- A cadastral or statutory survey
- Anything on a day the imagery does not cover, or under cloud
:::
::::
:::::

Our registers support your census and asset inventory under IFC Performance Standard 5; they do not replace them. They map land and assets, never people: we do not identify, count or follow individuals or vehicles.
::::

::::section{id="deliverables" eyebrow="10 · Deliverables" title="What comes out of the method"}
:::cards{cols="3"}
:::card{title="PDF report" icon="file-text"}
A cover with the key figures and the route rating; the survey details (method, imagery source and date, resolution, UTM zone, search edge); an overview map with chainage; the segment table; a photo crop of each structure in chainage order; the coordinate register; and the ground-check appendix. In English or Portuguese.
:::
:::card{title="GIS deliverables" icon="layers"}
GeoPackage, GeoJSON, KMZ and Shapefile layers that open in QGIS, ArcGIS and Google Earth. Each structure carries its distance to the line and its band; change surveys add layers of new, removed and unchanged structures.
:::
:::card{title="Interactive map file" icon="map"}
A self-contained interactive map of your results that your team can open in a web browser, including in the field. It is a file you keep, not an account to manage.
:::
:::

See a [reviewed sample](/results), the [questions buyers ask first](/faq), or [how we work](/how-we-work) with your team.
::::

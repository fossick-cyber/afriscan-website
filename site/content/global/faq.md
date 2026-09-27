---
key: faq
title: Encroachment Monitoring FAQ | AfriScan
description: "Answers for integrity, land and procurement teams: imagery dates, reviewer checks, buffer widths, file formats, data handling, permits and how proposals work."
h1: Frequently asked questions
crumb: FAQ
section: resources
nav_group: resources
nav_order: 20
nav_label: FAQ
nav_blurb: Imagery dates, review, formats, permits and proposals
eyebrow: Resources
lead: Straight answers for integrity engineers, land and resettlement teams, GIS leads and procurement. They are grouped by topic below; if your question is not here, ask it in your request and we will answer it in the proposal.
og:
  headline: Frequently asked questions
  subline: Imagery dates, review, formats, permits and proposals
cta:
  title: Didn't find your question?
  text: Ask it with your request, and send the route or boundary if you have it. We answer in the written proposal, with the scope, the imagery plan and the deliverables.
  button: Request a proposal
---

::::section{id="topics" eyebrow="Jump to" title="Topics"}
:::chips
- [Imagery and dates](#imagery)
- [Review and quality](#review)
- [Distances and bands](#measurement)
- [Change and re-surveys](#change)
- [Deliverables](#deliverables)
- [Drones and permits](#drones)
- [Land, people and data](#land-data)
- [Working with us](#working)
- [Country questions](#countries)
:::
::::

::::section{id="imagery" tone="alt" eyebrow="Imagery and dates" title="Imagery and dates"}
:::details{summary="How recent is the imagery in a survey?" open="true"}
It depends on what exists for your area and on the job. We check the archive before we scope the work and tell you in the proposal what is available. Every report names its imagery and, where the source provides it, the capture date.
:::
:::details{summary="How do we know when an image was taken?"}
Purchased satellite scenes come with their acquisition date and time in the metadata, and drone photos record their capture time. The report states the date of each image used, and the metadata stays with the imagery. Map-service basemaps carry no stated date, which is why we never use them for a dated record.
:::
:::details{summary="Do you use Google Earth or other web-map imagery?"}
To screen and plan, and to show the pipeline sample on this site. Google, Bing and Esri web-map layers have no stated capture date and their terms do not make them survey imagery, so they are never delivered and never the basis of a dated record. The pipeline sample was marked on Google satellite imagery during review; the [sample outputs](/results) show those marks on that imagery, credited "Imagery © Google" and undated, next to strip views of the register and the route on a dated Sentinel-2 scene.
:::
:::details{summary="Can you capture imagery on a date we choose?"}
Not with certainty from satellites: a new capture is requested for a window, and the date depends on satellite availability and weather. A drone survey, subject to the permits each job requires, gives the most control over the date.
:::
:::details{summary="What about cloud and the rainy season?"}
Optical satellite and drone imagery cannot see through cloud. Radar satellites can show larger ground changes, such as clearing and earthworks, under rainy-season cloud, and flagged areas are then checked on optical imagery or by drone.
:::
:::details{summary="Can you use the imagery we already have?"}
Yes. Send georeferenced drone orthophotos or satellite scenes and we run the same structure, change and corridor analysis on them. Results depend on the imagery's quality and resolution.
:::
:::details{summary="Should we choose satellite or drone imagery?"}
For most corridor and concession work, both, in order: satellite for the whole route and its history, drone detail where the satellite picture raises a question or a single date matters. The trade-offs are in [satellite or drone?](/insights/satellite-or-drone-corridor-surveys).
:::
::::

::::section{id="review" eyebrow="Review and quality" title="Review and quality"}
:::details{summary="Does a person check every result?" open="true"}
Yes. Open building datasets and segmentation models propose structures; a reviewer confirms, removes and adds structures on the imagery before anything is delivered. Anything the imagery cannot settle is listed for a ground check.
:::
:::details{summary="What does the reviewer actually check?"}
The whole route or site, in chainage order: whether each proposed structure is real, whether anything was missed (thatch, mud-brick and zinc roofs, small outbuildings and homesteads under trees are the usual misses), and whether a proposal is really a bush, a rock or a shadow. The register records, for each structure it lists, whether it was proposed automatically or added by the reviewer. The steps are set out in the [methodology](/methodology#detection).
:::
:::details{summary="Why don't you publish an accuracy percentage?"}
Because it would mislead. Detection quality depends on the imagery, the season, tree cover and local building styles, so one figure cannot describe your route. We state the imagery used, have a person review the results, and list what needs a ground check. On a large project we can agree a check of a sample against your own field data.
:::
:::details{summary="What if our field team disagrees with the register?"}
Tell us. Field checks are the best test a register can have, and a sample check against your own data can be built into the scope. Corrections go into the register with the same IDs, so later surveys build on the corrected record.
:::
:::details{summary="What is the encroachment-density rating?"}
Each 500 m of route is rated High (more than 5 structures inside the widest buffer), Medium (1 to 5) or Low (none). It is a count rule to help you prioritise, not a safety or integrity assessment. See the [methodology](/methodology#rating).
:::
::::

::::section{id="measurement" tone="alt" eyebrow="Distances and bands" title="Distances and bands"}
:::details{summary="What buffer widths can you report?" open="true"}
Up to six distances, set by you. The default is 50 m and 100 m; we add any width your concession, a statutory strip, a servitude or your company standard sets.
:::
:::details{summary="Which line are distances measured from?"}
From the route or boundary you supply, in metres on the ground, in the route's local UTM zone. Tell us whether your file is a pipe centreline, a line's axis, the edge of a servitude or a concession limit, or send the servitude polygon too, and we set the bands so the register reads the way your strip is defined.
:::
:::details{summary="Can you measure against an area rather than a line?"}
Yes. For a concession, lease, plant perimeter or project footprint we report structures inside the boundary, in a ring around it, or both, with an optional setback inside the boundary. Your own facilities are recorded as what they are and never counted as encroachment.
:::
:::details{summary="How precise are the distances?"}
They are measured precisely from the positions on the imagery, but every image has some georeferencing error, and drone surveys without ground control are positioned by GPS. Structures within a few metres of a band edge can fall on either side, so each distance is given and the close ones can be checked first. A register is not a cadastral survey.
:::
::::

::::section{id="change" eyebrow="Change and re-surveys" title="Change and re-surveys"}
:::details{summary="How do you find what has changed?" open="true"}
We compare structures between two or more surveys with known dates. New and removed structures are flagged automatically and confirmed by a reviewer on side-by-side and swipe views of the dates, each with the date it was first seen. See [change detection](key:change-detection).
:::
:::details{summary="Do you send alerts?"}
After each scheduled re-survey, your team gets an email saying what has changed and where, with the updated register and layers. Imagery supply sets the pace: a re-survey happens when there is new imagery to compare, so we agree a schedule that matches what can realistically be captured.
:::
:::details{summary="How often should a route be re-surveyed?"}
It depends on how fast the land around it changes. Fast-growing stretches near towns may justify frequent re-surveys; remote stretches much less. Each re-survey uses the same method and bands as the baseline, so the results can be compared like with like.
:::
::::

::::section{id="deliverables" tone="alt" eyebrow="Deliverables" title="Deliverables and formats"}
:::details{summary="Which file formats do you accept and deliver?" open="true"}
We accept KML, KMZ, GeoJSON, Shapefile, GPX and GeoPackage, or we draw the route with you. We deliver a PDF report, GeoPackage, GeoJSON, KMZ and Shapefile layers, and an interactive map file.
:::
:::details{summary="What is in the PDF report?"}
A cover with the key figures, the survey details (method, imagery and its date, resolution, coordinate system), an overview map with chainage, the 500 m segment table, a photo crop of each structure in chainage order, the coordinate register and a list of structures to verify on the ground. See the [reviewed sample](/results).
:::
:::details{summary="Is the report available in Portuguese?"}
Yes. Reports are available in English or Portuguese.
:::
:::details{summary="How are results delivered?"}
As files your team keeps: the PDF, the GIS layers and a self-contained interactive map file that opens in a web browser. There is no client portal or login to manage.
:::
::::

::::section{id="drones" eyebrow="Drones and permits" title="Drones and permits"}
:::details{summary="Do satellite surveys need a drone permit?" open="true"}
No drone flies in a satellite-based survey. Satellite-based surveys are available in Mozambique, South Africa and Nigeria.
:::
:::details{summary="Which permits does a drone survey need?"}
It depends on the country and the site: typically an operator approval from the civil aviation authority, flight permissions for the site and airspace, and in some countries separate authorisation for the survey itself or for handing over the images. Every drone proposal sets out what that flight needs. Drone surveys are subject to the permits and authorisations each job requires, and Afridrone is working towards the operator approvals each country requires. The three countries are compared in [drone rules by country](/drone-regulations).
:::
:::details{summary="Who flies the drones?"}
[Afridrone](https://afridr.one/), which flies AfriScan’s drone work, subject to the permits and authorisations each job requires.
:::
::::

::::section{id="land-data" tone="alt" eyebrow="Land, people and data" title="Land, people and data"}
:::details{summary="Can you tell us who lives in a structure, or whether it is legal?" open="true"}
No. We map structures on the land. Occupancy, ownership, use and whether a structure is authorised are for your field teams, your census and the authorities to establish.
:::
:::details{summary="Do you identify or follow people or vehicles?"}
No. Our surveys map land and assets: structures, cleared ground, excavations and tracks. We do not identify people, and we do not follow people or vehicles.
:::
:::details{summary="Can the register support a resettlement census or a dispute?"}
It can support them. A dated register supports IFC Performance Standard 5 cut-off-date records and your census and asset inventory, which it does not replace. Packaged with file fingerprints and an independent timestamp, it supports legal and community processes. See [dated imagery for cut-off dates](/insights/dated-imagery-cut-off-dates).
:::
:::details{summary="How do you handle personal data in imagery?"}
Imagery of homes and yards can hold personal information under laws such as South Africa's POPIA and Nigeria's Data Protection Act. Registers describe structures, not people, carry no names and go only to the contacts you name. How project data is stored and kept is set out in the proposal. For South Africa, see [POPIA and aerial imagery](/za/popia); for the website itself, the [privacy notice](/privacy).
:::
:::details{summary="Will you publish our route or results?"}
No. We never publish a client's route, imagery or results without written permission.
:::
::::

::::section{id="working" eyebrow="Working with us" title="Proposals, tenders and your data"}
:::details{summary="How does a project start?" open="true"}
Send the route or boundary, the country and province, the distances that matter and what the record is for. We reply with a written proposal: scope, imagery plan, method, deliverables and schedule. See [how we work](/how-we-work).
:::
:::details{summary="Can we start with part of the route?"}
Yes. Many projects start with one stretch, one site or one concession boundary, so your team can check the register against what it knows before the rest is surveyed.
:::
:::details{summary="Do you respond to tenders and supplier forms?"}
Yes. Send the tender documents, or your supplier or prequalification forms, with your request; the proposal states which registrations are in place for your contract and which are being arranged. Our guide to [scoping an encroachment survey](/insights/encroachment-survey-scope) sets out what a good scope asks for, whoever you buy from.
:::
:::details{summary="Can you work under an ESIA or engineering consultancy?"}
Yes. Consultancies can commission structure registers, route comparisons and GIS layers for their own reports, delivered in the formats their teams already use.
:::
::::

::::section{id="countries" tone="alt" eyebrow="Country questions" title="Mozambique, South Africa and Nigeria"}
:::details{summary="Mozambique: what is the 50 m partial protection zone?" open="true"}
The Land Law makes the land within 50 m on each side of oil, gas, water, electricity and telecommunications lines a partial protection zone, where no land-use right (DUAT) can be acquired ([Lei n.º 19/97, arts. 8 and 9](https://www.pdul.gov.mz/content/download/486/2635/file/Lei%20de%20Terras.pdf)). Our default bands report structures within 50 m and 100 m. The full guide: [the 50 m partial protection zone](/mz/50m-protection-zone) ([em português](/mz/pt/zona-de-proteccao-parcial-50-metros)).
:::
:::details{summary="South Africa: how do POPIA and servitudes affect a survey?"}
Imagery of homes can be personal information, and some uses of a register may need the Information Regulator's prior authorisation, so the purpose is settled before the first survey. Registers are measured against the servitude widths you supply and are not cadastral surveys. See the [servitude encroachment guide](/za/servitude-encroachment-guide) and [POPIA and aerial imagery](/za/popia).
:::
:::details{summary="Nigeria: what protects a pipeline right of way?"}
The pipeline licence, the Petroleum Industry Act's right-of-way provisions and, where one is made, an order under section 12 of the Oil Pipelines Act restricting buildings and cultivation within up to 100 feet of the boundary of the land in the pipeline licence. See the [pipeline right-of-way guide](/ng/pipeline-right-of-way-guide).
:::
::::

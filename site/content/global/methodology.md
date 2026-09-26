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
lead: The method behind every AfriScan register, written for the integrity engineers, GIS leads and ESIA authors who have to rely on it. It explains what is automatic, what a person checks, how distances and ratings are worked out, and where the limits are.
og:
  headline: How we detect, review and measure structures
  subline: Sources, reviewer checks, distance rules and limits
faq:
  - q: Do you publish an accuracy figure?
    a: No. Detection quality depends on the imagery, the season, tree cover and local building styles, so a single figure would mislead. Instead, a person reviews every result, and each report states the imagery used and lists the structures that need a ground check. On a large project we can agree a check of a sample of structures against your own field data.
  - q: Which structures do you count?
    a: Buildings and other roofed structures visible on the imagery, including thatch, mud-brick and metal-roofed houses, outbuildings and larger sheds. We do not count people, vehicles, crops or trees. On request, reviewers can tag structure categories that the imagery shows, such as main building or outbuilding.
  - q: How are structures near a buffer line handled?
    a: Each structure is placed in a band by its measured distance. Georeferencing of any imagery has some error, so structures within a few metres of a band edge can fall on either side; the report gives each distance so your team can see which ones are close to the edge.
  - q: Can you use our own survey data or imagery?
    a: Yes. Send georeferenced drone orthophotos or satellite scenes and we run the same analysis on them. Results depend on the imagery's quality and resolution, which we check before we scope the work.
  - q: What happens to low-confidence detections?
    a: They are not silently dropped or silently kept. A reviewer decides each one; anything the imagery cannot settle goes to a "verify on the ground" list in the report.
---

::::section{id="inputs" eyebrow="1 · Inputs" title="What the method starts from"}
- **A route or a boundary.** KML, KMZ, GeoJSON, Shapefile, GPX or GeoPackage, or a line or polygon we draw with you and you confirm.
- **Distances.** Up to six buffer distances, set by you; 50 m and 100 m by default. For areas: inside the boundary, in a ring around it, or both, with an optional setback inside the boundary.
- **A date, if the record needs one.** A resettlement cut-off date, the start of a construction season, or the date of a previous survey to compare against.
- **Imagery.** Chosen with you: see the next section.
::::

::::section{id="imagery" tone="alt" eyebrow="2 · Imagery and dates" title="Every report names its imagery"}
:::callout{tone="note" title="Imagery statement"}
Every report names its imagery and, where the source provides it, the capture date. Where the date matters (cut-off dates, change detection, evidence packs) we use dated imagery: a drone survey, a purchased satellite scene or your own georeferenced imagery.
:::

Map-service basemaps, the satellite layers behind common web maps, do not state when their images were captured. We use them only for internal screening and planning, never as delivered survey imagery. Automatic structure detection needs very-high-resolution imagery; coarser open satellite data (Copernicus Sentinel) is used only for larger patterns such as vegetation and land-cover change. Details and credits are on [imagery and data sources](/imagery).
::::

::::section{id="detection" eyebrow="3 · Detection, then review" title="Several sources propose; a person decides"}
:::steps
:::step{title="Open building datasets"}
Google Open Buildings, Microsoft Building Footprints and OpenStreetMap give a first set of mapped footprints, each credited as its licence requires. They reflect the imagery those projects used, which may be years old.
:::
:::step{title="Segmentation on current imagery"}
Open-source building-segmentation models from the humanitarian mapping community (RAMP and HOT fAIr) run on the survey imagery itself, to find structures the datasets miss or that are newer than them.
:::
:::step{title="Merge and record sources"}
Overlapping results are merged into one structure, and each structure keeps a record of which sources found it, so a reviewer can see what is agreed and what is not.
:::
:::step{title="Reviewer check"}
A reviewer goes through the whole route or site on the imagery: confirms real structures, removes false ones (bushes, rocks, shadows), marks structures that every source missed, and sends anything uncertain to a "verify on the ground" list.
:::
:::

Generic detectors miss thatch, mud-brick and zinc roofs, which is why no automatic result goes to a client unreviewed. Where automatic detection is not enough, for example in dense villages or mixed tree cover, reviewers mark structures by hand on the imagery. Every mark is attributed, so the register shows who confirmed what.
::::

::::section{id="distances" tone="alt" eyebrow="4 · Distances and bands" title="How each structure is measured"}
- Distances are measured from each structure to the route or boundary **as you supplied it**, in the route's local UTM zone, so they are true ground distances in metres.
- Bands are **cumulative**: "within 100 m" includes all the structures within 50 m. A "beyond" band reaches to the edge of the search area, which extends past the widest buffer.
- The register gives each structure an ID, its distance, its band, its chainage (distance along the route) and its coordinates in WGS84 and UTM.
- Georeferencing error in any imagery means structures within a few metres of a band edge can fall on either side. Those distances are shown so they can be checked.
::::

::::section{id="rating" eyebrow="5 · Encroachment density" title="The 500 m rating, and what it is not"}
The route is divided into 500 m segments. Each segment is rated from the number of structures inside the widest buffer (100 m by default):

:::facts{cols="3"}
- High: More than 5 structures
- Medium: 1 to 5 structures
- Low: No structures
:::

A structure close to a segment boundary counts in each segment it touches, so segment counts can add up to more than the route total. The rating is a density measure that helps you decide where to send people first. **It is not a safety, hazard or integrity assessment**, and it says nothing about whether a structure is authorised.
::::

::::section{id="change" tone="alt" eyebrow="6 · Change between dates" title="New and removed structures"}
Change detection compares structures between two or more surveys with known dates. New and removed structures are flagged automatically and then confirmed by a reviewer, with before-and-after views of each change. Two images of the same place are never perfectly aligned, and seasons change how roofs and vegetation look, so a flag is a prompt for review, not a finding in itself. Footprint datasets are not used for change detection, because they describe one moment, not two.
::::

::::section{id="limits" eyebrow="7 · Limits" title="What results can and cannot tell you"}
:::::columns{split="1-1"}
::::col
:::checklist
- Structures visible from above, their footprint and their distance on the imagery date
- Change between dated surveys, confirmed by a reviewer
- Where structures cluster along a route
:::
::::
::::col
:::checklist{tone="no"}
- Occupancy, ownership, use, value or eligibility for compensation
- Anything under dense canopy, inside buildings or underground
- The condition of a pipeline or line
- A cadastral or statutory survey
:::
::::
:::::

Our registers support your census and asset inventory under IFC Performance Standard 5; they do not replace them.
::::

::::section{id="deliverables" tone="alt" eyebrow="8 · Deliverables" title="What comes out of the method"}
:::cards{cols="3"}
:::card{title="PDF report" icon="file-text"}
Clear PDF reports with maps, route tables, a photo of each structure and a coordinate register, in English or Portuguese.
:::
:::card{title="GIS deliverables" icon="layers"}
GeoPackage, GeoJSON, KMZ and Shapefile layers that open in QGIS, ArcGIS and Google Earth.
:::
:::card{title="Interactive map file" icon="map"}
A self-contained interactive map of your results that your team can open in a web browser, including in the field.
:::
:::

See a [reviewed sample](/results), or read [how we work](/how-we-work) with your team.
::::

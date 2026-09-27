---
key: results
title: Sample Encroachment Survey Outputs | AfriScan
description: "What an AfriScan survey delivers: structures mapped along a route, distance to the line, 50 m and 100 m buffer registers, density ratings, PDF and GIS files."
h1: Sample encroachment survey outputs
crumb: Sample outputs
section: resources
nav_group: resources
nav_order: 10
nav_label: Sample outputs
nav_blurb: A reviewed sample from a high-pressure gas pipeline in Mozambique
eyebrow: Resources
lead: Real outputs from a reviewed sample on the route of a high-pressure gas pipeline in Mozambique, shown with the route owner's permission. Every mark is a reviewer's manual mark, drawn exactly as recorded, and the public version leaves out coordinates.
buttons:
  - {label: Ask for a redacted sample report, intent: sample-report}
  - {label: How the method works, key: methodology}
og:
  headline: Sample encroachment survey outputs
  subline: A high-pressure gas pipeline in Mozambique · 50 m and 100 m buffers · reviewer marks
cta:
  title: Want to see your own corridor?
  text: Send the route file (KML, GeoJSON or Shapefile) and the distances that matter to you, and we will scope a survey of your line or site.
  button: Send your route
  intent: proposal
faq:
  - q: Is this sample automatic detection?
    a: No. Every mark in this sample was placed by a reviewer; no automatic detection result is shown. On a client survey, open building datasets and segmentation models propose structures first and a reviewer confirms, corrects and adds to them. See [how we detect, review and measure structures](/methodology).
  - q: Why are the close-ups on Google imagery, and why is there no capture date?
    a: The reviewer marked this sample on Google satellite imagery in our review tool, so the close-ups show the marks on that imagery, credited "Imagery © Google". Google does not state when that imagery was captured, so it shows where each mark sits, not when a structure appeared. Client surveys that need a date use imagery that can be dated and delivered, such as a drone survey, a purchased satellite scene or your own georeferenced imagery. The overview of ratings further down uses a dated Copernicus Sentinel-2 scene.
  - q: Why do some rings sit beside a roof rather than on it?
    a: Each ring is centred on the single point the reviewer placed for a structure. A point does not trace the roof, and the imagery prepared for the review showed less detail than these close-ups, so a ring can sit beside the roof it refers to rather than on it. Distances in the register are measured from those points, exactly as recorded.
  - q: Why is the register drawn as a straight strip?
    a: A strip view straightens the route so that each marked structure sits at its chainage (the distance along the route) and its distance from the line, with the north side of the line at the top. Distances across the route are drawn at twice the scale of distances along it, so the 50 m and 100 m bands can be read. It is a chart of the register, not a map.
  - q: Could a structure be missing from the sample?
    a: Yes, and the close-ups show some. The review recorded the structures the reviewer marked at the time, inside the search area around the route, and roofs inside the bands with no ring were not marked in it. We publish the sample as it was recorded, with nothing added. Structures under tree cover, roofs that blend with the ground and anything outside the search area are also missed from the air. A delivered survey is reviewed against the imagery it names, and anything that cannot be settled from the air is listed for a check on the ground.
  - q: Are the plants and well pads on the route counted?
    a: No. The facilities near both ends of this route belong to the pipeline operator. An operator's own installations are part of the asset, not encroachment, and they are not in the register.
related: [oil-gas, right-of-way-monitoring, insight-survey-scope]
---

::::section{id="at-a-glance" eyebrow="The sample at a glance" title="A high-pressure gas pipeline in Mozambique"}
:::facts{cols="4"}
- Route length: 10.78 km
- Buffers: 50 m and 100 m
- Method: Reviewer marks (manual)
- Structures marked: 59
- Within 50 m of the line: 11
- Within 100 m of the line: 36
- 500 m segments: 22
- Rated high: 4
:::

Counts are cumulative: "within 100 m" includes the 11 structures within 50 m. Distances are measured to the route as supplied, in its local UTM zone (36S). The other 23 marks lie between 100 m and the edge of the search area.
::::

::::section{id="on-imagery" tone="alt" eyebrow="On satellite imagery" title="The reviewer's marks on Google satellite imagery" lead="The whole route, then close-ups of about 600 by 400 m of the stretches where the reviewer placed marks, each with the route, its 50 m and 100 m bands and a ring on each mark the reviewer placed. The letters on the overview show where each close-up sits."}
:::sample-gallery{data="sample-pipeline-google" priority="true"}
:::
::::

::::section{id="corridor-view" eyebrow="Register view" title="The densest stretch, structure by structure" lead="Km 5.0 to 6.5, where the route runs beside an existing track through farmland and homesteads. Each mark sits at its chainage and its distance from the line, and carries the register ID used in the table below."}
:::figure{src="samples/sample-pipeline-register-km5-6" alt="Strip view of the pipeline route between km 5.0 and 6.5: the route as a straight orange line with red 50 m and amber 100 m bands on both sides, and twenty reviewer marks, R20 to R39, placed by chainage and distance; the three 500 m segments below are rated medium (4), high (16) and high (7)" caption="The sample pipeline, km 5.0 to 6.5: reviewer marks by chainage and distance from the line, with the rating of each 500 m segment" badge="Reviewed · manual marks" size="wide" credit="Drawn by AfriScan from the sample register. No imagery; distances across the route drawn at twice the along-route scale."}
<span class="band band--a">Within 50 m</span> <span class="band band--b">50 to 100 m</span> <span class="band band--c">Beyond 100 m</span> The north side of the line is at the top.
:::

### Register excerpt for this stretch

Chainage is the distance along the route from its start. This public sample leaves out coordinates. A client's delivered register gives each structure in WGS84 and UTM, and goes only to the contacts the client names.

:::register{data="sample-pipeline"}
:::
::::

::::section{id="segments" tone="alt" eyebrow="Encroachment density" title="Every 500 m of route, rated" lead="The rating is a count rule that tells you where to send people first. It is not a safety or integrity assessment."}
:::figure{src="samples/sample-pipeline-route-ratings" alt="The pipeline route on a Sentinel-2 satellite scene, running about 10 km from a gas plant in the west, past a settlement, to a wetland in the east; the route is coloured by rating, red for the high stretches between km 4 and 6.5, amber for medium and grey for low" caption="The whole route, each 500 m coloured by its rating: red high, amber medium, grey low" size="wide" credit="Route on a Copernicus Sentinel-2 scene of 2 August 2026 (contains modified Copernicus Sentinel data 2026), shown for location. At 10 m per pixel the scene cannot show individual structures; the ratings come from the reviewed register, not from this scene."}
:::

:::segments{data="sample-pipeline"}
Each segment is rated from the structures within the widest buffer, 100 m on this route, measured from any point of the segment. A structure near a segment boundary counts in both segments it touches, so segment counts add up to more than the route total. In the three views below, the outlined area is what the rating counts; marks outside it are faded.
:::

:::cards{cols="3"}
:::figure{src="samples/sample-pipeline-register-high" alt="Strip view of km 5.5 to 6.0 rated high: sixteen reviewer marks inside the area within 100 m of the stretch, two of them within 50 m of the line; two marks outside that area are faded" caption="High · km 5.5 to 6.0" size="third" credit="Register view; same scale in all three"}
16 structures within 100 m of the stretch, where the route runs beside homesteads on both sides. Three of them stand just past its ends and count for the next stretch too.
:::
:::figure{src="samples/sample-pipeline-register-medium" alt="Strip view of km 9.0 to 9.5 rated medium: two reviewer marks inside the 50 m band on the north side of the line, just past km 9.0" caption="Medium · km 9.0 to 9.5" size="third" credit="Register view; same scale in all three"}
2 structures, both inside the 50 m band: few, but close to the line. They stand just past km 9.0, so km 8.5 to 9.0 is rated medium too.
:::
:::figure{src="samples/sample-pipeline-register-low" alt="Strip view of km 7.5 to 8.0 rated low: the 50 m and 100 m bands with no reviewer marks" caption="Low · km 7.5 to 8.0" size="third" credit="Register view; same scale in all three"}
No structures within 100 m: the line crosses bush and burnt grassland here.
:::
:::
::::

::::section{id="delivery" eyebrow="What a delivery contains" title="The same outputs, for your route"}
:::cards{cols="3"}
:::card{title="PDF report" icon="file-text"}
A cover with the key figures; survey details (method, imagery source and date, UTM zone); an overview map with buffers, chainage and a scale bar; the segment table; a photo of each structure in chainage order; and the full coordinate register. In English or Portuguese.
:::
:::card{title="GIS layers" icon="layers"}
GeoPackage, GeoJSON, KMZ and Shapefile. Each structure carries its distance to the line and its buffer band; the route is split into rated segments. They open in QGIS, ArcGIS and Google Earth.
:::
:::card{title="Interactive map file" icon="map"}
A self-contained map of the results that opens in a web browser. The result layers work offline; background imagery needs a connection unless your survey imagery is included.
:::
:::

:::cta{title="Want to open the files before a proposal?" text="Ask for the redacted sample: the report and GIS layers in the delivered format, with the route generalised, no coordinates and no basemap imagery, so your GIS and land teams can check how the files are built." button="Ask for a redacted sample" intent="sample-report"}
:::
::::

::::section{id="about-sample" tone="alt" eyebrow="About this sample" title="What this sample can and cannot show"}
:::::columns{split="1-1"}
::::col
:::checklist
- Where structures stand relative to the line, by band and by stretch of route
- How the density rating turns a register into a list of stretches to visit first
- What your GIS team receives and how the files are structured
:::
::::
::::col
:::checklist{tone="no"}
- Who lives in a structure, who owns it, or what it is used for
- Whether a structure is authorised: that is for the route owner and the authorities
- Anything about the pipeline's condition or integrity
:::
::::
:::::

:::callout{tone="scope" title="Shown with permission"}
We never publish a client's route, imagery or results without written permission. The route is shown here with its owner's permission. The facilities near both ends of the route are the operator's own installations: they are part of the asset, not encroachment, and are not in the register.
:::
::::

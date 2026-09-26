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
nav_blurb: A reviewed sample from the T-9 pipeline route in Mozambique
eyebrow: Resources
lead: Real outputs from a reviewed sample on the T-9 replacement pipeline route in Inhambane Province, Mozambique, shown with the route owner's permission. The marks are shown exactly as the reviewer recorded them.
buttons:
  - {label: Request a sample report, intent: sample-report}
  - {label: How the method works, key: methodology}
og:
  headline: Sample encroachment survey outputs
  subline: T-9 replacement pipeline, Inhambane · 50 m and 100 m buffers · reviewer marks
cta:
  title: Want to see your own corridor?
  text: Send the route file (KML, GeoJSON or Shapefile) and the distances that matter to you, and we will scope a survey of your line or site.
  button: Send your route
  intent: proposal
faq:
  - q: Is this sample automatic detection?
    a: No. Every mark in this sample was placed by a reviewer on satellite imagery; no automatic detection result is shown. On a client survey, open building datasets and segmentation models propose structures first and a reviewer confirms, corrects and adds to them. See [how we detect, review and measure structures](/methodology).
  - q: Why is the imagery not dated?
    a: This sample was reviewed on a Google satellite basemap, which does not state when its images were captured. That is fine for a sample and for internal screening, but not for a record that must reflect a date. Client surveys that need a date use dated imagery, such as a drone survey, a purchased satellite scene or your own georeferenced imagery.
  - q: Some structures on the imagery have no mark. Why?
    a: The review recorded the structures the reviewer confirmed at the time, inside the search area around the route. Structures under tree cover, roofs that blend with the ground and anything outside the search area are not marked. A delivered survey is reviewed against the imagery it names, and anything that cannot be settled from the air is listed for a check on the ground.
  - q: Why are the boxes slightly off some roofs?
    a: Reviewer marks are points. The boxes are drawn around each point so they can be seen at this scale, and a point placed at the edge of a roof puts its box partly on the ground next to it. Distances are measured from the point.
  - q: Are the plants and well pads on the route counted?
    a: No. The facilities near both ends of this route belong to the pipeline operator. An operator's own installations are part of the asset, not encroachment, and they are not in the register.
related: [oil-gas, right-of-way-monitoring, insight-survey-scope]
---

::::section{id="at-a-glance" eyebrow="The sample at a glance" title="T-9 replacement pipeline, Inhambane"}
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

::::section{id="corridor-view" tone="alt" eyebrow="Map view" title="What the register looks like on the ground" lead="The densest stretch of the route, where it runs beside an existing track through farmland and homesteads. Each mark carries the register ID used in the table below."}
:::figure{src="samples/t9-km5-6" alt="Satellite view of the T-9 route between km 5.0 and 6.3, with the route in orange, a red 50 m band, an amber 100 m band and reviewer marks R21 to R40 on homesteads on both sides of the line" caption="T-9 replacement pipeline, km 5.0 to 6.3: route, 50 m and 100 m buffers, and reviewer marks coloured by band" badge="Reviewed · manual marks" size="wide" priority="true" credit="Imagery © Google. The background is a Google satellite basemap shown for illustration: it has no capture date and is not delivered survey imagery. Route, buffers, marks and IDs drawn by AfriScan from the review data."}
Red marks are within 50 m of the line, amber marks between 50 m and 100 m, and teal marks beyond 100 m.
:::

### Register excerpt for this stretch

Chainage is the distance along the route from its start. The public sample leaves out coordinates; the delivered register gives each structure in WGS84 and UTM.

:::register{data="t9"}
:::
::::

::::section{id="segments" eyebrow="Encroachment density" title="Every 500 m of route, rated" lead="The rating is a count rule that tells you where to send people first. It is not a safety or integrity assessment."}
:::segments{data="t9"}
Each segment is rated from the structures within the widest buffer, 100 m on this route. A structure near a segment boundary counts in both segments it touches, so segment counts add up to more than the route total.
:::

:::cards{cols="3"}
:::figure{src="samples/t9-rating-high" alt="Close view of km 5.5 to 6.0: homesteads on both sides of the route inside the 50 m and 100 m bands" caption="High · km 5.5 to 6.0" size="third" credit="Imagery © Google"}
16 structures within 100 m, where the route runs beside homesteads on both sides.
:::
:::figure{src="samples/t9-rating-medium" alt="Close view around km 9.0: two reviewer marks on small plots inside the 50 m band, just north of the route" caption="Medium · around km 9.0" size="third" credit="Imagery © Google"}
2 structures, both inside the 50 m band: few, but close to the line. They sit on the boundary between two segments, so km 8.5 to 9.0 and km 9.0 to 9.5 are both rated medium.
:::
:::figure{src="samples/t9-rating-low" alt="Close view around km 8.0: the route runs through bush and burnt grassland with no structures inside either buffer" caption="Low · km 7.5 to 8.5" size="third" credit="Imagery © Google"}
No structures within 100 m in either segment: a cleared right of way through bush and burnt grassland.
:::
:::
::::

::::section{id="delivery" tone="alt" eyebrow="What a delivery contains" title="The same outputs, for your route"}
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

:::cta{title="Want the full sample report?" text="We can send the PDF and GIS files for this sample so your GIS and land teams can open them." button="Request a sample report" intent="sample-report"}
:::
::::

::::section{id="about-sample" eyebrow="About this sample" title="What this sample can and cannot show"}
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
We never publish a client's route, imagery or results without written permission. The T-9 route is shown here with the route owner's permission. The facilities near both ends of the route are the operator's own installations: they are part of the asset, not encroachment, and are not in the register.
:::
::::

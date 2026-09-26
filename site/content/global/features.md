---
key: how-it-works
hub: how
title: "How It Works: Structure Mapping & Buffer Zones | AfriScan"
description: Send a KML, GeoJSON or Shapefile route and buffer widths. We map the structures, measure each to the line, review them and deliver PDF and GIS files.
h1: How AfriScan maps structures along your corridor
crumb: How it works
nav_group: how
nav_order: 10
nav_blurb: From your route file to a reviewed register, step by step
eyebrow: How it works
lead: From your route or boundary file to a reviewed register of structures, with every stretch of the route rated and every result checked by a person. No site visit is needed to start.
buttons:
  - {label: Request a proposal, intent: proposal}
  - {label: See sample outputs, key: results}
og:
  headline: How AfriScan maps structures along your corridor
  subline: Route file in, reviewed register out, with PDF reports and GIS files
faq:
  - q: What file formats can we send?
    a: KML, KMZ, GeoJSON, Shapefile, GPX or GeoPackage for routes and boundaries. If you do not have a file, describe the location and we draw the route or boundary with you and send it back for your confirmation before anything is surveyed.
  - q: How many buffer distances can a survey report?
    a: Up to six, each set by you. The default is 50 m and 100 m, which matches the width of Mozambique's partial protection zone along pipelines and power lines and gives a second, wider band. Area surveys report structures inside a boundary, in a ring around it, or both.
  - q: Can we see the results before the report is final?
    a: Yes. We can share the register and map for your comments before the PDF is issued, so your team can flag anything it already knows about, such as your own facilities or structures already compensated.
  - q: Do you need access to the site?
    a: No, not for satellite work. Drone surveys need access to take-off points and are subject to the permits and authorisations each job requires, which we plan with you.
  - q: How often can a route be re-surveyed?
    a: On a schedule agreed with you. How often new imagery can be captured depends on satellite availability, cloud and, for drone work, on permits, so we set the cadence per project rather than promising a fixed interval.
---

::::section{id="steps" eyebrow="The process" title="Four steps from route file to register"}
:::steps{style="list"}
:::step{title="Scope the route or site"}
You send the route or boundary and tell us what the record is for: a baseline register, a resettlement cut-off date, a comparison of route options, or scheduled re-surveys. We agree the distances that matter (a statutory protection strip, your company standard, or a hazard zone your engineers have defined), the date the record must reflect, the imagery plan and the deliverables. That goes into a written proposal before any work starts.
:::
:::step{title="Screen from satellite"}
We map structures along the whole route or across the whole site. Open building datasets and open segmentation models run on dated satellite imagery, on drone orthophotos, or on imagery you already hold, and give a first set of structures. Overlapping results are merged and each structure records which sources found it.
:::
:::step{title="Check up close where it matters"}
Where the satellite picture is not enough, for example dense villages under tree cover or a stretch you need recorded in detail, a drone survey captures orthophotos and elevation models of just those stretches. Drone surveys are subject to the permits and authorisations each job requires.
:::
:::step{title="Review, measure and deliver"}
A reviewer checks each structure against the imagery, removes false detections and marks anything missed. Each structure is then measured to the line, placed in its buffer band and chainage, and each 500 m stretch is rated for encroachment density. You receive the PDF report, the GIS layers and an interactive map file, and re-surveys follow on the schedule you choose.
:::
:::
::::

::::section{id="inputs" tone="alt" eyebrow="What you send us" title="Everything a survey needs fits in one email"}
:::::columns{split="1-1"}
::::col
:::checklist
- **The route or boundary:** KML, KMZ, GeoJSON, Shapefile, GPX or GeoPackage, or a description we can draw from
- **Country and province**, and any stretches you already know are sensitive
- **The distances to report:** up to six, 50 m and 100 m by default
- **The date the record must reflect**, if there is one, such as a cut-off date
- **Imagery you already hold**, such as your own drone orthophotos or satellite scenes
- **Deliverables and deadline**, and whether the report should be in English or Portuguese
:::
::::
::::col
:::callout{tone="note" title="Corridors and areas"}
**Corridor mode** measures each structure to a line: a pipeline, power line, road or railway. **Area mode** works from a boundary: a concession, lease, plant perimeter or resettlement site, and counts structures inside it, in a ring around it, or both, with an optional setback inside the boundary.
:::
::::
:::::
::::

::::section{id="measurement" eyebrow="Measurement" title="Distances, bands and the density rating"}
Each structure's distance is measured from the route as you supplied it, in the route's local UTM zone, so distances are in true metres. Bands are cumulative: "within 100 m" includes everything within 50 m, and a "beyond" band reaches to the edge of the search area.

The route is split into 500 m segments. Each segment is rated from the structures inside the widest buffer:

:::facts{cols="3"}
- High: More than 5 structures
- Medium: 1 to 5 structures
- Low: No structures
:::

The rating is a count rule to help you decide where to send people first. It is not a safety, hazard or integrity assessment, and a single structure 10 m from the line can matter more than six at 90 m, which is why each structure is also listed with its own distance. The [methodology](/methodology) page sets out every rule, including how structures near segment boundaries are counted.
::::

::::section{id="repeat" tone="alt" eyebrow="Change over time" title="Re-surveys and change between dates"}
:::cards{cols="2"}
:::card{title="Change detection between dated surveys" icon="compare"}
New and removed structures between two or more dated surveys, flagged automatically and confirmed by a reviewer, with before-and-after views of each change. It needs imagery with known dates on both sides of the comparison.
:::
:::card{title="Scheduled re-surveys with change notices" icon="calendar"}
Your route or site re-surveyed on a schedule agreed with you, with an email to your team after each survey saying what has changed and where. The cadence depends on when new imagery can be captured, so it is set per project.
:::
:::
::::

::::section{id="drone" eyebrow="Drone surveys" title="Drone detail for the stretches that need it"}
:::::columns{split="2-1" align="center"}
::::col
Satellite screening narrows the search; drone flights over the flagged stretches then capture the detail your team needs: a georeferenced orthophoto, surface and terrain elevation models, a point cloud and a processing quality report. Drone work is carried out by [Afridrone](https://afridr.one/), AfriScan's sister drone-services brand, subject to the permits and authorisations each job requires. Afridrone is working towards the operator approvals needed in each country, and every drone job is planned around the permits that particular flight needs.
::::
::::col
:::callout{tone="legal" title="Permits are part of the plan"}
Every country regulates commercial drone flights differently, and some also control what may be done with the images. We plan the permits into the schedule rather than assuming them.
:::
::::
:::::
::::

::::section{id="limits" tone="alt" eyebrow="Honest scope" title="What results can and cannot tell you"}
:::::columns{split="1-1"}
::::col
### What a survey shows
:::checklist
- Where structures stand, their footprint on the imagery and their distance to the line or boundary
- What changed between two dated surveys, confirmed by a reviewer
- Which stretches of a route carry the most structures
- The imagery used and, where the source gives it, its capture date
:::
::::
::::col
### What it does not show
:::checklist{tone="no"}
- Who lives in or owns a structure, or what it is used for
- Whether a structure is authorised, or who is eligible for compensation
- The condition or integrity of your pipeline or line
- Anything hidden under dense canopy, or underground
:::
::::
:::::

Results support your census, your surveyors and your engineers. They do not replace them. Read the full [methodology](/methodology), the [imagery and data sources](/imagery) and [how we work](/how-we-work).
::::

---
key: data-sources
title: Open Data Sources, Licences & Credits | AfriScan
description: The open datasets and models behind AfriScan registers, the licence each one carries, the credit we give, and how basemaps and purchased imagery are treated.
h1: Data sources and credits
crumb: Data sources & credits
section: how
eyebrow: Imagery and data
lead: "AfriScan registers are built on imagery chosen for each job and on open datasets published by others. This page lists every open source we draw on, the licence it carries and the credit we give, so your GIS, legal and procurement teams can see exactly what sits inside a deliverable."
og:
  headline: Data sources and credits
  subline: Open datasets, their licences and how we credit them
faq:
  - q: Can we reuse the GIS layers you deliver in our own maps and reports?
    a: Yes. The layers are yours to use in your own work. Where a layer contains open data, its licence travels with it, so keep the credit line that comes with the layer. For OpenStreetMap-derived layers, the ODbL also has share-alike terms for databases you publish; your GIS team can check them at the ODbL link in the table above.
  - q: Do you deliver the basemap imagery shown in web maps?
    a: No. Map-service basemaps carry no capture date and their terms do not make them survey imagery. We use them only to screen and plan, and never as delivered imagery or as the basis of a dated record.
  - q: Which satellite operators do you buy from?
    a: We source very-high-resolution scenes from commercial archives and request new captures from commercial operators, choosing per job by date, season and cloud. The report names the imagery used, and the proposal sets out its licence terms for your use.
  - q: Why credit data that is free to use?
    a: Because the licences require it, and because it tells the reader where each part of a register came from. A footprint from an open dataset reflects the imagery that dataset used, which can be years older than your survey. The credit is part of reading the result correctly.
cta:
  title: Need to know which data a survey will use?
  text: Every proposal names the imagery and the open datasets the work will use, with their dates where the source gives them and their licence terms. Send the route or site to get one.
  button: Request a proposal
---

::::section{id="principles" eyebrow="How we treat data" title="Three rules"}
:::cards{cols="3"}
:::card{title="Named in every report" icon="file-text"}
Every report names its imagery and, where the source provides it, the capture date.
:::
:::card{title="Credited as each licence asks" icon="scale"}
The open datasets we use are listed below with the credit and licence text each one requires.
:::
:::card{title="Basemaps are not survey imagery" icon="eye"}
Google, Bing and Esri web-map layers are used only to screen and plan. They are never delivered, and never the basis of a dated record.
:::
:::
::::

::::section{id="footprints" tone="alt" eyebrow="Building footprints" title="Open building datasets"}
Footprint datasets give the first set of candidate structures before a reviewer checks them on the survey imagery. Each one reflects the imagery its publisher used, so it is a starting list, not a current count.

| Dataset | Publisher | What it gives us | Licence | Credit we give |
|---|---|---|---|---|
| [Open Buildings](https://sites.research.google/gr/open-buildings/) (version 3, May 2023) | Google Research | Building footprints across Africa, derived from high-resolution imagery | [CC BY 4.0](https://creativecommons.org/licenses/by/4.0/) (also offered under ODbL 1.0) | Google Open Buildings, CC BY 4.0 |
| [Global ML Building Footprints](https://github.com/microsoft/GlobalMLBuildingFootprints) | Microsoft | Building footprints detected from imagery dated 2014 to 2024 | [CDLA-Permissive-2.0](https://cdla.dev/permissive-2-0/) | Microsoft Building Footprints, CDLA-Permissive-2.0 |
| [OpenStreetMap](https://www.openstreetmap.org/copyright) | OpenStreetMap contributors | Buildings, roads and tracks mapped by volunteers | [ODbL 1.0](https://opendatacommons.org/licenses/odbl/1-0/) | © OpenStreetMap contributors, ODbL |

**What to know.** Google's footprints come from detection run in 2023 on imagery that is several years old in places. Microsoft's come from imagery dated between 2014 and 2024. OpenStreetMap coverage depends on what volunteers have mapped, and is patchy in many rural areas. None of them is used for change detection, because each describes a single moment. Published quality figures for these datasets are their publishers' own; we do not quote them as ours.
::::

::::section{id="earth-observation" eyebrow="Earth observation" title="Open satellite and derived data"}
Coarser open satellite data serves the vegetation, land-cover, terrain and fire work in the catalogue. It is used for larger patterns only, never to count individual buildings.

| Dataset | What it gives us | Licence or terms | Credit we give |
|---|---|---|---|
| Copernicus Sentinel-1 and Sentinel-2 | Radar and optical imagery for larger ground changes, vegetation and burnt areas | [Copernicus Sentinel data legal notice](https://sentinels.copernicus.eu/documents/247904/690755/Sentinel_Data_Legal_Notice) (free, full and open) | "Contains modified Copernicus Sentinel data [year]" |
| Copernicus DEM | Elevation for slope, relief and drainage screening | Copernicus DEM licence | The credit text the licence sets, in each report that uses it |
| [ESA WorldCover](https://esa-worldcover.org/en/data-access) | Land-cover classes | CC BY 4.0 | "© ESA WorldCover project [year] / Contains modified Copernicus Sentinel data ([year]) processed by ESA WorldCover consortium" |
| [WorldPop](https://www.worldpop.org/) | Population grids used, with stated assumptions, in household estimates | CC BY 4.0 | WorldPop, CC BY 4.0 |
| [NASA FIRMS](https://firms.modaps.eosdis.nasa.gov/) | Satellite fire hotspots | NASA open data | NASA Fire Information for Resource Management System (FIRMS) |

Radar can show larger ground changes, such as clearing and earthworks, even under rainy-season cloud. Optical satellite and drone imagery cannot see through cloud, and flagged areas are checked on optical imagery or by drone.
::::

::::section{id="models" tone="alt" eyebrow="Open-source models" title="The segmentation models behind the first pass"}
:::::columns{split="1-1"}
::::col
Two open-source building-segmentation models from the humanitarian mapping community run on the survey imagery to find structures the footprint datasets miss or that are newer than them:

- **[RAMP](https://rampml.global/)** (Replicable AI for Microplanning), an open building-segmentation model built for imagery of about 50 cm per pixel or finer.
- **[HOT fAIr](https://www.hotosm.org/en/tools-resources/tech-product-suite/fair/)**, the Humanitarian OpenStreetMap Team's open AI-assisted mapping tool, which builds on the RAMP approach and can be tuned to local imagery.
::::
::::col
:::callout{tone="scope" title="Models propose; people decide"}
Model output is a set of proposals. A reviewer confirms, removes and adds structures on the imagery before anything is delivered, and the register records, for each structure it lists, whether it was proposed automatically or added by the reviewer. How the sources are merged and checked is set out in the [methodology](/methodology#detection).
:::
::::
:::::
::::

::::section{id="imagery" eyebrow="Imagery" title="Purchased, drone and client imagery"}
:::cards{cols="3"}
:::card{title="Commercial satellite scenes" icon="satellite"}
Dated very-high-resolution scenes from commercial archives, and new captures requested from commercial operators. Each report names the scene and its capture date; the proposal sets out the terms under which you may use the imagery we deliver.
:::
:::card{title="Drone orthophotos" icon="drone"}
Flown by [Afridrone](https://afridr.one/), subject to the permits and authorisations each job requires. Some countries also regulate who may receive and publish aerial survey imagery; the proposal sets out what applies. See [drone rules by country](/drone-regulations).
:::
:::card{title="Your own imagery" icon="layers"}
Orthophotos and scenes you send stay yours, and we never publish them without your written permission. Results depend on their resolution, georeferencing and date.
:::
:::
::::

::::section{id="site-images" tone="alt" eyebrow="On this website" title="Images shown on this site"}
- **The T-9 sample.** The T-9 replacement pipeline route in Inhambane, Mozambique, is shown with the route owner's permission. Its reviewer marks appear only as register strip views, drawn by AfriScan from the sample register: no imagery and no coordinates. The route views and the page headers use a Copernicus Sentinel-2 scene of 2 August 2026 (contains modified Copernicus Sentinel data 2026), shown for location; at 10 m per pixel it cannot show individual structures.
- **No map-service imagery.** No Google, Bing or Esri basemap imagery is shown on this site.
- **Schematics.** Diagrams labelled "Schematic" are drawn by AfriScan with invented geometry to explain how a register is read. They are not real routes or sites.
- **No people.** We show no photographs that identify anyone, and no client logos.

The source of every image is stated with the image. For the imagery options behind a survey, see [imagery and data sources](/imagery).
::::

---
key: vegetation-land-cover-fire
template: solution
title: Vegetation, Land Cover & Fire Mapping | AfriScan
description: Where vegetation has been cleared, has regrown or has burnt along servitudes, estates and concessions, compared between dates from Copernicus Sentinel data.
h1: Vegetation, land cover and fire around your assets
eyebrow: Change and environment
lead: Clearing is often the first sign of new settlement or new works, regrowth is the measure of rehabilitation, and fire near a line or a plantation is a risk worth knowing about. We map where vegetation has been cleared, has regrown or has burnt along your servitudes, estates and concessions, compare it between dates, and send regular notices of fire hotspots and forest loss near your assets.
used_in: [mining, power-utilities, agriculture-forestry-nature]
buttons:
  - {label: Request a proposal, intent: proposal}
  - {label: Open data and credits, key: imagery-data, anchor: "#open-data"}
related: [change-detection, encroachment-surveys, drone-surveys]
service:
  name: Vegetation, land cover and fire mapping
  type: Vegetation and land-cover change mapping
  description: Vegetation and land-cover change between dates along servitudes, estates and concessions, regular forest-loss and ground-disturbance notices checked by our team, fire hotspot notices and burnt-area maps, rehabilitation and revegetation tracking, vegetation height in servitudes and a land-use baseline, from Copernicus Sentinel and other open data, credited as each licence requires.
og:
  headline: Vegetation, land cover and fire around your assets
  subline: Clearing, regrowth and burnt areas compared between dates, from open satellite data
cta:
  title: Tell us which land and which changes matter to you.
  text: A servitude, an estate, a concession or a rehabilitation area, and whether you need clearing, regrowth, fire or all three. We reply with a written proposal.
  button: Request a proposal
faq:
  - q: How small a change can you see?
    a: The open satellite data behind this service shows larger patches of change, not individual trees or small plots. Where you need detail, such as vegetation height under a line or a rehabilitated slope, a drone survey adds it, subject to the permits each flight requires.
  - q: Is the fire notice an early-warning service?
    a: No. It is a regular notice of satellite-detected fire hotspots near your lines, pipelines, plantations and sites, drawn from NASA FIRMS, plus maps of burnt areas afterwards. Small, short-lived or cloud-covered fires can be missed, and it is not an emergency or fire-fighting service.
  - q: Does this measure clearance to our conductors?
    a: No. Vegetation height from drone elevation models shows where tall vegetation stands inside a servitude, and open canopy-height data gives the wider picture. Neither is a measured clearance to conductors; your line-maintenance standards stay the reference.
  - q: Does cloud stop the service in the rainy season?
    a: It can delay it. Optical satellite imagery cannot see through cloud, so results wait for clear images. Radar screening can show larger clearing even under rainy-season cloud, and we follow those areas up on optical imagery.
---

::::section{id="what" eyebrow="What it is" title="The vegetation signals around your land"}
:::::columns{split="2-1"}
::::col
Vegetation tells you a lot about what is happening on land you cannot visit every week. A cleared patch at a concession edge often comes before a new homestead or a new field. Regrowth inside a pipeline or power-line servitude builds up maintenance work and fire load. Revegetation on a closed dump or a restored borrow pit is the evidence a closure plan needs. And a burn scar near a line or across a plantation tells you where to look after the dry season.

We map those signals from Copernicus Sentinel satellite data and other open datasets, compare them between dates, and report them against your servitudes, estates, concessions and rehabilitation areas. Global forest-loss and ground-disturbance alert systems add regular notices for estates and concessions, each checked by our team before it reaches you.

It works alongside our structure registers: the same boundaries, the same reports, and the same review.
::::
::::col
:::callout{tone="note" title="Larger patches, not single trees"}
Open satellite data shows larger patches of change. It is not a species survey, a tree count or a clearance measurement, and cloud can delay results. Where detail matters, a drone survey adds it.
:::
::::
:::::
::::

::::section{id="how" tone="alt" eyebrow="How it works" title="Six kinds of vegetation and land-cover work"}
:::cards{cols="3"}
:::card{title="Clearing and regrowth" icon="leaf"}
Maps of where vegetation has been cleared, has regrown or has changed along servitudes, around plantations and estates, and across concessions, compared between dates.
:::
:::card{title="Forest-loss and disturbance notices" icon="alert"}
Regular notices of tree-cover loss and ground disturbance inside or near your estate or concession, drawn from global satellite alert systems and checked by our team.
:::
:::card{title="Fire hotspots and burnt areas" icon="target"}
Notices of satellite-detected fire hotspots near your lines, pipelines, plantations and sites, and maps of the burnt area after the fire.
:::
:::card{title="Rehabilitation and revegetation" icon="history"}
The trend in vegetation cover on rehabilitated mine land, closed borrow pits and restored servitudes, tracked between dates.
:::
:::card{title="Vegetation height in servitudes" icon="power"}
Where tall vegetation stands inside a power-line or pipeline servitude, from drone elevation models and open canopy-height data.
:::
:::card{title="Land-use baseline" icon="grid"}
Built-up land, cropland, bare ground, water and tree cover around your project, combined with our structure register.
:::
:::

Notices are regular, never real-time: they follow new satellite passes and clear imagery, and each is checked before it is sent. The cadence is agreed per project.
::::

::::section{id="deliverables" eyebrow="What you receive" title="Maps and notices, tied to your boundaries"}
:::checklist
- **Change maps** of cleared, regrown and burnt areas, by servitude segment, estate block or concession zone
- **Regular notices** of forest loss, ground disturbance and fire hotspots near your assets, checked by our team
- **Revegetation trends** for rehabilitation areas, compared between dates
- **Vegetation-height layers** for servitudes, where scoped
- **A land-use and land-cover baseline** of your project area
- **A PDF report** in English or Portuguese and **GIS layers** in GeoPackage, GeoJSON, KMZ and Shapefile
:::

### The services behind it

Every proposal names the services it includes. These are the ones this solution draws on.

:::catalogue{services="S11,S12,S13,S15,S30,S36"}
:::

:::callout{tone="note" title="Open data, credited"}
This work uses open data, credited in every report as its licence requires: contains modified Copernicus Sentinel data; fire hotspots from NASA FIRMS; alerts from Global Forest Watch; ESA WorldCover (CC BY 4.0); and the Meta and WRI canopy-height map (CC BY 4.0).
:::
::::

::::section{id="limits" tone="alt" eyebrow="Limits" title="What vegetation mapping is not"}
:::::columns{split="1-1"}
::::col
### What it shows
:::checklist
- Larger patches of clearing, regrowth and burnt ground between dates
- Where fire hotspots were detected near your assets, and the burnt area afterwards
- The trend in vegetation cover on rehabilitated land
:::
::::
::::col
### What it is not
:::checklist{tone="no"}
- Not a species, ecological or agricultural survey
- Not a measured clearance to conductors, or a tree count
- Not an early-warning, emergency or fire-fighting service
- Not a closure certificate or a regulator's sign-off
:::
::::
:::::
::::

::::section{id="who" eyebrow="Who uses it" title="Where vegetation is part of the land question"}
:::cards{cols="3"}
:::card{title="Mining" icon="mine" key="mining#rehabilitation" cta="Rehabilitation and closure"}
Closure and environmental teams tracking revegetation on rehabilitated land, and clearing at concession edges.
:::
:::card{title="Power & utilities" icon="power" key="power-utilities#vegetation-fire" cta="Vegetation and fire"}
Line-maintenance teams planning vegetation work in servitudes, and watching fire near lines after the dry season.
:::
:::card{title="Agriculture, forestry & nature" icon="tree" key="agriculture-forestry-nature#estates" cta="Estates and plantations"}
Estate, plantation and conservation managers tracking clearing, fire and boundary change across large holdings.
:::
:::

Also used by oil and gas operators along pipeline servitudes, ESIA teams preparing land-cover baselines, and renewables developers around their sites.
::::

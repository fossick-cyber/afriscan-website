---
key: custom-detection
template: solution
title: Custom Object Detection on Your Imagery | AfriScan
description: Detection tuned to your imagery and to what matters on your land, from local building styles to tanks, containers and plant, checked on part of your area first.
h1: Detection tuned to your imagery and to what matters on your land
eyebrow: Drone, imagery and custom detection
lead: Every landscape has its own roofs, and every project has its own list of things worth finding. We tune detection to your imagery and to the objects that matter on your land, check it on part of your area before it runs on the rest, and report each object type as its own layer alongside your structures. A person reviews every result.
used_in: [oil-gas, mining, renewables]
buttons:
  - {label: Request a proposal, intent: proposal}
  - {label: How we detect and review, key: methodology, anchor: "#detection"}
related: [imagery, excavation-mapping, drone-surveys]
service:
  name: Custom object detection
  type: Detection tuned to client imagery
  description: A structure detector tuned to the client's own drone or satellite imagery and local building styles, and detection of other visible objects such as tanks, containers, stockpiles, laydown areas and heavy plant on work sites, trained on examples from the client's imagery, checked on part of the area first and reviewed by a person.
og:
  headline: Detection tuned to your imagery
  subline: Structures, tanks, containers, plant and other visible objects, checked on part of your area first
cta:
  title: Tell us what you need to find, and on what imagery.
  text: Describe the objects that matter and send a sample of your imagery. We reply with what can realistically be detected at its resolution, and a written proposal.
  button: Request a proposal
faq:
  - q: How do you know the detector works on our imagery?
    a: We tune it on examples from your own imagery and check it on part of your area before it runs on the rest. You see the results on that part first. Every result on the full area is still reviewed by a person before delivery.
  - q: What about people and vehicles?
    a: We never detect people. Heavy plant such as excavators and trucks is recorded only as equipment visible on work sites and laydown yards on each survey date, never tracked, and never identified to an owner or operator.
  - q: Can you tell us how full a tank is, or what a container holds?
    a: No. We count and map what the imagery shows, such as tanks, containers, stockpiles and laydown areas, and compare them between dates. Contents and fill levels are not visible.
  - q: What does "describe what to find" mean?
    a: You tell us in words what to look for, such as "new fenced plots" or "water-filled pits", and we search your imagery for it, with every result checked by a reviewer. Some objects are too small or too varied to find reliably; we tell you what we find on part of your area before running it on the rest.
  - q: Do you publish accuracy figures for custom detection?
    a: No. Results depend on your imagery, the season and the objects themselves, so a single figure would mislead. The check on part of your area shows you what to expect, and the review catches what the detector misses or gets wrong.
---

::::section{id="what" eyebrow="What it is" title="When generic detection is not enough"}
:::::columns{split="2-1"}
::::col
Open building datasets and general-purpose detection models are a good start, but they were built on someone else's imagery and someone else's buildings. Thatch, mud-brick and zinc roofs, homesteads under trees, and the particular camera and flying height of your own drone programme can all trip them up. And many projects need to find things that are not buildings at all: storage tanks and containers at a site, laydown areas along a corridor, excavators and trucks at earthworks nobody announced.

Custom detection closes both gaps. We tune a structure detector to your imagery and local building styles, or train detection of the other objects that matter to you on examples from your own imagery. Each object type is reported as its own layer in your GIS files and reports, alongside the structure register, so it never inflates the structure counts.

Anything visible at the imagery's resolution is a candidate. Anything smaller, hidden or underground is not.
::::
::::col
:::callout{tone="scope" title="Tuned, checked, then run"}
Every custom detection job is tuned and checked on part of your area before it runs on the rest, and every result is reviewed by a person before delivery. We never detect or track people.
:::
::::
:::::
::::

::::section{id="how" tone="alt" eyebrow="How it works" title="From examples to a reviewed layer"}
:::steps
:::step{title="What to find" icon="target"}
We agree the objects and the imagery. If an object is too small for the imagery's resolution, we say so before any work starts.
:::
:::step{title="Examples" icon="grid"}
Reviewers label examples of each object on your imagery, and we tune or train detection on them.
:::
:::step{title="Check on part of the area" icon="search"}
The tuned detection runs on part of your area first, and you see the results before it runs on the rest.
:::
:::step{title="Run and review" icon="user-check"}
Detection runs across the whole area; a reviewer confirms, removes and adds results, and each object type is delivered as its own layer.
:::
:::

:::cards{cols="3"}
:::card{title="Structures, your way" icon="houses"}
A building detector tuned to your drone or satellite imagery and local building styles, for more complete registers where generic models struggle.
:::
:::card{title="Tanks, containers and stored materials" icon="layers"}
An inventory of tanks, containers, stockpiles and laydown areas on your sites and along your corridors, compared between dates.
:::
:::card{title="Plant and equipment on site" icon="excavation"}
Heavy equipment visible at construction sites, laydown yards and unexplained earthworks, recorded on each survey date as activity on land.
:::
:::
::::

::::section{id="deliverables" eyebrow="What you receive" title="Separate layers, reviewed like everything else"}
:::checklist
- **One layer per object type**, in GeoPackage, GeoJSON, KMZ and Shapefile, alongside the structure register
- **Counts by zone or segment** for each object type, compared between dates where re-surveyed
- **The check on part of your area**, with the examples used, so your team can see what was tuned
- **A PDF report** in English or Portuguese and an interactive map file
:::

### The services behind it

Every proposal names the services it includes. These are the ones this solution draws on.

:::catalogue{services="S19,S20,S24,S25,S26"}
:::
::::

::::section{id="limits" tone="alt" eyebrow="Limits" title="What custom detection is not"}
:::::columns{split="1-1"}
::::col
### What it does
:::checklist
- Finds objects visible at the imagery's resolution, tuned to your imagery
- Reports each object type separately from structures
- Compares objects between dated surveys
:::
::::
::::col
### What it does not do
:::checklist{tone="no"}
- It never detects, counts or tracks people
- It never tracks vehicles or identifies an owner or operator
- It does not measure contents or fill levels
- It does not promise that any object can be found
:::
::::
:::::
::::

::::section{id="who" eyebrow="Who uses it" title="Where the list of things to find is your own"}
:::cards{cols="3"}
:::card{title="Oil & gas" icon="pipeline" key="oil-gas#lng-sites" cta="Plants and field sites"}
Tanks, containers, laydown areas and plant visible at plants, well pads and construction spreads, compared between surveys.
:::
:::card{title="Mining" icon="mine" key="mining#your-imagery" cta="Your drone imagery"}
Detection tuned to the drone imagery your survey teams already fly, and objects of interest around pits, dumps and haul roads.
:::
:::card{title="Renewables" icon="sun" key="renewables#construction" cta="Construction"}
Laydown, plant and earthworks on solar and wind sites, recorded on each survey date.
:::
:::
::::

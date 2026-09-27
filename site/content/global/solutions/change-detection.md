---
key: change-detection
template: solution
title: Change Detection Between Dated Surveys | AfriScan
description: New and removed structures between dated surveys, flagged automatically and confirmed by a reviewer, with before-and-after views and a notice after each survey.
h1: Change detection between dated surveys
eyebrow: Change and environment
lead: One survey tells you what stands on your land. Two dated surveys tell you what has changed. We compare your route or site between dates, flag new and removed structures automatically, and have a reviewer confirm each change with before-and-after views, on a re-survey schedule agreed with you.
used_in: [oil-gas, mining, power-utilities, project-finance-esia, rail-roads, telecom-fibre, government]
buttons:
  - {label: Request a proposal, intent: proposal}
  - {label: How change is checked, key: methodology, anchor: "#change"}
related: [right-of-way-monitoring, encroachment-surveys, imagery]
service:
  name: Change detection and scheduled re-surveys
  type: Change detection between dated surveys
  description: New and removed structures between two or more dated surveys of a route or site, flagged automatically and confirmed by a reviewer with before-and-after views, with scheduled re-surveys, a change notice after each, settlement growth trends, radar screening for larger changes under rainy-season cloud, and construction progress from drone surveys.
og:
  headline: Change detection between dated surveys
  subline: New and removed structures, flagged automatically and confirmed by a reviewer
cta:
  title: Tell us what you need to watch, and how often.
  text: Send the route or site and the dates that matter. We reply with an imagery plan, a realistic re-survey cadence and a written proposal.
  button: Request a proposal
faq:
  - q: How quickly will we know about a new structure?
    a: After the next scheduled survey that captures it. Change detection compares dated imagery, so it is scheduled, never real-time. The cadence is agreed per project and depends on when new imagery of your area is available, on cloud and, for drones, on permits.
  - q: Why does a reviewer have to confirm each change?
    a: Because two images of the same place are never perfectly aligned, and seasons change how roofs, shadows and vegetation look. An automatic flag is a prompt to look; a reviewer compares the before-and-after views and decides whether the change is real.
  - q: Can you compare against the imagery we flew last year?
    a: Yes, if it is georeferenced and dated. Your own drone orthophotos or satellite scenes can be the first date, the second, or both. Results depend on the quality and resolution of each image.
  - q: What about the rainy season?
    a: Optical satellite and drone imagery cannot see through cloud. Radar satellites can show larger ground changes, such as clearing, earthworks and new large buildings, even under rainy-season cloud; we flag those areas and follow up on optical imagery or with a drone check.
  - q: Can you use open building datasets for change?
    a: No. A footprint dataset describes one moment, so it gives the same answer for both dates. Change detection always compares two or more dated images of your area.
---

::::section{id="what" eyebrow="What it is" title="What changed, where, and between which dates"}
:::::columns{split="2-1"}
::::col
Most land problems are changes: a structure that was not there at the last survey, a patch of cleared ground at the concession edge, earthworks beside a pipeline, a village that has doubled along a new access road. The question your team has to answer is not only what stands on the land but what has arrived since a date that matters, whether that is the last patrol, the start of construction or a resettlement cut-off date.

Change detection answers it on imagery with known capture dates. We compare your route or site between two or more dated surveys, flag new and removed structures automatically, and a reviewer confirms each one against the before-and-after views. Unchanged structures carry forward with their IDs, so the register stays a single record from one survey to the next.

Between re-surveys, two wider signals help you plan: how built-up land around your assets has grown year by year, and radar comparisons that show larger ground changes when rainy-season cloud blocks optical imagery.
::::
::::col
:::callout{tone="note" title="Flagged, then confirmed"}
Every change we report has been flagged automatically **and** confirmed by a reviewer. An automatic flag on its own is a prompt, not a finding.
:::
::::
:::::
::::

::::section{id="how" tone="alt" eyebrow="How it works" title="From two dates to a confirmed change list"}
:::steps
:::step{title="Dates and imagery" icon="calendar"}
We agree the dates to compare and the imagery for each: dated satellite scenes, drone surveys or your own georeferenced imagery. We check that each image is good enough to compare before we start.
:::
:::step{title="Flag" icon="compare"}
Structures are mapped on each date and matched between them. Unmatched structures are flagged as new or removed; matched ones carry their IDs forward.
:::
:::step{title="Confirm" icon="user-check"}
A reviewer checks each flag on side-by-side, swipe and multi-date views, removes false changes caused by alignment, shadow or season, and marks any change the automatic pass missed.
:::
:::step{title="Notify" icon="send"}
You receive the change list, maps and layers, and a notice to the contacts you name. The next survey follows on the schedule agreed with you.
:::
:::

:::cards{cols="3"}
:::card{title="Scheduled re-surveys" icon="history"}
Your route or site re-surveyed on a cadence agreed per project, with an email after each survey saying what changed and where.
:::
:::card{title="Construction and earthworks" icon="excavation"}
Dated drone orthophotos and elevation models compared side by side, with maps of where ground has been cut or filled between visits, subject to the permits each flight requires.
:::
:::card{title="Boundaries and perimeters" icon="boundary"}
New structures, cleared ground and new tracks at your boundary between dates, with fence lines and gates checked by a reviewer on drone imagery.
:::
:::
::::

::::section{id="deliverables" eyebrow="What you receive" title="A change list your field team can act on"}
:::checklist
- **A confirmed change list:** each new or removed structure with its ID, coordinates, chainage or zone, and the dates between which it changed
- **Before-and-after views** of each change, from the dated imagery named in the report
- **Change layers** in GeoPackage and GeoJSON: new, removed and unchanged, alongside KMZ and Shapefile
- **An updated register and density table** where the survey is along a route
- **A change notice** by email after each re-survey
- **A PDF report** in English or Portuguese and an interactive map file
:::

### The services behind it

Every proposal names the services it includes. These are the ones this solution draws on.

:::catalogue{services="S08,S09,S10,S14,S16,S17"}
:::
::::

::::section{id="limits" tone="alt" eyebrow="Limits" title="What change detection is not"}
:::::columns{split="1-1"}
::::col
### What it shows
:::checklist
- New and removed structures between two dated images, confirmed by a reviewer
- Construction, cut and fill between dated drone surveys
- Where built-up land has grown around your assets over the years
- Larger ground changes under cloud, from radar, for follow-up
:::
::::
::::col
### What it does not show
:::checklist{tone="no"}
- Anything between the two dates: it is scheduled, never continuous
- A change on a date without imagery, or a guaranteed capture date
- Small changes on radar, which shows larger changes only
- Who made a change, or why
:::
::::
:::::
::::

::::section{id="who" eyebrow="Who uses it" title="Anyone who needs to know what has arrived since a date"}
:::cards{cols="3"}
:::card{title="Oil & gas" icon="pipeline" key="oil-gas#rights-of-way" cta="Pipeline rights of way"}
New structures in pipeline protection zones and around plants between surveys, reported against the baseline register.
:::
:::card{title="Mining" icon="mine" key="mining#concessions" cta="Concession edges"}
Change at concession edges and around haul roads, dumps and tailings facilities, before and during expansions.
:::
:::card{title="Power & utilities" icon="power" key="power-utilities#servitudes" cta="Line servitudes"}
New building under and beside lines, and in bulk-water servitudes, between maintenance cycles.
:::
:::card{title="Project finance & ESIA" icon="clipboard" key="project-finance-esia#independent-monitoring" cta="Dated evidence between visits"}
Lenders and independent monitors comparing a project footprint and its resettlement sites between site visits.
:::
:::card{title="Rail & roads" icon="rail" key="rail-roads#reserves" cta="Road and rail reserves"}
New structures and crossings in road and rail reserves between inspections.
:::
:::card{title="Government & public programmes" icon="building" key="government#municipal-land" cta="Public land"}
Municipal and public landholders tracking building on land held for a purpose.
:::
:::

Also used by telecom operators along [buried fibre routes](key:telecom-fibre#routes), and by lenders' advisers comparing [each monitoring period](key:project-finance-esia#lender-reporting) with the last.
::::

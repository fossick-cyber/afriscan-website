---
key: household-estimates
template: solution
title: Settlement Mapping & Household Estimates | AfriScan
description: Structure layers and estimated households for mini-grid sizing, service planning, consultation and microplanning, with every assumption stated in the report.
h1: Structure layers and estimated households for planning
eyebrow: Baselines and records
lead: "Planning a mini-grid, a consultation programme, a census or a health campaign starts with a simple question: how many households are out there, and where? We map the structures across your area, estimate households from them with the assumptions stated, and deliver layers your planners can load straight into their own systems. The figures are always estimates, never a census."
used_in: [power-utilities, project-finance-esia, government]
buttons:
  - {label: Request a proposal, intent: proposal}
  - {label: How we detect and review, key: methodology, anchor: "#detection"}
related: [resettlement-cut-off-baselines, route-site-selection, encroachment-surveys]
service:
  name: Settlement mapping and household estimates
  type: Structure layers and household estimates
  description: Area-wide structure layers, reviewer-led mapping and structure categories, settlement growth trends, and estimated households and people from structure counts with the assumptions stated, for mini-grid sizing, service and consultation planning, and census and health-campaign microplanning.
og:
  headline: Structure layers and estimated households
  subline: For mini-grids, consultation and microplanning, with every assumption stated
cta:
  title: Tell us the area you are planning for.
  text: A district, a concession, a list of candidate villages or a project footprint, and what the estimate is for. We reply with a written proposal.
  button: Request a proposal
faq:
  - q: Is this a census?
    a: No. We map structures, not people, and estimate households from the structure count using stated assumptions, such as a persons-per-household figure from national census data and the share of structures that are likely to be dwellings. Every report sets out the assumptions so your team can change them. It is not a census and never replaces one.
  - q: Why not just use an open building dataset?
    a: Open datasets are a good start, and we use and credit them, but they reflect the imagery those projects used, which may be years old, and they miss many thatch, mud-brick and zinc roofs. We check them against current imagery and have reviewers mark structures by hand where automatic detection is not enough.
  - q: Can the estimates be used for a grant or programme application?
    a: They can support one, as a stated, repeatable method. Whether they meet a particular funder's data requirements is for the programme to decide, so we recommend checking the requirements first and we can shape the assumptions table to match.
  - q: Can you separate houses from other buildings?
    a: Reviewers can tag what the imagery shows, such as main building, outbuilding, livestock enclosure or under construction, which improves the estimate. Categories describe the imagery; they do not establish how a building is used.
---

::::section{id="what" eyebrow="What it is" title="Where the households are, before anyone visits"}
:::::columns{split="2-1"}
::::col
Energy-access programmes size mini-grids village by village. Consultation teams plan how many meetings a route needs. Census and health-campaign microplanners split districts into workloads. Each of them needs a number of households, and a map of where they are, before the first field visit.

We start from the structures. We map them across your area from current imagery and open building datasets, have reviewers check them and mark by hand where automatic detection is not enough, and tag what the imagery shows where categories help. Then we estimate households and people from the structure count using stated assumptions, and deliver both the layer and the estimate.

Year-by-year trends in built-up land show where settlement is growing, so your plan reflects where people are moving, not only where they were.
::::
::::col
:::callout{tone="scope" title="Always estimated"}
Household and population figures are estimates from structure counts, with the assumptions stated in every report. We map structures, not people. It is not a census, and it does not count or identify anyone.
:::
::::
:::::
::::

::::section{id="how" tone="alt" eyebrow="How it works" title="From structures to an estimate you can check"}
:::steps
:::step{title="Area and purpose" icon="map"}
You send the area, or the candidate villages or sites, and what the estimate is for. The purpose sets the detail: a regional screen, a village list or a household-level layer.
:::
:::step{title="Structure layer" icon="houses"}
We map structures from current imagery and open building datasets, with reviewers confirming them and marking by hand in dense villages or under tree cover.
:::
:::step{title="Categories" icon="grid"}
Where scoped, reviewers tag each structure with what the imagery shows, so outbuildings and enclosures do not inflate the household estimate.
:::
:::step{title="Estimate" icon="clipboard"}
Households and people are estimated from the structure count with stated persons-per-household assumptions from national census data, cross-checked against open population grids.
:::
:::
::::

::::section{id="deliverables" eyebrow="What you receive" title="Layers and estimates your planners can use"}
:::checklist
- **A structure layer** for the whole area, in GeoPackage, GeoJSON, KMZ and Shapefile, ready for your own GIS
- **Estimated households and people** by village, grid cell, candidate site or zone
- **An assumptions table:** the persons-per-household figures and their source, the categories counted as dwellings, and how to change them
- **Settlement growth trends**, as an area-level picture
- **A PDF report** in English or Portuguese where scoped
:::

### The services behind it

Every proposal names the services it includes. These are the ones this solution draws on.

:::catalogue{services="S05,S35,S10,S07,S06"}
:::

Open data used is credited as its licence requires: Google Open Buildings (CC BY 4.0), Microsoft Building Footprints (CDLA-Permissive-2.0), © OpenStreetMap contributors (ODbL) and WorldPop population grids (CC BY 4.0).
::::

::::section{id="limits" tone="alt" eyebrow="Limits" title="What household estimates are not"}
:::::columns{split="1-1"}
::::col
### What they give you
:::checklist
- A structure layer your planners can trust as a starting point, reviewed by a person
- Estimated households and people with every assumption stated
- Where settlement is growing
:::
::::
::::col
### What they are not
:::checklist{tone="no"}
- Not a census, and not a count of people
- Not a resettlement census or asset inventory under IFC Performance Standard 5
- Not a statement of how any building is used, or who lives in it
- Not complete where structures are hidden under dense canopy
:::
::::
:::::
::::

::::section{id="who" eyebrow="Who uses it" title="Planners who need a number before the field visit"}
:::cards{cols="3"}
:::card{title="Power & utilities" icon="power" key="power-utilities#energy-access" cta="Energy access"}
Mini-grid developers and electrification programmes sizing demand and choosing sites village by village.
:::
:::card{title="Project finance & ESIA" icon="clipboard" key="project-finance-esia#esia-rap-baselines" cta="ESIA and RAP baselines"}
ESIA and RAP teams planning consultation and census effort along a route or across a footprint.
:::
:::card{title="Government & public programmes" icon="building" key="government#energy-access-programmes" cta="Public programmes"}
Public agencies and development partners planning services, energy access and census or health-campaign microplanning.
:::
:::

Also used by oil and gas and rail projects planning community engagement along new routes, and by estates planning services for surrounding villages.
::::

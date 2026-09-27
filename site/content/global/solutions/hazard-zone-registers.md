---
key: hazard-zone-registers
template: solution
title: Structures in Hazard & Safety Zones | AfriScan
description: Structures inside the tailings, blast, dam-inundation and plant safety zones your engineers define, each listed with its location and distance to the source.
h1: Structures in your hazard, safety and exclusion zones
eyebrow: Protect corridors and sites
lead: Your engineers have drawn the zones. The question is what stands inside them today, and what has arrived since the last check. We map the structures inside each safety, tailings, blast or inundation zone you define, list each one with its location and distance to the source, and re-survey the zones on a schedule agreed with you.
used_in: [oil-gas, mining, power-utilities]
buttons:
  - {label: Request a proposal, intent: proposal}
  - {label: How we measure structures, key: methodology, anchor: "#distances"}
related: [encroachment-surveys, household-estimates, terrain-flood-post-event]
service:
  name: Structures in hazard and safety zones
  type: Hazard and safety-zone structure register
  description: A register of the structures inside safety, exclusion, tailings, blast, flood and dam-inundation zones defined by the client's engineers, each with its location and distance to the source, with settlement growth trends, scheduled re-surveys, flood-exposure screening, estimated households and reviewer categories.
og:
  headline: Structures in your hazard and safety zones
  subline: A current register of what stands inside the zones your engineers define
cta:
  title: Send us the zones your engineers have drawn.
  text: Safety, exclusion, tailings, blast or inundation zones as polygons or distances from a source. We reply with an imagery plan and a written proposal.
  button: Request a proposal
faq:
  - q: Do you model the dam break, blast or flood?
    a: No. The zones come from your engineers, your consultants or the regulator. We map what stands inside them. Where you have several zones, such as nested inundation extents or blast radii, each gets its own band in the register.
  - q: Can you tell us how many people are at risk?
    a: We report structures in the zone, not people. Where planning needs it, we add estimated households from the structure count with the assumptions stated, and say clearly that it is an estimate, not a count of people.
  - q: Does this show that our facility conforms to a standard?
    a: No. The register is an input to your emergency planning, stakeholder engagement and reporting. Conformance with any tailings, community health and safety or other standard comes from your whole management system, not from a map.
  - q: How often should the zones be re-surveyed?
    a: As often as settlement around them changes, which varies a great deal by site. We agree the cadence with you and add settlement growth trends around the zones, so you can see where pressure is building between surveys.
---

::::section{id="what" eyebrow="What it is" title="A current list of what stands inside each zone"}
:::::columns{split="2-1"}
::::col
Hazard zones are drawn once, carefully, by engineers. Settlement around them keeps moving. A tailings dam-break inundation area, a blast exclusion radius, a safety setback around a gas plant or the land below a dam can each gather new homesteads and other buildings years after the zone was set, and the emergency plan is only as good as its list of what is inside.

We keep that list current. You give us the zones, as polygons or as distances from a source, and we map the structures inside each one from dated imagery, list each with its location and its distance to the source, and tag what the imagery shows where categories are scoped. Re-surveys on an agreed schedule show what has arrived since the last list.

In some countries the zone is set in law. Mozambique's Land Law makes the land within 250 m of dams and reservoirs a partial protection zone, and a 2001 decree set a 200 m safety zone along a gas pipeline corridor, where building needs the operator's consent. Elsewhere it comes from your engineers or your licence conditions.
::::
::::col
:::callout{tone="legal" title="Sources, as of 26 September 2026"}
[Lei n.º 19/97, de 1 de Outubro, art. 8(e)](https://www.pdul.gov.mz/content/download/486/2635/file/Lei%20de%20Terras.pdf): 250 m around dams and reservoirs. [Decreto n.º 36/2001, de 20 de Novembro, arts. 2 and 3](https://faolex.fao.org/docs/pdf/moz50003.pdf): 200 m safety zone and operator consent; whether it continues under the 2026 Petroleum Law is for counsel to confirm. A summary for information, not legal advice.
:::
::::
:::::
::::

::::section{id="how" tone="alt" eyebrow="How it works" title="From your zones to a register"}
:::steps
:::step{title="Zones and sources" icon="shield"}
You send the zones and the source each is measured from: the dam crest, the blast point, the plant boundary, the flood line. Nested zones become bands, from the innermost out.
:::
:::step{title="Structures in each zone" icon="houses"}
We map the structures from dated imagery and open building datasets. A reviewer confirms each one and, where scoped, tags what the imagery shows: main building, outbuilding, livestock enclosure, under construction.
:::
:::step{title="Context around the zones" icon="water"}
Where useful, we add historically flooded and low-lying ground, how built-up land around the zones has grown year by year, and estimated households with the assumptions stated.
:::
:::step{title="Re-surveys" icon="calendar"}
The zones are re-surveyed on a schedule agreed with you, and a notice after each survey says what has arrived, by zone.
:::
:::
::::

::::section{id="deliverables" eyebrow="What you receive" title="A register your emergency planners can use"}
:::checklist
- **A structure register by zone:** each structure's ID, coordinates, zone or band, distance to the source and, where scoped, category
- **Counts per zone and band**, with estimated households where scoped and the assumptions stated
- **Settlement growth** around the zones, as an area-level trend
- **Change after each re-survey:** new and removed structures by zone, confirmed by a reviewer
- **A PDF report** in English or Portuguese, **GIS layers** (GeoPackage, GeoJSON, KMZ, Shapefile) and an interactive map file
:::

### The services behind it

Every proposal names the services it includes. These are the ones this solution draws on.

:::catalogue{services="S03,S10,S09,S28,S35,S07"}
:::
::::

::::section{id="limits" tone="alt" eyebrow="Limits" title="What a hazard-zone register is not"}
:::::columns{split="1-1"}
::::col
### What it gives you
:::checklist
- The structures visible from above inside each zone you define, on the imagery date
- Distance from each structure to the source
- What has arrived in each zone since the last survey
:::
::::
::::col
### What it is not
:::checklist{tone="no"}
- Not a dam-break, blast, flood or consequence model: the zones are yours
- Not a count of people or a measure of harm: estimated households are estimates
- Not a statement that a facility conforms to any standard
- Not an early-warning or emergency service
:::
::::
:::::
::::

::::section{id="who" eyebrow="Who uses it" title="Owners of zones that must stay clear"}
:::cards{cols="3"}
:::card{title="Oil & gas" icon="pipeline" key="oil-gas#lng-sites" cta="Plants and LNG sites"}
HSE and land teams at gas plants, LNG sites, compressor stations and well pads, measuring structures against safety and exclusion zones.
:::
:::card{title="Mining" icon="mine" key="mining#tailings" cta="Tailings and hazard zones"}
Tailings, geotechnical and social-performance teams keeping a current list of structures in dam-break inundation areas and blast zones.
:::
:::card{title="Power & utilities" icon="power" key="power-utilities#water" cta="Dams and reservoirs"}
Owners of dams, reservoirs and bulk-water schemes tracking structures in inundation and reservoir protection zones.
:::
:::

Also used on concessions where [open and water-filled pits are listed against the zones you draw](key:mining#gemstone-concessions), by project finance and ESIA teams checking community health and safety commitments, and by public agencies with flood lines to keep clear.
::::

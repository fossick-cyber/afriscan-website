---
key: za-water-utilities
template: industry
title: Bulk-Water & Municipal Servitude Mapping, SA | AfriScan
description: A current register of structures and ground disturbance along bulk-water mains, reservoirs and municipal servitudes in South Africa, compared between surveys.
h1: Bulk-water and municipal servitudes, a current register instead of an old snapshot
crumb: Bulk water and municipal services
nav_group: industries
nav_order: 25
nav_label: Bulk water & municipal services
nav_blurb: Water mains, reservoirs and municipal servitudes in South Africa
summary: A register of structures along bulk-water mains and municipal servitudes, the history of when they appeared, and re-surveys on a schedule.
icon: water
eyebrow: Power & utilities · South Africa
lead: Where the last full encroachment survey of your mains is years old, nobody can say which stretches have changed. We build a dated register of the structures, cleared ground and excavations along your bulk-water pipelines, around reservoirs and pump stations and across municipal servitudes, use archive imagery to show roughly when each appeared, and re-survey on a schedule you set. A person reviews every result.
buttons:
  - {label: Send us your pipeline network, intent: proposal}
  - {label: See sample outputs, key: results}
service:
  name: Bulk-water and municipal servitude encroachment survey, South Africa
  type: Servitude encroachment survey, imagery history and scheduled re-surveys
  description: Structures, cleared ground and excavations along bulk-water pipelines, around reservoirs and pump stations and across municipal servitudes in South Africa, each measured to the pipe route, dated from archive imagery where it exists, compared between surveys and reviewed by a person, with PDF reports and GIS layers.
og:
  headline: Bulk-water servitudes in South Africa, mapped from the air
  subline: A dated register of structures along your mains, and what changed since the last survey
related: [change-detection, imagery-history-due-diligence, hazard-zone-registers]
faq:
  - q: Our pipeline network is thousands of kilometres long. Where do you start?
    a: With the whole network at screening level, from satellite imagery and open building data, so every stretch gets a first count and an encroachment-density rating. The report then ranks the stretches, and the detailed review, drone checks or field visits go where the counts are highest or the change is fastest. Very long networks are processed in sections, and the proposal says how.
  - q: Can you tell us when a structure first appeared on the servitude?
    a: Roughly, where archive imagery exists. We compare dated satellite scenes of the stretch and report the first scene in which each structure is visible and the last in which it was not. How far back that goes, and how often, depends on the archive for your area, which we check before we commit.
  - q: Do you need the servitude diagrams, or just the pipe route?
    a: Either works. With the pipe route and the servitude widths, we buffer the route; with the servitude polygons from your GIS, we measure inside them. If widths differ from one registration to the next, send them per section and the register uses the right width for each stretch.
  - q: Can you see our valve chambers, air valves or leaks?
    a: No. We map what is visible at the surface around the line, such as structures, cleared ground, fresh digging, spoil and new tracks. Chambers, valves and the pipe itself are yours to inspect, and water losses are outside what imagery can show.
  - q: Can a municipality buy this?
    a: Yes, through its normal supply-chain process under the Municipal Finance Management Act and the Central Supplier Database. Our [procurement notes](/za/procurement) set out what applies and what to ask any supplier for.
  - q: Will the register be used against the people living on the servitude?
    a: That is your decision, and POPIA shapes it. Our registers describe structures, not people, and carry no names. If you plan to use one as evidence about unlawful occupation, POPIA's prior-authorisation rule (s57) may apply to your organisation, so take advice first. See [POPIA and aerial imagery](/za/popia).
cta:
  title: Close the gap between your last encroachment survey and today.
  text: "Send the pipeline network or the servitude polygons (Shapefile, GeoPackage, KML or GeoJSON), the widths that apply and the year of your last full survey. We reply with a scope, an imagery plan and a written proposal."
  button: Request a proposal
  secondary: See sample outputs
  secondary_href: /results
---

::::section{id="problem" eyebrow="The problem, in your words" title="Long mains, growing metros, and a picture that is years old"}
:::cards{cols="2"}
:::card{title="“Our last full encroachment survey is years old.”" icon="history"}
A baseline is a snapshot. Every year after it, new structures go up and old ones come down, and the register slowly stops describing the ground. Nobody can say which stretches have changed without looking again.
:::
:::card{title="“We find out when a crew or a burst finds it.”" icon="alert"}
A structure on top of a main is found when a maintenance team cannot reach a valve, or when the pipe fails under a foundation. By then it is occupied, and the conversation is harder.
:::
:::card{title="“The metro keeps growing towards our servitudes.”" icon="houses"}
Bulk mains laid across open land decades ago now run beside townships, industrial parks and new estates. The pressure is highest where the city is growing fastest, and that moves year by year.
:::
:::card{title="“A wayleave request arrived and nobody has checked the ground.”" icon="clipboard"}
Other services want to cross or share your servitude: roads, power lines, fibre, housing projects. A current picture of what already stands there makes the wayleave conversation factual.
:::
:::

### Who this is for

:::chips
- Servitude and land managers
- Property and encroachment committees
- Asset and GIS teams
- Pipeline maintenance planners
- Dam-safety and reservoir engineers
- Municipal water, sanitation and electricity departments
- Engineering consultancies working for the utility
:::
::::

::::section{id="context" tone="alt" eyebrow="The South African picture" title="What a bulk-water utility's own paper says" lead="A 2026 paper by an author at a large bulk-water utility sets out the problem more clearly than any sales brochure. It is the author's account, not an official statement by the utility."}
:::::columns{split="2-1"}
::::col
The paper, presented at the [FIG Congress 2026](https://fig.net/resources/proceedings/fig_proceedings/fig2026/papers/ts05h/TS05H_singh_13637.pdf), describes a network of about 3,660 km of pipelines, with servitude encroachment ranked second on the utility's strategic risk register. A 2016 baseline by SANSA identified 22 informal encroachments; after that there is a gap in encroachment data from 2017 to 2025, and in December 2025 a consultant was appointed for six months to map the servitudes. The paper also describes the utility's engagement with other servitude holders, Eskom and Transnet among them.

Its central point applies to every water utility and municipality: once a structure is occupied, resolving an unlawful occupation goes through the courts under the Prevention of Illegal Eviction from and Unlawful Occupation of Land Act (PIE Act), and that takes a long time. The value is in finding a new structure while it is new.
::::
::::col
:::callout{tone="note" title="From a snapshot to a series"}
- **Once:** a baseline register of the whole network
- **Back in time:** archive imagery to show roughly when each structure appeared, where scenes exist
- **Forward:** re-surveys on a schedule you set, with a notice after each
- **Where it matters:** drone checks or field visits on the stretches that changed
:::
::::
:::::
::::

::::section{id="register" eyebrow="What we map along the main" title="A register of what stands on the servitude, then what changed"}
:::::columns{split="2-1"}
::::col
We buffer the pipe route you supply, or work inside your servitude polygons, and list each structure the review confirms with its distance to the pipe, the band it falls in, its chainage and its coordinates. Each 500 m stretch is rated for **encroachment density**, high, medium or low, by a count rule that tells your servitude team where to go first. It is not a pipe-condition or safety rating.

Beyond structures, we flag what else is visible at the surface and matters to a buried main: fresh digging and spoil on or near the servitude, new tracks and crossings, cleared ground that often comes before building, and stockpiles or containers standing on the line.

After the baseline, the network is re-surveyed on the schedule you agree with us, by section. New and removed structures between dated surveys are flagged automatically and confirmed by a reviewer, with before-and-after views, and your team receives a notice after each survey saying what changed and where.
::::
::::col
### What you receive

:::checklist
- A structure register: ID, distance to the pipe, band, chainage and coordinates
- A density rating for each 500 m stretch, and a ranked list of stretches
- Fresh excavation, spoil, cleared ground and new tracks near the line, flagged for checking
- For each structure, roughly when it appeared, where archive imagery exists
- New and removed structures between surveys, with before-and-after views
- A PDF report and GeoPackage, GeoJSON, KMZ and Shapefile layers
:::
::::
:::::

:::solutions{keys="right-of-way-monitoring,change-detection,imagery-history-due-diligence" cols="3"}
:::
::::

::::section{id="reservoirs" tone="alt" eyebrow="Reservoirs, dams and works" title="Around reservoirs, pump stations and treatment works"}
:::::columns{split="1-1"}
::::col
Around a reservoir, pump station or treatment works the line becomes a boundary, and the question becomes a ring: what stands inside the site, what stands in a band around it, and how that band is changing. We map all three from the boundary file you supply.

For dams and reservoirs, your engineers may already hold inundation or safety zones. We list the structures inside the zones they supply, each with its location and distance to the source, and show how settlement inside them changes between surveys. We do not model dam breaks or floods: the zones come from your engineers.
::::
::::col
:::callout{tone="scope" title="Security at your sites"}
Pump stations, reservoirs and works can be strategic installations. Our deliverables leave the security measures at your sites out of every image and report, and drone flights near them need SACAA notice on form CA 101-20 and your security team's permission. See [drone law in South Africa](/za/drone-regulations#sensitive-sites).
:::
::::
:::::
::::

::::section{id="municipal" eyebrow="Municipal services" title="Servitudes a metro or municipality holds"}
Municipal water, sanitation and electricity departments hold servitudes over thousands of erven: bulk and reticulation mains, outfall sewers, stormwater channels and overhead lines. The same register works for each, measured to the line or inside the servitude polygon, with the imagery date on every record.

For municipalities, a structure register is a planning and service-delivery layer as much as an encroachment one: it shows where the network runs under homes that will need access, relocation of the service or a wayleave, and it helps the land-use, human-settlements and infrastructure departments work from the same map. Municipalities buy through their supply-chain processes under the Municipal Finance Management Act, with suppliers on the Central Supplier Database ([procurement notes](/za/procurement)).
::::

::::section{id="imagery" tone="alt" eyebrow="The imagery" title="Satellite first, your drone imagery where you fly, archive to fill the gap"}
:::cards{cols="3"}
:::card{title="Satellite screening" icon="satellite"}
Dated very-high-resolution satellite scenes and open building datasets give every stretch of the network a first count, with no site visit and no drone flight.
:::
:::card{title="Archive imagery" icon="history"}
Where archive scenes exist for your area, we show roughly when structures appeared between your last survey and today. Coverage varies by place and year; we tell you what exists before you commit.
:::
:::card{title="Your own drone imagery" icon="drone"}
If your teams already fly sections of the network, we run the same analysis on your orthophotos. Drone surveys by us are subject to the approvals and security clearances each job requires.
:::
:::
::::

::::section{id="how" eyebrow="How it works" title="From network file to reviewed register"}
:::steps
:::step{title="Scope"}
You send the pipe network or the servitude polygons, the widths that apply, the year of your last full survey and what the record is for. We agree the sections, the imagery and the deliverables in a written proposal.
:::
:::step{title="Screen the whole network"}
Structures are mapped along every section from dated satellite imagery, open building datasets or your own imagery, and each 500 m stretch is rated.
:::
:::step{title="Go back, then look closer"}
Archive scenes show roughly when structures appeared. Drone checks or your field teams cover the stretches that need detail.
:::
:::step{title="Review, deliver and repeat"}
A person checks every result before the report and GIS layers go out. Re-surveys follow the schedule you set, with a notice after each.
:::
:::

More on [how it works](/features), the [methodology](/methodology) and [imagery and data sources](/imagery).
::::

::::section{id="scope" tone="alt" eyebrow="Honest scope" title="What we do, and what we don't"}
:::::columns{split="1-1"}
::::col
### What we do

:::checklist
- Map structures, cleared ground, fresh excavations and tracks visible from the air
- Measure each structure to the pipe route or inside the servitude you supply
- Show roughly when structures appeared, where archive imagery exists
- Flag change between dated surveys, confirmed by a reviewer
:::
::::
::::col
### What we don't do

:::checklist{tone="no"}
- Inspect the pipe, detect leaks or measure water losses
- Identify, count or follow people
- Replace a professional land surveyor or a Surveyor-General diagram
- Decide whether a structure is lawful, or who must move
:::
::::
:::::
::::

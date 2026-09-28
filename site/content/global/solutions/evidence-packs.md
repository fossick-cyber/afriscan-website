---
key: evidence-packs
template: solution
title: Evidence Packs for Legal & Community Processes | AfriScan
description: Dated imagery, structure registers and maps packaged with file fingerprints and an independent timestamp, and location checks for single grievances and claims.
h1: Evidence packs and location checks for legal and community processes
eyebrow: Baselines and records
lead: When a record may be questioned months or years later, it helps to show that it has not changed since the day it was made. We package dated imagery, structure registers and maps with a fingerprint of each file and an independent timestamp, and for a single grievance we assemble every dated image of that location. They support legal and community processes. They do not decide them.
used_in: [oil-gas, mining, project-finance-esia]
buttons:
  - {label: Request a proposal, intent: proposal}
  - {label: Imagery and dates, key: imagery-data}
related: [resettlement-cut-off-baselines, imagery-history-due-diligence, change-detection]
service:
  name: Evidence packs and grievance location checks
  type: Dated evidence packs for legal and community processes
  description: Dated imagery, structure registers and maps packaged with a SHA-256 fingerprint of each file and an independent timestamp, location sheets of every dated image covering a single point for grievances and claims, imagery history, cut-off-date registers and very-high-resolution archive scenes, supporting legal and community processes.
og:
  headline: Evidence packs and location checks
  subline: Dated records, sealed with file hashes and an independent timestamp, that support legal and community processes
cta:
  title: Tell us what the record has to show, and to whom.
  text: A cut-off date, a handover, a contested location or a grievance log. We reply with the imagery that exists for your dates and a written proposal.
  button: Request a proposal
faq:
  - q: Will a court or tribunal accept an evidence pack?
    a: Whether a court or tribunal accepts a pack is for your lawyers; it supports legal and community processes. The fingerprints and timestamp show that the files existed at that time and have not changed since; they do not prove how or when the imagery was captured, which the pack records separately from each image's source.
  - q: What does the timestamp actually prove?
    a: That the manifest of file fingerprints existed at the time stated by an independent timestamping authority, and therefore that each file in the pack is unchanged since then. The capture date of each image comes from its source, such as the drone survey log or the satellite scene metadata, and is recorded separately.
  - q: Can a location check settle a grievance?
    a: No. It shows your grievance officer what each dated image shows at that point, and how that compares with the cut-off register. It supports the investigation; the decision stays with your grievance mechanism.
  - q: Can you build a pack on map-service basemaps?
    a: "No. Basemaps carry no capture date, so they cannot show what was on the ground on a particular day. Evidence packs use dated imagery only: drone surveys, dated satellite scenes or your own imagery with a known date."
  - q: Who sees the pack?
    a: Only the contacts you name. We never publish a client's evidence, registers or imagery without your written permission, and what you share with authorities or communities is your decision.
---

::::section{id="what" eyebrow="What it is" title="A record that can show when it was made"}
:::::columns{split="2-1"}
::::col
Land questions often come back long after the fieldwork. A community claims a structure predates the cut-off date. A new owner of an onshore asset needs to show what was around the wells at handover. A regulator asks when clearing began. By then the question is not only what the imagery shows but whether the record has been altered since.

An evidence pack answers the second question. We assemble the dated imagery, the structure register, the maps and the method notes for the date that matters, compute a SHA-256 fingerprint of each file, and have the manifest timestamped by an independent timestamping authority. Anyone can later recompute the fingerprints and check them against the manifest and its timestamp.

For a single grievance or claim, a location check gathers every dated image that covers the point, in date order, alongside the cut-off register where one exists, so your team can see what was there on each date.
::::
::::col
:::callout{tone="scope" title="What the timestamp does, and doesn't, show"}
It shows the files existed, unchanged, at the timestamped time. It does not show how or when the imagery was captured; each image's capture date comes from its source and is recorded in the pack. The pack supports your process. It does not decide a claim.
:::
::::
:::::
::::

::::section{id="how" tone="alt" eyebrow="How it works" title="Two kinds of record"}
:::cards{cols="2"}
:::card{title="Evidence pack" icon="file-check" eyebrow="For a date that matters"}
1. We agree the date, the area and what the record must show.
2. We source dated imagery: a drone survey on the day, subject to the approvals and security clearances each job requires, or dated very-high-resolution scenes from commercial archives.
3. A reviewer builds the structure register and maps; unclear structures are listed for a ground check.
4. We fingerprint every file, timestamp the manifest independently and deliver the pack with a record of who prepared each part.
:::
:::card{title="Location check" icon="search" eyebrow="For one point or claim"}
1. You send the coordinates, or the grievance reference and a sketch.
2. We gather every dated image in the library, your imagery and the archive that covers the point.
3. The location sheet shows a crop from each date, its source and capture date, and the matching entry in the cut-off register, if there is one.
4. Where the dates matter, the sheet itself can be packaged and timestamped.
:::
:::

Where an archive search is needed, we check what dated imagery exists before we scope the work, so you know in advance which dates the record can and cannot reach.
::::

::::section{id="deliverables" eyebrow="What you receive" title="What an evidence pack contains"}
:::checklist
- **The dated imagery** used, with its source and capture date
- **The structure register and maps** for the date, prepared and reviewed by named roles
- **A method note:** the widths, sources and review steps used
- **A manifest** listing every file with its SHA-256 fingerprint
- **An independent timestamp** on the manifest
- **Location sheets** for single grievances and claims, where scoped
- **A PDF report** in English or Portuguese and GIS layers in GeoPackage, GeoJSON, KMZ and Shapefile
:::

### The services behind it

Every proposal names the services it includes. These are the ones this solution draws on.

:::catalogue{services="S32,S34,S33,S31,S41"}
:::
::::

::::section{id="limits" tone="alt" eyebrow="Limits" title="What an evidence pack is not"}
:::::columns{split="1-1"}
::::col
### What it supports
:::checklist
- Community engagement, grievance handling and legal processes, with a record that can show it is unchanged
- Your IFC Performance Standard 5 cut-off-date records
- Handover baselines and dated land histories
:::
::::
::::col
### What it does not do
:::checklist{tone="no"}
- It does not decide a claim, or attribute a change to anyone
- It does not prove how or when an image was captured
- It does not replace your lawyers' view of what a court or tribunal will accept
- It is never built on undated basemap imagery
:::
::::
:::::
::::

::::section{id="who" eyebrow="Who uses it" title="Teams whose records may be tested later"}
:::cards{cols="3"}
:::card{title="Oil & gas" icon="pipeline" key="oil-gas#evidence" cta="Evidence and grievances"}
Land and community-relations teams recording cut-off dates, asset handovers and contested locations along pipelines and around sites.
:::
:::card{title="Mining" icon="mine" key="mining#evidence" cta="Evidence and grievances"}
Social-performance and tenure teams documenting expansion footprints, grievances and the state of land at closure.
:::
:::card{title="Project finance & ESIA" icon="clipboard" key="project-finance-esia#grievances" cta="Grievances and claims"}
Resettlement managers, grievance officers and lenders' monitors who need a record they can check against later claims.
:::
:::

Also used for [the record of each monitoring period](key:project-finance-esia#lender-reporting) on financed projects and [cut-off-date registers for new lines and interconnectors](key:power-utilities#interconnectors), and by power utilities, rail and road agencies and public landholders with disputed stretches of servitude or reserve.
::::

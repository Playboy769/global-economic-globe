# STMicroelectronics N.V. — SEC Cross-Check Notes (FY2025 20-F + Form 4/3)

CIK: 0000932787. STM is a Foreign Private Issuer (Netherlands-incorporated, Schiphol Airport NL registered office) — files Form 20-F (annual) and Form 6-K (quarterly earnings/other events), not 10-K/10-Q.

## 20-F (FY ended December 31, 2025; filed 2026-02-26, accession 0000932787-26-000009)

### Customer concentration (Item 3.D Risk Factors / Item 4)
- Largest customer, **Apple**, accounted for **17.7%, 14.5%, 12.3%** of total net revenues in **2025, 2024, 2023** respectively — a rising single-customer concentration trend across three consecutive years.
- No aggregate "top 5 customers" percentage is disclosed (unlike ARM's 20-F) — Apple is the only customer whose revenue percentage is broken out under `CustomerConcentrationRiskMember` XBRL tagging.
- Company states it serves "over 200,000 customers" and explicitly says it does "not believe to be dependent on any one customer or group of customers," despite disclosing Apple's rising concentration.
- Major named customers (Item 4, "in alphabetical order"): **Apple, Bosch, Continental, Denso, HP, Mobileye, Samsung, SpaceX, Tesla, Vitesco**. This is the highest-confidence source for automotive/industrial/consumer customer relationships (company's own 20-F business description, not analyst inference).

### AWS strategic engagement — includes an equity-warrant structure not mentioned on the earnings call
- On Feb 9, 2026, STM announced an "expanded strategic collaboration with Amazon Web Services (AWS) through a multi-year, multi-billion USD commercial engagement" — this is the deal referenced on the earnings call.
- **Not disclosed on the call, but disclosed in the 20-F (Item 10.C Material Contracts):** on Feb 6, 2026, STM issued **warrants to AWS for the acquisition of up to 24.8 million ST ordinary shares**, vesting in tranches "substantially tied to payments for ST products and services purchased by AWS and its affiliates," exercisable over a **7-year period** from issuance at an **initial exercise price of $28.38**.
- This is a vendor-financing/customer-equity-incentive structure (AWS earns discounted equity as it buys more silicon) — economically similar in spirit to a large customer's capacity-reservation deposit (as seen in Micron's SCA framework), but structured as dilutive share warrants rather than cash deposits. High confidence (verbatim from Item 10.C and Exhibit 4.1).

### NXP MEMS sensor business acquisition
- First announced July 24, 2025; closing announced Feb 2, 2026 — consistent with the "completed in February" timing stated on the Q1 call.
- 20-F describes the acquired business as focused on "automotive safety and non-safety products and sensors for industrial applications."

### Manufacturing footprint / supply chain (Item 4.D)
- STM **owns all of its manufacturing facilities**, but land underlying certain sites (Muar, Malaysia; Shenzhen, China; Kirkop, Malta; Toa Payoh and Ang Mo Kio, Singapore) is leased.
- Total front-end (fab) capacity as of Dec 31, 2025: **~140,000 wafer starts per week (200mm equivalent)**.
- **Fab 3 is shared with Tower Semiconductor.**
- **Joint 300mm fab with GlobalFoundries Inc. in Crolles, France** — EU-Commission-approved, projected cost **€7.5 billion** (capex + maintenance + ancillary), with up to **€2.9 billion** in French State support under the EU Chips Act.
- **Shenzhen, China front-end facility jointly owned with Shenzhen SEG Hi-Tech Industrial Company Limited** (a subsidiary of Shenzhen Electronics Group) — this is the most likely SEC-disclosed counterpart to the "China for China" wafer strategy referenced on the call (the call names local partner "Wayon" for STM32 wafers made in China; the 20-F's disclosed JV partner of record is Shenzhen SEG Hi-Tech, not "Wayon" — the two names may refer to related or overlapping arrangements, but the 20-F does not use the name "Wayon" anywhere, so this link is not fully confirmed and is flagged low-confidence).
- **Sanan-STMicroelectronics JV** (the SiC joint venture referenced on the call, based in Chongqing) is consolidated by STM as a **Variable Interest Entity (VIE), with STM as primary beneficiary** — meaning STM controls and fully consolidates the JV's assets/liabilities despite not owning 100% of it.
- **Silicon carbide substrates plant in Catania, Italy**: Italian State aid of a **€2 billion direct grant** to support STM's planned **€5 billion** capital investment (approved under EU State Aid rules, May 31, 2024).
- General risk-factor language: STM "currently use[s] external silicon foundries and back-end subcontractors for a portion of our manufacturing activities" — generic dependency language, **no single external foundry or subcontractor is named as sole-sourced** in the risk-factor text itself (foundry names above come from the facilities table and Other Developments section, not the risk factors).

### R&D intensity
- Approximately **19.25% of employees** work in R&D; STM spent approximately **17.3% of net revenues on R&D in 2025**.

### FY2025 full-year margin context (for reference — Q1 FY2026 figures used elsewhere in this report come from the earnings call transcript only)
- FY2025 gross margin: 33.9%, down 540bps from 39.3% in FY2024, mainly due to lower manufacturing efficiencies, sales price/mix, lower capacity reservation fees, negative currency effect, and higher unused capacity charges.
- Q4 2025 gross margin: 35.2%, 20bps above guidance midpoint on better product mix.

### Confidence grading used in Analysis.html supply-chain tab
- **High confidence (verbatim from 20-F)**: Apple = 17.7%/14.5%/12.3% of revenue (2025/2024/2023); named major customers list; AWS warrant terms (24.8M shares, $28.38, 7-year); GlobalFoundries Crolles JV terms; Tower Semiconductor Fab 3 sharing; Sanan JV VIE consolidation; Catania SiC grant; ~140,000 wafer starts/week; R&D % of revenue/headcount.
- **Medium confidence (generic risk-factor language, no named single-source)**: "external silicon foundries and back-end subcontractors" dependency.
- **Low confidence / unconfirmed link**: whether the 20-F's "Shenzhen SEG Hi-Tech Industrial Company Limited" JV is the same arrangement as the "Wayon" China wafer partner named on the earnings call — the two names do not co-occur in either source document.

## Form 4 / Form 3 insider-trading signal — none available (structural, not a data gap)

A full-text search of STM's SEC filing history (CIK 0000932787) turns up only **one Form 3 and two Form 4 filings, all dated 2003** — no insider Form 4 activity around this Q1 FY2026 earnings call exists to analyze.

**Read**: this is not a missing-data problem to flag as "unverified" — it is structural. As a Dutch Foreign Private Issuer, STM's officers and directors are **generally exempt from Section 16 short-swing-profit and beneficial-ownership reporting requirements** (Exchange Act Section 3(a)(12) / Rule 3a12-3), unlike ARM Holdings (also an FPI) whose officers do voluntarily file Form 4s. Unlike the Micron/ARM playbooks in this framework, there is no insider-trading signal available for STM's risk matrix; this should be stated as a structural absence rather than treated as an open due-diligence question.

# Lease Wizard — Specification

A local tool for generating Washington State residential leases (and condition reports) for single-family rental properties. Designed to compress a recurring 30–60 minute task (re-keying data into Zillow Rental Manager + adding renewal language and addenda) into a sub-two-minute wizard run.

## Why this exists

Zillow Rental Manager generates initial leases but doesn't handle renewals well. Each renewal cycle requires re-entering all the same data (property details, tenant info, rent), generating the lease, downloading the PDF, modifying it to add renewal language and custom addenda (lawn care, etc.), and re-saving for e-signing. This tool eliminates the re-entry: property and landlord profiles are stored once, prior tenancies are reusable as templates for renewals, and the entire packet (lease + addenda + condition report) regenerates from a single tenancy file.

## Scope (in / out)

**In scope:**
- Generate a Washington-compliant residential lease (core + addenda) from JSON inputs.
- Support multiple single-family properties and arbitrary tenants.
- Handle renewals via a "prior tenancy" pre-fill pattern.
- Generate a property condition report driven by a per-property room list.
- Assemble a final packet PDF (lease + mold pamphlet + lead paint pamphlet + trash calendar).
- Streamlit-based wizard UI for the full new-lease and renewal flow.

**Out of scope (for now):**
- Multi-state support (Washington only).
- E-signing integration (handled externally).
- Tenant payment processing (handled via Stessa).
- Move-in checklist generator (data model ready; generator not yet built).

## Architecture

Three layers, intentionally simple:

**Data layer** — JSON files in `data/`. Three profile types: `landlord.json` (one file), `properties/<id>.json` (one per property), `tenancies/<id>.json` (one per signed lease). The data layer is the single source of truth; everything else is a transformation. Real data files are gitignored; anonymized samples live in `data/samples/`.

**Generation layer** — Python scripts in `build_scripts/`. Pure functions of `(landlord, property, tenancy) → docx`. No state, no UI dependencies. Importable as modules and runnable from the command line.

**UI layer** — `app.py` (Streamlit). Reads the data layer, walks the user through five steps, writes a new tenancy JSON, and calls the generation layer to produce the final documents.

```
┌─────────────────────────┐
│ app.py (Streamlit UI)   │
└────────────┬────────────┘
             │
             ▼
┌─────────────────────────┐    ┌────────────────────────┐
│ build_scripts/          │◄───┤ data/                  │
│   generate_lease.py     │    │   landlord.json        │
│   generate_condition_…  │    │   properties/*.json    │
│   assemble_packet.py    │    │   tenancies/*.json     │
└─────────────────────────┘    └────────────────────────┘
             │                           ▲
             ▼                           │
          output/                    static/
          *.docx, *.pdf          pamphlets (appended)
```

## Data model

Three JSON profiles. Documented in detail in `README.md`; summarized here.

`landlord.json` — Contact info, bank for security deposit, default fees, accepted payment methods. Edited rarely.

`properties/<id>.json` — One per rental. Address, jurisdiction (city/county), year_built, lead paint disclosure flag, appliances, utilities (electricity, gas, water source, sewer, trash with provider + pickup day + calendar URL, heating, snow, landscaping, telephone/cable/internet), storage, parking, smoke/CO detector status, furnace filter location, room list for condition report. Edited when a property's setup changes.

`tenancies/<id>.json` — One per signed lease. References a property_id. Lease kind (`new` / `renewal`), term (fixed/month-to-month, dates), tenants list, occupants, rent, security deposit (lump sum or installments, already-held flag for renewals), payment methods override, rental insurance, addenda toggles (pets with full financial terms, parking, lawn care with equipment provider, rules with custom rules list, mold/lead-paint/fire-safety auto-derived).

## Conditional logic

A renewal flips the deposit-due, adds the renewal recital, sets the recital's reference to the prior lease date. Fixed term writes Start + End in Section 1.4; month-to-month writes Start only with notice language. Pets-on adds the Pet Addendum and the cross-reference in Section 1.11; pets-off says "Pets are not permitted" inline and skips the addendum. Same pattern for parking, lawn care, rules. Storage has three states: not available, included in rent, included for an extra fee. Telephone/cable/internet each have three setup states: tenant pays, landlord pays, landlord arranges and charges a fixed fee. Pet financial terms (rent, non-refundable fee, refundable deposit) each have on/off; if any are on, a "Pet Financial Terms" block appears in the Pet Addendum. Rental insurance has on/off plus an optional minimum coverage amount that only renders if set. Lead Paint Disclosure attaches if and only if the property profile says it's required. Mold and Fire Safety Addenda are always-on (Washington-required). Security deposit installment payments render when `payment_schedule` is `installments`; lump sum is the default. Lawn care equipment provider is `landlord` or `tenant`. Payment methods are landlord defaults unless the tenancy specifies an override.

## Washington-specific provisions included

- RCW 59.18.575 domestic violence / sexual assault / harassment early-termination clause (Section 1.4).
- Mold Addendum + EPA pamphlet acknowledgment (always-on).
- Fire Safety Addendum (always-on, references RCW Landlord-Tenant Act smoke detector requirements).
- Lead-Based Paint Hazard Disclosure (federal, conditional on pre-1978 housing).
- Utility-termination prohibition preamble (Section 1.6).
- Casualty Damage section with rent abatement and termination on uninhabitable certification (Section 2.11).
- WA-permitted electronic-only payment under RCW 59.18.063 (via tenancy-level payment override).
- WA-permitted security deposit installment payments under RCW 59.18.610.

## What's done

- Full data model (landlord, properties, tenancies) with sample/anonymous versions for testing.
- Lease generator (`generate_lease.py`) producing 22–25 page WA-compliant leases including all required addenda.
- Condition report generator (`generate_condition_report.py`) producing room-by-room walkthrough forms.
- Packet assembler (`assemble_packet.py`) stitching lease PDF + mold pamphlet + lead paint pamphlet + optional trash calendar PDF.
- Streamlit wizard (`app.py`) with five steps: Property, Tenants, Money, Addenda, Review.
- Pre-fill from prior tenancy for renewals (one click loads everything and bumps the term forward by one year).
- Auto-conversion to PDF via LibreOffice if available; graceful fallback to DOCX-only with download button in UI.

## What's left

Roughly in order of value:

**Move-in checklist generator.** Data is already captured in the property profile (thermostat brand, furnace filter location, trash day, landlord contact). Need a `generate_move_in_checklist.py` producing a property-specific checklist. Mostly applicable to new leases.

**Property profile TODOs.** A few fields in `data/properties/` are still incomplete: CO detector status (both properties), rooms list for the county property, trash pickup day confirmation.

**Editing existing tenancies in the UI.** Currently the wizard only creates new ones; editing an existing tenancy requires hand-editing the JSON file.

**Trash calendar PDF attachment.** `assemble_packet.py` already supports it — just set `utilities.trash.calendar_pdf_path` in the property profile once the PDF is downloaded locally.

**Tests** for the generator scripts.

**Deferred:**
- Multi-landlord / multi-state support.
- Versioning of generated documents.
- Cloud storage / backup of the data folder.
- Apartment-lease review (bedbug disclosure, just-cause enumeration, rent-increase notice formalism). Low priority for single-family rentals.

## PDF generation approach

**Current approach:** The lease is generated as `.docx` via python-docx, then converted to PDF using LibreOffice headless (cross-platform) with fallback to `docx2pdf` (drives Microsoft Word on macOS/Windows via AppleScript). pypdf stitches the static pamphlets onto the back.

**If this becomes burdensome** (e.g. neither LibreOffice nor Word is available), the cleanest migration paths are:

- **ReportLab** — draws to a PDF canvas directly. Requires rewriting `generate_lease.py` to use ReportLab's Platypus layout engine. Most work, most control.
- **WeasyPrint** — renders HTML/CSS to PDF. The lease would be templated in HTML (Jinja2). Adds an HTML layer but removes all external tool dependencies.

Either option keeps pypdf for the pamphlet-stitching step. The data model and JSON profiles would not change.

## Running it

```bash
cd /path/to/lease-wizard
uv venv && uv pip install -r requirements.txt
uv run streamlit run app.py
```

Streamlit opens a browser tab. Walk through the five steps; click Generate on the Review step. Files land in `output/`.

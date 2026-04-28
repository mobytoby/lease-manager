# Lease Wizard — Specification

A small local tool for generating Washington State residential leases (and condition reports) for two single-family rental properties owned by Toby Buckley. Designed to compress a recurring 30–60 minute task (re-keying data into Zillow Rental Manager + adding renewal language and addenda) into a sub-two-minute wizard run.

## Why this exists

Zillow Rental Manager generates initial leases but doesn't handle renewals well. Each renewal cycle requires re-entering all the same data (property details, tenant info, rent), generating the lease, downloading the PDF, modifying it to add renewal language and custom addenda (lawn care, etc.), and re-saving for e-signing. This tool eliminates the re-entry: property and landlord profiles are stored once, prior tenancies are reusable as templates for renewals, and the entire packet (lease + addenda + condition report) regenerates from a single tenancy file.

## Scope (in / out)

**In scope:**
- Generate a Washington-compliant residential lease (core + addenda) from JSON inputs.
- Support two known properties (one Spokane city, one Spokane County) and arbitrary tenants.
- Handle renewals via a "prior tenancy" pre-fill pattern.
- Generate a property condition report driven by a per-property room list.
- Streamlit-based wizard UI for the renewal flow.

**Out of scope (for now):**
- Multi-state support (Washington only).
- Multi-landlord support (single landlord, hardcoded as Toby).
- E-signing integration (user already handles e-signing externally).
- Tenant payment processing (handled via Stessa).
- Auto-attachment of static PDFs (mold pamphlet, lead paint pamphlet, trash calendar) to the final packet — planned but not built.

## Architecture

Three layers, intentionally simple:

**Data layer** — JSON files in `data/`. Three profile types: `landlord.json` (one file), `properties/<id>.json` (one per property), `tenancies/<id>.json` (one per signed lease). The data layer is the single source of truth; everything else is a transformation.

**Generation layer** — Python scripts in `build_scripts/`. Pure functions of `(landlord, property, tenancy) → docx`. No state, no UI dependencies. Importable as modules and runnable from the command line.

**UI layer** — `app.py` (Streamlit). Reads the data layer, walks the user through five steps, writes a new tenancy JSON, and calls the generation layer to produce the final documents. The UI is a thin layer on top of the data and generation layers; everything important is defined below it.

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
└─────────────────────────┘    │   tenancies/*.json     │
             │                 └────────────────────────┘
             ▼
       sample_renders/
         *.docx, *.pdf
```

## Data model

Three JSON profiles. Documented in detail in `README.md`; summarized here.

`landlord.json` — Toby's contact info, bank for security deposit, default fees, accepted payment methods. Edited rarely.

`properties/<id>.json` — One per rental. Address, jurisdiction (city/county), year_built, lead paint disclosure flag, appliances, utilities (electricity, gas, water source, sewer, trash with provider+pickup day+calendar URL, heating, snow, landscaping, telephone/cable/internet), storage, parking, smoke/CO detector status, furnace filter location, room list for condition report. Edited when a property's setup changes.

`tenancies/<id>.json` — One per signed lease. References a property_id. Lease kind (`new` / `renewal`), term (fixed/month-to-month, dates), tenants list, occupants, rent, security deposit (lump sum or installments, already-held flag for renewals), payment methods override (per-tenancy override of landlord defaults), rental insurance, addenda toggles (pets with full financial terms, parking, lawn care with equipment provider, rules with custom rules list, mold/lead-paint/fire-safety auto-derived).

## Conditional logic

Most logic in the generator is one of:

A renewal flips the deposit-due, adds the renewal recital, sets the recital's reference to the prior lease date. Fixed term writes Start + End in Section 1.4; month-to-month writes Start only with notice language. Pets-on adds the Pet Addendum and the cross-reference in Section 1.11; pets-off says "Pets are not permitted" inline and skips the addendum. Same pattern for parking, lawn care, rules. Storage has three states: not available, included in rent, included for an extra fee. Telephone/cable/internet each have three setup states: tenant pays, landlord pays, landlord arranges and charges a fixed fee. Pet financial terms (rent, non-refundable fee, refundable deposit) each have on/off; if any are on, a "Pet Financial Terms" block appears in the Pet Addendum. Rental insurance has on/off plus an optional minimum coverage amount that only renders if set. Lead Paint Disclosure attaches if and only if the property profile says it's required (typically pre-1978 housing). Mold and Fire Safety Addenda are always-on (Washington-required). Security deposit installment payments render when `payment_schedule` is `installments`; lump sum is the default. Lawn care equipment provider is `landlord` or `tenant`. Payment methods are landlord defaults unless the tenancy specifies an override (e.g., electronic-only).

## Washington-specific provisions included

- RCW 59.18.575 domestic violence / sexual assault / harassment early-termination clause (Section 1.4).
- Mold Addendum + EPA pamphlet acknowledgment (always-on).
- Fire Safety Addendum (always-on, references RCW Landlord-Tenant Act smoke detector requirements).
- Lead-Based Paint Hazard Disclosure (federal, conditional on pre-1978 housing).
- Utility-termination prohibition preamble (Section 1.6).
- Casualty Damage section with rent abatement and termination on uninhabitable certification (Section 2.11).
- WA-permitted electronic-only payment under RCW 59.18.063 (via tenancy-level payment_methods_override).
- WA-permitted security deposit installment payments under RCW 59.18.610.

## What's done (as of this iteration)

- Full data model (landlord, two properties, two tenancies all populated and demonstrating different paths through the wizard).
- Lease generator (`generate_lease.py`) producing 22-25 page WA-compliant leases including all required addenda.
- Condition report generator (`generate_condition_report.py`) producing room-by-room walkthrough forms.
- Streamlit wizard (`app.py`) with five steps: Property, Tenants, Money, Addenda, Review.
- Pre-fill from prior tenancy for renewals (one click loads everything and bumps the term forward by one year).
- Auto-conversion to PDF via LibreOffice if available; graceful fallback to docx-only.
- Sample renders for both properties' upcoming renewals (Bolton 2026, Passanise 2026).

## What's left

Roughly in order of value:

**Move-in checklist generator.** Currently the move-in checklist (PDF in `uploads/` for the 18th Ave property) hasn't been templated. Need a `generate_move_in_checklist.py` that produces a property-specific checklist using the property profile's thermostat brand, furnace filter location, trash day, and landlord contact. Mostly applicable to new leases, not renewals.

**Static-asset packet assembly.** When generating a new lease, the final packet should include the pamphlets and trash calendar:
- Mold pamphlet PDF (always)
- Lead Paint pamphlet PDF (when lead paint disclosure is included)
- Trash calendar PDF (property-specific, for new leases)
The pamphlets are public-domain federal documents; the trash calendar comes from the city or Waste Management. Need a `assemble_packet.py` that takes the lease docx, converts to PDF, and concatenates the static PDFs onto the back.

**Sumac condition report.** Fill in `rooms_for_condition_report` in `data/properties/sumac.json`. Currently empty (TODO).

**Property profile TODOs.** CO detector status (both properties), furnace filter location (Sumac), thermostat brand (Sumac), trash pickup day (Sumac).

**Apartment-lease review (deferred).** A pass through Toby's apartment lease to surface any WA-specific provisions or addenda we're missing (bedbug disclosure, just-cause enumeration, rent-increase notice formalism). Low priority for single-family rentals.

**Polish.** Better validation on the wizard ("you forgot to add a tenant"). Editing existing tenancies (currently the wizard only creates new ones; you'd reload by hand-editing the JSON). A "duplicate this tenancy and bump dates" button as a renewal shortcut without going through the wizard.

**Deferred for v1:**
- Tests for the generator.
- Multi-landlord / multi-state support.
- Versioning of generated documents (so we can regenerate old leases years later with the original generator code).
- Cloud storage / backup of the data folder.

## PDF generation approach and future options

**Current approach:** The lease is generated as a `.docx` via python-docx, then converted to PDF using `docx2pdf`, which drives Microsoft Word on macOS via AppleScript. pypdf then stitches the static pamphlets onto the back. This requires Word to be installed and triggers a one-time macOS Automation permission grant (System Settings → Privacy & Security → Automation → allow the terminal app to control Word). Once granted it does not prompt again.

This is intentionally Mac-only. The SPEC is single-landlord, single-machine — portability is not a current requirement.

**If this ever becomes burdensome** (e.g. macOS tightens AppleScript permissions, Word is no longer available, or the tool needs to run on other platforms), the cleanest migration path is to drop python-docx and generate PDFs directly in pure Python:

- **ReportLab** — mature, battle-tested, draws to a PDF canvas directly. Requires rewriting `generate_lease.py` to use ReportLab's Platypus layout engine instead of python-docx. Most work, most control.
- **WeasyPrint** — renders HTML/CSS to PDF. The lease would be templated in HTML (Jinja2), which many find more maintainable than programmatic document construction. Adds an HTML layer but removes all external dependencies.

Either option keeps pypdf for the final pamphlet-stitching step, which is already clean and portable. The data model and JSON profiles would not change.

## Running it

```bash
cd /Users/mobish/development/src/lease-wizard
pip install -r requirements.txt
streamlit run app.py
```

Streamlit opens a browser tab. Walk through the five steps; click Generate on the Review step. Files land in `sample_renders/`.

## Re-running for next year's renewal

For Bolton's 2027 renewal: open the wizard, pick 18th Ave in Step 1, pick the 2026 renewal as the prior tenancy, click "Pre-fill". Everything's filled in. Adjust any fields that changed (rent, etc.). Generate. Total time: under two minutes.

# Lease Wizard

A Streamlit-based tool for generating Washington State residential lease packets (lease + addenda + property condition report) from JSON data files.

See `SPEC.md` for the full specification: scope, architecture, what's done, what's left.

## Getting started

The app reads three types of JSON profiles from `data/`:

```
data/
├── landlord.json
├── properties/<id>.json
└── tenancies/<id>.json
```

These files are gitignored (they contain real names, addresses, and financial terms). The app has built-in forms for creating and editing all three profile types — the fastest path is:

1. Copy the sample landlord profile as a starting point, then edit it in the UI:
   ```bash
   cp data/samples/landlord.json data/landlord.json
   ```
2. Run the app (`streamlit run app.py`) and open **Landlord Profile** in the sidebar to fill in your details.
3. Open **Properties** to add your rental property (or copy a sample and edit it there).
4. You're ready — start a new lease from the **New Lease** wizard.

Alternatively, skip the copy step entirely and fill everything in from scratch using the UI forms. See the **Data model** section below for field-by-field documentation.

## Quick start

**With uv (recommended):**

```bash
uv venv && uv pip install -r requirements.txt
uv run streamlit run app.py
```

Install uv once with `curl -LsSf https://astral.sh/uv/install.sh | sh` (or `brew install uv`).

**With standard venv:**

```bash
python -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt
streamlit run app.py
```

A browser tab opens with the wizard. Walk through five steps (Property, Tenants, Money, Addenda, Review) and click Generate. Output lands in `output/`.

For renewals: in Step 1, pick the property, pick the prior tenancy from the dropdown, click "Pre-fill from this tenancy." All fields populate and the term bumps forward by one year. Adjust whatever changed (rent, etc.) and generate.

## Generating from the CLI (without the UI)

The generators are also runnable directly. Useful for scripting or when you want to bypass the wizard.

```bash
cd build_scripts
python3 generate_lease.py \
  ../data/landlord.json \
  ../data/properties/my_property.json \
  ../data/tenancies/my_tenancy.json \
  ../output/Lease-MyTenant-2026.docx

python3 generate_condition_report.py \
  ../data/landlord.json \
  ../data/properties/my_property.json \
  ../data/tenancies/my_tenancy.json \
  ../output/Condition-Report-MyTenant-2026.docx
```

## Layout

```
lease-wizard/
├── app.py                               # Streamlit wizard UI
├── requirements.txt
├── build_scripts/
│   ├── lease_helpers.py                 # Shared formatting helpers
│   ├── generate_lease.py                # Lease + addenda generator
│   ├── generate_condition_report.py     # Property condition report generator
│   └── assemble_packet.py               # Stitches lease PDF + static pamphlets
├── data/
│   ├── landlord.json                    # Your landlord profile (gitignored)
│   ├── properties/                      # One file per rental (gitignored)
│   ├── tenancies/                       # One file per signed lease (gitignored)
│   └── samples/                        # Anonymized samples for bootstrapping
│       ├── landlord.json
│       ├── properties/sample_property.json
│       └── tenancies/sample_tenancy.json
├── static/
│   ├── mold-pamphlet.pdf                # Appended to every packet
│   └── lead-paint-pamphlet.pdf          # Appended when pre-1978 housing
└── output/                              # Generated documents (gitignored)
    ├── Lease-*.{docx,pdf}
    ├── Packet-*.pdf
    └── Condition-Report-*.{docx,pdf}
```

## Data model

Three JSON profiles drive everything. Each has a clear role.

### `landlord.json`

Stable per-landlord data (contact info, bank for security deposit, default fees, accepted payment methods). You'll edit this rarely. Notable fields:

- `payment_methods.check_or_money_order` — Washington requires a non-electronic option.
- `payment_methods.electronic` — list; add/remove services here.
- `defaults.late_fee`, `defaults.grace_period_days`, `defaults.smoking_violation_fee`, `defaults.variable_charge_payment_window_days`, `defaults.month_to_month_tenant_notice_days` — applied across all leases unless overridden in the tenancy file.

### `properties/<id>.json`

Stable per-property data (one file per rental). The wizard reads this when you pick a property in Step 1. Notable fields:

- `jurisdiction` — `"city"` or `"county"`. Drives trash service defaults.
- `year_built` and `lead_paint_disclosure_required` — drives whether the Lead Paint Disclosure auto-attaches (required for pre-1978 housing under 42 USC 4852d).
- `appliances` — list of `{name}`. The condition report uses the appliance items; the lease lists them in Section 1.3.
- `utilities.*` — every utility has a block describing who pays, who provides, and (for trash) the pickup day and calendar URL.
- `furnace_filter.location` — used in the move-in checklist (next iteration).
- `storage.*`, `parking.*` — `available`, `included_in_rent`, `monthly_fee`. Drives Section 1.9 and 1.10 of the lease.
- `rooms_for_condition_report` — list of room names. Drives the condition report's per-room tables.
- `smoke_detectors.battery_operated`, `fire_alarm_system` — drives the Fire Safety Addendum text.

### `tenancies/<id>.json`

Per-lease data (one file per signed lease). Notable fields:

- `lease_kind` — `"new"` or `"renewal"`. When `"renewal"`, a renewal recital is added at the top and the security deposit section reflects continuation rather than re-collection.
- `term.type` — `"fixed"` or `"month_to_month"`. Drives Section 1.4 language.
- `tenants` — list. Joint-and-several language is unconditional; signature blocks scale to the list.
- `rent.first_rent_due` — `"at_signing"`, `"at_start"`, or `"first_due_date_after_start"`.
- `security_deposit.already_held` — `true` for renewals where you keep the existing deposit.
- `security_deposit.payment_schedule` — `"lump_sum"` (default) or `"installments"`. When installments, set `installments.count`, `installments.amount`, and `installments.first_due` (`"signing"` or `"start"`).
- `rental_insurance.required` + `minimum_coverage` — on/off plus optional minimum.
- `addenda.*.include` — boolean toggle for each addendum. Pet/parking/lawn-care/rules are user choice; mold and fire safety are always-on for WA; lead paint is property-driven.
- `addenda.pets.pets` — list of `{name, description}`.
- `addenda.pets.pet_rent_monthly`, `non_refundable_pet_fee`, `refundable_pet_deposit`, `insurance_required` — financial + insurance variants.
- `addenda.rules.custom_rules` — extra one-off rules to append to the Rules Addendum.
- `utilities_overrides` — per-tenancy override of landlord payment method defaults (e.g., electronic-only via Stessa).

## Conditional logic — what flips with what

A renewal flips the deposit-due, adds the renewal recital, and sets the recital's reference to the prior lease date. A fixed term writes Start + End in Section 1.4; month-to-month writes Start only with notice language. Pets-on adds the Pet Addendum and the cross-reference in Section 1.11; pets-off says "Pets are not permitted" inline and skips the addendum. Same pattern for parking, lawn care, rules. Storage has three states: not available, included in rent, included for an extra fee. Telephone/cable/internet each have three states: tenant pays, landlord pays, landlord arranges and charges a fixed fee. Pet financial terms each have on/off; if any are on, a "Pet Financial Terms" block appears in the Pet Addendum. Rental insurance has on/off plus an optional minimum coverage amount that only renders if set. Lead Paint Disclosure attaches if and only if the property profile says it's required (typically pre-1978 housing).

The Table of Contents on the cover page is built dynamically from the same toggles, so it always reflects what's actually in the packet.

## Always-included for WA

- Mold Addendum (with mold pamphlet appended to the final packet PDF).
- Fire Safety Addendum (smoke detector + fire alarm language adjusts to the property profile).

## Washington-specific provisions

**Section 1.4 — Domestic violence / sexual assault / harassment early termination** (RCW 59.18.575). Required by Washington law. Always-on.

**Section 1.5 — Manner of Payment** respects per-tenancy overrides. A tenancy file may include a `utilities_overrides` block to restrict accepted payment methods (e.g., electronic-only). Permitted under RCW 59.18.063 if the tenant agrees in writing in the lease.

**Section 1.6 — Utility termination prohibition preamble.** Forbids either party from intentionally cutting off utilities except for repairs or normal occupancy.

**Section 1.7 — Security deposit installment payments.** Supports `payment_schedule: "installments"` with `count`, `amount`, and `first_due`. Permitted under RCW 59.18.610.

**Section 2.9 — Default remedies.** Spells out specific remedies available on default: unpaid rent, holdover rent, future rent through Expiration Date with mitigation obligation, repair costs, attorneys' fees.

**Section 2.11 — Casualty Damage.** Covers uninhabitable events: termination on 30 days' notice with rent abatement, partial habitability with proportional rent reduction, tenant liability for tenant-caused damage.

**Lawn Care Addendum** — `equipment_provider` field: `"landlord"` or `"tenant"`.

## Renewing a tenancy

Open the wizard, pick the property in Step 1, pick the prior tenancy as the base, click "Pre-fill from this tenancy." All fields populate with last year's data and the term bumps forward by one year. Adjust whatever changed (rent, etc.) and generate. Total time: under two minutes.

For a CLI-only renewal: copy the prior tenancy JSON to a new file, update `effective_date`, `prior_lease_effective_date`, `term.start_date`, `term.end_date`, and `rent.monthly_rent`. Re-run `generate_lease.py`. You should never need to touch the property or landlord profiles for a routine renewal.

## What's not yet built

- **Move-in checklist generator.** Data is already in the property profile (thermostat, filter location, trash day, landlord contact); generator comes next.
- **Editing existing tenancies in the UI.** Currently the wizard only creates new ones. Workaround: hand-edit the JSON and re-run.
- **Tests** for the generator scripts.

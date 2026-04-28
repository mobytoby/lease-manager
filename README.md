# Lease Wizard

A small Streamlit-based system for generating residential lease packets (lease + addenda + property condition report) from JSON data, for properties in Washington State.

See `SPEC.md` for the full specification: scope, architecture, what's done, what's left.

## Quick start

```bash
pip install -r requirements.txt
streamlit run app.py
```

A browser tab opens with the wizard. Walk through five steps (Property, Tenants, Money, Addenda, Review) and click Generate. Output lands in `sample_renders/`.

For renewals: in Step 1, pick the property, pick the prior tenancy from the dropdown, click "Pre-fill from this tenancy." All the fields populate and the term bumps forward by one year. Adjust whatever changed (rent, etc.) and generate.

## Generating from the CLI (without the UI)

The generators are also runnable directly. Useful for scripting or when you want to bypass the wizard.

```bash
cd build_scripts
python3 generate_lease.py \
  ../data/landlord.json \
  ../data/properties/18th_ave.json \
  ../data/tenancies/bolton_2026_renewal.json \
  ../sample_renders/Lease-Bolton-Renewal-2026.docx
```

## Layout

```
lease-wizard/
├── data/
│   ├── landlord.json                    # Single landlord (Toby) profile
│   ├── properties/
│   │   ├── 18th_ave.json                # City property (Bolton's)
│   │   └── sumac.json                   # County property (Passanise's)
│   └── tenancies/
│       ├── bolton_2026_renewal.json     # Bolton 2026 renewal
│       └── passanise_2026.json          # Passanise 2026 — installment deposit, electronic-only payment
├── build_scripts/
│   ├── lease_helpers.py                 # Shared formatting helpers
│   ├── generate_lease.py                # Lease + addenda generator
│   └── generate_condition_report.py     # Property condition report generator
├── sample_renders/
│   ├── Lease-Bolton-Renewal-2026.{docx,pdf}
│   ├── Lease-Passanise-2026.{docx,pdf}
│   └── Property-Condition-Report-18th-Ave.{docx,pdf}
└── templates/                           # (reserved for Jinja templates if/when we extract them)
```

## Generating documents

Run from `build_scripts/`:

```bash
python3 generate_lease.py \
  ../data/landlord.json \
  ../data/properties/18th_ave.json \
  ../data/tenancies/bolton_2026_renewal.json \
  ../sample_renders/Lease-Bolton-Renewal-2026.docx

python3 generate_condition_report.py \
  ../data/landlord.json \
  ../data/properties/18th_ave.json \
  ../data/tenancies/bolton_2026_renewal.json \
  ../sample_renders/Property-Condition-Report-18th-Ave.docx
```

## Data model

Three JSON profiles drive everything. Each has a clear role.

### `landlord.json`

Stable per-landlord data (your contact info, bank for security deposit, default fees, accepted payment methods). You'll edit this rarely. Notable fields:

- `payment_methods.check_or_money_order` — Washington requires a non-electronic option.
- `payment_methods.electronic` — list; you can add/remove services here.
- `defaults.late_fee`, `defaults.grace_period_days`, `defaults.smoking_violation_fee`, `defaults.variable_charge_payment_window_days`, `defaults.month_to_month_tenant_notice_days` — applied across all leases unless overridden in the tenancy file.

### `properties/<id>.json`

Stable per-property data (one file per rental). The wizard reads this when you pick a property in step 1. Notable fields:

- `jurisdiction` — `"city"` or `"county"`. Drives trash service defaults.
- `year_built` and `lead_paint_disclosure_required` — drives whether the Lead Paint Disclosure auto-attaches.
- `appliances` — list of `{name, make?, model?, serial?}`. The condition report uses the appliance items; the lease lists them in Section 1.3.
- `utilities.*` — every utility has a small block describing who pays, who provides, and (for trash) the pickup day and calendar URL.
- `utilities.heating.thermostat` — used in the move-in checklist (next iteration).
- `furnace_filter.location` — used in the move-in checklist (next iteration).
- `storage.*`, `parking.*` — `available`, `included_in_rent`, `monthly_fee`. Drives Section 1.9 and 1.10 of the lease.
- `rooms_for_condition_report` — list of room names. Drives the condition report's per-room tables.
- `smoke_detectors.battery_operated`, `fire_alarm_system` — drives the Fire Safety Addendum text.

### `tenancies/<id>.json`

Per-renewal/per-lease data (one file per signed lease). Notable fields:

- `lease_kind` — `"new"` or `"renewal"`. When `"renewal"`, a renewal recital is added at the top and the security deposit section reflects continuation rather than re-collection.
- `term.type` — `"fixed"` or `"month_to_month"`. Drives Section 1.4 language.
- `tenants` — list. Joint-and-several language is unconditional; signature blocks scale to the list.
- `rent.first_rent_due` — `"at_signing"`, `"at_start"`, or `"first_due_date_after_start"`. (Currently the lease assumes the third option; the wizard will surface the others.)
- `security_deposit.already_held` — true for renewals where you keep the existing deposit; flips the "due at signing" line.
- `rental_insurance.required` + `minimum_coverage` — on/off plus optional minimum.
- `addenda.*.include` — boolean toggle for each addendum. Pet/parking/lawn-care/rules are user choice; mold and fire safety are always-on for WA; lead paint is property-driven.
- `addenda.pets.pets` — list of `{name, description}`.
- `addenda.pets.pet_rent_monthly`, `non_refundable_pet_fee`, `refundable_pet_deposit`, `insurance_required` — financial + insurance variants.
- `addenda.rules.custom_rules` — extra one-off rules to append to the Rules Addendum.

## Conditional logic — what flips with what

Most of the if/then logic in the generator follows a small number of rules:

A renewal flips the deposit-due, adds the renewal recital, and sets the recital's reference to the prior lease date. A fixed term writes Start + End in Section 1.4; month-to-month writes Start only with notice language. Pets-on adds the Pet Addendum and the cross-reference in Section 1.11; pets-off says "Pets are not permitted" inline and skips the addendum. Same pattern for parking, lawn care, rules. Storage has three states: not available, included in rent, included for an extra fee. Telephone/cable/internet each have three states: tenant pays, landlord pays, landlord arranges and charges a fixed fee. Pet financial terms each have on/off; if any are on, a "Pet Financial Terms" block appears in the Pet Addendum. Rental insurance has on/off plus an optional minimum coverage amount that only renders if set. Lead Paint Disclosure attaches if and only if the property profile says it's required (typically pre-1978 housing).

The Table of Contents on the cover page is built dynamically from the same toggles, so it always reflects what's actually in the packet.

## Always-included for WA

- Mold Addendum (with mold pamphlet appended at assembly time — the wizard will append the static PDF).
- Fire Safety Addendum (smoke detector + fire alarm language adjusts to the property profile).

## Provisions added in the second iteration (after reviewing the Passanise lease)

The following were added to bring the generator in line with WA practice and the existing Zillow templates we benchmarked against:

**Section 1.4 — Domestic violence / sexual assault / harassment early termination clause** (RCW 59.18.575). Required by Washington law. Always-on. Allows tenant to terminate with 30 days' written notice; with documentation (protection order or police report), tenant is discharged from rent and entitled to deposit return.

**Section 1.5 — Manner of Payment now respects tenancy override.** A tenancy file may include a `payment_methods_override` block to restrict accepted payment methods (e.g., electronic-only via Stessa, with check/money-order disabled). Without an override, the landlord profile defaults are used. This is permitted under RCW 59.18.063 if the tenant agrees in writing in the lease.

**Section 1.6 — Utility termination prohibition preamble.** New first paragraph forbids either party from intentionally cutting off utilities except for repairs or normal occupancy.

**Section 1.7 / 1.1.1 — Security deposit installment payments.** A tenancy may specify `payment_schedule: "installments"` with `count`, `amount`, and `first_due` (`signing` or `start`). Section 1.1.1 (upfront amounts) and Section 1.7 (security deposit terms) both render the installment language correctly. Defaults to `lump_sum` for backward compatibility.

**Section 2.9 — Default remedies expanded.** Spells out the specific remedies available on default (unpaid rent, holdover rent, future rent through Expiration Date with mitigation, repair costs, attorneys' fees), matching the language in the Passanise lease.

**Section 2.11 — Casualty Damage (new section).** Covers what happens if fire, storm, or other casualty makes the Property uninhabitable: termination on 30 days' notice with rent abatement, partial habitability with proportional rent reduction, and tenant liability for damage caused by tenant negligence. Renumbered subsequent sections (2.12 Tenant's Property, 2.13 General, 2.14 Disclosures, 2.15 Execution, 2.16 Contact Information).

**Lawn Care Addendum — equipment provider field.** `addenda.lawn_care.equipment_provider` is `"landlord"` (Bolton) or `"tenant"` (Passanise). Defaults to `"landlord"`.

## Property profile additions in the second iteration

`utilities.water_source` and `utilities.sewer` — each has `type` (`"city"` or `"well"` / `"city"` or `"septic"`) and `provider`. Forward-looking fields for properties that aren't on municipal services. The legacy `water_sewer` block is kept for billing/payment-side info.

`co_detectors` — `installed`, `count`, mirrors `smoke_detectors`. Required by RCW 19.27.530 in WA residential rentals; both current property profiles have a TODO to confirm.

## What's not in this iteration

- Move-in checklist generator — base content captured in the data model already (thermostat, filter location, trash day, gas utility); generator and template come next.
- Static-asset assembly: appending the trash calendar PDF, mold pamphlet PDF, lead paint pamphlet PDF to the final packet.
- A Jinja-templated `.docx` — current generators build the document directly with `python-docx`. Once you've reviewed the output and the structure is stable, we can extract the boilerplate into a single `core_lease.docx` template (with `docxtpl`) so editing prose doesn't mean editing Python.
- The Streamlit wizard UI itself.
- Tests.

## Renewing the Boltons in 12 months

For their 2027 renewal: copy `tenancies/bolton_2026_renewal.json` to `tenancies/bolton_2027_renewal.json`, change `effective_date`, `prior_lease_effective_date`, `term.start_date`, `term.end_date`, and (if applicable) `rent.monthly_rent`. Re-run `generate_lease.py`. That's the whole renewal — you should never touch the property or landlord profiles for a routine renewal.

## Trash calendar URLs

- City of Spokane (city property): the link in `18th_ave.json` is a placeholder pointing to the city's residential trash schedule index page; replace with the current annual residential collection PDF URL when you confirm it.
- Waste Management Northwest (county property): `wmnorthwest.com` link captured in `county_property.json`. Confirm it actually applies to the specific service area and pickup day.

The trash calendar PDF is intended to be a static file appended to the move-in packet (next iteration), not template content.

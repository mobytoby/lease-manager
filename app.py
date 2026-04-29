"""Lease Wizard — Streamlit UI for generating residential leases.

Run from the lease-wizard directory:
    streamlit run app.py

The wizard reads property and landlord profiles from data/, walks through
five steps, writes a tenancy JSON to data/tenancies/, and produces the
assembled lease (and condition report, if the property has a room list)
in output/.
"""

import json
import shutil
import subprocess
import sys
from copy import deepcopy
from datetime import date, timedelta
from pathlib import Path

import streamlit as st

# Make our generator modules importable
ROOT = Path(__file__).parent.resolve()
BUILD_SCRIPTS = ROOT / "build_scripts"
sys.path.insert(0, str(BUILD_SCRIPTS))

from generate_lease import generate as generate_lease  # noqa: E402
from generate_condition_report import generate as generate_condition_report  # noqa: E402
from assemble_packet import assemble as assemble_packet  # noqa: E402

DATA_DIR = ROOT / "data"
PROPERTIES_DIR = DATA_DIR / "properties"
TENANCIES_DIR = DATA_DIR / "tenancies"
OUTPUT_DIR = ROOT / "output"


# ---------------------------------------------------------------------------
# Session state
# ---------------------------------------------------------------------------

def init_session():
    st.session_state.setdefault("step", 1)
    st.session_state.setdefault("tenancy", empty_tenancy())
    st.session_state.setdefault("last_generated", None)


def empty_tenancy():
    return {
        "id": "",
        "property_id": "",
        "lease_kind": "renewal",
        "prior_lease_effective_date": None,
        "effective_date": str(date.today()),
        "term": {
            "type": "fixed",
            "start_date": str(date.today()),
            "end_date": str(date.today() + timedelta(days=365)),
        },
        "tenants": [],
        "additional_occupants": [],
        "rent": {
            "monthly_rent": 0,
            "first_rent_due": "first_due_date_after_start",
            "prorated_first_month": 0,
            "monthly_rent_due_day": 1,
            "grace_period_days": 5,
            "late_fee": 200,
            "nsf_fee": "variable",
        },
        "security_deposit": {
            "amount": 0,
            "already_held": True,
            "additional_due_at_signing": 0,
            "payment_schedule": "lump_sum",
        },
        "rental_insurance": {"required": False, "minimum_coverage": None},
        "addenda": {
            "pets": {
                "include": False, "pets": [],
                "pet_rent_monthly": 0, "non_refundable_pet_fee": 0,
                "refundable_pet_deposit": 0, "insurance_required": False,
            },
            "parking": {"include": True},
            "lawn_care": {"include": True, "equipment_provider": "landlord"},
            "rules": {"include": True, "custom_rules": []},
            "mold_disclosure": {"include": True},
            "lead_paint_disclosure": {"include": False},
            "fire_safety": {"include": True},
        },
    }


def strip_scratch(obj):
    """Remove wizard-only scratch keys (those starting with underscore)
    from the tenancy dict before persisting it."""
    if isinstance(obj, dict):
        return {k: strip_scratch(v) for k, v in obj.items() if not k.startswith("_")}
    if isinstance(obj, list):
        return [strip_scratch(v) for v in obj]
    return obj


# ---------------------------------------------------------------------------
# Data loading
# ---------------------------------------------------------------------------

def load_landlord():
    return json.loads((DATA_DIR / "landlord.json").read_text())


def load_property(prop_id):
    return json.loads((PROPERTIES_DIR / f"{prop_id}.json").read_text())


def list_properties():
    out = []
    for path in sorted(PROPERTIES_DIR.glob("*.json")):
        data = json.loads(path.read_text())
        addr = data["address"]
        label = f"{addr['street']}, {addr['city']}, {addr['state']} {addr['zip']}"
        out.append((data["id"], label))
    return out


def list_tenancies_for(property_id):
    out = []
    for path in sorted(TENANCIES_DIR.glob("*.json")):
        data = json.loads(path.read_text())
        if data.get("property_id") != property_id:
            continue
        tenants = ", ".join(t["name"] for t in data.get("tenants", []))
        term = data.get("term", {})
        label = (
            f"{data['id']} — {tenants} "
            f"({term.get('start_date', '?')} → {term.get('end_date', '?') or 'm-to-m'})"
        )
        out.append((data["id"], label, data))
    return out


# ---------------------------------------------------------------------------
# Step 1: Property & Lease Type
# ---------------------------------------------------------------------------

def step_1_property_and_kind():
    st.header("Step 1 — Property & Lease Type")

    properties = list_properties()
    if not properties:
        st.error("No property profiles found in data/properties/")
        return

    prop_ids = [p[0] for p in properties]
    prop_labels_by_id = {p[0]: p[1] for p in properties}
    current = st.session_state.tenancy["property_id"]
    idx = prop_ids.index(current) if current in prop_ids else 0

    new_prop_id = st.selectbox(
        "Which property?",
        options=prop_ids,
        format_func=lambda i: prop_labels_by_id[i],
        index=idx,
    )
    st.session_state.tenancy["property_id"] = new_prop_id

    lease_kind = st.radio(
        "What kind of lease?",
        options=["renewal", "new"],
        format_func=lambda x: (
            "Renewal of an existing tenancy" if x == "renewal" else "New lease (new tenant)"
        ),
        index=0 if st.session_state.tenancy["lease_kind"] == "renewal" else 1,
        horizontal=True,
    )
    st.session_state.tenancy["lease_kind"] = lease_kind

    if lease_kind == "renewal":
        prior_options = list_tenancies_for(new_prop_id)
        if not prior_options:
            st.warning(
                "No prior tenancies found for this property. "
                "Either pick 'New lease' or add a tenancy file to data/tenancies/."
            )
        else:
            ids = [t[0] for t in prior_options]
            labels = {t[0]: t[1] for t in prior_options}
            current_prior = st.session_state.tenancy.get("_prior_picker", ids[0])
            idx_p = ids.index(current_prior) if current_prior in ids else 0
            picked = st.selectbox(
                "Which prior tenancy is being renewed?",
                options=ids,
                format_func=lambda i: labels[i],
                index=idx_p,
            )
            st.session_state.tenancy["_prior_picker"] = picked
            st.caption(
                "Click the button below to pre-fill the wizard with this tenancy's data — "
                "tenants, addenda, pet list, etc. — and bump the term forward by one year. "
                "You can edit anything in the following steps."
            )
            if st.button("Pre-fill from this tenancy", type="primary"):
                pre_fill_from(prior_options, picked)
                st.rerun()

    st.divider()
    st.session_state.tenancy["id"] = st.text_input(
        "Tenancy file ID (saved as data/tenancies/<id>.json)",
        value=st.session_state.tenancy.get("id", ""),
        help="Auto-suggested when you pre-fill. Free-form otherwise.",
    )


def pre_fill_from(prior_options, picked_id):
    prior = next(t[2] for t in prior_options if t[0] == picked_id)
    new = deepcopy(prior)
    new["lease_kind"] = "renewal"
    prior_eff = prior.get("effective_date") or prior.get("term", {}).get("start_date")
    new["prior_lease_effective_date"] = prior_eff
    new["effective_date"] = str(date.today())

    # Bump term forward by one year, keeping the same span
    try:
        s = date.fromisoformat(prior["term"]["start_date"])
        new_start = s.replace(year=s.year + 1)
        new["term"]["start_date"] = str(new_start)
        if prior["term"].get("end_date"):
            e = date.fromisoformat(prior["term"]["end_date"])
            new["term"]["end_date"] = str(e.replace(year=e.year + 1))
    except (ValueError, KeyError):
        pass

    # New ID like "bolton_2027_renewal"
    seed = picked_id.split("_")[0] if "_" in picked_id else picked_id
    yr = date.fromisoformat(new["term"]["start_date"]).year
    new["id"] = f"{seed}_{yr}_renewal"

    # Renewal: deposit continues
    new["security_deposit"]["already_held"] = True
    new["security_deposit"]["payment_schedule"] = "lump_sum"
    new["security_deposit"]["additional_due_at_signing"] = 0
    new["security_deposit"].pop("installments", None)

    # Strip any private notes from the source
    new = strip_scratch(new)

    # Re-attach the picker so the dropdown stays on this choice
    new["_prior_picker"] = picked_id
    st.session_state.tenancy = new


# ---------------------------------------------------------------------------
# Step 2: Tenants & Occupants
# ---------------------------------------------------------------------------

def step_2_tenants():
    st.header("Step 2 — Tenants & Occupants")

    st.subheader("Tenants")
    st.caption("Everyone who signs the lease. Multiple tenants are jointly and severally liable.")

    tenants = st.session_state.tenancy["tenants"]

    for i, t in enumerate(tenants):
        with st.container(border=True):
            cols = st.columns([5, 1])
            with cols[0]:
                t["name"] = st.text_input("Name", value=t.get("name", ""), key=f"tname_{i}")
            with cols[1]:
                st.markdown("&nbsp;")  # vertical alignment hack
                if st.button("Remove", key=f"rm_tenant_{i}"):
                    tenants.pop(i)
                    st.rerun()
            cols = st.columns(2)
            with cols[0]:
                t["phone"] = st.text_input("Phone", value=t.get("phone", ""), key=f"tphone_{i}")
            with cols[1]:
                t["email"] = st.text_input("Email", value=t.get("email", ""), key=f"temail_{i}")
            t["address_for_notice_pre_start"] = st.text_input(
                "Address for notice (prior to start date)",
                value=t.get("address_for_notice_pre_start", ""),
                key=f"taddr_{i}",
            )

    if st.button("+ Add tenant"):
        tenants.append(
            {"name": "", "phone": "", "email": "", "address_for_notice_pre_start": ""}
        )
        st.rerun()

    st.subheader("Additional Occupants")
    st.caption(
        "Children, dependents, or others who'll live in the property but aren't on the lease. Optional."
    )

    occupants = st.session_state.tenancy["additional_occupants"]
    for i, o in enumerate(occupants):
        cols = st.columns([3, 1, 0.6])
        with cols[0]:
            o["name"] = st.text_input("Name", value=o.get("name", ""), key=f"oname_{i}")
        with cols[1]:
            age_val = "" if o.get("age") in (None, "") else str(o["age"])
            o["age"] = st.text_input("Age", value=age_val, key=f"oage_{i}")
        with cols[2]:
            st.markdown("&nbsp;")
            if st.button("X", key=f"rm_occ_{i}"):
                occupants.pop(i)
                st.rerun()

    if st.button("+ Add occupant"):
        occupants.append({"name": "", "age": ""})
        st.rerun()


# ---------------------------------------------------------------------------
# Step 3: Term, Rent, Deposit
# ---------------------------------------------------------------------------

def step_3_money():
    st.header("Step 3 — Term, Rent & Deposit")

    t = st.session_state.tenancy["term"]
    rent = st.session_state.tenancy["rent"]
    sd = st.session_state.tenancy["security_deposit"]

    st.subheader("Term")
    t["type"] = st.radio(
        "Term type",
        options=["fixed", "month_to_month"],
        index=0 if t["type"] == "fixed" else 1,
        format_func=lambda x: "Fixed" if x == "fixed" else "Month-to-month",
        horizontal=True,
    )
    cols = st.columns(2)
    with cols[0]:
        start = st.date_input(
            "Start date",
            value=date.fromisoformat(t["start_date"]),
            key="term_start",
        )
        t["start_date"] = str(start)
    if t["type"] == "fixed":
        with cols[1]:
            end_default = (
                date.fromisoformat(t["end_date"]) if t.get("end_date")
                else start + timedelta(days=365)
            )
            end = st.date_input("End date", value=end_default, key="term_end")
            t["end_date"] = str(end)
    else:
        t["end_date"] = None

    st.subheader("Rent")
    cols = st.columns(2)
    with cols[0]:
        rent["monthly_rent"] = st.number_input(
            "Monthly rent ($)", value=float(rent["monthly_rent"]),
            step=25.0, min_value=0.0, format="%.2f",
        )
    with cols[1]:
        rent["late_fee"] = st.number_input(
            "Late fee ($)", value=float(rent["late_fee"]),
            step=25.0, min_value=0.0, format="%.2f",
        )
    cols = st.columns(2)
    with cols[0]:
        rent["grace_period_days"] = st.number_input(
            "Grace period (days)", value=int(rent["grace_period_days"]),
            step=1, min_value=0,
        )
    with cols[1]:
        rent["monthly_rent_due_day"] = st.number_input(
            "Rent due on day of month", value=int(rent["monthly_rent_due_day"]),
            step=1, min_value=1, max_value=28,
        )

    st.subheader("Security Deposit")
    sd["amount"] = st.number_input(
        "Total security deposit amount ($)",
        value=float(sd["amount"]), step=100.0, min_value=0.0, format="%.2f",
    )
    sd["already_held"] = st.checkbox(
        "Already held by Landlord (renewal — deposit continues without re-collection)",
        value=sd.get("already_held", False),
    )
    if not sd["already_held"]:
        sched = st.radio(
            "Payment schedule",
            options=["lump_sum", "installments"],
            index=0 if sd.get("payment_schedule", "lump_sum") == "lump_sum" else 1,
            format_func=lambda x: (
                "Lump sum at signing" if x == "lump_sum"
                else "Monthly installments"
            ),
            horizontal=True,
        )
        sd["payment_schedule"] = sched
        if sched == "installments":
            inst = sd.setdefault(
                "installments", {"count": 3, "amount": 0, "first_due": "start"}
            )
            cols = st.columns(3)
            with cols[0]:
                inst["count"] = st.number_input(
                    "Number of installments", value=int(inst.get("count", 3)),
                    step=1, min_value=2,
                )
            with cols[1]:
                inst["amount"] = st.number_input(
                    "Installment amount ($)", value=float(inst.get("amount", 0)),
                    step=25.0, min_value=0.0, format="%.2f",
                )
            with cols[2]:
                inst["first_due"] = st.radio(
                    "First installment due",
                    options=["signing", "start"],
                    index=0 if inst.get("first_due") == "signing" else 1,
                    format_func=lambda x: (
                        "At signing" if x == "signing" else "At Start Date"
                    ),
                )
        else:
            sd.pop("installments", None)


# ---------------------------------------------------------------------------
# Step 4: Addenda
# ---------------------------------------------------------------------------

def step_4_addenda():
    st.header("Step 4 — Addenda")

    a = st.session_state.tenancy["addenda"]
    prop = load_property(st.session_state.tenancy["property_id"])

    st.subheader("Optional addenda")

    # Pets
    a["pets"]["include"] = st.checkbox(
        "Pet Addendum (tenant has pets)", value=a["pets"]["include"]
    )
    if a["pets"]["include"]:
        with st.container(border=True):
            pets = a["pets"]["pets"]
            for i, p in enumerate(pets):
                cols = st.columns([2, 4, 0.6])
                with cols[0]:
                    p["name"] = st.text_input(
                        "Pet name", value=p.get("name", ""), key=f"pet_n_{i}"
                    )
                with cols[1]:
                    p["description"] = st.text_input(
                        "Description (e.g., 'Female chihuahua')",
                        value=p.get("description", ""),
                        key=f"pet_d_{i}",
                    )
                with cols[2]:
                    st.markdown("&nbsp;")
                    if st.button("X", key=f"rm_pet_{i}"):
                        pets.pop(i)
                        st.rerun()
            if st.button("+ Add pet"):
                pets.append({"name": "", "description": ""})
                st.rerun()

            cols = st.columns(3)
            with cols[0]:
                a["pets"]["pet_rent_monthly"] = st.number_input(
                    "Pet rent / month ($)",
                    value=float(a["pets"].get("pet_rent_monthly", 0)),
                    step=10.0, min_value=0.0, format="%.2f",
                )
            with cols[1]:
                a["pets"]["non_refundable_pet_fee"] = st.number_input(
                    "Non-refundable pet fee ($)",
                    value=float(a["pets"].get("non_refundable_pet_fee", 0)),
                    step=25.0, min_value=0.0, format="%.2f",
                )
            with cols[2]:
                a["pets"]["refundable_pet_deposit"] = st.number_input(
                    "Refundable pet deposit ($)",
                    value=float(a["pets"].get("refundable_pet_deposit", 0)),
                    step=25.0, min_value=0.0, format="%.2f",
                )
            a["pets"]["insurance_required"] = st.checkbox(
                "Pet insurance required",
                value=a["pets"].get("insurance_required", False),
            )

    # Parking
    a["parking"]["include"] = st.checkbox(
        "Parking Addendum", value=a["parking"]["include"]
    )

    # Lawn care
    a["lawn_care"]["include"] = st.checkbox(
        "Tenant Yard Care Addendum", value=a["lawn_care"]["include"]
    )
    if a["lawn_care"]["include"]:
        a["lawn_care"]["equipment_provider"] = st.radio(
            "Lawn care equipment provided by",
            options=["landlord", "tenant"],
            index=0 if a["lawn_care"].get("equipment_provider", "landlord") == "landlord" else 1,
            format_func=lambda x: (
                "Landlord (provides equipment)" if x == "landlord"
                else "Tenant (uses own equipment)"
            ),
            horizontal=True,
        )

    # Rules
    a["rules"]["include"] = st.checkbox(
        "Rules Addendum", value=a["rules"]["include"]
    )
    if a["rules"]["include"]:
        rules_text = st.text_area(
            "Custom rules (one per line; appended to the standard rules)",
            value="\n".join(a["rules"].get("custom_rules", [])),
            height=120,
        )
        a["rules"]["custom_rules"] = [
            line.strip() for line in rules_text.split("\n") if line.strip()
        ]

    # Auto-derived
    st.subheader("Auto-derived (Washington requirements)")
    st.markdown("Mold Addendum — always included")
    st.markdown("Fire Safety Addendum — always included")
    lp = prop.get("lead_paint_disclosure_required", False)
    yb = prop.get("year_built")
    if lp:
        st.markdown(
            f"Lead-Based Paint Hazard Disclosure — **included** "
            f"(property year_built {yb}, pre-1978)"
        )
    else:
        st.markdown(
            "Lead-Based Paint Hazard Disclosure — not required (post-1978 housing)"
        )
    a["lead_paint_disclosure"]["include"] = lp


# ---------------------------------------------------------------------------
# Step 5: Review & Generate
# ---------------------------------------------------------------------------

def step_5_review():
    st.header("Step 5 — Review & Generate")

    tenancy = strip_scratch(st.session_state.tenancy)
    prop = load_property(tenancy["property_id"])

    addr = prop["address"]
    st.markdown(f"**Property:** {addr['street']}, {addr['city']}, {addr['state']} {addr['zip']}")
    st.markdown(
        f"**Lease type:** "
        f"{'Renewal' if tenancy['lease_kind'] == 'renewal' else 'New lease'}"
    )
    if tenancy["lease_kind"] == "renewal":
        st.markdown(
            f"**Prior lease effective date:** "
            f"{tenancy.get('prior_lease_effective_date', '?')}"
        )

    term = tenancy["term"]
    if term["type"] == "fixed":
        st.markdown(f"**Term:** {term['start_date']} → {term['end_date']}")
    else:
        st.markdown(f"**Term:** month-to-month from {term['start_date']}")

    tenants = tenancy.get("tenants", [])
    if tenants:
        st.markdown(f"**Tenant(s):** {', '.join(t['name'] for t in tenants)}")
    st.markdown(f"**Monthly rent:** ${tenancy['rent']['monthly_rent']:,.2f}")

    sd = tenancy["security_deposit"]
    if sd.get("already_held"):
        st.markdown(f"**Security deposit:** ${sd['amount']:,.2f} (already held under prior lease)")
    elif sd.get("payment_schedule") == "installments":
        inst = sd.get("installments", {})
        st.markdown(
            f"**Security deposit:** ${sd['amount']:,.2f} in {inst.get('count')} installments "
            f"of ${inst.get('amount', 0):,.2f} (first due at "
            f"{inst.get('first_due', 'start')})"
        )
    else:
        st.markdown(f"**Security deposit:** ${sd['amount']:,.2f} lump sum at signing")

    addenda_on = []
    for name, v in tenancy["addenda"].items():
        if v.get("include"):
            addenda_on.append(name.replace("_", " "))
    st.markdown(f"**Addenda included:** {', '.join(addenda_on)}")

    with st.expander("Show raw tenancy JSON"):
        st.json(tenancy)

    st.divider()
    if not tenancy.get("id"):
        st.warning("Set a tenancy ID in Step 1 before generating.")
    elif not tenants:
        st.warning("Add at least one tenant in Step 2 before generating.")
    else:
        if st.button("Generate Lease & Condition Report", type="primary"):
            with st.spinner("Generating documents..."):
                result = run_generate(tenancy)
            st.session_state.last_generated = result
            st.rerun()

    for warn in st.session_state.pop("pdf_warnings", []):
        st.warning(warn)

    if st.session_state.get("last_generated"):
        st.success("Generated successfully.")
        result = st.session_state.last_generated
        for label, path in result.items():
            if path and Path(path).exists():
                with open(path, "rb") as f:
                    st.download_button(
                        label=f"Download {label}",
                        data=f.read(),
                        file_name=Path(path).name,
                        key=f"dl_{label}",
                    )
                st.caption(f"Saved to: `{path}`")


def run_generate(tenancy):
    tenancy_path = TENANCIES_DIR / f"{tenancy['id']}.json"
    tenancy_path.write_text(json.dumps(tenancy, indent=2))

    landlord = load_landlord()
    property_ = load_property(tenancy["property_id"])

    OUTPUT_DIR.mkdir(exist_ok=True)
    last_name = (
        tenancy["tenants"][0]["name"].split()[-1] if tenancy.get("tenants") else "Tenant"
    )
    base = f"Lease-{last_name}-{tenancy['id']}"
    lease_docx = OUTPUT_DIR / f"{base}.docx"
    cr_docx = OUTPUT_DIR / f"Condition-Report-{last_name}-{tenancy['id']}.docx"

    generate_lease(landlord, property_, tenancy, str(lease_docx))

    if property_.get("rooms_for_condition_report"):
        generate_condition_report(landlord, property_, tenancy, str(cr_docx))
    else:
        cr_docx = None

    lease_pdf = try_pdf(lease_docx)
    cr_pdf = try_pdf(cr_docx) if cr_docx else None

    # Assemble the full packet (lease PDF + static pamphlets) when possible
    packet_pdf = None
    if lease_pdf:
        packet_path = lease_docx.with_name(
            lease_docx.stem.replace("Lease-", "Packet-") + ".pdf"
        )
        try:
            packet_pdf = assemble_packet(lease_pdf, property_, tenancy, packet_path)
        except Exception as e:
            # Non-fatal: packet assembly failing shouldn't block the lease download
            st.warning(f"Packet assembly skipped: {e}")

    return {
        "Lease (DOCX)": lease_docx,
        "Lease (PDF)": lease_pdf,
        "Full Packet (PDF)": packet_pdf,
        "Condition Report (DOCX)": cr_docx,
        "Condition Report (PDF)": cr_pdf,
        "Tenancy JSON": tenancy_path,
    }


def try_pdf(docx_path):
    """Convert a .docx to PDF, trying LibreOffice then docx2pdf (Word on macOS)."""
    if not docx_path:
        return None
    pdf_path = docx_path.with_suffix(".pdf")
    errors = []

    # 1. LibreOffice (headless, cross-platform)
    soffice = shutil.which("soffice")
    if not soffice:
        candidate = "/Applications/LibreOffice.app/Contents/MacOS/soffice"
        if Path(candidate).exists():
            soffice = candidate
    if soffice:
        try:
            subprocess.run(
                [
                    soffice, "--headless", "--convert-to", "pdf",
                    "--outdir", str(docx_path.parent), str(docx_path),
                ],
                check=True, capture_output=True, timeout=60,
            )
            if pdf_path.exists():
                return pdf_path
        except Exception as e:
            errors.append(f"LibreOffice: {e}")
    else:
        errors.append("LibreOffice: not found")

    # 2. docx2pdf — uses Microsoft Word on macOS/Windows.
    #    Copy docx into a temp dir so Word can read AND write freely, then move result.
    try:
        import tempfile, shutil as _shutil
        from docx2pdf import convert
        with tempfile.TemporaryDirectory() as tmp:
            tmp_docx = Path(tmp) / docx_path.name
            tmp_pdf  = Path(tmp) / pdf_path.name
            _shutil.copy2(str(docx_path), str(tmp_docx))
            convert(str(tmp_docx), str(tmp_pdf))
            if tmp_pdf.exists():
                _shutil.move(str(tmp_pdf), str(pdf_path))
                return pdf_path
            else:
                errors.append("docx2pdf: ran but no PDF produced (even in tmp dir)")
    except ImportError:
        errors.append("docx2pdf: not installed (run: pip install docx2pdf)")
    except Exception as e:
        errors.append(f"docx2pdf: {e}")

    msg = (
        "PDF conversion failed — DOCX only. To get PDFs, install LibreOffice or "
        "ensure Microsoft Word is installed. Details: " + " | ".join(errors)
    )
    print(f"[try_pdf] {msg}", file=sys.stderr)
    st.session_state.setdefault("pdf_warnings", []).append(msg)
    return None


# ---------------------------------------------------------------------------
# Landlord Profile page
# ---------------------------------------------------------------------------

def empty_landlord():
    return {
        "name": "",
        "address": {"street": "", "unit": "", "city": "", "state": "WA", "zip": ""},
        "phone": "",
        "phone_for_tenants": "",
        "email": "",
        "security_deposit_bank": {"name": "", "address": ""},
        "payment_methods": {
            "check_or_money_order": {
                "enabled": True,
                "payable_to": "",
                "mail_to": {"street": "", "unit": "", "city": "", "state": "WA", "zip": ""},
            },
            "electronic": [],
        },
        "defaults": {
            "late_fee": 200.0,
            "grace_period_days": 5,
            "smoking_violation_fee": 250.0,
            "variable_charge_payment_window_days": 15,
            "monthly_rent_due_day": 1,
            "month_to_month_tenant_notice_days": 20,
        },
    }


def page_landlord():
    st.header("Landlord Profile")
    landlord_path = DATA_DIR / "landlord.json"
    if "landlord_data" not in st.session_state:
        st.session_state.landlord_data = (
            json.loads(landlord_path.read_text()) if landlord_path.exists()
            else empty_landlord()
        )
    data = st.session_state.landlord_data

    st.subheader("Personal Info")
    data["name"] = st.text_input("Full name", value=data.get("name", ""))
    data["email"] = st.text_input("Email", value=data.get("email", ""))
    cols = st.columns(2)
    with cols[0]:
        data["phone"] = st.text_input("Phone (general)", value=data.get("phone", ""))
    with cols[1]:
        data["phone_for_tenants"] = st.text_input(
            "Phone for tenants", value=data.get("phone_for_tenants", "")
        )

    st.subheader("Mailing Address")
    addr = data.setdefault("address", {})
    addr["street"] = st.text_input("Street", value=addr.get("street", ""), key="ll_street")
    addr["unit"] = st.text_input("Unit / Apt (optional)", value=addr.get("unit", ""), key="ll_unit")
    cols = st.columns(3)
    with cols[0]:
        addr["city"] = st.text_input("City", value=addr.get("city", ""), key="ll_city")
    with cols[1]:
        addr["state"] = st.text_input("State", value=addr.get("state", "WA"), key="ll_state")
    with cols[2]:
        addr["zip"] = st.text_input("ZIP", value=addr.get("zip", ""), key="ll_zip")

    st.subheader("Security Deposit Bank")
    bank = data.setdefault("security_deposit_bank", {})
    cols = st.columns(2)
    with cols[0]:
        bank["name"] = st.text_input("Bank name", value=bank.get("name", ""))
    with cols[1]:
        bank["address"] = st.text_input(
            "Bank address (one line)", value=bank.get("address", "")
        )

    st.subheader("Payment Methods")
    pm = data.setdefault("payment_methods", {})
    chk = pm.setdefault("check_or_money_order", {})
    chk["enabled"] = st.checkbox(
        "Accept check / money order", value=chk.get("enabled", True)
    )
    if chk["enabled"]:
        chk["payable_to"] = st.text_input(
            "Make checks payable to", value=chk.get("payable_to", "")
        )
        st.caption("Mail-to address for checks:")
        mail = chk.setdefault("mail_to", {})
        mail["street"] = st.text_input("Street", value=mail.get("street", ""), key="chk_street")
        mail["unit"] = st.text_input("Unit (optional)", value=mail.get("unit", ""), key="chk_unit")
        cols = st.columns(3)
        with cols[0]:
            mail["city"] = st.text_input("City", value=mail.get("city", ""), key="chk_city")
        with cols[1]:
            mail["state"] = st.text_input("State", value=mail.get("state", "WA"), key="chk_state")
        with cols[2]:
            mail["zip"] = st.text_input("ZIP", value=mail.get("zip", ""), key="chk_zip")

    st.caption("Electronic payment services")
    electronic = pm.setdefault("electronic", [])
    to_remove = []
    for i, svc in enumerate(electronic):
        cols = st.columns([3, 3, 1])
        with cols[0]:
            svc["service"] = st.text_input(
                "Service", value=svc.get("service", ""), key=f"esvc_{i}"
            )
        with cols[1]:
            svc["notes"] = st.text_input(
                "Notes", value=svc.get("notes", ""), key=f"enotes_{i}"
            )
        with cols[2]:
            st.write("")
            if st.button("✕", key=f"edel_{i}"):
                to_remove.append(i)
    for i in reversed(to_remove):
        electronic.pop(i)
    if to_remove:
        st.rerun()
    if st.button("+ Add electronic payment service"):
        electronic.append({"service": "", "notes": ""})
        st.rerun()

    st.subheader("Lease Defaults")
    d = data.setdefault("defaults", {})
    cols = st.columns(3)
    with cols[0]:
        d["late_fee"] = st.number_input(
            "Late fee ($)", value=float(d.get("late_fee", 200)), step=25.0, min_value=0.0
        )
    with cols[1]:
        d["grace_period_days"] = st.number_input(
            "Grace period (days)", value=int(d.get("grace_period_days", 5)),
            step=1, min_value=0,
        )
    with cols[2]:
        d["monthly_rent_due_day"] = st.number_input(
            "Rent due day", value=int(d.get("monthly_rent_due_day", 1)),
            step=1, min_value=1, max_value=28,
        )
    cols = st.columns(3)
    with cols[0]:
        d["smoking_violation_fee"] = st.number_input(
            "Smoking violation fee ($)",
            value=float(d.get("smoking_violation_fee", 250)), step=25.0, min_value=0.0,
        )
    with cols[1]:
        d["variable_charge_payment_window_days"] = st.number_input(
            "Variable charge window (days)",
            value=int(d.get("variable_charge_payment_window_days", 15)), step=1, min_value=1,
        )
    with cols[2]:
        d["month_to_month_tenant_notice_days"] = st.number_input(
            "M-to-M tenant notice (days)",
            value=int(d.get("month_to_month_tenant_notice_days", 20)), step=1, min_value=1,
        )

    st.divider()
    if st.button("Save Landlord Profile", type="primary"):
        landlord_path.write_text(json.dumps(data, indent=2))
        st.success("Landlord profile saved.")


# ---------------------------------------------------------------------------
# Properties page
# ---------------------------------------------------------------------------

def empty_property():
    return {
        "id": "",
        "address": {"street": "", "city": "Spokane", "state": "WA", "zip": ""},
        "jurisdiction": "city",
        "property_type": "single-family residence",
        "year_built": 2000,
        "lead_paint_disclosure_required": False,
        "appliances": [],
        "utilities": {
            "electricity": {"provider": "Avista", "paid_by": "tenant", "included_in_rent": False},
            "natural_gas": {"provider": "Avista", "paid_by": "tenant", "included_in_rent": False},
            "water_sewer": {"provider": "", "paid_by": "landlord", "included_in_rent": True},
            "trash": {
                "provider": "", "paid_by": "landlord", "included_in_rent": True,
                "pickup_day": "", "calendar_url": "", "calendar_pdf_path": None,
                "calendar_last_refreshed": None, "notes": "",
            },
            "heating": {"type": "natural_gas", "thermostat": {"brand": "", "notes": ""}},
            "snow_removal": {"provided_by": "tenant"},
            "landscaping": {"provided_by": "tenant", "tenant_addendum_required": True},
            "telephone": {"setup": "tenant"},
            "cable_tv": {"setup": "tenant"},
            "internet": {"setup": "tenant"},
        },
        "storage": {"available": False, "location": "", "included_in_rent": True, "monthly_fee": 0},
        "parking": {"available": True, "type": "designated", "included_in_rent": True, "monthly_fee": 0},
        "smoke_detectors": {"installed": True, "battery_operated": True},
        "co_detectors": {"installed": None},
        "fire_alarm_system": False,
        "furnace_filter": {"location": "", "instructions": ""},
        "rooms_for_condition_report": [],
    }


def page_properties():
    st.header("Properties")
    editing = st.session_state.get("prop_editing")  # None | "new" | property_id

    if editing is None:
        # ------------------------------------------------------------------ list view
        properties = list_properties()
        if not properties:
            st.info("No properties yet. Add one to get started.")
        else:
            for prop_id, label in properties:
                cols = st.columns([5, 1])
                with cols[0]:
                    st.write(f"**{label}**")
                with cols[1]:
                    if st.button("Edit", key=f"edit_{prop_id}"):
                        st.session_state.prop_editing = prop_id
                        st.session_state.prop_data = load_property(prop_id)
                        st.rerun()
        st.divider()
        if st.button("＋ Add Property", type="primary"):
            st.session_state.prop_editing = "new"
            st.session_state.prop_data = empty_property()
            st.rerun()

    else:
        # ------------------------------------------------------------------ edit / new
        data = st.session_state.prop_data
        is_new = editing == "new"

        if is_new:
            st.subheader("New Property")
            data["id"] = st.text_input(
                "Property ID (filename slug, e.g. maple_st)",
                value=data.get("id", ""),
                help="Lowercase, underscores only. Saved as data/properties/<id>.json",
            )
        else:
            st.subheader(data["address"].get("street", editing))
            st.caption(f"ID: `{editing}`")

        # Basic info
        st.subheader("Address")
        addr = data.setdefault("address", {})
        addr["street"] = st.text_input("Street", value=addr.get("street", ""))
        cols = st.columns(3)
        with cols[0]:
            addr["city"] = st.text_input("City", value=addr.get("city", "Spokane"))
        with cols[1]:
            addr["state"] = st.text_input("State", value=addr.get("state", "WA"))
        with cols[2]:
            addr["zip"] = st.text_input("ZIP", value=addr.get("zip", ""))

        cols = st.columns(3)
        with cols[0]:
            data["jurisdiction"] = st.selectbox(
                "Jurisdiction", options=["city", "county"],
                index=0 if data.get("jurisdiction", "city") == "city" else 1,
            )
        with cols[1]:
            data["year_built"] = st.number_input(
                "Year built", value=int(data.get("year_built") or 1990),
                min_value=1800, max_value=date.today().year, step=1,
            )
        with cols[2]:
            data["lead_paint_disclosure_required"] = st.checkbox(
                "Lead paint disclosure required (pre-1978)",
                value=data.get("lead_paint_disclosure_required", False),
            )
        data["property_type"] = st.text_input(
            "Property type", value=data.get("property_type", "single-family residence")
        )

        # Utilities
        with st.expander("Utilities"):
            utils = data.setdefault("utilities", {})

            st.markdown("**Electricity**")
            el = utils.setdefault("electricity", {})
            cols = st.columns(2)
            with cols[0]:
                el["provider"] = st.text_input("Provider", value=el.get("provider", ""), key="el_prov")
            with cols[1]:
                el["paid_by"] = st.selectbox(
                    "Paid by", ["tenant", "landlord"],
                    index=0 if el.get("paid_by", "tenant") == "tenant" else 1, key="el_paid",
                )

            st.markdown("**Natural Gas**")
            ng = utils.setdefault("natural_gas", {})
            cols = st.columns(2)
            with cols[0]:
                ng["provider"] = st.text_input("Provider", value=ng.get("provider", ""), key="ng_prov")
            with cols[1]:
                ng["paid_by"] = st.selectbox(
                    "Paid by", ["tenant", "landlord"],
                    index=0 if ng.get("paid_by", "tenant") == "tenant" else 1, key="ng_paid",
                )

            st.markdown("**Water & Sewer**")
            ws = utils.setdefault("water_sewer", {})
            cols = st.columns(2)
            with cols[0]:
                ws["provider"] = st.text_input("Provider", value=ws.get("provider", ""), key="ws_prov")
            with cols[1]:
                ws["paid_by"] = st.selectbox(
                    "Paid by", ["landlord", "tenant"],
                    index=0 if ws.get("paid_by", "landlord") == "landlord" else 1, key="ws_paid",
                )

            st.markdown("**Trash**")
            tr = utils.setdefault("trash", {})
            cols = st.columns(2)
            with cols[0]:
                tr["provider"] = st.text_input("Provider", value=tr.get("provider", ""), key="tr_prov")
            with cols[1]:
                tr["pickup_day"] = st.text_input("Pickup day", value=tr.get("pickup_day", ""), key="tr_day")
            cols = st.columns(2)
            with cols[0]:
                tr["paid_by"] = st.selectbox(
                    "Paid by", ["landlord", "tenant"],
                    index=0 if tr.get("paid_by", "landlord") == "landlord" else 1, key="tr_paid",
                )
            with cols[1]:
                tr["notes"] = st.text_input("Notes", value=tr.get("notes", ""), key="tr_notes")
            tr["calendar_url"] = st.text_input(
                "Calendar URL", value=tr.get("calendar_url", ""), key="tr_url"
            )
            tr["calendar_pdf_path"] = (
                st.text_input(
                    "Calendar PDF path (relative to repo root, once downloaded)",
                    value=tr.get("calendar_pdf_path") or "", key="tr_pdf",
                ) or None
            )

            st.markdown("**Heating**")
            ht = utils.setdefault("heating", {})
            ht_options = ["natural_gas", "electric", "oil", "heat_pump", "other"]
            cols = st.columns(2)
            with cols[0]:
                ht_val = ht.get("type", "natural_gas")
                ht["type"] = st.selectbox(
                    "Type", ht_options,
                    index=ht_options.index(ht_val) if ht_val in ht_options else 0,
                    key="ht_type",
                )
            therm = ht.setdefault("thermostat", {})
            with cols[1]:
                therm["brand"] = st.text_input(
                    "Thermostat brand", value=therm.get("brand", ""), key="therm_brand"
                )
            therm["notes"] = st.text_input(
                "Thermostat notes", value=therm.get("notes", ""), key="therm_notes"
            )

            st.markdown("**Other Services**")
            cols = st.columns(3)
            with cols[0]:
                utils.setdefault("snow_removal", {})["provided_by"] = st.selectbox(
                    "Snow removal", ["tenant", "landlord"],
                    index=0 if utils.get("snow_removal", {}).get("provided_by", "tenant") == "tenant" else 1,
                    key="snow",
                )
            land = utils.setdefault("landscaping", {})
            with cols[1]:
                land["provided_by"] = st.selectbox(
                    "Landscaping", ["tenant", "landlord"],
                    index=0 if land.get("provided_by", "tenant") == "tenant" else 1,
                    key="land_by",
                )
            with cols[2]:
                land["tenant_addendum_required"] = st.checkbox(
                    "Lawn care addendum",
                    value=land.get("tenant_addendum_required", True),
                )

            setup_opts = ["tenant", "landlord", "landlord_fixed_fee"]
            cols = st.columns(3)
            for svc_key, svc_label, col in [
                ("telephone", "Telephone", cols[0]),
                ("cable_tv", "Cable TV", cols[1]),
                ("internet", "Internet", cols[2]),
            ]:
                with col:
                    svc = utils.setdefault(svc_key, {"setup": "tenant"})
                    val = svc.get("setup", "tenant")
                    svc["setup"] = st.selectbox(
                        svc_label, setup_opts,
                        index=setup_opts.index(val) if val in setup_opts else 0,
                        key=f"svc_{svc_key}",
                    )

        # Appliances
        with st.expander("Appliances"):
            appliances = data.setdefault("appliances", [])
            to_remove = []
            for i, appl in enumerate(appliances):
                cols = st.columns([5, 1])
                with cols[0]:
                    appl["name"] = st.text_input(
                        "Name", value=appl.get("name", ""), key=f"appl_{i}",
                        label_visibility="collapsed",
                    )
                with cols[1]:
                    if st.button("✕", key=f"appl_del_{i}"):
                        to_remove.append(i)
            for i in reversed(to_remove):
                appliances.pop(i)
            if to_remove:
                st.rerun()
            if st.button("+ Add appliance"):
                appliances.append({"name": ""})
                st.rerun()

        # Storage & Parking
        with st.expander("Storage & Parking"):
            stor = data.setdefault("storage", {})
            stor["available"] = st.checkbox("Storage available", value=stor.get("available", False))
            if stor["available"]:
                cols = st.columns(3)
                with cols[0]:
                    stor["location"] = st.text_input("Location", value=stor.get("location", ""))
                with cols[1]:
                    stor["included_in_rent"] = st.checkbox(
                        "Included in rent", value=stor.get("included_in_rent", True), key="stor_incl"
                    )
                with cols[2]:
                    stor["monthly_fee"] = st.number_input(
                        "Monthly fee ($)", value=float(stor.get("monthly_fee", 0)),
                        min_value=0.0, step=10.0, key="stor_fee",
                    )

            st.divider()
            park = data.setdefault("parking", {})
            park["available"] = st.checkbox("Parking available", value=park.get("available", True))
            if park["available"]:
                park_types = ["designated", "street", "garage", "lot"]
                park_val = park.get("type", "designated")
                cols = st.columns(3)
                with cols[0]:
                    park["type"] = st.selectbox(
                        "Type", park_types,
                        index=park_types.index(park_val) if park_val in park_types else 0,
                    )
                with cols[1]:
                    park["included_in_rent"] = st.checkbox(
                        "Included in rent", value=park.get("included_in_rent", True), key="park_incl"
                    )
                with cols[2]:
                    park["monthly_fee"] = st.number_input(
                        "Monthly fee ($)", value=float(park.get("monthly_fee", 0)),
                        min_value=0.0, step=10.0, key="park_fee",
                    )

        # Safety
        with st.expander("Safety"):
            sd = data.setdefault("smoke_detectors", {})
            cols = st.columns(2)
            with cols[0]:
                sd["installed"] = st.checkbox(
                    "Smoke detectors installed", value=sd.get("installed", True)
                )
            with cols[1]:
                if sd["installed"]:
                    sd["battery_operated"] = st.checkbox(
                        "Battery operated", value=sd.get("battery_operated", True)
                    )

            co = data.setdefault("co_detectors", {})
            co_installed = co.get("installed")
            co_val = st.selectbox(
                "CO detectors installed",
                ["Yes", "No", "Unknown"],
                index=0 if co_installed is True else (1 if co_installed is False else 2),
            )
            co["installed"] = True if co_val == "Yes" else (False if co_val == "No" else None)

            data["fire_alarm_system"] = st.checkbox(
                "Fire alarm system", value=data.get("fire_alarm_system", False)
            )

        # Maintenance
        with st.expander("Maintenance"):
            ff = data.setdefault("furnace_filter", {})
            ff["location"] = st.text_input(
                "Furnace filter location", value=ff.get("location", "")
            )
            ff["instructions"] = st.text_area(
                "Filter instructions", value=ff.get("instructions", ""), height=80
            )

        # Rooms for condition report
        with st.expander("Rooms for Condition Report"):
            rooms = data.setdefault("rooms_for_condition_report", [])
            to_remove = []
            for i, room in enumerate(rooms):
                cols = st.columns([5, 1])
                with cols[0]:
                    rooms[i] = st.text_input(
                        "Room", value=room, key=f"room_{i}", label_visibility="collapsed"
                    )
                with cols[1]:
                    if st.button("✕", key=f"room_del_{i}"):
                        to_remove.append(i)
            for i in reversed(to_remove):
                rooms.pop(i)
            if to_remove:
                st.rerun()
            if st.button("+ Add room"):
                rooms.append("")
                st.rerun()

        # Save / Cancel
        st.divider()
        cols = st.columns([1, 1, 4])
        with cols[0]:
            if st.button("← Cancel"):
                st.session_state.prop_editing = None
                st.session_state.pop("prop_data", None)
                st.rerun()
        with cols[1]:
            label = "Save New Property" if is_new else "Save Changes"
            if st.button(label, type="primary"):
                prop_id = data.get("id", "").strip()
                if not prop_id:
                    st.error("Property ID is required.")
                else:
                    PROPERTIES_DIR.mkdir(parents=True, exist_ok=True)
                    path = PROPERTIES_DIR / f"{prop_id}.json"
                    path.write_text(json.dumps(data, indent=2))
                    st.session_state.prop_editing = None
                    st.session_state.pop("prop_data", None)
                    st.success(f"Saved {path.name}")
                    st.rerun()


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

STEPS = [
    ("Property", step_1_property_and_kind),
    ("Tenants", step_2_tenants),
    ("Money", step_3_money),
    ("Addenda", step_4_addenda),
    ("Review", step_5_review),
]


def main():
    st.set_page_config(page_title="Lease Wizard", layout="centered")
    init_session()

    st.sidebar.markdown("## Lease Wizard")
    section = st.sidebar.radio(
        "Navigate",
        options=["lease", "properties", "landlord"],
        format_func=lambda x: {
            "lease": "📋  New Lease",
            "properties": "🏡  Properties",
            "landlord": "👤  Landlord Profile",
        }[x],
        key="section",
        label_visibility="collapsed",
    )
    st.sidebar.divider()

    if section == "lease":
        for i, (name, _) in enumerate(STEPS, 1):
            if i < st.session_state.step:
                st.sidebar.markdown(f"✓ {i}. {name}")
            elif i == st.session_state.step:
                st.sidebar.markdown(f"**▶ {i}. {name}**")
            else:
                st.sidebar.markdown(f"&nbsp;&nbsp;&nbsp;{i}. {name}")
        st.sidebar.divider()
        if st.sidebar.button("Reset wizard"):
            for k in ["step", "tenancy", "last_generated"]:
                st.session_state.pop(k, None)
            st.rerun()

        _, render_fn = STEPS[st.session_state.step - 1]
        render_fn()

        st.divider()
        cols = st.columns([1, 1, 4])
        with cols[0]:
            if st.session_state.step > 1:
                if st.button("← Back"):
                    st.session_state.step -= 1
                    st.rerun()
        with cols[1]:
            if st.session_state.step < len(STEPS):
                if st.button("Next →", type="primary"):
                    st.session_state.step += 1
                    st.rerun()

    elif section == "properties":
        page_properties()

    elif section == "landlord":
        page_landlord()


if __name__ == "__main__":
    main()

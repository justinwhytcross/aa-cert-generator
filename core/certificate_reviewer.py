"""Review returned certificate PDFs against checklist requirements."""

import re
from .pdf_reader import extract_text
from .checklist_parser import extract_ncc_clauses, extract_standards


def review_against_checklist(
    cert_paths: list,
    cert_filenames: list,
    checklist_items: list,
    psol_matches: dict,
    uploaded_reports: dict = None,
) -> list:
    """Review certificates against the checklist. Returns one row per checklist item.

    Args:
        cert_paths: List of temp paths to uploaded certificate PDFs
        cert_filenames: Original filenames for each cert
        checklist_items: Master checklist items list
        psol_matches: Dict mapping item id -> matched PSOLs
        uploaded_reports: Dict mapping report type -> list of report dicts

    Returns:
        List of dicts, one per checklist item:
        [{item_name, item_id, category, certificate, status, issues}]
    """
    uploaded_reports = uploaded_reports or {}

    # Extract text from all uploaded certs
    certs = []
    for path, filename in zip(cert_paths, cert_filenames):
        text = extract_text(path)
        certs.append({"path": path, "filename": filename, "text": text})

    # Skip admin items that don't generate templates
    skip_ids = {
        "cou_form", "government_levies", "trade_clearances",
        "survey_certificate", "fire_brigade_clearance",
        "developers_consent", "performance_solutions",
        "tccs_operational_acceptance", "access_consultant_signoff",
        "weatherproofing_facade", "fire_door_schedule",
    }

    results = []
    used_certs = set()  # Track which certs have been matched

    for item in checklist_items:
        item_id = item.get("id", "")
        if item_id in skip_ids:
            continue

        item_name = item.get("name", "")
        category = item.get("category", "")

        # Find the best matching cert for this checklist item
        matched_cert = _find_best_cert(item, certs, used_certs)

        if not matched_cert:
            results.append({
                "item_name": item_name,
                "item_id": item_id,
                "category": category,
                "certificate": "—",
                "status": "NOT RECEIVED",
                "issues": [],
            })
            continue

        used_certs.add(matched_cert["filename"])

        # Build expected requirements
        matched_psols = psol_matches.get(item_id, [])
        report_refs = set()
        for p in matched_psols:
            ref = p.get("report_reference", "")
            if ref:
                report_refs.add(ref)

        report_types = item.get("report_types", [])
        has_report = bool(report_refs)
        for rtype in report_types:
            if uploaded_reports.get(rtype):
                has_report = True
                for r in uploaded_reports[rtype]:
                    ref = r.get("reference", "")
                    if ref:
                        report_refs.add(ref)

        # Run checks
        issues = _check_certificate(
            matched_cert["text"],
            ncc_clauses=item.get("ncc_clauses", []),
            standards=item.get("standards", []),
            has_report=has_report,
            report_refs=report_refs,
        )

        if not issues:
            status = "PASS"
        else:
            # If only warnings, status is REVIEW; if any fails, FAIL
            has_fail = any(i["severity"] == "FAIL" for i in issues)
            status = "FAIL" if has_fail else "REVIEW"

        results.append({
            "item_name": item_name,
            "item_id": item_id,
            "category": category,
            "certificate": matched_cert["filename"],
            "status": status,
            "issues": issues,
        })

    # Check for unmatched certs (uploaded but didn't match any item)
    for cert in certs:
        if cert["filename"] not in used_certs:
            results.append({
                "item_name": "— Unmatched",
                "item_id": "",
                "category": "",
                "certificate": cert["filename"],
                "status": "REVIEW",
                "issues": [{"text": "Could not match to a checklist item", "severity": "WARN"}],
            })

    return results


def _check_certificate(text: str, ncc_clauses: list, standards: list,
                       has_report: bool, report_refs: set) -> list:
    """Run all checks on a certificate's text. Returns list of issues found."""
    issues = []

    # 1. NCC clauses
    if ncc_clauses:
        found_clauses = extract_ncc_clauses(text)
        missing = [c for c in ncc_clauses if c not in found_clauses]
        if missing:
            issues.append({
                "text": f"Missing NCC clause: {', '.join(missing)}",
                "severity": "FAIL",
            })

    # 2. Australian Standards
    if standards:
        found_standards = extract_standards(text)
        for exp_std in standards:
            exp_num = re.search(r'(\d+\.?\d*)', exp_std)
            if not exp_num:
                continue
            exp_num_str = exp_num.group(1)
            matched = False
            for fs in found_standards:
                if exp_num_str in fs:
                    matched = True
                    # Check year
                    exp_year = re.search(r'(\d{4})', exp_std)
                    found_year = re.search(r'(\d{4})', fs)
                    if exp_year and found_year and exp_year.group(1) != found_year.group(1):
                        issues.append({
                            "text": f"Wrong AS year: {exp_std} (found {found_year.group(1)})",
                            "severity": "FAIL",
                        })
                    break
            if not matched:
                issues.append({
                    "text": f"AS not referenced: {exp_std}",
                    "severity": "FAIL",
                })

    # 3. Report reference (FER, Access, etc.)
    if has_report and report_refs:
        has_any_ref = bool(re.search(
            r'[Pp]erformance\s+[Ss]olution|[Ff]ire\s+[Ee]ngineering|FER\b|'
            r'[Aa]ccess\s+[Ss]olution|[Ff]acade\s+[Ee]ngineer',
            text
        ))
        if not has_any_ref:
            issues.append({
                "text": "No report reference found (FER/Access/Facade expected)",
                "severity": "FAIL",
            })

    # 4. ABN
    abn_match = re.search(r'\b\d{2}\s?\d{3}\s?\d{3}\s?\d{3}\b', text)
    if not abn_match:
        issues.append({
            "text": "No ABN detected",
            "severity": "WARN",
        })

    # 5. Signature
    has_sig = bool(re.search(r'[Ss]ign(?:ed|ature)', text))
    if not has_sig:
        issues.append({
            "text": "No signature detected",
            "severity": "WARN",
        })

    return issues


def _find_best_cert(item: dict, certs: list, used_certs: set) -> dict:
    """Find the best matching certificate for a checklist item."""
    item_id = item.get("id", "")
    item_name = item.get("name", "").lower()

    keyword_map = {
        "fire_hydrants": ["fire hydrant", "hydrant"],
        "fire_hose_reels": ["fire hose reel", "hose reel"],
        "sprinklers": ["sprinkler"],
        "sprinklers_independent": ["sprinkler", "independent"],
        "extinguishers": ["extinguisher"],
        "smoke_detection": ["smoke detection", "smoke alarm", "smoke detector", "as 1670"],
        "emergency_lights": ["emergency light", "emergency exit light"],
        "exit_signs": ["exit sign"],
        "sound_system_intercom": ["sound system", "intercom", "ewis"],
        "bows": ["building occupant warning", "bows", "occupant warning"],
        "mechanical_services": ["air-conditioning", "ventilation", "hvac", "mechanical service"],
        "stair_pressurisation": ["stair pressurisation", "stair pressuri"],
        "zone_smoke_control": ["zone smoke control", "smoke control"],
        "fire_rated_penetrations": ["fire rated penetration", "penetration seal"],
        "fire_rated_penetrations_independent": ["fire rated penetration", "independent"],
        "fire_rated_doorsets": ["fire door", "doorset", "fire rated door"],
        "lightweight_fire_rating": ["lightweight fire", "fire rating wall", "lightweight wall"],
        "waterproofing_internal": ["waterproofing", "wet area", "as 3740"],
        "waterproofing_external": ["external waterproofing", "as 4654"],
        "glazing": ["glazing", "as 1288"],
        "insulation": ["thermal insulation", "r-value"],
        "insulation_eer": ["energy efficiency", "energy rating", "eer"],
        "acoustics": ["acoustic", "sound insulation"],
        "plumbing_acoustics": ["plumbing acoustic", "pipe acoustic"],
        "balustrades": ["balustrade", "handrail"],
        "structural_engineers_clearance": ["structural engineer", "structural compliance", "structural clearance"],
        "truss_and_bracing": ["truss", "bracing", "wall frame"],
        "truss_and_bracing_installation": ["truss install", "framing carpenter"],
        "roofing": ["roofing", "roof covering", "roof installation"],
        "sarking": ["sarking"],
        "flooring": ["floor covering", "flooring"],
        "termite_protection": ["termite"],
        "slip_resistance": ["slip resistance", "slip test", "as 4586"],
        "hearing_augmentation": ["hearing augmentation", "hearing loop"],
        "fire_resisting_construction": ["fire resisting construction", "fire resistant construction"],
        "fire_hazard_properties": ["fire hazard properties", "fire hazard"],
        "fire_test_reports": ["fire test report", "as 1530"],
        "internal_linings": ["internal lining", "wall lining", "ceiling lining"],
        "carpet_thickness": ["carpet thickness"],
        "structural_restraint": ["structural restraint", "seismic restraint"],
        "dincel_installation": ["dincel"],
        "lift_certificate": ["lift certificate", "elevator", "workcover"],
        "roof_ventilation": ["roof ventil"],
        "fire_engineers_clearance": ["fire engineer clearance", "fire engineering clearance"],
    }

    keywords = keyword_map.get(item_id, [])
    best_cert = None
    best_score = 0

    for cert in certs:
        if cert["filename"] in used_certs:
            continue

        text_lower = cert["text"].lower()
        filename_lower = cert["filename"].lower()
        score = 0

        # Check filename first (most reliable)
        for kw in keywords:
            if kw in filename_lower:
                score += 5
            if kw in text_lower:
                score += 1

        # Also match on item name in filename
        # e.g., "FIRE HYDRANTS.pdf" matches fire_hydrants
        name_words = item_name.split()
        for word in name_words:
            if len(word) > 3 and word in filename_lower:
                score += 3

        # NCC clause match
        for clause in item.get("ncc_clauses", []):
            if clause in cert["text"]:
                score += 2

        # AS standard match
        for std in item.get("standards", []):
            std_num = re.search(r'(\d{4,})', std)
            if std_num and std_num.group(1) in cert["text"]:
                score += 1

        if score > best_score:
            best_score = score
            best_cert = cert

    return best_cert if best_score >= 2 else None

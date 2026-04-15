"""Absolute Approvals - Certificate Template Generator & Reviewer"""

import os
import sys
import json
import tempfile
import streamlit as st

# Add parent to path for imports
sys.path.insert(0, os.path.dirname(__file__))

from core.checklist_parser import parse_checklist
from core.psol_extractor import extract_psols_from_pdf, extract_psols_from_multiple, match_psols_to_checklist, detect_report_type
from core.template_generator import generate_all_templates, generate_zip
from core.certificate_reviewer import review_against_checklist

DATA_DIR = os.path.join(os.path.dirname(__file__), "data")
PROJECTS_DIR = os.path.join(os.path.dirname(__file__), "projects")


st.set_page_config(
    page_title="Absolute Approvals - Certificate Tools",
    page_icon="AA",
    layout="wide",
)

st.title("Absolute Approvals")
st.subheader("Certificate Template Generator & Reviewer")

tab1, tab2 = st.tabs(["Generate Certificate Templates", "Review Returned Certificates"])


# ============================================================
# TAB 1: Generate Certificate Templates
# ============================================================
with tab1:
    st.markdown("Upload a COU checklist and Performance Solution reports to generate certificate templates for trades.")

    col1, col2 = st.columns(2)

    with col1:
        st.markdown("#### 1. Upload COU Checklist")
        checklist_file = st.file_uploader(
            "Upload the Commercial Certificate Checklist (.docx)",
            type=["docx"],
            key="checklist_upload",
        )

    with col2:
        st.markdown("#### 2. Upload Performance Solution Reports")
        psol_files = st.file_uploader(
            "Upload Performance Solution Reports (FER, Access, Facade, etc.)",
            type=["pdf", "docx"],
            accept_multiple_files=True,
            key="psol_upload",
        )

    if checklist_file:
        # Save checklist to temp file and parse
        with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as tmp:
            tmp.write(checklist_file.read())
            tmp_path = tmp.name

        try:
            checklist_data = parse_checklist(tmp_path)
            project_info = checklist_data["project"]
            checklist_items_parsed = checklist_data["items"]

            st.markdown("---")
            st.markdown("#### Project Details")
            col_a, col_b = st.columns(2)
            with col_a:
                st.text_input("Job Number", value=project_info.get("job_number", ""), key="job_no")
                st.text_input("Attention", value=project_info.get("attention", ""), key="attention")
                st.text_input("Company", value=project_info.get("company", ""), key="company")
            with col_b:
                st.text_input("Building", value=project_info.get("building", ""), key="building")
                st.text_area("Address", value=project_info.get("address", ""), key="address", height=80)

            st.markdown(f"**Checklist items found:** {len(checklist_items_parsed)}")

            # Load master items for PSOL matching and template generation
            with open(os.path.join(DATA_DIR, "checklist_items.json"), "r") as f:
                master_items = json.load(f)

            # Process PSOL PDFs
            all_psols = []
            uploaded_reports = {}  # {type: [{reference, company, date}]}
            if psol_files:
                psol_temp_paths = []
                psol_original_names = []
                for pf in psol_files:
                    with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
                        tmp.write(pf.read())
                        psol_temp_paths.append(tmp.name)
                        psol_original_names.append(pf.name)

                with st.spinner("Extracting Performance Solutions from PDFs..."):
                    # Extract PSOLs and detect report types
                    for temp_path, orig_name in zip(psol_temp_paths, psol_original_names):
                        result = extract_psols_from_pdf(temp_path)
                        report = result["report"]
                        report["filename"] = orig_name

                        # Detect report type from content and filename
                        rtype = detect_report_type(temp_path, orig_name)

                        # Add report to uploaded_reports by type
                        if rtype:
                            if rtype not in uploaded_reports:
                                uploaded_reports[rtype] = []
                            report_entry = {
                                "reference": report.get("reference", ""),
                                "company": report.get("company", ""),
                                "date": report.get("date", ""),
                                "filename": orig_name,
                            }
                            # Avoid duplicate entries
                            if report_entry not in uploaded_reports[rtype]:
                                uploaded_reports[rtype].append(report_entry)

                        # Add PSOLs with report metadata
                        for psol in result["psols"]:
                            psol["report_filename"] = orig_name
                            psol["report_reference"] = report.get("reference", "")
                            psol["report_company"] = report.get("company", "")
                            psol["report_date"] = report.get("date", "")
                            all_psols.append(psol)

                st.markdown(f"**Performance Solutions extracted:** {len(all_psols)}")

                # Show detected report types
                if uploaded_reports:
                    for rtype, reports in uploaded_reports.items():
                        for r in reports:
                            ref_str = r.get("reference", r.get("filename", ""))
                            st.markdown(f"**{rtype.title()} report detected:** {ref_str}")

                # Clean up temp files
                for p in psol_temp_paths:
                    try:
                        os.unlink(p)
                    except Exception:
                        pass

            # Match PSOLs to checklist items
            psol_matches = match_psols_to_checklist(all_psols, master_items)

            # Show PSOL matching preview
            if all_psols:
                st.markdown("---")
                st.markdown("#### Performance Solution Matching")
                match_data = []
                for item in master_items:
                    matched = psol_matches.get(item["id"], [])
                    if matched:
                        ps_nums = ", ".join(p["ps_number"] for p in matched)
                        match_data.append({
                            "Certificate": item["name"],
                            "Matched PSOLs": ps_nums,
                            "Count": len(matched),
                        })
                if match_data:
                    st.dataframe(match_data, use_container_width=True)
                else:
                    st.info("No Performance Solutions matched to checklist items. Templates will be generated without PSOL references.")

            # Generate button
            st.markdown("---")
            if st.button("Generate Certificate Templates", type="primary", use_container_width=True):
                # Update project info from form
                project_info_final = {
                    "job_number": st.session_state.get("job_no", project_info.get("job_number", "")),
                    "attention": st.session_state.get("attention", project_info.get("attention", "")),
                    "company": st.session_state.get("company", project_info.get("company", "")),
                    "building": st.session_state.get("building", project_info.get("building", "")),
                    "address": st.session_state.get("address", project_info.get("address", "")),
                    "date": project_info.get("date", ""),
                    "description": project_info.get("description", ""),
                    "ncc_year": project_info.get("ncc_year", "2022"),
                    "ncc_amendment": project_info.get("ncc_amendment", ""),
                }

                with st.spinner("Generating certificate templates..."):
                    zip_bytes = generate_zip(master_items, psol_matches, project_info_final, uploaded_reports)

                job_no = project_info_final.get("job_number", "certificates")
                st.download_button(
                    label=f"Download All Templates ({job_no}).zip",
                    data=zip_bytes,
                    file_name=f"{job_no} - Certificate Templates.zip",
                    mime="application/zip",
                    type="primary",
                    use_container_width=True,
                )
                st.success(f"Certificate templates generated for job {job_no}")

                # Also save to projects folder
                project_dir = os.path.join(PROJECTS_DIR, job_no, "templates_out")
                os.makedirs(project_dir, exist_ok=True)
                generated = generate_all_templates(master_items, psol_matches, project_info_final, project_dir, uploaded_reports)
                st.info(f"Templates also saved to: {project_dir}")

                # Save project data for review tab
                project_data = {
                    "project_info": project_info_final,
                    "uploaded_reports": uploaded_reports,
                    "psol_matches": {
                        k: [{"ps_number": p["ps_number"], "clauses": p["clauses"], "summary": p.get("summary", ""),
                             "report_reference": p.get("report_reference", ""),
                             "report_company": p.get("report_company", ""),
                             "report_date": p.get("report_date", "")}
                            for p in v]
                        for k, v in psol_matches.items()
                    },
                }
                with open(os.path.join(PROJECTS_DIR, job_no, "project.json"), "w") as f:
                    json.dump(project_data, f, indent=2)

        except Exception as e:
            st.error(f"Error parsing checklist: {e}")
        finally:
            try:
                os.unlink(tmp_path)
            except Exception:
                pass


# ============================================================
# TAB 2: Review Returned Certificates
# ============================================================
with tab2:
    st.markdown("Upload returned certificate PDFs to review against the project checklist.")

    # List existing projects
    existing_projects = []
    if os.path.exists(PROJECTS_DIR):
        for d in os.listdir(PROJECTS_DIR):
            project_json = os.path.join(PROJECTS_DIR, d, "project.json")
            if os.path.exists(project_json):
                existing_projects.append(d)

    if not existing_projects:
        st.info("No projects found. Generate certificate templates first (Tab 1) to create a project.")
    else:
        selected_project = st.selectbox("Select Project", existing_projects)

        if selected_project:
            project_json_path = os.path.join(PROJECTS_DIR, selected_project, "project.json")
            with open(project_json_path, "r") as f:
                project_data = json.load(f)

            project_info = project_data.get("project_info", {})
            st.markdown(f"**Project:** {project_info.get('building', selected_project)}")
            st.markdown(f"**Job Number:** {selected_project}")

            # Load master items
            with open(os.path.join(DATA_DIR, "checklist_items.json"), "r") as f:
                master_items = json.load(f)

            # Reconstruct project data
            psol_matches = project_data.get("psol_matches", {})
            proj_uploaded_reports = project_data.get("uploaded_reports", {})

            # Upload certificates
            cert_files = st.file_uploader(
                "Upload returned certificates",
                type=["pdf", "docx"],
                accept_multiple_files=True,
                key="cert_upload",
            )

            if cert_files and st.button("Review Certificates", type="primary", use_container_width=True):
                cert_temp_paths = []
                cert_original_names = []
                for cf in cert_files:
                    with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
                        tmp.write(cf.read())
                        cert_temp_paths.append(tmp.name)
                        cert_original_names.append(cf.name)

                with st.spinner("Reviewing certificates against checklist..."):
                    results = review_against_checklist(
                        cert_temp_paths, cert_original_names,
                        master_items, psol_matches, proj_uploaded_reports,
                    )

                # Clean up
                for p in cert_temp_paths:
                    try:
                        os.unlink(p)
                    except Exception:
                        pass

                # Summary counts
                pass_count = sum(1 for r in results if r["status"] == "PASS")
                fail_count = sum(1 for r in results if r["status"] == "FAIL")
                review_count = sum(1 for r in results if r["status"] == "REVIEW")
                missing_count = sum(1 for r in results if r["status"] == "NOT RECEIVED")

                st.markdown("---")
                st.markdown("#### Checklist Review")
                col_a, col_b, col_c, col_d = st.columns(4)
                col_a.metric("Pass", pass_count)
                col_b.metric("Fail", fail_count)
                col_c.metric("Review", review_count)
                col_d.metric("Not Received", missing_count)

                st.markdown("---")

                # Build table data
                table_rows = []
                for r in results:
                    issues_text = "; ".join(i["text"] for i in r.get("issues", []))
                    table_rows.append({
                        "Status": r["status"],
                        "Checklist Item": r["item_name"],
                        "Certificate": r["certificate"],
                        "Issues": issues_text if issues_text else ("—" if r["status"] == "PASS" else ""),
                    })

                st.dataframe(
                    table_rows,
                    use_container_width=True,
                    hide_index=True,
                    column_config={
                        "Status": st.column_config.TextColumn(width="small"),
                        "Checklist Item": st.column_config.TextColumn(width="medium"),
                        "Certificate": st.column_config.TextColumn(width="medium"),
                        "Issues": st.column_config.TextColumn(width="large"),
                    },
                )

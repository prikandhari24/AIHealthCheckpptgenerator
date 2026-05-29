from pathlib import Path

import streamlit as st

from src.config import OUTPUT_DIR, UPLOAD_DIR
from src.env_loader import load_env_file
from src.excel_parser import normalize_label, parse_excel
from src.logo_processor import process_logos
from src.ppt_builder import build_ppt


APP_DIR = Path(__file__).resolve().parent
load_env_file(APP_DIR / ".env")


def suggest_company_for_file(filename: str, companies: list[str], fallback_index: int = 0) -> str:
    normalized_filename = normalize_label(Path(filename).stem)
    exact_match = next(
        (company for company in companies if normalize_label(company) == normalized_filename),
        None,
    )
    if exact_match:
        return exact_match

    filename_tokens = set(normalized_filename.split())
    best_company = companies[0]
    best_score = -1
    for company in companies:
        company_tokens = set(normalize_label(company).split())
        score = len(filename_tokens & company_tokens)
        if score > best_score:
            best_company = company
            best_score = score
    if best_score <= 0:
        return companies[min(fallback_index, len(companies) - 1)]
    return best_company


st.set_page_config(
    page_title="AI Health Check PPT Generator",
    layout="wide",
)

st.title("AI Health Check PPT Generator")

OUTPUT_DIR.mkdir(exist_ok=True)
UPLOAD_DIR.mkdir(exist_ok=True)

excel_file = st.file_uploader(
    "Upload Excel file",
    type=["xlsx"],
)

front_slide_logo = st.file_uploader(
    "Upload front slide logo (optional)",
    type=["png", "jpg", "jpeg"],
    help="Use this if the cover slide should show a different logo than the client logo used in the score slides.",
)

client_logo = st.file_uploader(
    "Upload client logo",
    type=["png", "jpg", "jpeg"],
)

peer_logos = st.file_uploader(
    "Upload peer logos",
    type=["png", "jpg", "jpeg"],
    accept_multiple_files=True,
)

parsed_data = None
if excel_file:
    try:
        parsed_data = parse_excel(excel_file)
        st.success("Excel parsed successfully")
        st.dataframe(parsed_data.preview_df)
    except Exception as error:
        st.error(str(error))

client_company = None
peer_logo_assignments: dict[str, object] = {}
include_callouts = True

if parsed_data:
    client_default = 0
    if client_logo is not None:
        suggested_client = suggest_company_for_file(
            client_logo.name,
            parsed_data.companies,
        )
        client_default = parsed_data.companies.index(suggested_client)

    client_company = st.selectbox(
        "Which Excel company is the client?",
        options=parsed_data.companies,
        index=client_default,
    )

    peer_companies = [company for company in parsed_data.companies if company != client_company]
    st.caption(
        "Upload one logo for each peer company and map each upload to the correct company."
    )

    if peer_logos:
        for index, uploaded_logo in enumerate(peer_logos):
            suggested_company = suggest_company_for_file(
                uploaded_logo.name,
                peer_companies,
                fallback_index=index,
            )
            default_index = peer_companies.index(suggested_company)
            selected_company = st.selectbox(
                f"Peer logo mapping for {uploaded_logo.name}",
                options=peer_companies,
                index=default_index,
                key=f"peer_logo_mapping_{index}",
            )
            peer_logo_assignments[selected_company] = uploaded_logo

        missing_peer_companies = [
            company for company in peer_companies if company not in peer_logo_assignments
        ]
        duplicate_mappings = len(peer_logo_assignments) != len(peer_logos)

        if duplicate_mappings:
            st.warning("Each uploaded peer logo must be mapped to a different peer company.")
        elif missing_peer_companies:
            st.warning(
                "Missing peer logos for: " + ", ".join(missing_peer_companies)
            )

    st.divider()
    include_callouts = st.checkbox(
        "Include callouts",
        value=True,
    )

if st.button("Generate PPT"):
    if parsed_data is None:
        st.error("Please upload a valid Excel file.")
    elif client_logo is None:
        st.error("Please upload the client logo.")
    elif client_company is None:
        st.error("Please choose the client company from the Excel file.")
    else:
        peer_companies = [company for company in parsed_data.companies if company != client_company]

        if not peer_logos:
            st.error("Please upload peer logos.")
        elif len(peer_logos) != len(peer_companies):
            st.error(
                f"Please upload {len(peer_companies)} peer logos, one for each peer company."
            )
        elif set(peer_logo_assignments) != set(peer_companies):
            st.error("Please map each peer logo to a different peer company.")
        else:
            try:
                title_logo_path, client_logo_path, peer_logo_paths = process_logos(
                    front_slide_logo,
                    client_logo,
                    peer_logo_assignments,
                )
                ppt_path = build_ppt(
                    parsed_data,
                    client_company,
                    title_logo_path,
                    client_logo_path,
                    peer_logo_paths,
                    include_callouts=include_callouts,
                )

                with open(ppt_path, "rb") as file_handle:
                    st.download_button(
                        label="Download generated PPT",
                        data=file_handle.read(),
                        file_name="AI_Health_Check_Output.pptx",
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    )
            except Exception as error:
                st.error(str(error))

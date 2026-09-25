
import io
import re
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Extraction Opposite Party",
    page_icon="📱",
    layout="wide"
)

EXPECTED_COLUMNS = [
    "Receipt No.",
    "Completion Time",
    "Initiation Time",
    "Details",
    "Transaction Status",
    "Currency",
    "Paid In",
    "Withdrawn",
    "Balance",
    "Reason Type",
    "Opposite Party",
    "Linked Transaction ID",
    "Remark",
]


def read_transaction_file(uploaded_file):
    """Read one MVola export with the same structure as the source files."""
    filename = uploaded_file.name.lower()

    if filename.endswith(".xls"):
        engine = "xlrd"
    elif filename.endswith(".xlsx"):
        engine = "openpyxl"
    else:
        raise ValueError("Format non supporté. Utilisez un fichier .xls ou .xlsx.")

    df = pd.read_excel(
        uploaded_file,
        header=5,
        engine=engine
    )

    # Clean column labels
    df.columns = [
        str(col).replace("\t", " ").strip()
        if pd.notna(col)
        else ""
        for col in df.columns
    ]

    # Some exports put all header labels inside the first cell.
    # If the expected names are not available but the file has the expected
    # number/order of columns, restore the known schema by position.
    if (
        "Transaction Status" not in df.columns
        or "Opposite Party" not in df.columns
    ):
        if df.shape[1] >= len(EXPECTED_COLUMNS):
            new_columns = list(df.columns)
            for i, name in enumerate(EXPECTED_COLUMNS):
                new_columns[i] = name
            df.columns = new_columns
        else:
            raise ValueError(
                "Impossible d'identifier les colonnes "
                "'Transaction Status' et 'Opposite Party'."
            )

    return df


def extract_numbers_from_files(uploaded_files):
    """Filter Completed transactions and extract the first 10-digit number."""
    frames = []
    errors = []

    for uploaded_file in uploaded_files:
        try:
            df = read_transaction_file(uploaded_file)

            completed = df[
                df["Transaction Status"]
                .astype("string")
                .str.replace("\t", "", regex=False)
                .str.strip()
                .str.casefold()
                .eq("completed")
            ].copy()

            opposite_party = (
                completed["Opposite Party"]
                .astype("string")
                .str.replace("\t", "", regex=False)
                .str.strip()
            )

            # Extract the first sequence of 10 digits.
            # Example:
            # "2694486407 - Saindou Razda" -> "2694486407"
            numbers = opposite_party.str.extract(r"(\d{10})", expand=False)

            result = pd.DataFrame({
                "Numero": numbers
            }).dropna(subset=["Numero"])

            frames.append(result)

        except Exception as exc:
            errors.append(f"{uploaded_file.name}: {exc}")

    if frames:
        final_df = pd.concat(frames, ignore_index=True)
    else:
        final_df = pd.DataFrame(columns=["Numero"])

    return final_df, errors


def dataframe_to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Numeros")
    output.seek(0)
    return output.getvalue()


st.title("📱 Extraction des numéros Opposite Party")

st.write(
    "Importez un ou plusieurs exports Excel. "
    "L'application garde uniquement les transactions **Completed**, "
    "récupère les **10 premiers chiffres** de `Opposite Party`, "
    "identifie les numéros commençant par **411** comme numéros staff, "
    "puis permet de visualiser et télécharger le résultat."
)

uploaded_files = st.file_uploader(
    "Sélectionnez vos fichiers",
    type=["xls", "xlsx"],
    accept_multiple_files=True
)

if "result_df" not in st.session_state:
    st.session_state.result_df = None

if "staff_removed" not in st.session_state:
    st.session_state.staff_removed = False


if st.button(
    "⚙️ Traitement",
    type="primary",
    disabled=not uploaded_files
):
    result_df, errors = extract_numbers_from_files(uploaded_files)

    st.session_state.result_df = result_df
    st.session_state.staff_removed = False

    if errors:
        for error in errors:
            st.warning(error)


if st.session_state.result_df is not None:

    current_df = st.session_state.result_df.copy()

    staff_mask = current_df["Numero"].astype("string").str.startswith("411")
    staff_df = current_df.loc[staff_mask].copy()

    col1, col2, col3 = st.columns(3)

    col1.metric("Numéros extraits", f"{len(current_df):,}".replace(",", " "))
    col2.metric(
        "Numéros uniques",
        f"{current_df['Numero'].nunique():,}".replace(",", " ")
    )
    col3.metric(
        "Numéros staff",
        f"{len(staff_df):,}".replace(",", " ")
    )

    st.divider()

    st.subheader("Numéro staff")

    if not staff_df.empty:
        st.dataframe(
            staff_df,
            use_container_width=True,
            hide_index=True
        )

        if st.button(
            "🗑️ Supprimer les numéros staff",
            type="secondary"
        ):
            st.session_state.result_df = (
                current_df.loc[~staff_mask]
                .reset_index(drop=True)
            )
            st.session_state.staff_removed = True
            st.rerun()
    else:
        if st.session_state.staff_removed:
            st.success("Les numéros staff ont été supprimés du résultat.")
        else:
            st.info("Aucun numéro commençant par 411 n'a été trouvé.")

    st.divider()

    final_df = st.session_state.result_df.copy()

    st.subheader("Résultat")

    st.dataframe(
        final_df,
        use_container_width=True,
        hide_index=True,
        height=500
    )

    csv_data = final_df.to_csv(index=False).encode("utf-8-sig")
    excel_data = dataframe_to_excel(final_df)

    download_col1, download_col2 = st.columns(2)

    download_col1.download_button(
        label="⬇️ Télécharger en CSV",
        data=csv_data,
        file_name="opposite_party_numeros.csv",
        mime="text/csv",
        use_container_width=True
    )

    download_col2.download_button(
        label="⬇️ Télécharger en Excel",
        data=excel_data,
        file_name="opposite_party_numeros.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
        use_container_width=True
    )

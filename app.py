import io
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
    """
    Lit un fichier Excel.

    Certains fichiers ont l'extension .xls
    alors que leur vrai format interne est XLSX.
    On essaie donc openpyxl puis xlrd.
    """

    raw = uploaded_file.getvalue()

    df = None
    errors = []

    for engine in ("openpyxl", "xlrd"):
        try:
            df = pd.read_excel(
                io.BytesIO(raw),
                header=5,
                engine=engine
            )
            break

        except Exception as exc:
            errors.append(f"{engine}: {exc}")

    if df is None:
        raise ValueError(
            "Impossible de lire ce fichier Excel. "
            + " | ".join(errors)
        )

    # Nettoyage des noms de colonnes
    df.columns = [
        str(col).replace("\t", " ").strip()
        if pd.notna(col)
        else ""
        for col in df.columns
    ]

    # Certains exports ont un problème avec la ligne des headers.
    # Si les colonnes attendues ne sont pas reconnues,
    # on restaure leurs noms selon leur position.
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

    all_results = []
    errors = []

    for uploaded_file in uploaded_files:

        try:

            df = read_transaction_file(uploaded_file)

            # -----------------------------------
            # 1. Garder uniquement Completed
            # -----------------------------------

            completed = df[
                df["Transaction Status"]
                .astype("string")
                .str.replace("\t", "", regex=False)
                .str.strip()
                .str.casefold()
                .eq("completed")
            ].copy()


            # -----------------------------------
            # 2. Nettoyer Opposite Party
            # -----------------------------------

            opposite_party = (
                completed["Opposite Party"]
                .astype("string")
                .str.replace("\t", "", regex=False)
                .str.strip()
            )


            # -----------------------------------
            # 3. Extraire les 10 premiers chiffres
            # -----------------------------------

            numbers = opposite_party.str.extract(
                r"(\d{10})",
                expand=False
            )


            # -----------------------------------
            # 4. Créer le résultat
            # -----------------------------------

            result = pd.DataFrame({
                "Numero": numbers
            })

            result = result.dropna(
                subset=["Numero"]
            )

            all_results.append(result)


        except Exception as exc:

            errors.append(
                f"{uploaded_file.name}: {exc}"
            )


    # -----------------------------------
    # Concaténation de tous les fichiers
    # -----------------------------------

    if all_results:

        final_df = pd.concat(
            all_results,
            ignore_index=True
        )

    else:

        final_df = pd.DataFrame(
            columns=["Numero"]
        )


    return final_df, errors


def dataframe_to_excel(df):

    output = io.BytesIO()

    with pd.ExcelWriter(
        output,
        engine="openpyxl"
    ) as writer:

        df.to_excel(
            writer,
            index=False,
            sheet_name="Numeros"
        )

    output.seek(0)

    return output.getvalue()


# ======================================================
# INTERFACE
# ======================================================

st.title("📱 Extraction des numéros Opposite Party")

st.write(
    """
    Importez un ou plusieurs fichiers Excel.

    L'application va :

    - garder uniquement les transactions **Completed**
    - récupérer `Opposite Party`
    - extraire les **10 premiers chiffres**
    - identifier les numéros commençant par **411**
    - permettre de supprimer les numéros staff
    - générer un fichier final téléchargeable
    """
)


# ======================================================
# UPLOAD
# ======================================================

uploaded_files = st.file_uploader(
    "Sélectionnez vos fichiers Excel",
    type=["xls", "xlsx"],
    accept_multiple_files=True
)


# ======================================================
# SESSION STATE
# ======================================================

if "result_df" not in st.session_state:

    st.session_state.result_df = None


if "staff_removed" not in st.session_state:

    st.session_state.staff_removed = False


# ======================================================
# TRAITEMENT
# ======================================================

if st.button(
    "⚙️ Traitement",
    type="primary",
    disabled=not uploaded_files
):

    result_df, errors = extract_numbers_from_files(
        uploaded_files
    )

    st.session_state.result_df = result_df

    st.session_state.staff_removed = False


    # Afficher les erreurs éventuelles
    if errors:

        for error in errors:

            st.warning(error)


# ======================================================
# RESULTATS
# ======================================================

if st.session_state.result_df is not None:

    current_df = (
        st.session_state.result_df.copy()
    )


    # ==================================================
    # NUMEROS STAFF
    # ==================================================

    staff_mask = (
        current_df["Numero"]
        .astype("string")
        .str.startswith("411")
    )

    staff_df = current_df.loc[
        staff_mask
    ].copy()


    # ==================================================
    # KPI
    # ==================================================

    col1, col2, col3 = st.columns(3)


    col1.metric(
        "Numéros extraits",
        f"{len(current_df):,}".replace(",", " ")
    )


    col2.metric(
        "Numéros uniques",
        f"{current_df['Numero'].nunique():,}".replace(",", " ")
    )


    col3.metric(
        "Numéros staff",
        f"{len(staff_df):,}".replace(",", " ")
    )


    st.divider()


    # ==================================================
    # AFFICHAGE STAFF
    # ==================================================

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
                current_df.loc[
                    ~staff_mask
                ]
                .reset_index(drop=True)
            )

            st.session_state.staff_removed = True

            st.rerun()


    else:

        if st.session_state.staff_removed:

            st.success(
                "Les numéros staff ont été supprimés."
            )

        else:

            st.info(
                "Aucun numéro commençant par 411 n'a été trouvé."
            )


    st.divider()


    # ==================================================
    # RESULTAT FINAL
    # ==================================================

    final_df = (
        st.session_state.result_df.copy()
    )


    st.subheader("Résultat final")


    st.dataframe(
        final_df,
        use_container_width=True,
        hide_index=True,
        height=500
    )


    # ==================================================
    # EXPORT CSV
    # ==================================================

    csv_data = (
        final_df
        .to_csv(index=False)
        .encode("utf-8-sig")
    )


    # ==================================================
    # EXPORT EXCEL
    # ==================================================

    excel_data = dataframe_to_excel(
        final_df
    )


    col_csv, col_excel = st.columns(2)


    col_csv.download_button(
        label="⬇️ Télécharger en CSV",
        data=csv_data,
        file_name="opposite_party_numeros.csv",
        mime="text/csv",
        use_container_width=True
    )


    col_excel.download_button(
        label="⬇️ Télécharger en Excel",
        data=excel_data,
        file_name="opposite_party_numeros.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

import streamlit as st
import pandas as pd
import io

# Configuration de la page
st.set_page_config(page_title="Comparateur de Cartes / Ledgers", page_icon="📊", layout="wide")

st.title("📊 השוואת כרטיסיות / Ledgers Comparison")
st.write("העלה קובץ Excel עם 2 גיליונות כדי לבצע הצלבה ולהפיק דוח דלתאות.")

# 1. Widget pour téléverser le fichier Excel
uploaded_file = st.file_uploader("בחר קובץ Excel (עם לפחות 2 גיליונות)", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        xls = pd.ExcelFile(uploaded_file)
        sheet_names = xls.sheet_names

        if len(sheet_names) < 2:
            st.error("⚠️ הקובץ חייב להכיל לפחות 2 גיליונות!")
        else:
            # Sélection des deux premières feuilles
            df1 = pd.read_excel(xls, sheet_name=0)
            df2 = pd.read_excel(xls, sheet_name=1)

            st.success(f"✅ הקובץ נטען בהצלחה! משווים בין: **{sheet_names[0]}** ל-**{sheet_names[1]}**")

            # 2. Matching selon 'תאריך', 'אסמכתא', 'חובה', 'זכות'
            cols_to_match = ['תאריך', 'אסמכתא', 'חובה', 'זכות']

            # Vérification que les colonnes existent bien dans les deux feuilles
            missing_cols1 = [c for c in cols_to_match if c not in df1.columns]
            missing_cols2 = [c for c in cols_to_match if c not in df2.columns]

            if missing_cols1 or missing_cols2:
                st.error(f"⚠️ חסרות עמודות חובה במעקב: {missing_cols1 or missing_cols2}")
            else:
                # Mouvement présents dans Feuille 1 mais manquants dans Feuille 2
                merged1 = df1.merge(df2, on=cols_to_match, how='left', indicator=True)
                only_in_sheet1 = merged1[merged1['_merge'] == 'left_only'].drop(columns=['_merge'])

                # Mouvement présents dans Feuille 2 mais manquants dans Feuille 1
                merged2 = df2.merge(df1, on=cols_to_match, how='left', indicator=True)
                only_in_sheet2 = merged2[merged2['_merge'] == 'left_only'].drop(columns=['_merge'])

                # 3. Résumé
                summary_data = {
                    'מדד': [
                        'סה"כ שורות',
                        'סה"כ חובה',
                        'סה"כ זכות',
                        'תנועות ייחודיות (חסרות בגיליון השני)'
                    ],
                    sheet_names[0]: [
                        len(df1),
                        df1['חובה'].sum(),
                        df1['זכות'].sum(),
                        len(only_in_sheet1)
                    ],
                    sheet_names[1]: [
                        len(df2),
                        df2['חובה'].sum(),
                        df2['זכות'].sum(),
                        len(only_in_sheet2)
                    ]
                }
                df_summary = pd.DataFrame(summary_data)

                # Affichage des Métriques / Résumé
                st.subheader("📋 סיכום מנהלים")
                st.dataframe(df_summary, use_container_width=True)

                col1, col2 = st.columns(2)
                col1.metric(f"חורגים ב-{sheet_names[0]}", len(only_in_sheet1))
                col2.metric(f"חורגים ב-{sheet_names[1]}", len(only_in_sheet2))

                # Affichage des détails sous forme d'onglets
                tab1, tab2 = st.tabs([f"קיימות ב-{sheet_names[0]} בלבד", f"קיימות ב-{sheet_names[1]} בלבד"])

                with tab1:
                    st.dataframe(only_in_sheet1, use_container_width=True)

                with tab2:
                    st.dataframe(only_in_sheet2, use_container_width=True)

                # 4. Génération du rapport Excel en mémoire pour téléchargement
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    df_summary.to_excel(writer, sheet_name="סיכום מנהלים", index=False)
                    only_in_sheet1.to_excel(writer, sheet_name=f"רק ב-{sheet_names[0]}", index=False)
                    only_in_sheet2.to_excel(writer, sheet_name=f"רק ב-{sheet_names[1]}", index=False)

                # Bouton de téléchargement
                st.download_button(
                    label="📥 הורד דוח השוואה ב-Excel",
                    data=buffer.getvalue(),
                    file_name="report_comparison_summary.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

    except Exception as e:
        st.error(f"שגיאה בעיבוד הקובץ: {e}")

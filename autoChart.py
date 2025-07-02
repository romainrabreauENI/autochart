# app.py
import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from pptx import Presentation
from pptx.util import Inches
from io import BytesIO

st.set_page_config(page_title="Générateur de slide", layout="centered")
st.title("Générateur de slide PowerPoint")

uploaded = st.file_uploader("Déposez votre fichier Excel (.xlsx)", type="xlsx")
if uploaded:
    df = pd.read_excel(uploaded)
    df['Mois'] = pd.to_datetime(df['Mois'])
    df = df.sort_values('Mois')
    df['Durée'] = (
        df['Temps effectif total']
          .str.split(':')
          .apply(lambda x: int(x[0]) + int(x[1]) / 60)
    )
    pivot = (
        df
        .pivot_table(index='Mois',
                     columns='E-formation',
                     values='Durée',
                     aggfunc='sum')
        .reindex(sorted(df['Mois'].unique()), fill_value=0)
        .cumsum()
    )
    pivot = pivot.reset_index()
    pivot['Mois_str'] = pivot['Mois'].dt.strftime('%b %Y')

    # --- Tracé Matplotlib ---
    fig, ax = plt.subplots(figsize=(10,6))
    ax.plot(pivot['Mois_str'], pivot['Bibliothèque Numérique ENI'],
            marker='o', label='BNE')
    ax.plot(pivot['Mois_str'], pivot['Bibliothèque Numérique ENI e-formations'],
            marker='s', label='BNE e-formations')
    ax.fill_between(pivot['Mois_str'], pivot['Bibliothèque Numérique ENI'], alpha=0.2)
    ax.fill_between(pivot['Mois_str'], pivot['Bibliothèque Numérique ENI e-formations'], alpha=0.2)
    ax.set_xlabel("Date")
    ax.set_ylabel("Heures cumulées")
    ax.set_title("Temps effectif cumulé par e-formation")
    plt.xticks(rotation=45)
    plt.tight_layout()

    st.pyplot(fig)

    # --- Préparation du buffer image pour PowerPoint ---
    img_buf = BytesIO()
    fig.savefig(img_buf, format="png", dpi=200)
    img_buf.seek(0)

    prs = Presentation("TemplateGraph.pptx")
    slide = prs.slides[0]
    w, h = Inches(9), Inches(6)
    left = (prs.slide_width - w) // 2
    top  = (prs.slide_height - h) // 2
    slide.shapes.add_picture(img_buf, left, top, width=w, height=h)

    out_buf = BytesIO()
    prs.save(out_buf)
    out_buf.seek(0)

    st.success("Votre diapositive est prête :")
    st.download_button(
        "Télécharger la slide PPTX",
        data=out_buf.getvalue(),
        file_name="Présentation_Générée.pptx",
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
    )
else:
    st.info("En attente de votre fichier…")


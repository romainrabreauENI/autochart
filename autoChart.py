import streamlit as st
import pandas as pd
import plotly.express as px
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
    df['Durée'] = df['Temps effectif total'] \
        .str.split(':') \
        .apply(lambda x: int(x[0]) + int(x[1]) / 60)
    pivot = (
        df
        .pivot_table(index='Mois', columns='E-formation', values='Durée', aggfunc='sum')
        .reindex(sorted(df['Mois'].unique()), fill_value=0)
        .cumsum()
        .reset_index()
    )
    pivot['Mois_str'] = pivot['Mois'].dt.strftime('%b %Y')

    fig = px.area(
        pivot,
        x='Mois_str',
        y=['Bibliothèque Numérique ENI', 'Bibliothèque Numérique ENI e-formations'],
        title="Temps effectif cumulé par e-formation",
        labels={'Mois_str':'Date','value':'Heures','variable':'E-formation'},
        template='plotly_white'
    )
    fig.update_traces(mode='lines+markers')
    fig.update_layout(
        xaxis_tickangle=-45,
        margin=dict(l=60, r=60, t=80, b=120),
        legend=dict(orientation="h", x=0.5, xanchor="center", y=1.1)
    )

    st.plotly_chart(fig, use_container_width=True)

    png_bytes = fig.to_image(format='png', width=1200, height=600, scale=2)
    img_buf = BytesIO(png_bytes)

    prs = Presentation("TemplateGraph.pptx")
    slide = prs.slides[0]
    slide_w, slide_h = prs.slide_width, prs.slide_height
    w, h = Inches(9), Inches(6)
    left = (slide_w - w) // 2
    top  = (slide_h - h) // 2
    slide.shapes.add_picture(img_buf, left, top, width=w, height=h)

    out_buf = BytesIO()
    prs.save(out_buf)
    out_buf.seek(0)

    st.download_button(
        "Télécharger la slide PPTX",
        data=out_buf.getvalue(),
        file_name="Présentation_Générée.pptx",
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
    )
else:
    st.info("En attente de votre fichier…")

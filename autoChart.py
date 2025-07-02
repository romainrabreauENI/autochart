import streamlit as st
import pandas as pd
from io import BytesIO
from pptx import Presentation
from pptx.chart.data import ChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.util import Inches, Pt

st.set_page_config(page_title="Générateur de slide", layout="centered")
st.title("Générateur de slide PowerPoint")

uploaded = st.file_uploader("Déposez votre fichier Excel (.xlsx)", type="xlsx")
if uploaded:
    df = pd.read_excel(uploaded)
    df['Mois'] = pd.to_datetime(df['Mois'])
    df = df.sort_values('Mois')
    df['Durée'] = df['Temps effectif total'].str.split(':').apply(lambda x: int(x[0]) + int(x[1]) / 60)
    pivot = (
        df
        .pivot_table(index='Mois', columns='E-formation', values='Durée', aggfunc='sum')
        .reindex(sorted(df['Mois'].unique()), fill_value=0)
        .cumsum()
        .reset_index()
    )
    pivot['Mois_str'] = pivot['Mois'].dt.strftime('%b %Y')

    prs = Presentation("TemplateGraph.pptx")
    slide = prs.slides[0]
    chart_data = ChartData()
    chart_data.categories = pivot['Mois_str']
    chart_data.add_series("Bibliothèque Numérique ENI", pivot.get("Bibliothèque Numérique ENI", []))
    chart_data.add_series("Bibliothèque Numérique ENI e-formations", pivot.get("Bibliothèque Numérique ENI e-formations", []))
    slide_w, slide_h = prs.slide_width, prs.slide_height
    w, h = Inches(9), Inches(5)
    left = (slide_w - w) // 2
    top = (slide_h - h) // 2
    chart = slide.shapes.add_chart(XL_CHART_TYPE.LINE, left, top, w, h, chart_data).chart
    chart.has_title = True
    chart.chart_title.text_frame.text = "Temps effectif cumulé par e-formation"
    chart.has_legend = True
    chart.legend.position = XL_LEGEND_POSITION.BOTTOM
    chart.legend.include_in_layout = False
    cat_ax = chart.category_axis
    cat_ax.tick_label_rotation = -45
    cat_ax.tick_labels.font.size = Pt(10)
    val_ax = chart.value_axis
    val_ax.has_major_gridlines = True

    out = BytesIO()
    prs.save(out)
    out.seek(0)

    st.download_button(
        "Télécharger la slide PPTX",
        data=out.getvalue(),
        file_name="Présentation_Générée.pptx",
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
    )
else:
    st.info("En attente du fichier Excel")

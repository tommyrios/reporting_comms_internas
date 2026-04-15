from pathlib import Path
from typing import Any
from pptx import Presentation
from analyzer import validate_report_json

def create_pptx(report: dict[str, Any], output_path: Path):
    prs = Presentation()
    safe_report = validate_report_json(report)
    
    try:
        slide = prs.slides.add_slide(prs.slide_layouts[0])
        s1 = safe_report["slide_1_cover"]
        slide.shapes.title.text = "Informe de Gestión\n" + str(s1.get("area", ""))
        slide.placeholders[1].text = str(s1.get("period", ""))
    except: pass

    try:
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        slide.shapes.title.text = "2. Visión General y Audiencia"
        tf = slide.placeholders[1].text_frame
        s2 = safe_report["slide_2_overview"]
        tf.text = f"Evolución del Volumen:\n- Actual: {s2.get('volume_current')}\n- Anterior: {s2.get('volume_previous')}\n- Variación: {s2.get('volume_change')}"
        p = tf.add_paragraph()
        p.text = "\nSegmentación:"
        for seg in s2.get("audience_segments", []):
            p = tf.add_paragraph()
            p.text = f"- {seg.get('label')}: {seg.get('value')}%"
            p.level = 1
        p = tf.add_paragraph()
        p.text = f"\nConclusión: {s2.get('conclusion_message')}"
    except: pass

    try:
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        slide.shapes.title.text = "3. Ejes Estratégicos"
        tf = slide.placeholders[1].text_frame
        s3 = safe_report["slide_3_strategy"]
        tf.text = "Temáticas:"
        for dist in s3.get("content_distribution", []):
            p = tf.add_paragraph()
            p.text = f"- {dist.get('theme')}: {dist.get('weight')}%"
            p.level = 1
        p = tf.add_paragraph()
        p.text = "\nBalance:"
        bal = s3.get('canal_balance', {})
        p.text = f"- Institucional: {bal.get('institutional', 0)}%\n- Transaccional: {bal.get('transactional_talent', 0)}%"
    except: pass

    try:
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        slide.shapes.title.text = "4. Ranking Push"
        tf = slide.placeholders[1].text_frame
        s4 = safe_report["slide_4_push_ranking"]
        tf.text = "Top Comunicaciones:"
        for i, comm in enumerate(s4.get("top_communications", [])):
            p = tf.add_paragraph()
            p.text = f"#{i+1}: {comm.get('name')} (Clics: {comm.get('clicks')})"
            p.level = 1
        p = tf.add_paragraph()
        p.text = f"\nAprendizaje: {s4.get('key_learning')}"
    except: pass

    try:
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        slide.shapes.title.text = "5. Desempeño Pull"
        tf = slide.placeholders[1].text_frame
        s5 = safe_report["slide_5_pull_performance"]
        tf.text = f"Publicaciones: {s5.get('pub_current')}\nPromedio Lecturas: {s5.get('avg_reads')} | Vistas: {s5.get('total_views')}"
        p = tf.add_paragraph()
        p.text = "\nTop Notas:"
        for note in s5.get("top_notes", []):
            p = tf.add_paragraph()
            p.text = f"- {note.get('title')} (Únicas: {note.get('unique_reads')})"
            p.level = 1
    except: pass

    try:
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        slide.shapes.title.text = "6. Hitos"
        tf = slide.placeholders[1].text_frame
        tf.text = "Destacados:"
        for hit in safe_report["slide_6_hitos"]:
            p = tf.add_paragraph()
            p.text = f"{hit.get('quarter')}: {hit.get('description')}"
            p.level = 1
    except: pass

    try:
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        slide.shapes.title.text = "7. Eventos"
        tf = slide.placeholders[1].text_frame
        s7 = safe_report["slide_7_events"]
        tf.text = f"Eventos: {s7.get('total_events')} | Participaciones: {s7.get('total_participants')}\n\nCierre: {s7.get('conclusion')}"
    except: pass

    prs.save(str(output_path))
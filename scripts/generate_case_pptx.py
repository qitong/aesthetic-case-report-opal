#!/usr/bin/env python3
"""
Aesthetic Case Report Opal - PPTX Generator Engine
Generates clinical-grade 16:9 PowerPoint presentations for medical aesthetics case reports,
supporting both Literature-Integrated Mode and Standalone Clinical Patient Journey Mode.
"""

import json
import os
import sys
from pathlib import Path
from PIL import Image, ImageOps
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE


def hex_to_rgb(hex_str: str) -> RGBColor:
    hex_str = hex_str.lstrip("#")
    return RGBColor(int(hex_str[0:2], 16), int(hex_str[2:4], 16), int(hex_str[4:6], 16))


DEFAULT_PALETTES = {
    "academic": {
        "bg": "#F9F8F6",
        "card_bg": "#FFFFFF",
        "card_border": "#E5E0D8",
        "text_primary": "#1A1D20",
        "text_secondary": "#5A626A",
        "accent": "#2C4C64",
        "accent_soft": "#EBF0F5",
        "highlight": "#C0392B",
        "font_cn": "PingFang SC",
        "font_en": "Helvetica Neue",
    },
    "editorial": {
        "bg": "#FAF7F2",
        "card_bg": "#FFFFFF",
        "card_border": "#E8DFD5",
        "text_primary": "#231F20",
        "text_secondary": "#7A7067",
        "accent": "#B8926A",
        "accent_soft": "#F3ECE3",
        "highlight": "#8C6D53",
        "font_cn": "PingFang SC",
        "font_en": "Helvetica Neue",
    },
}


class OpalCaseReportDeck:
    def __init__(self, theme_mode="editorial", custom_theme=None):
        self.prs = Presentation()
        self.prs.slide_width = Inches(13.333)
        self.prs.slide_height = Inches(7.5)
        self.blank_layout = self.prs.slide_layouts[6]

        base = DEFAULT_PALETTES.get(theme_mode, DEFAULT_PALETTES["editorial"]).copy()
        if custom_theme:
            base.update(custom_theme)
        self.theme = base
        self._temp_files = []

    def _apply_background(self, slide, bg_hex=None):
        bg_color = hex_to_rgb(bg_hex or self.theme["bg"])
        bg = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(7.5)
        )
        bg.fill.solid()
        bg.fill.fore_color.rgb = bg_color
        bg.line.fill.background()
        return bg

    def _add_header(self, slide, title_text: str, category_tag: str = "", subtitle_text: str = ""):
        if category_tag:
            tag_box = slide.shapes.add_textbox(Inches(0.6), Inches(0.38), Inches(10), Inches(0.3))
            tf = tag_box.text_frame
            tf.word_wrap = True
            tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
            p = tf.paragraphs[0]
            p.text = category_tag.upper()
            p.font.size = Pt(10)
            p.font.bold = True
            p.font.color.rgb = hex_to_rgb(self.theme["accent"])
            p.font.name = self.theme["font_cn"]

        title_top = Inches(0.65) if category_tag else Inches(0.45)
        title_box = slide.shapes.add_textbox(Inches(0.6), title_top, Inches(12.1), Inches(0.55))
        tf = title_box.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
        p = tf.paragraphs[0]
        p.text = title_text
        p.font.size = Pt(24)
        p.font.bold = True
        p.font.color.rgb = hex_to_rgb(self.theme["text_primary"])
        p.font.name = self.theme["font_cn"]

        if subtitle_text:
            sub_box = slide.shapes.add_textbox(Inches(0.6), title_top + Inches(0.48), Inches(12.1), Inches(0.35))
            tf_sub = sub_box.text_frame
            tf_sub.word_wrap = True
            tf_sub.margin_left = tf_sub.margin_top = tf_sub.margin_right = tf_sub.margin_bottom = 0
            p_sub = tf_sub.paragraphs[0]
            p_sub.text = subtitle_text
            p_sub.font.size = Pt(12)
            p_sub.font.color.rgb = hex_to_rgb(self.theme["text_secondary"])
            p_sub.font.name = self.theme["font_cn"]

    def _add_footer(self, slide, footer_text: str = "", slide_num: int = None):
        if footer_text:
            f_box = slide.shapes.add_textbox(Inches(0.6), Inches(7.05), Inches(10.5), Inches(0.3))
            tf = f_box.text_frame
            tf.word_wrap = True
            tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
            p = tf.paragraphs[0]
            p.text = footer_text
            p.font.size = Pt(9)
            p.font.color.rgb = hex_to_rgb(self.theme["text_secondary"])
            p.font.name = self.theme["font_cn"]

        if slide_num is not None:
            n_box = slide.shapes.add_textbox(Inches(11.8), Inches(7.05), Inches(0.933), Inches(0.3))
            tf = n_box.text_frame
            tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.RIGHT
            p.text = f"{slide_num:02d}"
            p.font.size = Pt(9)
            p.font.bold = True
            p.font.color.rgb = hex_to_rgb(self.theme["accent"])
            p.font.name = self.theme["font_en"]

    def _add_image_fitted(self, slide, img_path: str, left: float, top: float, width: float, height: float, crop_to_fill: bool = True):
        """
        Inserts an image into (left, top, width, height) in inches without distorting aspect ratio.
        If crop_to_fill is True, center-crops the image to match the target aspect ratio.
        Otherwise, fits inside the box and centers it.
        """
        if not img_path or not os.path.exists(img_path):
            placeholder = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE, Inches(left), Inches(top), Inches(width), Inches(height)
            )
            placeholder.fill.solid()
            placeholder.fill.fore_color.rgb = hex_to_rgb(self.theme["accent_soft"])
            placeholder.line.color.rgb = hex_to_rgb(self.theme["card_border"])
            tf = placeholder.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER
            p.text = f"[Image Placeholder]\n{os.path.basename(img_path) if img_path else 'No Image'}"
            p.font.size = Pt(11)
            p.font.color.rgb = hex_to_rgb(self.theme["text_secondary"])
            return placeholder

        with Image.open(img_path) as im:
            im = ImageOps.exif_transpose(im)
            if crop_to_fill:
                target_ratio = width / height
                img_ratio = im.width / im.height
                if img_ratio > target_ratio:
                    new_width = int(im.height * target_ratio)
                    offset = (im.width - new_width) // 2
                    im_cropped = im.crop((offset, 0, offset + new_width, im.height))
                else:
                    new_height = int(im.width / target_ratio)
                    offset = (im.height - new_height) // 2
                    im_cropped = im.crop((0, offset, im.width, offset + new_height))

                temp_path = f"/tmp/opal_crop_{len(self._temp_files)}_{os.path.basename(img_path)}.png"
                im_cropped.save(temp_path, format="PNG")
                self._temp_files.append(temp_path)
                return slide.shapes.add_picture(temp_path, Inches(left), Inches(top), Inches(width), Inches(height))
            else:
                target_ratio = width / height
                img_ratio = im.width / im.height
                if img_ratio > target_ratio:
                    w = width
                    h = width / img_ratio
                    x = left
                    y = top + (height - h) / 2
                else:
                    h = height
                    w = height * img_ratio
                    x = left + (width - w) / 2
                    y = top
                return slide.shapes.add_picture(img_path, Inches(x), Inches(y), Inches(w), Inches(h))

    def add_cover_slide(self, title: str, subtitle: str = "", author: str = "", affiliation: str = "", tag: str = "AESTHETIC CASE REPORT"):
        slide = self.prs.slides.add_slide(self.blank_layout)
        self._apply_background(slide)

        accent_bar = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(0.6), Inches(2.2), Inches(0.8), Inches(0.06)
        )
        accent_bar.fill.solid()
        accent_bar.fill.fore_color.rgb = hex_to_rgb(self.theme["accent"])
        accent_bar.line.fill.background()

        tag_box = slide.shapes.add_textbox(Inches(0.6), Inches(1.75), Inches(10), Inches(0.35))
        p_tag = tag_box.text_frame.paragraphs[0]
        p_tag.text = tag.upper()
        p_tag.font.size = Pt(11)
        p_tag.font.bold = True
        p_tag.font.color.rgb = hex_to_rgb(self.theme["accent"])
        p_tag.font.name = self.theme["font_en"]

        t_box = slide.shapes.add_textbox(Inches(0.6), Inches(2.5), Inches(11.5), Inches(1.8))
        tf = t_box.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = title
        p.font.size = Pt(36)
        p.font.bold = True
        p.font.color.rgb = hex_to_rgb(self.theme["text_primary"])
        p.font.name = self.theme["font_cn"]

        if subtitle:
            p2 = tf.add_paragraph()
            p2.text = subtitle
            p2.space_before = Pt(14)
            p2.font.size = Pt(18)
            p2.font.color.rgb = hex_to_rgb(self.theme["text_secondary"])
            p2.font.name = self.theme["font_cn"]

        if author or affiliation:
            meta_box = slide.shapes.add_textbox(Inches(0.6), Inches(5.8), Inches(11.5), Inches(0.9))
            tf_m = meta_box.text_frame
            tf_m.word_wrap = True
            if author:
                pa = tf_m.paragraphs[0]
                pa.text = author
                pa.font.size = Pt(14)
                pa.font.bold = True
                pa.font.color.rgb = hex_to_rgb(self.theme["text_primary"])
                pa.font.name = self.theme["font_cn"]
            if affiliation:
                paf = tf_m.add_paragraph() if author else tf_m.paragraphs[0]
                paf.text = affiliation
                paf.font.size = Pt(11)
                paf.font.color.rgb = hex_to_rgb(self.theme["text_secondary"])
                paf.font.name = self.theme["font_cn"]

    def add_section_divider(self, index_str: str, title_cn: str, title_en: str = ""):
        slide = self.prs.slides.add_slide(self.blank_layout)
        self._apply_background(slide)

        idx_box = slide.shapes.add_textbox(Inches(1.0), Inches(2.4), Inches(3.0), Inches(1.2))
        p_idx = idx_box.text_frame.paragraphs[0]
        p_idx.text = index_str
        p_idx.font.size = Pt(54)
        p_idx.font.bold = True
        p_idx.font.color.rgb = hex_to_rgb(self.theme["accent"])
        p_idx.font.name = self.theme["font_en"]

        t_box = slide.shapes.add_textbox(Inches(1.0), Inches(3.6), Inches(10.5), Inches(1.5))
        tf = t_box.text_frame
        tf.word_wrap = True
        p_cn = tf.paragraphs[0]
        p_cn.text = title_cn
        p_cn.font.size = Pt(32)
        p_cn.font.bold = True
        p_cn.font.color.rgb = hex_to_rgb(self.theme["text_primary"])
        p_cn.font.name = self.theme["font_cn"]

        if title_en:
            p_en = tf.add_paragraph()
            p_en.text = title_en
            p_en.space_before = Pt(8)
            p_en.font.size = Pt(16)
            p_en.font.color.rgb = hex_to_rgb(self.theme["text_secondary"])
            p_en.font.name = self.theme["font_en"]

    def add_assessment_slide(
        self,
        title: str,
        patient_label: str,
        photo_path: str,
        scores: list,
        clinical_notes: list,
        footer: str = "",
        slide_num: int = None,
    ):
        """
        Creates a Pre- or Post-Treatment Facial Assessment Slide.
        scores: list of dicts [{"domain": "Expression (Static / Dynamic)", "score": 2, "max": 3, "note": "..."}, ...]
        """
        slide = self.prs.slides.add_slide(self.blank_layout)
        self._apply_background(slide)
        self._add_header(slide, title, category_tag="CLINICAL ASSESSMENT", subtitle_text=patient_label)

        # Left Photo Card
        card_left = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(0.6), Inches(1.5), Inches(4.5), Inches(5.2)
        )
        card_left.fill.solid()
        card_left.fill.fore_color.rgb = hex_to_rgb(self.theme["card_bg"])
        card_left.line.color.rgb = hex_to_rgb(self.theme["card_border"])

        self._add_image_fitted(slide, photo_path, 0.75, 1.65, 4.2, 4.4, crop_to_fill=True)

        lbl_box = slide.shapes.add_textbox(Inches(0.75), Inches(6.15), Inches(4.2), Inches(0.4))
        p_lbl = lbl_box.text_frame.paragraphs[0]
        p_lbl.alignment = PP_ALIGN.CENTER
        p_lbl.text = patient_label
        p_lbl.font.size = Pt(11)
        p_lbl.font.bold = True
        p_lbl.font.color.rgb = hex_to_rgb(self.theme["text_primary"])
        p_lbl.font.name = self.theme["font_cn"]

        # Right Assessment Scales & Notes
        y_cursor = 1.55
        for item in scores[:5]:
            domain = item.get("domain", "")
            score = int(item.get("score", 0))
            max_score = int(item.get("max", 3))
            note = item.get("note", "")

            row_card = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE, Inches(5.4), Inches(y_cursor), Inches(7.333), Inches(0.72)
            )
            row_card.fill.solid()
            row_card.fill.fore_color.rgb = hex_to_rgb(self.theme["card_bg"])
            row_card.line.color.rgb = hex_to_rgb(self.theme["card_border"])

            tb = slide.shapes.add_textbox(Inches(5.6), Inches(y_cursor + 0.1), Inches(3.8), Inches(0.5))
            tf = tb.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.text = domain
            p.font.size = Pt(12)
            p.font.bold = True
            p.font.color.rgb = hex_to_rgb(self.theme["text_primary"])
            p.font.name = self.theme["font_cn"]
            if note:
                p2 = tf.add_paragraph()
                p2.text = note
                p2.font.size = Pt(9.5)
                p2.font.color.rgb = hex_to_rgb(self.theme["text_secondary"])
                p2.font.name = self.theme["font_cn"]

            # Draw 0..max_score blocks
            for s_idx in range(max_score + 1):
                bx = 9.6 + s_idx * 0.72
                bar = slide.shapes.add_shape(
                    MSO_SHAPE.ROUNDED_RECTANGLE, Inches(bx), Inches(y_cursor + 0.18), Inches(0.6), Inches(0.36)
                )
                bar.fill.solid()
                is_active = s_idx == score
                bar.fill.fore_color.rgb = hex_to_rgb(
                    self.theme["accent"] if is_active else self.theme["accent_soft"]
                )
                bar.line.fill.background()
                bp = bar.text_frame.paragraphs[0]
                bp.alignment = PP_ALIGN.CENTER
                bp.text = str(s_idx)
                bp.font.size = Pt(11)
                bp.font.bold = is_active
                bp.font.color.rgb = hex_to_rgb("#FFFFFF" if is_active else self.theme["text_secondary"])

            y_cursor += 0.82

        if clinical_notes:
            notes_card = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE, Inches(5.4), Inches(y_cursor + 0.05), Inches(7.333), Inches(6.7 - y_cursor - 0.05)
            )
            notes_card.fill.solid()
            notes_card.fill.fore_color.rgb = hex_to_rgb(self.theme["accent_soft"])
            notes_card.line.fill.background()

            ntb = slide.shapes.add_textbox(
                Inches(5.6), Inches(y_cursor + 0.12), Inches(6.9), Inches(6.6 - y_cursor - 0.12)
            )
            ntf = ntb.text_frame
            ntf.word_wrap = True
            for idx, note_line in enumerate(clinical_notes):
                np = ntf.paragraphs[0] if idx == 0 else ntf.add_paragraph()
                np.text = f"• {note_line}"
                np.space_after = Pt(4)
                np.font.size = Pt(11)
                np.font.color.rgb = hex_to_rgb(self.theme["text_primary"])
                np.font.name = self.theme["font_cn"]

        self._add_footer(slide, footer, slide_num)

    def add_treatment_plan_slide(
        self,
        title: str,
        subtitle: str,
        diagram_path: str,
        steps: list,
        footer: str = "",
        slide_num: int = None,
    ):
        """
        Creates a Treatment Plan & Rheological Mapping slide.
        steps: list of dicts [{"step_title": "Step 1: Deep Support", "product": "Restylane Lyft", "plane": "DMCFP / Supraperiosteal", "instrument": "22G 50mm Cannula", "dosage": "Left 0.7ml / Right 0.6ml", "rationale": "..."}]
        """
        slide = self.prs.slides.add_slide(self.blank_layout)
        self._apply_background(slide)
        self._add_header(slide, title, category_tag="TREATMENT PLAN & RHEOLOGY", subtitle_text=subtitle)

        # Left Diagram / Annotated Patient Photo
        self._add_image_fitted(slide, diagram_path, 0.6, 1.55, 4.6, 5.15, crop_to_fill=False)

        # Right Step Cards
        card_height = min(2.45, 5.15 / max(len(steps), 1) - 0.12)
        y_cursor = 1.55
        for step in steps:
            card = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE, Inches(5.5), Inches(y_cursor), Inches(7.233), Inches(card_height)
            )
            card.fill.solid()
            card.fill.fore_color.rgb = hex_to_rgb(self.theme["card_bg"])
            card.line.color.rgb = hex_to_rgb(self.theme["card_border"])

            stripe = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE, Inches(5.5), Inches(y_cursor), Inches(0.12), Inches(card_height)
            )
            stripe.fill.solid()
            stripe.fill.fore_color.rgb = hex_to_rgb(self.theme["accent"])
            stripe.line.fill.background()

            tb = slide.shapes.add_textbox(Inches(5.8), Inches(y_cursor + 0.12), Inches(6.7), Inches(card_height - 0.2))
            tf = tb.text_frame
            tf.word_wrap = True

            p_title = tf.paragraphs[0]
            p_title.text = step.get("step_title", "")
            p_title.font.size = Pt(14)
            p_title.font.bold = True
            p_title.font.color.rgb = hex_to_rgb(self.theme["accent"])
            p_title.font.name = self.theme["font_cn"]

            details = [
                f"产品与流变学 (Product): {step.get('product', 'N/A')}",
                f"靶区与层次 (Target & Plane): {step.get('plane', 'N/A')}",
                f"器械与剂量 (Instrument & Dosage): {step.get('instrument', '')} | {step.get('dosage', '')}",
            ]
            if step.get("rationale"):
                details.append(f"操作要点 (Rationale): {step.get('rationale')}")

            for line in details:
                p = tf.add_paragraph()
                p.text = f"• {line}"
                p.space_before = Pt(3)
                p.font.size = Pt(11)
                p.font.color.rgb = hex_to_rgb(self.theme["text_primary"])
                p.font.name = self.theme["font_cn"]

            y_cursor += card_height + 0.15

        self._add_footer(slide, footer, slide_num)

    def add_before_after_slide(
        self,
        title: str,
        view_label: str,
        before_img: str,
        after_img: str,
        before_caption: str = "Before (术前)",
        after_caption: str = "After (术后)",
        summary_metrics: str = "",
        footer: str = "",
        slide_num: int = None,
    ):
        slide = self.prs.slides.add_slide(self.blank_layout)
        self._apply_background(slide)
        self._add_header(slide, title, category_tag=f"BEFORE & AFTER COMPARISON · {view_label}")

        img_w = 5.85
        img_h = 4.55
        top_y = 1.35

        # Before Container
        self._add_image_fitted(slide, before_img, 0.6, top_y, img_w, img_h, crop_to_fill=True)
        b_lbl = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(0.6), Inches(top_y + img_h), Inches(img_w), Inches(0.42)
        )
        b_lbl.fill.solid()
        b_lbl.fill.fore_color.rgb = hex_to_rgb(self.theme["card_bg"])
        b_lbl.line.color.rgb = hex_to_rgb(self.theme["card_border"])
        bp = b_lbl.text_frame.paragraphs[0]
        bp.alignment = PP_ALIGN.CENTER
        bp.text = before_caption
        bp.font.size = Pt(12)
        bp.font.bold = True
        bp.font.color.rgb = hex_to_rgb(self.theme["text_secondary"])
        bp.font.name = self.theme["font_cn"]

        # After Container
        right_x = 13.333 - 0.6 - img_w
        self._add_image_fitted(slide, after_img, right_x, top_y, img_w, img_h, crop_to_fill=True)
        a_lbl = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(right_x), Inches(top_y + img_h), Inches(img_w), Inches(0.42)
        )
        a_lbl.fill.solid()
        a_lbl.fill.fore_color.rgb = hex_to_rgb(self.theme["accent"])
        a_lbl.line.fill.background()
        ap = a_lbl.text_frame.paragraphs[0]
        ap.alignment = PP_ALIGN.CENTER
        ap.text = after_caption
        ap.font.size = Pt(12)
        ap.font.bold = True
        ap.font.color.rgb = hex_to_rgb("#FFFFFF")
        ap.font.name = self.theme["font_cn"]

        if summary_metrics:
            m_box = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE, Inches(0.6), Inches(6.42), Inches(12.133), Inches(0.48)
            )
            m_box.fill.solid()
            m_box.fill.fore_color.rgb = hex_to_rgb(self.theme["accent_soft"])
            m_box.line.fill.background()
            mp = m_box.text_frame.paragraphs[0]
            mp.alignment = PP_ALIGN.CENTER
            mp.text = summary_metrics
            mp.font.size = Pt(11)
            mp.font.bold = True
            mp.font.color.rgb = hex_to_rgb(self.theme["text_primary"])
            mp.font.name = self.theme["font_cn"]

        self._add_footer(slide, footer, slide_num)

    def add_patient_journey_slide(
        self,
        title: str,
        subtitle: str,
        milestones: list,
        footer: str = "",
        slide_num: int = None,
    ):
        """
        Creates a multi-timepoint longitudinal Patient Journey slide (supports 3 to 5 milestones).
        milestones: [{"date": "2024 (Baseline)", "image": "path.jpg", "note": "Initial assessment"}, ...]
        """
        slide = self.prs.slides.add_slide(self.blank_layout)
        self._apply_background(slide)
        self._add_header(slide, title, category_tag="LONGITUDINAL PATIENT JOURNEY", subtitle_text=subtitle)

        n = max(1, min(len(milestones), 5))
        total_w = 12.133
        gap = 0.22
        col_w = (total_w - gap * (n - 1)) / n
        img_h = 4.15
        top_y = 1.65

        for idx, m in enumerate(milestones[:5]):
            x = 0.6 + idx * (col_w + gap)
            self._add_image_fitted(slide, m.get("image", ""), x, top_y, col_w, img_h, crop_to_fill=True)

            badge = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE, Inches(x), Inches(top_y + img_h + 0.08), Inches(col_w), Inches(0.38)
            )
            badge.fill.solid()
            is_last = idx == n - 1
            badge.fill.fore_color.rgb = hex_to_rgb(self.theme["accent"] if is_last else self.theme["card_bg"])
            badge.line.color.rgb = hex_to_rgb(self.theme["accent"] if is_last else self.theme["card_border"])
            bp = badge.text_frame.paragraphs[0]
            bp.alignment = PP_ALIGN.CENTER
            bp.text = m.get("date", f"T{idx}")
            bp.font.size = Pt(11)
            bp.font.bold = True
            bp.font.color.rgb = hex_to_rgb("#FFFFFF" if is_last else self.theme["text_primary"])
            bp.font.name = self.theme["font_cn"]

            if m.get("note"):
                ntb = slide.shapes.add_textbox(Inches(x), Inches(top_y + img_h + 0.5), Inches(col_w), Inches(0.65))
                ntf = ntb.text_frame
                ntf.word_wrap = True
                np = ntf.paragraphs[0]
                np.alignment = PP_ALIGN.CENTER
                np.text = m.get("note")
                np.font.size = Pt(9.5)
                np.font.color.rgb = hex_to_rgb(self.theme["text_secondary"])
                np.font.name = self.theme["font_cn"]

        self._add_footer(slide, footer, slide_num)

    def add_consensus_table_slide(
        self,
        title: str,
        subtitle: str,
        headers: list,
        rows: list,
        takeaway: str = "",
        footer: str = "",
        slide_num: int = None,
    ):
        slide = self.prs.slides.add_slide(self.blank_layout)
        self._apply_background(slide)
        self._add_header(slide, title, category_tag="EVIDENCE & ANATOMICAL RATIONALE", subtitle_text=subtitle)

        n_rows = len(rows) + 1
        n_cols = len(headers)
        table_shape = slide.shapes.add_table(
            n_rows, n_cols, Inches(0.6), Inches(1.65), Inches(12.133), Inches(4.5)
        )
        table = table_shape.table

        for c_idx, h_text in enumerate(headers):
            cell = table.cell(0, c_idx)
            cell.fill.solid()
            cell.fill.fore_color.rgb = hex_to_rgb(self.theme["accent"])
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            p = cell.text_frame.paragraphs[0]
            p.text = h_text
            p.font.size = Pt(12)
            p.font.bold = True
            p.font.color.rgb = hex_to_rgb("#FFFFFF")
            p.font.name = self.theme["font_cn"]

        for r_idx, row_data in enumerate(rows):
            for c_idx, val in enumerate(row_data):
                cell = table.cell(r_idx + 1, c_idx)
                cell.fill.solid()
                cell.fill.fore_color.rgb = hex_to_rgb(
                    self.theme["card_bg"] if r_idx % 2 == 0 else self.theme["accent_soft"]
                )
                cell.vertical_anchor = MSO_ANCHOR.MIDDLE
                p = cell.text_frame.paragraphs[0]
                p.text = str(val)
                p.font.size = Pt(11)
                p.font.bold = c_idx == 0
                p.font.color.rgb = hex_to_rgb(self.theme["text_primary"])
                p.font.name = self.theme["font_cn"]

        if takeaway:
            t_box = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE, Inches(0.6), Inches(6.35), Inches(12.133), Inches(0.5)
            )
            t_box.fill.solid()
            t_box.fill.fore_color.rgb = hex_to_rgb(self.theme["accent_soft"])
            t_box.line.fill.background()
            tp = t_box.text_frame.paragraphs[0]
            tp.alignment = PP_ALIGN.LEFT
            tp.text = f"  核心提示: {takeaway}"
            tp.font.size = Pt(11)
            tp.font.bold = True
            tp.font.color.rgb = hex_to_rgb(self.theme["accent"])
            tp.font.name = self.theme["font_cn"]

        self._add_footer(slide, footer, slide_num)

    def save(self, output_path: str):
        os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
        self.prs.save(output_path)
        for tmp in self._temp_files:
            if os.path.exists(tmp):
                try:
                    os.remove(tmp)
                except OSError:
                    pass
        return output_path


if __name__ == "__main__":
    if len(sys.argv) > 1 and os.path.exists(sys.argv[1]):
        with open(sys.argv[1], "r", encoding="utf-8") as f:
            spec = json.load(f)
        deck = OpalCaseReportDeck(theme_mode=spec.get("theme_mode", "editorial"), custom_theme=spec.get("custom_theme"))
        out_file = spec.get("output_path", "case_report.pptx")
        deck.save(out_file)
        print(f"Saved {out_file}")

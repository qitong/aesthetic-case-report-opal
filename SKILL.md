---
name: aesthetic-case-report-opal
description: Generate high-aesthetic clinical case report PPTX presentations for medical aesthetics and cosmetic dermatology. Use when the user asks to create a medical aesthetics case report slide deck, convert clinical patient records or before-and-after photos into a presentation, synthesize a clinical paper with a case report, or run the aesthetic-case-report-opal workflow.
---

# Aesthetic Case Report Opal

Generate clinical-grade, high-aesthetic PowerPoint (`.pptx`) presentations for medical aesthetics case reports, combining patient journeys, multi-angle before-and-after photography, anatomical and rheological rationale, and optional literature appraisal.

## When to Use

- The user wants to create a medical aesthetics or cosmetic dermatology case report PPTX.
- The user provides patient progress notes (病程记录), injection records, and before/after or longitudinal patient journey photos.
- The user wants to combine a clinical reference paper with a practical clinical case report into a cohesive presentation.
- The user invokes `aesthetic-case-report-opal` or asks for an aesthetic case presentation using default clinical templates or a custom `design.md`.

## Steps

### Step 1: Interactive Intake and Clarification

Before generating slides, inspect the files already provided in the conversation and ask a concise, structured set of questions to confirm any missing inputs:

1. **Input Sources Inventory**:
   - **Patient Record & Treatment Details**: Confirm the patient demographics, chief complaint, baseline assessment (e.g., Hirmand classification, 8-step lower eyelid checklist, or 5-axis Facial Assessment Scale), product selection, rheology, dosage per anatomical site (left vs. right), and instrument (needle vs. cannula gauge/length).
   - **Image Assets / Photo Directory**: Ask for the folder path or uploaded images containing Before & After photos (frontal, oblique 45-degree, lateral 90-degree, static vs. dynamic expression) and longitudinal Patient Journey photos (multi-timepoint follow-ups).
   - **Reference Literature (Optional)**: Ask whether the presentation is a **Standalone Clinical Case Report** or a **Literature-Integrated Case Report** (combining a published paper or anatomical technique study with the clinical case).

2. **Template & Visual Specification Mode**:
   - Ask whether the user wants to use the **Built-in Default Template** (and choose between **Mode A: Academic Literature & Anatomy Deep-Dive** or **Mode B: Longitudinal Holistic Aesthetic Journey** as defined in `references/default_template_spec.md`) OR provide a custom **`design.md`** file to govern the slide deck structure, color palette, typography, and layout.

3. **Privacy & De-identification**:
   - Confirm whether the supplied clinical photos are already de-identified. If the orbital/eye region is visible, apply a solid eye-mask band with `scripts/apply_eye_mask.py` **before any photo enters the deck**, unless the user explicitly provides written image consent and asks to keep the eyes visible.
   - This applies to AI-generated / simulated portraits too: a synthetic face is still shown as "the patient" in a clinical deck, and leaving it identifiable sets a wrong precedent for the workflow.
   - State the de-identification status in the slide footer and in the case-record documentation.

*Note: If the user provides all inputs upfront and specifies the template preference in their initial prompt, proceed directly to Step 2 without redundant questioning.*

### Step 2: Clinical Narrative & Structural Mapping

Read `references/default_template_spec.md` (or the user's custom `design.md`) and map the clinical data into a structured slide-by-slide blueprint:

- **Scenario 1: Literature + Case Report (Hybrid Academic Mode)**
  1. **Title & Metadata**: Bilingual title, authors/clinician, institution, evidence level or clinical focus.
  2. **Clinical Problem & Consensus**: The aesthetic/anatomical challenge (e.g., dynamic sausage deformity, retaining ligament tethering), mainstream consensus table (sequence, plane, dosage, rheology, instrument, danger zones).
  3. **Anatomical Mechanism & Classification**: Bone-to-skin layering, ligamentous attachments (e.g., AM vs. TTL/ORL), gliding spaces (preseptal, premaxillary, prezygomatic, DMCFP), and patient classification criteria (e.g., Innie vs. Outie, vector analysis).
  4. **Patient Baseline & Assessment**: Patient profile, pre-treatment multi-angle or dynamic assessment, checklist or severity scoring.
  5. **Two-Step / Multi-Layer Technique & Rheology**: Entry points (e.g., midface viaduct), deep structural support vs. superficial blending passes, cannula gauge, and product rheological matching.
  6. **Before & After Visual Comparison**: Standardized side-by-side static and dynamic comparisons with exact volume breakdowns (L/R ml).
  7. **Patient Journey / Follow-Up & Critical Appraisal**: Longitudinal timeline outcomes, complication monitoring, and clinical takeaways.

- **Scenario 2: Standalone Clinical Case Report (Holistic Clinical Mode)**
  1. **Cover Slide**: High-aesthetic editorial cover, patient case title, clinician credentials.
  2. **Aging Rationale & Anatomical Analysis**: Multi-layer facial aging assessment (bone, deep fat pads, SMAS/ligaments, superficial fat, dermis).
  3. **Patient Profile & Pre-Treatment Assessment**: Baseline photos paired with structured multi-axis scoring (Expression, Symmetry, Skin Quality, Proportions, Volume Loss/Sagging on a 0-3 scale).
  4. **Holistic Treatment Plan**: Visual facial mapping with product portfolio, exact injection sites, target planes, and L/R volumes (e.g., Zygomatic arch, Maxilla, NLF, Chin, Cheeks).
  5. **Post-Treatment Assessment Shift**: Pre- vs. Post-treatment severity scale comparison demonstrating objective improvement.
  6. **Multi-Angle Before & After Gallery**: Dedicated slides for Frontal, 45-degree Oblique, and 90-degree Profile views (static and animated/smiling).
  7. **Longitudinal Patient Journey**: Multi-timepoint chronological progression (e.g., Baseline -> Immediate / 1M -> 6M -> 1Y -> 2Y+) showcasing long-term tissue integration and natural rejuvenation.

### Step 3: Visual Design & Layout Execution

Apply strict medical aesthetics visual standards (detailed in `references/default_template_spec.md` unless overridden by `design.md`):

- **Aspect Ratio**: 16:9 widescreen (`13.333` x `7.5` inches).
- **Color Palette**: High-contrast editorial minimalism. Warm Ivory/Alabaster background (`#FAF8F5` or `#FFFFFF`), Deep Charcoal text (`#1E2022`), Muted Champagne/Rose Gold accents (`#C5A059` or `#8C6D53`), and Clinical Slate (`#4A607A`) for anatomical callouts.
- **Image Integrity**: Never stretch or distort clinical photographs. Preserve original aspect ratios using center-cropped containers or uniform height alignment. Always align interpupillary lines (Frankfurt horizontal plane) across Before and After pairs.
- **Typography**: Clean sans-serif hierarchy (e.g., `Helvetica Neue` / `Arial` / `PingFang SC` / `Microsoft YaHei`), generous whitespace, and clear source citations in 9-10pt footer text.

### Step 4: Patient Photo De-identification (Mandatory)

Clinical photography is identifiable personal data. Before building any slide, run `scripts/apply_eye_mask.py` on every patient photo that shows the eye region:

```bash
# Inspect the plan first, then apply in place with untouched originals preserved
python3 scripts/apply_eye_mask.py <photo_dir> --dry-run
python3 scripts/apply_eye_mask.py <photo_dir> --in-place --backup-dir <photo_dir>/_unmasked
```

- The default band spans **33%–49% of image height**, full width, solid `#1E2022`, fully opaque. On standardized frontal and 45-degree oblique portraits this covers brows, eyes and the nasal root without covering the neck or other chief-complaint anatomy.
- For lateral 90-degree profiles, pediatric cases, or unusual cropping, inspect one image first and override `--top` / `--bottom`. Use `--output-dir` for a non-destructive copy instead of `--in-place`.
- **Never** use a feathered, translucent, blurred, or rounded band for de-identification: soft edges can leak the orbital contour. Keep `--opacity 1.0` and `--radius 0`.
- Keep unmasked originals in a separate backup directory outside the deck's image paths; never embed them in the PPTX. Re-running `--in-place` never overwrites an existing backup, so the pristine originals survive.
- Add `眼部已脱敏 / Eyes masked` to the slide footer and to the case-record documentation.

### Step 5: Programmatic PPTX Generation

Use `scripts/generate_case_pptx.py` as the foundation or write a tailored `python-pptx` script to build the presentation:

1. Load and validate **only the de-identified images from Step 4** using `PIL.Image` to read dimensions and prevent aspect-ratio distortion.
2. Construct each slide programmatically using the layout helpers in `scripts/generate_case_pptx.py`.
3. Save the output `.pptx` file and make it downloadable for the user.
4. Verify that every slide has zero overlapping text boxes, properly scaled clinical images, and complete dosage/anatomical annotations.

## Gotchas

- **Clinical Photo Distortion**: Passing both `width` and `height` to `slide.shapes.add_picture()` in `python-pptx` without pre-cropping stretches the patient's face, which invalidates clinical Before & After comparisons. Always calculate aspect ratio via `PIL` and crop to a standardized bounding box or specify only one dimension.
- **Identifiable Eyes in Clinical Photos**: A full-face clinical photo exposes the orbital region, a biometric identifier. Never place an unmasked eye region into a shareable case deck, even for AI-generated/simulated portraits. Run `scripts/apply_eye_mask.py` and state the de-identification status in the footer.
- **Mask Geometry vs Target Anatomy**: A band that is too low covers the nose/mouth and weakens facial assessment; one that is too high leaves the lower lid or brow visible. Verify on one frontal and one oblique image that the band fully covers the orbital region while leaving the chief-complaint anatomy (e.g. neck lines) unobstructed.
- **Missing Dynamic vs. Static Distinction**: In periorbital and midface filler cases, static photos alone hide dynamic ridges ("sausage deformity"). Explicitly label whether clinical photos are taken at rest (静态) or in animation/smiling (动态/微笑), and note when dynamic photos are absent.
- **Bilateral Dosage Ambiguity**: Never list a single total volume without specifying whether it is unilateral (per side) or bilateral (total), or breaking down Left (L) vs. Right (R) volumes for facial asymmetry correction.
- **Anatomical Layer Precision**: Avoid vague terms like "injected into the cheek." Specify the exact anatomical compartment or gliding space (e.g., DMCFP, preseptal space, prezygomatic space, supraperiosteal vs. sub-orbicularis) and whether ligaments (TTL/ORL) were crossed or respected.

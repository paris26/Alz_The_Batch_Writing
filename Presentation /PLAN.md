# Thesis Presentation: AI Alzheimer and Dementia Classification

## Context

Build a 32-slide presentation from the 88-page thesis "AI Alzheimer and Dementia Classification" (University of West Attica, Gavriilidis Paraskevas). The user wants a story-driven, image-heavy, citation-rich, aesthetically unique presentation using Spectral font and editorial/web design principles. The existing 4-slide example in `Presentation /` shows the desired style: large images, visible citations, clean layout with decorative separators.

---

## Design Philosophy: "Editorial Magazine" Style

What makes this different from typical academic slides:
- **Asymmetric layouts** -- visual tension, not centered boredom
- **Oversized hero numbers** (72-96pt) as design anchors for data slides
- **Full-bleed images** -- brain scans dominate, not trapped in tiny boxes
- **Alternating dark/light backgrounds** -- dark (#0D1117) for neuroimaging, light (#F7F5F0) for data/text
- **No bullet points anywhere** -- prose, labeled panels, or visuals only
- **Architectural citation strips** -- every slide has a bottom citation bar, grounding credibility
- **Color-coded severity** -- Copper (neutral), Red (limitations), Green (future)

---

## Color Palette

| Role | Hex | Usage |
|------|-----|-------|
| Dark Background | #0D1117 | Brain scan slides, section dividers |
| Light Background | #F7F5F0 | Data slides, text-heavy slides |
| Copper/Amber | #C17F3A | Titles, accent lines, highlights |
| Clinical Blue | #3B82B6 | Imaging/MRI content |
| Alert Red | #DC4A4A | Limitations, warnings |
| Sage Green | #5B8C6B | Future directions, positive results |
| Text (dark bg) | #E8E4DE | Body text on dark |
| Text (light bg) | #2D2D2D | Body text on light |
| Citation Gray | #8B8680 | All citation text |

## Typography (Spectral throughout)

| Element | Weight | Size |
|---------|--------|------|
| Title | Bold | 28-32pt |
| Hero Number | Bold | 72-96pt |
| Body | Regular | 14-16pt |
| Callout | SemiBold Italic | 18-20pt |
| Citation | Light Italic | 9-10pt |
| Section Label | Bold CAPS | 10pt |

---

## 5 Layout Templates

1. **Hero Statistic** -- Oversized number (left 40%), context text (right), thin accent line between
2. **Full-Bleed Image** -- Image covers 60-65% of slide, text overlays gradient
3. **Split Compare** -- Two columns with thin vertical divider
4. **Dark Canvas** -- Near-black bg, centered image, text beside/below
5. **Statement** -- One powerful sentence centered, minimal supporting text

---

## 32-Slide Outline

### ACT I: THE PROBLEM (Slides 1-6)

| # | Title | Layout | Background | Image(s) | Key Content |
|---|-------|--------|------------|----------|-------------|
| 1 | Title Slide | Custom | Dark | `logo_en.png` | Title, subtitle, author, university |
| 2 | "The Silent Epidemic" | Hero Stat | Light | Generated: prevalence chart | "7.2M" hero number, economic burden |
| 3 | "The 20-Year Window" | Statement | Dark | None | "Pathology begins 15-20 years before first symptom" |
| 4 | "Biomarker Criteria" | Hero variant | Light | None | AT(N) framework: Amyloid, Tau, Neurodegeneration |
| 5 | "Why Neuroimaging" | Full-Bleed | Image-driven | `nihms-137059-f0004.jpg` | Non-invasive, quantifiable, objective |
| 6 | "The Datasets" | Split (3-col) | Light | None | ADNI, AIBL, OASIS panels |

### ACT II: THE SCIENCE (Slides 7-21)

| # | Title | Layout | Background | Image(s) | Key Content |
|---|-------|--------|------------|----------|-------------|
| 7 | Section Divider | Statement | Dark | None | "SEEING THE BRAIN" |
| 8 | "MRI Fundamentals" | Full-Bleed | Dark | `IntensityNormalization1.png` | MRI physics, Larmor equation |
| 9 | "CT and PET" | Split Compare | Light | `pmp-32-1-1-f1.png`, `pmp-32-1-1-f4.png` | CT vs PET comparison |
| 10 | "PET Biomarkers" | Hero variant | Light | None | FDG-PET (90% sensitivity), Amyloid PET, Tau PET |
| 11 | "Preprocessing Pipeline" | Dark Canvas | Dark | None | Visual pipeline flow with arrows |
| 12 | "Intensity Normalization" | Split Compare | Light | `IntensityNormalization2.png`, `IntensityNormalization3.png` | Z-Score, Histogram Matching, White Stripe |
| 13 | "Denoising" | Hero variant | Light | None | Mathematical formulation, NLM -> BM3D -> DL |
| 14 | "Skull Stripping" | Full-Bleed | Image-driven | `Skull Stripping image.png`, `Skull Stripping Techniques.png` | Deep learning methods, artifact warning |
| 15 | "Voxel-Based Morphometry" | Full-Bleed | Dark | `nihms154848f1.jpg` | Whole-brain analysis, AUC >0.90 |
| 16 | "Classical ML" | Hero Stat | Light | Generated: ML comparison chart | "94.5%" SVM accuracy, MCI cliff |
| 17 | "Deep Learning Revolution" | Dark Canvas | Dark | None | CNN architecture diagram, 2D vs 3D |
| 18 | "Vision Transformers & Hybrids" | Split Compare | Light | None | ViT vs Hybrid CNN+SVM |
| 19 | "Measuring Performance" | Hero variant | Light | Generated: ROC curve | Accuracy paradox, F1/AUC/Sensitivity |
| 20 | "Explainability" | Full-Bleed | Dark | `Grad-CAMVBM.png` | Grad-CAM, LIME, SHAP, 3 XAI pillars |
| 21 | "Our Experiments" | Full-Bleed | Dark | `limitations_gradcam.png` | Grad-CAM on 4 dementia stages |

### ACT III: THE RECKONING (Slides 22-30)

| # | Title | Layout | Background | Image(s) | Key Content |
|---|-------|--------|------------|----------|-------------|
| 22 | Section Divider | Statement | Dark | None | "THE INCONVENIENT TRUTHS" |
| 23 | "Data Leakage" | Hero Stat | Light | Generated: leakage chart | "-28%" hero, only 4.5% proper methodology |
| 24 | "Accuracy Paradox" | Full-Bleed | Light | `limitations_class_imbalance.png` | Class imbalance, 1% Moderate class |
| 25 | "Domain Shift" | Hero Stat | Light | None | "71%" accuracy on external data |
| 26 | "Shortcut Learning" | Dark Canvas | Dark | None | Correct vs shortcut reasoning diagram |
| 27 | "Label Uncertainty" | Hero Stat | Light | None | "71%" mixed pathology at autopsy |
| 28 | "Recent Advances" | Split Compare | Light | None | Multimodal fusion + augmentation/transfer learning |
| 29 | "Future Directions" | Dark Canvas | Dark | None | 3 pillars: Benchmarks, Interpretability, Multi-stage |
| 30 | "Methodological Triad" | Statement | Dark | None | Subject splitting, external validation, confounder control |

### BOOKENDS (Slides 31-32)

| # | Title | Layout | Background | Image(s) | Key Content |
|---|-------|--------|------------|----------|-------------|
| 31 | Conclusion | Statement | Dark | None | "Models powerful, data insufficient, trust unearned" |
| 32 | Thank You / Q&A | Custom | Dark | `logo_en.png` | Author info, GitHub link, Questions? |

---

## Charts to Generate (matplotlib)

1. **Slide 2**: AD Prevalence Projection -- line chart 7.2M (2025) to 13.8M (2060)
2. **Slide 16**: Classical ML Comparison -- grouped bars AD-vs-HC (~94%) vs MCI (~70%)
3. **Slide 19**: Conceptual ROC Curve -- good classifier vs chance
4. **Slide 23**: Data Leakage Impact -- two bars: with leakage (~95%) vs without (~67%)

All charts: 300 DPI, transparent background, presentation color palette, serif font.

---

## Image Assignment (from `/images/`)

| Image | Slide(s) |
|-------|----------|
| `logo_en.png` | 1, 32 |
| `nihms-137059-f0004.jpg` | 5 |
| `IntensityNormalization1.png` | 8 |
| `pmp-32-1-1-f1.png` | 9 |
| `pmp-32-1-1-f4.png` | 9 |
| `IntensityNormalization2.png` | 12 |
| `IntensityNormalization3.png` | 12 |
| `Skull Stripping image.png` | 14 |
| `Skull Stripping Techniques.png` | 14 |
| `nihms154848f1.jpg` | 15 |
| `Grad-CAMVBM.png` | 20 |
| `limitations_gradcam.png` | 21 |
| `limitations_class_imbalance.png` | 24 |
| `Graph.png` | 28 (transfer learning categorizations) |

Unused but available for supplementary: `nihms154848f2-4.jpg`, `nihms-137059-f0005.jpg`, `pmp-32-1-1-f2.png`

---

## Citation Strategy

- Every slide gets a **citation strip** at the bottom (full width, 9pt Spectral Light Italic, gray)
- Format: "Author et al., Year" style (not numbered)
- Sources: extracted from `references.bib` and supplemented with internet sources where needed
- The strip has a subtle background tint for visual consistency

---

## Implementation Notes

### Dependencies Required
```bash
# Python
uv pip install matplotlib seaborn --system

# Node.js (global)
npm install -g pptxgenjs playwright sharp
```

### Step 1: Generate charts
Create `Presentation /generated_charts/` directory. Run matplotlib scripts to produce 4 chart PNGs.

### Step 2: Build presentation with html2pptx workflow
Use the pptx skill's html2pptx workflow:
1. Create HTML files for each slide (720pt x 405pt, 16:9)
2. Use html2pptx.js to convert HTML to PowerPoint
3. Add charts via PptxGenJS to placeholder areas
4. Save to `Presentation /AI_Alzheimer_Thesis_Presentation.pptx`

### Step 3: Verification
- Generate thumbnails and inspect
- Check all images are embedded
- Verify citation strips on every content slide
- Confirm narrative flow

---

## Critical Constraints

- **Font**: Spectral is NOT web-safe. Use Georgia (closest web-safe serif) in HTML, then the PPTX will render with whatever serif is available. Alternatively, embed Spectral via raw XML post-processing.
- **No bullet points**: Use prose, labeled panels, or visuals only
- **No # prefix on hex colors** in PptxGenJS calls
- **All text must be in p/h1-h6/ul/ol tags** in HTML
- **No CSS gradients** - rasterize with Sharp first
- **Images**: Use absolute paths to `/home/paris/Code/Thesis_Finale/images/`

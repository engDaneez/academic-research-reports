---
name: academic-research-reports
description: Build, edit, audit, and professionally format Arabic or English academic reports, research papers, graduation research, and laboratory reports. Use for university reports, بحوث التخرج, التقارير الجامعية, writing and organizing sections, research methodology, verifiable citations, Word DOCX output, proofreading and preserving supplied content.
compatibility: Works with Codex and Claude Code agents that can read local files. Optional DOCX builder requires Python 3 and python-docx. For file rendering/visual checks use the host environment's available document tools.
metadata:
  version: "1.0.0"
---

# Academic Research & Reports Builder

An academically disciplined workflow for producing and reviewing **reports, research papers, graduation research, and lab reports** in Arabic or English. Treat the university / instructor's rubric as higher priority than this skill's default style.

## Activation and first move

Activate when asked to create, revise, organize, proofread, format, or check a university research/report document. Avoid generic filler. If files were supplied, read their entire relevant contents and requirements first. Distinguish **content creation**, **reorganization without shortening**, **format-only**, and **audit-only** requests.

Collect only details that materially affect delivery and are not already provided. See [intake](references/intake.md). In a new project ask for the essentials in one concise turn: exact topic/title, document type, language, university/instructor requirements or model, desired length, citation style, available sources/data, and requested output. If only one essential detail is missing, ask only that. When enough information exists, work immediately. When the user requests step-by-step collaboration, perform one task at a time rather than flooding them with future steps.

## Non-negotiable rules

1. **Instructor rubric wins.** Never present default font size, cover sequence, margins, pagination, section count, or citation style as universally mandated by universities.
2. **Preserve supplied work.** In edit/proofread/format mode, retain title, argument, facts, paragraphs and material scope. Do not shorten or remove sections unless explicitly requested; remove clear duplicates only if authorized. Never silently replace the original. Provide a changes summary.
3. **No fictional evidence.** Never manufacture quotations, research data, survey findings, bibliographic entries, DOI/URLs, page numbers, statistics, or references. If source verification isn't possible, label placeholders clearly, request sources or use citation-needed markers.
4. **Research integrity.** Distinguish confirmed facts from suggestions, user-provided results from generated illustration, and a plan from executed research. Do not claim plagiarism scores, Turnitin results, or approval without actual verification.
5. **References traceability.** Every substantive externally sourced assertion needs a traceable reference in the style requested; each bibliography item must be genuinely cited or an explicitly labeled further-reading entry. Consult [sources](references/sources-and-citations.md).
6. **Language and tone.** Write formal academic Arabic or English, not promotional language. Maintain consistent terminology and logical transitions. Use formal Iraqi academic conventions only when they match course requirements.
7. **Word is not just text.** Use real heading styles, section breaks, tables, captions, dynamic table of contents when possible, stable page numbering, and ensure right-to-left layout for Arabic. Never fake structural styles with font-size-only changes.
8. **Final QA.** Validate content, citation coverage, source authenticity, spelling, formatting, pagination, heading hierarchy, and output file readability before delivery. Visual-inspect a rendered DOCX/PDF using the host's available tools; a generated DOCX is not verified until inspected. See [QA](references/quality-control.md).
9. **Sensitive data.** Do not include the original student's personal details, institutional IDs, or non-public material in a public reusable skill or sample.

## Workflow: choose the correct lane

### A. New academic report or research
1. Review the prompt, rubric, any assigned readings or supplied research. Identify the document type and scope.
2. Build an **outline** matching the document type; consult [structures](references/structures.md). Confirm high-impact ambiguities when necessary but don't block a sufficiently defined task.
3. Create an evidence inventory: source title/author/date/URL, factual claims supported, gaps. Use reliable sources; prefer primary or scholarly sources when available. Do not fabricate publications.
4. Draft to the requested length: meaningful introduction, clear problem/objectives, suitable methodology, coherent chapters, analysis/discussion grounded in evidence, conclusions answering objectives, actionable recommendations, appropriate references and appendices.
5. Avoid repeating the same material between the introduction, theory, analysis, and conclusion. Do not pad to meet page count.
6. Apply the user's citation standard (APA/MLA/Chicago/IEEE, footnotes, numbered, etc.). Never change to APA automatically if the rubric specifies another style.
7. Assemble DOCX/PDF as requested; use real styles, title page, table of contents, captions, page numbering, and QA.

### B. Graduation thesis / extensive research
Use the same rules plus a structured front matter and methodology. A common **optional** sequence is cover → Qur'anic verse (if requested) → dedication → acknowledgments → abstract → contents → introduction and research methodology → chapters/sections → conclusions/recommendations → conclusion → references → appendices. Some rubrics order these elements differently; adapt exactly. See [structures](references/structures.md).

For Arabic research, headings may be `الباب` / `الفصل` / `المبحث` / `المطلب`, as instructed. Differentiate **الاستنتاجات** (findings) from **التوصيات** (actions), and **الخاتمة** (closing synthesis). Some university rubrics expect paired conclusions/recommendations; enforce equal counts only when requested.

### C. Laboratory report
Never invent experiments, measurements, apparatus, samples, statistical significance, safety conditions or outcomes. Prefer objective → background → materials → procedure → raw data → calculations → results → discussion of error/limitations → conclusions → references, customized for instructor requirements. Clearly mark missing measured values as required input.

### D. Edit/format existing document
1. Inventory and preserve original content and heading hierarchy. Identify exact rubric deviations.
2. Separate changes into spelling/grammar, organization, content gaps, citation problems, and layout fixes. Do **not** introduce new factual claims as if sourced.
3. Make minimal supported corrections. Preserve the author's argument and substantive content; ask before large conceptual revisions when the scope is editing only.
4. Produce edited copy and (when practical) a short before/after audit or changes report.

## Formatting defaults (only when user/instructor provided no rule)

Arabic academic baseline used in this workflow: **A4; Traditional Arabic; body 14 pt; Heading 1 = 18 pt bold; Heading 2 = 16 pt bold; 1.5 line spacing; justified text; 2.5 cm margins; right-to-left paragraphs; consistent page numbering**. Some universities require right margin 3 cm, page numbering below left, and preliminary pages with separate labels. Apply rubric; see [formatting](references/formatting.md). These are *defaults for this skill*, not universal academic standards. For English, use sensible requested/institutional typography instead of imposing Arabic defaults.

## Word artifact output

- Prefer using a trusted document authoring capability in the environment, with actual paragraph styles, captions, automatic numbering/table of contents, and visual PDF/PNG review.
- A **fallback DOCX generator** is bundled in `scripts/build_docx.py`: accepts UTF-8 JSON with sections and optional front matter and builds a Word file. Install from `requirements.txt`; see [README](README.md). For example:

  `python scripts/build_docx.py --input assets/example-report.json --output output/example-report.docx`

- The sample JSON is illustrative material, **not** a sourced research paper. Do not submit it as genuine academic evidence.
- The fallback generator uses a TOC field that may need refreshing inside Microsoft Word. Its advanced pagination capabilities are limited: for lettered preliminary pages, custom bibliography styles, precise citation fields, complex tables, Arabic footnotes, and cross-references, use advanced Word authoring/manual finishing instead of pretending they are fully implemented.

## Pre-delivery acceptance gate

Read [quality-control](references/quality-control.md). Deliver:
1. Final requested artifact (DOCX and/or PDF), not an unrendered promise.
2. A clear brief description of what was included or changed, including any source or data gaps.
3. Explicit notes about placeholders, unverified citations, TOC refresh, formatting differences or content that requires author/instructor review.

If operating without filesystem/tool access, deliver Word-ready structured text and explain which artifact cannot be generated.
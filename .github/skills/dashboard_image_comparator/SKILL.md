---
name: enterprise_dashboard_comparator
description: "Compare complete dashboard screenshots by business meaning, visible metrics, visual structure, layout, filters, rendering quality, and atomic element unrolling using pure spatial schema discovery."
version: 4.11.0
---

# Enterprise Dashboard Image Comparator

This skill enforces one single authoritative behavior contract:

[dashboard_image_comparator.prompt.md](../../../dashboard_image_comparator.prompt.md)

Load that canonical prompt before comparing dashboard images. It defines pure schema discovery, zero-hardcode visual extraction, semantic structural matching, weighted similarity scoring, anomaly severity rules, and report formats. Pixel diffing and raw OCR parsing are supporting techniques only, never primary decision rules.

## Inputs

- `image_a`: baseline or reference dashboard image
- `image_b`: current or test dashboard image
- `comparison_mode`: `strict`, `business`, or `regression`; default `business`
- `similarity_threshold`: percentage threshold; default `95`
- `output_format`: `markdown` or `json`; default `markdown`

Do not introduce parallel extraction schemas, visual-type classification branching, or conflicting output contracts here. Maintain all operational behavior within the canonical prompt.

## Consistent Power BI Text Extraction

- Use deterministic visual reading order, object ownership, naming, and null
	handling on every run for the same screenshot.
- Do not silently drop visible Power BI dashboard text or values. Preserve all
	visible KPI labels, filter selections, chart labels, legend labels, table
	headers, row labels, values, units, dates, and status text.
- Extract bracketed and parenthesized text exactly as displayed, including the
	brackets: `[All]`, `(Blank)`, `(>40 days)`, `[N/A]`, and similar text are
	meaningful content, not formatting noise.
- Do not replace bracketed text with an empty value, remove the brackets, or
	treat it as cache/debug text without visual evidence that it is non-business
	metadata. Preserve it in the owning object and comparison row.
- Logos, brand marks, watermarks, and decorative text are excluded unless the
	text independently functions as a dashboard title, filter, header, or
	business label.
- Partial OCR or crop fragments are not independent business objects. Do not
	emit fragments such as `Pri`, `eek`, `ate`, `Ag`, or `didates` as separate
	comparison rows when they belong to one bounded visual object. Keep the
	fragment in a `[TEXT_LOSS]` note and retain the owning object's value; preserve
	it as a label only when the dashboard boundary genuinely clips the object.
- If visible important text is clipped, unreadable, or missing from extraction,
	record a highlighted `[TEXT_LOSS]` note with the visible fragment, issue,
	bounding box when available, and severity. Never invent missing characters.

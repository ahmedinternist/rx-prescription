# Glass dashboard refinement plan

## Current white/light-gray palette — supersedes the charcoal update below

The user's attached `original-07b89ef4ec4344efd01fe788b22e53f5.webp`
is the current color reference: white shell, light neutral-gray panels,
saturated-blue accents and charcoal type. Approximate tokens are background
`#F6F7F9`, cards `#E9EAEC`, fields `#FFFFFF`, sidebar `#EEF0F3`,
text `#25272A`, secondary `#454C56`, muted `#59616C`, accent `#0960C7`,
primary `#0964DC`, selection `#E1EBFB`, edges `#C3C9D1`.
White-gradient reflections and position-aware shared-backdrop sampling remain;
shadows are softened for the light surfaces. Warning/delete/reference colors
return to readable light-background variants. Layouts, fonts, records and
prescription exports are unchanged. Darker blue than the reference is used
for small text/icons to maintain contrast. Prior palette notes are historical.

## Current palette update — 2026-09-17

The user's attached `original-b377d98007782961d562d78f81044127.webp`
supersedes the light palette below. Its charcoal glass, slate-blue shading,
electric-blue actions and light type are adapted without copying illustrations
or changing page architecture. Colors are visual approximations:
background `#141A22`, cards `#222B37`, fields `#181F29`, sidebar `#1C2531`,
text `#F2F5FA`, secondary `#C3CEDD`, muted `#AAB8CA`, icon/text accent
`#91ADFF`, action fill `#214EE8`, selection `#2B3C5C`, edge `#45556B`.
Action fill and text accent are separate to preserve contrast. Status and
reference colors retain their meaning with dark backgrounds/light foregrounds.
Editable fields stay opaque. The supplied edit silhouette is tinted blue using
its original alpha mask. Layout, fonts, data and prescription exports are unchanged.
The following light-color section is retained as historical design context.

White-gradient refinement: the shared charcoal backdrop has a broad white haze;
panels/sidebar use cached low-opacity white reflections with an upper glint and
diffuse lower gradient. Bright cool edge highlights reinforce simulated glass.
Fields remain opaque, and rendering keeps existing debounce/visibility/cache
limits. Normal-text contrast is tested against actual gradient pixels.

Palette verification: 76 automated tests passed, including normal-text contrast
for primary actions, fields, warning/delete symbols and all five reference-card
color pairs. The English/Arabic scaling matrix passed 136 hidden page checks.
These are not screenshot-based pixel checks or all-Windows-version certification.

## Source

- Gallery: https://dribbble.com/search/Dashboard-UI-Glass-Effect-Concept
- Selected light-blue example: https://dribbble.com/shots/25346029-Glassmorphism-Dashboard-UI-Design
- Designer: Leon Abramovic. Reviewed 2026-09-17.
- Evidence: the selected shot and description were viewed through the user's
  browser. Firecrawl screenshot/branding captures returned a temporary error;
  no valid full-page screenshot asset was retained. Do not treat the error
  output in `.firecrawl/` as a reference image.
- This is an original adaptation for the existing CustomTkinter prescription
  app, not a copy of the designer's logo, illustration, content or assets.

## Design summary

Observed: light-blue background, frosted rounded panels, subtle highlights,
soft visual depth and subdued navigation. The creator describes transparent
panels and subtle blur. Exact opacity, spacing and colors were not measured.

Proposed: refine the current simulated-glass dashboard into a cohesive light
glass shell. Preserve its single sidebar, compact prescribing workflow and
readable fields. Do not add the example's extra icon rail, analytics widgets,
profile imagery or large decorative spaces.

## Proposed design tokens (approximations, not extracted CSS)

- Background: pale blue `#DCEEF8`, soft cyan `#CFEAF6`, teal `#DDF3EF`.
- Text: charcoal `#1B2435`; secondary text `#49576B`.
- Icons: blue `#1464DC`; selected navigation teal `#127D79`.
- Input fields: opaque white with visible `#C5DCE9` border.
- Frosted white veil: approximately 70–80% sidebar, 85–92% content cards.
- Glass edges: 1 logical-pixel white highlight plus subtle blue edge.
- Shadows: restrained, static pre-rendered blue-gray shadow, 4–6 px visual spread.
- Corners: 12–16 px cards; 8–10 px fields and compact controls.
- Spacing: 4/6/8/10 px rhythm; no additional large header margins.
- Typography: preserve existing approved font sizes, bold brand names and
  regular scientific names. Use Segoe UI and the existing Arabic rendering.

## Components and page patterns

1. **Shared backdrop:** use one coherent background across the visible shell;
   avoid every card restarting the same local gradient. Prototype coordinate-
   aligned sampling and ensure scrolling/resizing does not produce seams.
2. **Frosted sidebar:** translucent-looking pale surface, slightly raised edge,
   stable selected teal row and consistent existing icons. Keep collapse mode.
3. **Search bar:** frosted white strip, visible focus outline, aligned dropdown.
   Preserve autocomplete, keyboard selection and current dismissal behavior.
4. **Cards:** white-veiled surfaces with small highlights and subtle shadows.
   Retain natural card height and current two-/three-column responsive layouts.
5. **Fields:** keep actual editable text areas opaque. Do not lower text opacity
   or tint dosage/notes enough to reduce readability.
6. **Actions:** compact white outlined text buttons; existing plain edit/delete/
   plus/star symbols. Hover changes color only, never size/position. No tooltips.
7. **Consistency:** apply shared visual tokens to Settings and other pages,
   preserving existing layouts, removed elements and scrollbar behavior.

## Implementation order

1. Prototype the coherent backdrop and sidebar on the existing dashboard.
2. Refine search, medication cards and action surfaces without changing commands.
3. Apply approved component styling to Favorites, Drug Classes, Treatment
   Template, Patient/Prescriber pages, reference cards and Settings.
4. Verify visual readability and screen-fit before rebuilding v4.82.

## Rendering constraints

CustomTkinter does not provide CSS backdrop-filter. Use simulated glass/static
compositing in the existing renderer; do not promise real desktop transparency.
Native Acrylic/Mica is not part of this plan. Prefer cached textures, weak
interpreter-local image references and debounced resize updates. Coordinate-
dependent surfaces must have position-aware cache keys; limit cache size and
render visible surfaces only to avoid memory growth and scroll lag.

## Verification

- English and Arabic; names, mixed-script fields and dropdowns.
- 100%, 125%, 150% and 200% scaling; minimum window and maximized view.
- Long medication lists, Favorites grids, class browser and template selection.
- No hover flicker, layout shift, dropdown clipping or background seams.
- Stable memory after repeated navigation, resizing and card deletion.
- Existing automated tests and hidden page checks; separate visible/pixel QA.
- Prescriptions, Word/PDF/QR exports, imports and saved records remain unchanged.
- Local Windows testing does not certify untested Windows versions.

## Rerun inputs

workflow: firecrawl-website-design-clone
source_url: https://dribbble.com/shots/25346029-Glassmorphism-Dashboard-UI-Design
target_stack: existing Python / CustomTkinter / Pillow desktop app
output: DESIGN.md

Status: implemented. The existing simulated-glass renderer now
samples position-aware backdrop regions, adds static inset depth/highlights,
preserves selected-card tints and releases hidden/off-screen image references.
The same surfaces are applied to the dashboard, cards and Settings;
prescription commands, typography and data formats are unchanged.

Verification: 75 automated tests and 136 hidden page/section checks passed
(English/Arabic, 100–200% widget scaling). Desktop screenshot verification was
unavailable because the computer-use helper reported a window ownership
mismatch. These checks do not certify pixel layout on every display or Windows
version. The effect remains simulated glass, not real desktop transparency.

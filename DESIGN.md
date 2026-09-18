# Glass dashboard refinement plan

## Cloud QR Settings and export states — 2026-09-18

Reference lock: preserve the existing compact Clinical Glass Settings cards,
opaque white inputs, blue action text and native keyboard focus. The approved
cloud QR plan owns content and failure choices; Refero craft-details owns visible
submit states and disabled controls while requests run. Replace static viewer/key
controls with the fixed endpoint, a masked API-key input and a clear upload notice.
Use the existing busy footer for progress and a compact three-choice failure dialog.
No new imagery, tooltips, palette or dashboard redesign is introduced.

## Uncapped top search width — 2026-09-18

Reference lock: the user's annotated search screenshot requests a wider result
surface extending to the red line, without a screen-edge cap. Preserve the approved
white glass style, typography and search-bar left alignment. Only top search opts
out of horizontal screen fitting, keeping the full doubled width request even if
it extends beyond a monitor. Vertical row fitting, scrolling and dismissal remain
unchanged; all other autocomplete menus keep their existing screen limits.

Verification: 85 automated tests passed. The actual App top-search renderer was
checked with 14 synthetic class results and a vertical scrollbar: it retained the
full doubled request, with its list filling the widened surface and its left edge
matching the search bar. The captured surface was reviewed; any portion beyond
the physical monitor is naturally not visible, but no application width cap applies.

## Double-width top search results — 2026-09-17

Reference lock: preserve the approved popup surface, typography, left alignment,
keyboard controls and overflow behavior. The user's explicit 2× width request
changes only the main search results; shared medicine/favorite autocomplete keeps
its existing width. Double the larger of the field width and content-based request,
then fit it to the screen's right edge. Existing horizontal scrolling remains
available when a result exceeds the physical screen width.

Verification: 84 automated tests passed, including an exact 160→320 px width
check and screen clipping/overflow checks. Isolated English 100% and Arabic 200%
layout probes passed; the wider popup capture was reviewed.

## Dropdown scope and compact controls — 2026-09-17

Reference lock: preserve the approved white Clinical Glass surfaces, blue outline
icons and existing layout. The user's requested refinements own the menu scope and
density; Refero's bundled color guidance supplies the foreground/background pairing
rule. Font overrides (10–56 px) apply outside Settings only. Settings descendants
retain their original menu fonts both on creation and when preferences are reapplied.
Top search popups measure complete result strings, expand within available screen
width and expose horizontal scrolling only when the text cannot fit. Their left
edge stays aligned with the search bar. Detailed-class Add Drug uses the existing
blue plus on white with a pale-blue hover; the command and icon size are unchanged.
The Settings footer decreases by 16 px, keeping one line of bottom padding. No new
tooltips, record migrations or export-format changes are introduced.

Verification: 83 automated tests passed, including maximum-size persistence,
Settings-menu isolation, content-measured popup widths, horizontal overflow access,
screen bounds with scrollbar height, and the Add Drug command. Four isolated
English/Arabic probes at 100%/200% passed; Settings, detailed-class controls and
56 px search results were visually reviewed using synthetic data.

## Search and font controls — 2026-09-17

Reference lock: the supplied screenshot identifies the broken three-sided focus
outline. Preserve the current white Clinical Glass theme; add vertical inset so
the transparent entry canvas cannot cover the search frame border. Replace the
command-key decoration with a matching 18 px blue outline magnifier, following
Refero's bundled icon grid guidance. General gains one compact two-column font
card. Default retains existing dropdown typography; explicit pixel-size choices
apply to native dropdowns, autocomplete results and context menus. Patient-name
size affects the on-screen input, not prescription typography. Save updates
existing controls without clearing entered data; restore/reset include the new
preferences. No tooltips are added.

Verification: 80 automated tests passed. Four isolated English/Arabic layout
checks at 100% and 200% confirmed the focused search outline, font preferences,
and General-card visibility. Actual search and Settings captures were reviewed.
Six synthetic Word files (A5/A4 with header, without header, and medications only)
were reopened and their OOXML section sizes verified. Word rendering was not
available because the bundled document runtime has no LibreOffice; physical
printer-driver behavior is not certified. Live records were not used.

## Eight approved visual refinements — 2026-09-17

Reference lock: preserve the user's Care white/light-gray palette and the
approved Medication Entry / Settings layout. Refero's bundled icon, typography
and form craft guidance informs execution; no new live Refero research is claimed.

Decision ledger: native single hairline borders replace inset double outlines;
white reflections are quieter; 12 px field labels retain 36 px input heights;
cached 18 px action artwork balances plus, star, trash and arrows while retaining
the supplied edit silhouette. Number/action anchors stay at the name-input baseline
even when a linked-brand picker opens. Favorites and Template toolbars share
36 px search/text-action heights and 9 px corners. The footer uses a borderless
white sheet. Neutral entry/dropdown borders turn blue on focus without resizing;
validation colors and existing input bindings are preserved. No tooltips are added.

Scope: visual only, v4.82.0 retained; no record or export-format changes.

Verification: 78 automated checks passed, including action command retention,
star state changes, color-only focus geometry and preservation of validation
colors. Eight isolated English/Arabic layout probes at 100%, 125%, 150% and
200% passed; updated Medication Entry and Clinic Identity screenshots were
reviewed. Testing does not access the user's live records.

## Approved Medication Entry and Settings implementation — 2026-09-17

Build target: the two Clinical Glass mockups approved by the user. The app
remains v4.82.0; this change does not alter records or prescription output.
The user's Care reference owns the white/light-gray canvas, blue accents,
charcoal type and compact rounded panels. Earlier Linear settings research
contributes grouping only; Healthie workflow research contributes separation
of clinical tasks only. Live Refero tools were unavailable; bundled Refero
craft guidance and the approved visual targets were used instead.

Decision ledger:

| Decision | Source / role | Implementation |
| --- | --- | --- |
| White outlined medication actions | Approved Medication Entry preview | All four toolbar actions keep their commands and use white surfaces |
| Direct reorder arrows | Approved numbered medicine rows | Up/down buttons; drag and context-menu sorting retained on the plain number |
| Opaque white details | Care reference and Refero form guidance | Frequency/notes match the entry fields, with pale-blue dropdown triggers |
| White Settings canvas, gray grouped cards | Approved Settings preview | New cached sheet surface; existing shared glass renderer retained |
| Compact Clinic Identity | Approved contact/logo/preview layout | Three balanced contact fields, inline logo thumbnail, visible live preview |
| Bottom action breathing room | User requirement and approved preview | Existing 32 px bottom inset retained; restore/reset/save remain distinct |
| Visible empty-field search hints | Approved search controls | Non-value overlay hints disappear on focus or typing; StringVar search values stay clean |

Glass remains simulated using static gradients and highlights, not real
desktop transparency or live blur. Hover changes colors only. Existing page
navigation, plain edit/delete symbols, typography roles, Arabic text bindings,
save/discard warnings and data formats are preserved.

Implementation evidence: `output/medication-approved-actual.png` and
`output/settings-approved-actual.png`, captured from an isolated test instance.
Layout probe: `output/approved_glass_probe.py`, using fresh interpreters for
English/Arabic and paired widget/window scaling at 100–200%.

Verification: 77 automated tests passed. All eight English/Arabic scaling
combinations passed layout checks, including clinic preview visibility and
footer button separation. Screenshot review found and corrected unnecessary
logo/preview spacing and missing search hints. Tests use isolated application
data, not the user's records. This is local layout verification, not a claim
of compatibility certification for every Windows version.

Earlier palette sections below are historical, not the current build target.

## Current white/light-gray palette — supersedes the charcoal update below

The user's attached `original-07b89ef4ec4344efd01fe788b22e53f5.webp`
is the current color reference: white shell, light neutral-gray panels,
saturated-blue accents and charcoal type. Approximate tokens are background
`#F6F7F9`, cards `#E9EAEC`, fields `#FFFFFF`, sidebar `#EEF0F3`,
text `#25272A`, secondary `#454C56`, muted `#59616C`, accent `#0960C7`,
primary `#0964DC`, selection `#E1EBFB`, edges `#CDD2D9`.
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

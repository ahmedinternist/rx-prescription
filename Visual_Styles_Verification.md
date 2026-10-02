# Compact visual consistency — 7.5

## Version 7.8 follow-up

Favorite Drugs cards display only trade/scientific names and the category at
the end of the action row. The regimen-summary row is omitted. Editing fields,
stored dosage/frequency/duration/notes and prescription insertion remain unchanged.
An isolated GUI regression verifies visible names/category, absence of regimen
text, category placement and unchanged favorite data.

Verification: **205 tests passed in 69.31 seconds**. The separate 7.8 executable
passed an isolated launch check with no nonempty application error log. Earlier
versioned executables and existing application data are retained.

## Version 7.7 follow-up

Favorite and treatment-editor headings wrap long names inside constrained columns. Detailed-class medicine rows reserve a separate action area before allocating width to brand/scientific names. Resize callbacks update wrapping only when the available text width changes; logical names and export content remain unchanged.

Compact English/Arabic Saving/Saved/Failed labels accompany template, favorite, mapping and medication-prescription saves. Favorite/template success remains visible after returning from the editor. Patient-save feedback uses its existing status area. Mapping and prescription worker results update these labels on the UI thread. Cancelled similar-patient selection clears the pending indicator; a completed older prescription snapshot does not mark newer edits Saved. Existing error propagation and dialogs remain intact.

Regression coverage includes constrained long-name geometry, action visibility, successful writes, rejected writes and exceptions. No API removal, storage format or document-layout change.

Final 7.7 regression suite: **204 passed in 109.28 seconds**. The rebuilt executable includes the reserved detailed-class action area and passed an isolated launch check without an application error log. Previous executable versions remain available.

## Version 7.6 follow-up

Treatment card summaries contain dosage, frequency, duration and note values without field-name prefixes. The editing controls retain their labels. Medication Entry expands only the active row and hides other rows' editing fields, including incomplete rows, without discarding their values. Deferred focus callbacks ignore rows that have since collapsed.

The mapping list displays scientific and trade names side by side, using larger 22-pixel trade text rather than the small metadata font. Classification/status stays below. Long names are visually ellipsized to prevent overlap; underlying database names remain unchanged. The selected-medicine heading also shows scientific name followed by trade name, without the former newline. Existing MappingMedicineList calls remain valid; `trade_name` is an optional keyword.

Verification: 203 regression tests passed in 96.78 seconds, followed by all 11 focused performance/UI tests on the final source. Archive inspection confirmed bounded inline-name layout in the 7.6 executable; the isolated packaged launch remained running without an application error log.

- Standard bordered glass cards use a 12-pixel corner radius; detailed drug rows use the same radius and 4-pixel card gap.
- Favorite, expanded saved-template and detailed-class medicine names use a shared 15-point hierarchy: brand names bold, scientific names regular. Regimen details use regular 12-point secondary text. Existing user-configurable field fonts and exported document styling are unchanged.
- Add, Edit, Delete and Star actions share 30-by-30 hit areas, 18-pixel artwork and transparent borderless backgrounds. Command semantics, confirmation dialogs and star state remain unchanged. No tooltips are added.
- Entry and editable-combobox focus changes only border color. Blur restores its original color unless validation has changed it. Danger/warning borders are preserved, and repeated focus events do not overwrite the restoration state.
- Behavioral tests cover shared styles, icon geometry, symbol recoloring, repeated focus/blur, unchanged field dimensions, and error-border preservation. Existing regressions protect medication values, persistence, exports and navigation.

Release verification uses isolated application data. This update does not change storage schemas, cloud payloads, clinical behavior or export layouts. Previous executable versions are retained.

Final regression suite: **202 passed in 161.79 seconds**. The packaged 7.5 app stayed running during an isolated launch check without an application error log. Archive inspection confirmed the final shared edit-icon sizing is included.

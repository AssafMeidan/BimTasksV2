# Filter to Legend — Design

## Summary
Command that reads view filters + color fill overrides from the active view, lets user pick/reorder which to include via a checklist dialog, then generates a Legend View with color swatches and labels.

## Flow
1. Click "Filter Legend" button (Tools toolbar group)
2. Handler reads ParameterFilterElements applied to active view + OverrideGraphicSettings
3. Checklist dialog: filter name + color preview, checkboxes, up/down arrow buttons for reordering
4. Creates new Legend View named `Legend - {active view name}`
5. Per selected filter (in user-defined order):
   - FilledRegion color swatch (~10mm x 10mm default)
   - TextNote: filter name + filter value (just the value, e.g., "Zone A")
6. Vertical layout, one row per filter

## Dialog Features
- Checklist with color preview swatches
- Check/uncheck individual filters
- Up/down arrow buttons to manually reorder
- Collapsible settings section: swatch size, text height, spacing

## Legend Elements
- FilledRegion for swatches (new FilledRegionType per unique color/pattern)
- TextNote for name + value

## Edge Cases
- No color override → hollow rectangle + "no override"
- Pattern overrides → use actual fill pattern
- Duplicate legend name → append number
- No filters → TaskDialog message

## Files
- NEW: Commands/Handlers/FilterToLegendHandler.cs
- NEW: Views/FilterToLegendDialog.xaml + .xaml.cs
- NEW: Helpers/LegendBuilder.cs
- MODIFY: ViewModels/FloatingToolbarViewModel.cs (add button)
- MODIFY: Views/FloatingToolbarWindow.xaml (add button)

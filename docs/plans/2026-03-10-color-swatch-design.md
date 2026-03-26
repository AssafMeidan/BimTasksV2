# Color Swatch by Shared Parameter — Design Document

**Date:** 2026-03-10
**Status:** Approved — ready for implementation

## Overview

Two-part feature:
1. **Panel Tab Infrastructure** — persistent tab strip for the dockable panel so views survive switching
2. **Color Swatch Command** — temporarily color elements by shared parameter values, with interactive legend in the dockable panel

---

## Part 1: Panel Tab Infrastructure

### Problem
The dockable panel re-creates views on every switch. Users lose state (filter tree selections, calculation results) and have no way to navigate back to a previously opened view.

### Solution
Add a tab strip to the dockable panel. Cache views in a dictionary so they persist across switches.

### Panel Layout
```
┌──────────────────────────────────────────┐
│  ▄▄▄▄▄▄▄▄▄▄  BimTasks Panel  ▄▄▄▄▄▄▄▄  │  header
├──────────────────────────────────────────┤
│ [Filter Tree] [Calculations] [Colors ×] │  tab strip
├──────────────────────────────────────────┤
│         active tab content               │
└──────────────────────────────────────────┘
```

### Behavior
- Tabs appear only when a view is opened (empty panel shows no tabs)
- Tab strip hidden when only one view exists (no clutter)
- Clicking a tab shows the cached view — no re-creation
- Some tabs have (×) close button for cleanup (e.g., Color Swatch clears overrides)
- `PanelTitle` updates to match active tab
- `SwitchDockablePanelEvent` still works — cache-first, create on miss

### Tab Style
- Flat compact buttons, `#1565C0` theme
- Active tab highlighted, inactive tabs muted

### Files to Modify
- `ViewModels/BimTasksDockablePanelViewModel.cs` — `Dictionary<string, FrameworkElement>` cache, `ObservableCollection<TabInfo>` for tab strip, switch logic
- `Views/BimTasksDockablePanel.xaml` — add `ItemsControl` tab row between header and content

---

## Part 2: Color Swatch Command

### Purpose
Temporarily color elements in the active Revit view by shared parameter values. Show an interactive legend in the BimTasks dockable panel.

### Panel Layout
```
┌──────────────────────────────────────────┐
│  Parameter: [Assembly Code        ▼]     │
│  Show with: [Description          ▼]     │  optional, default "None"
│  Categories: [Walls, Floors, ...  ▼]     │  multi-select, auto-picks top
├──────────────────────────────────────────┤
│  [Apply Colors]          [Clear Colors]  │
├──────────────────────────────────────────┤
│  ■ A1.10  Exterior Wall         (42)    │
│  ■ A1.20  Interior Wall         (18)    │
│  ■ B2.30  Concrete Floor        (12)    │
│  ■ (empty)                       (3)    │
└──────────────────────────────────────────┘
```

### Flow
1. User opens command → panel switches to Color Swatch view
2. Parameter dropdown: all shared parameters on visible model elements (no annotation categories)
3. Category multi-select: auto-checks most populated model categories
4. User picks parameter → legend rows show unique values with auto-assigned colors (16-color palette)
5. **Apply Colors** → `View.SetElementOverrides()` per element
6. Click swatch → inline preset palette (8-10 colors) + "Custom..." button → Windows `ColorDialog`
7. Changing parameter/categories → re-scan, rebuild legend, clear previous overrides
8. **Clear Colors** → restore original `OverrideGraphicSettings` from stored backup
9. Closing tab (×) also clears overrides

### Color Assignment
- 16-color distinguishable categorical palette (reds, blues, greens, oranges, purples, teals, pinks, browns)
- Colors cycle if more unique values than palette size
- "(empty)" group gets gray swatch

### Override Storage
- Before applying: store each element's original `OverrideGraphicSettings` in `Dictionary<ElementId, OverrideGraphicSettings>`
- On clear/close: restore originals in a single transaction
- Only override surface foreground pattern + color

### Secondary Parameter ("Show with")
- Same shared parameter list minus the selected primary parameter
- Default: "None"
- If elements with same primary value have different secondary values → use first found
- Displayed as second column in legend row

### Edge Cases
- No value → grouped under "(empty)" with gray swatch
- More values than colors → palette cycles
- View changes in Revit → legend becomes stale, user clicks Apply again (no auto-refresh)

### New Files
- `Commands/Handlers/ColorSwatchHandler.cs` — entry handler
- `Commands/Proxies/ToolCommands.cs` — add `ColorSwatchCommand` proxy
- `Views/ColorSwatchView.xaml` + `.xaml.cs` — panel view
- `ViewModels/ColorSwatchViewModel.cs` — logic, override management, color assignment
- `Ribbon/RibbonConstants.cs` — add button to ToolButtons
- `Services/ColorOverrideService.cs` — apply/clear/restore override logic (keeps ViewModel clean)

### Revit API Notes
- `View.SetElementOverrides(ElementId, OverrideGraphicSettings)` — per-element override
- `OverrideGraphicSettings.SetSurfaceForegroundPatternId()` + `SetSurfaceForegroundPatternColor()`
- Must find solid fill pattern via `FillPatternElement.GetFillPattern().IsSolidFill`
- `color.IsValid` guard before accessing RGB channels
- All override operations require a `Transaction`

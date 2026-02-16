using System;
using Autodesk.Revit.UI;

namespace BimTasksV2.Ribbon
{
    /// <summary>
    /// Builds the BimTasks ribbon tab with 5 domain-based panels.
    /// Called from BimTasksApp.OnStartup (default ALC — no Prism/Unity refs).
    /// Icons are generated at runtime by <see cref="IconGenerator"/>.
    /// </summary>
    public static class RibbonBuilder
    {
        public static void Build(UIControlledApplication app, string assemblyPath)
        {
            // Create the top-level tab
            app.CreateRibbonTab(RibbonConstants.TabName);

            // Create 5 domain-based panels (order matters — determines ribbon layout)
            var panelWalls     = app.CreateRibbonPanel(RibbonConstants.TabName, RibbonConstants.WallPanel);
            var panelStructure = app.CreateRibbonPanel(RibbonConstants.TabName, RibbonConstants.StructurePanel);
            var panelSchedules = app.CreateRibbonPanel(RibbonConstants.TabName, RibbonConstants.SchedulePanel);
            var panelBOQ       = app.CreateRibbonPanel(RibbonConstants.TabName, RibbonConstants.BOQPanel);
            var panelTools     = app.CreateRibbonPanel(RibbonConstants.TabName, RibbonConstants.ToolPanel);

            // --- Panel 1: Walls ---
            // Buttons 0-3 are standalone, 4-6 are a SplitButton (cladding), 7 is standalone
            for (int i = 0; i < RibbonConstants.CladdingSplitStart; i++)
            {
                AddPushButton(panelWalls, RibbonConstants.WallButtons[i], assemblyPath);
            }

            // Cladding SplitButton (indices 4, 5, 6)
            AddSplitButton(panelWalls, "WallCladdingSplit", "Wall Cladding",
                assemblyPath,
                RibbonConstants.WallButtons[RibbonConstants.CladdingSplitStart],
                RibbonConstants.WallButtons[RibbonConstants.CladdingSplitStart + 1],
                RibbonConstants.WallButtons[RibbonConstants.CladdingSplitEnd]);

            // Remaining wall buttons after split (index 7: CreateWindowFamilies)
            for (int i = RibbonConstants.CladdingSplitEnd + 1; i < RibbonConstants.WallButtons.Length; i++)
            {
                AddPushButton(panelWalls, RibbonConstants.WallButtons[i], assemblyPath);
            }

            // --- Panel 2: Structure ---
            foreach (var btn in RibbonConstants.StructureButtons)
            {
                AddPushButton(panelStructure, btn, assemblyPath);
            }

            // --- Panel 3: Schedules ---
            foreach (var btn in RibbonConstants.ScheduleButtons)
            {
                AddPushButton(panelSchedules, btn, assemblyPath);
            }

            // --- Panel 4: BOQ ---
            foreach (var btn in RibbonConstants.BOQButtons)
            {
                AddPushButton(panelBOQ, btn, assemblyPath);
            }

            // --- Panel 5: Tools ---
            // SplitButton first: CreateSelectedItem + AddClunasFromDwg
            AddSplitButton(panelTools, "ToolActionsSplit", "Create Similar",
                assemblyPath,
                RibbonConstants.ToolSplitButtons);

            // Regular tool buttons
            foreach (var btn in RibbonConstants.ToolButtons)
            {
                AddPushButton(panelTools, btn, assemblyPath);
            }
        }

        // =====================================================================
        // Helpers
        // =====================================================================

        private static PushButton? AddPushButton(RibbonPanel panel, ButtonDef def, string assemblyPath)
        {
            var data = new PushButtonData(def.Name, def.Text, assemblyPath, def.ClassName);

            try
            {
                var btn = panel.AddItem(data) as PushButton;
                if (btn != null)
                    ApplyIcon(btn, def.Name);
                return btn;
            }
            catch (Exception)
            {
                // Log or silently continue — button might already exist
                return null;
            }
        }

        private static void AddSplitButton(
            RibbonPanel panel,
            string splitName,
            string splitText,
            string assemblyPath,
            params ButtonDef[] members)
        {
            if (members == null || members.Length == 0)
                return;

            var splitData = new SplitButtonData(splitName, splitText);

            try
            {
                var splitButton = panel.AddItem(splitData) as SplitButton;
                if (splitButton == null)
                    return;

                foreach (var member in members)
                {
                    var btnData = new PushButtonData(member.Name, member.Text, assemblyPath, member.ClassName);

                    try
                    {
                        var pb = splitButton.AddPushButton(btnData);
                        if (pb != null)
                            ApplyIcon(pb, member.Name);
                    }
                    catch (Exception)
                    {
                        // Skip this member if it fails
                    }
                }
            }
            catch (Exception)
            {
                // Fall back: add first member as a standalone button
                if (members.Length > 0)
                {
                    AddPushButton(panel, members[0], assemblyPath);
                }
            }
        }

        private static void ApplyIcon(RibbonButton btn, string name)
        {
            var large = IconGenerator.GetIcon(name, 32);
            var small = IconGenerator.GetIcon(name, 16);
            if (large != null) btn.LargeImage = large;
            if (small != null) btn.Image = small;
        }
    }
}

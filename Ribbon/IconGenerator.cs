using System;
using System.Collections.Generic;
using System.Globalization;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace BimTasksV2.Ribbon
{
    /// <summary>
    /// Generates 32×32 and 16×16 ribbon button icons at runtime using WPF drawing.
    /// Each icon is a panel-themed colored rounded rectangle with a white symbol.
    /// </summary>
    public static class IconGenerator
    {
        // Frozen brushes and pens (thread-safe, allocated once)
        private static readonly SolidColorBrush White = Freeze(new SolidColorBrush(Colors.White));
        private static readonly Pen Thick = Freeze(new Pen(White, 2.2)
        {
            StartLineCap = PenLineCap.Round,
            EndLineCap = PenLineCap.Round,
            LineJoin = PenLineJoin.Round
        });
        private static readonly Pen Thin = Freeze(new Pen(White, 1.5)
        {
            StartLineCap = PenLineCap.Round,
            EndLineCap = PenLineCap.Round,
            LineJoin = PenLineJoin.Round
        });

        // Panel background colors
        private static readonly SolidColorBrush WallBg   = Freeze(new SolidColorBrush(Color.FromRgb(59, 130, 246)));
        private static readonly SolidColorBrush StructBg = Freeze(new SolidColorBrush(Color.FromRgb(245, 158, 11)));
        private static readonly SolidColorBrush SchedBg  = Freeze(new SolidColorBrush(Color.FromRgb(16, 185, 129)));
        private static readonly SolidColorBrush BoqBg    = Freeze(new SolidColorBrush(Color.FromRgb(139, 92, 246)));
        private static readonly SolidColorBrush ToolBg   = Freeze(new SolidColorBrush(Color.FromRgb(6, 182, 212)));

        private static T Freeze<T>(T f) where T : Freezable { f.Freeze(); return f; }

        // Button name → (background brush, foreground draw action)
        private static readonly Dictionary<string, (SolidColorBrush Bg, Action<DrawingContext> Draw)> Icons = new()
        {
            // Walls (blue)
            ["btnSetWindowDoorPhase"]   = (WallBg, DrawSync),
            ["btnCopyLinkedWalls"]      = (WallBg, DrawCopyRects),
            ["btnAdjustWallHeight"]     = (WallBg, DrawWallFloor),
            ["btnChangeWallHeight"]     = (WallBg, DrawWallHeight),
            ["btnAddChipuyToWall"]      = (WallBg, DrawCladding),
            ["btnAddChipuyExternal"]    = (WallBg, dc => { DrawCladding(dc); Label(dc, "E", 22); }),
            ["btnAddChipuyInternal"]    = (WallBg, dc => { DrawCladding(dc); Label(dc, "I", 23); }),
            ["btnCreateWindowFamilies"] = (WallBg, DrawWindow),
            // Structure (amber)
            ["btnSetBeamsToGround"]        = (StructBg, DrawBeamDown),
            ["btnJoinConcreteWallsFloors"] = (StructBg, DrawJoinGeom),
            // Schedules (green)
            ["btnExportScheduleToExcel"]   = (SchedBg, DrawExportGrid),
            ["btnEditScheduleInExcel"]     = (SchedBg, DrawEditGrid),
            ["btnImportKeySchedules"]      = (SchedBg, DrawImportGrid),
            // BOQ (violet)
            ["btnBOQSetup"]            = (BoqBg, DrawGear),
            ["btnBOQCreateSchedules"]  = (BoqBg, DrawTable),
            ["btnBOQAutoFillQty"]      = (BoqBg, DrawBolt),
            ["btnBOQCalcEffUnitPrice"] = (BoqBg, dc => CenterText(dc, "$", 18)),
            ["btnCreateSeifeiChoze"]   = (BoqBg, DrawDoc),
            // Tools (cyan)
            ["btnShowFilterTree"]       = (ToolBg, DrawFunnel),
            ["btnCheckWallsAreaVolume"] = (ToolBg, DrawArea),
            ["btnShowUniformat"]        = (ToolBg, DrawHash),
            ["btnCopyCategoryFromLink"] = (ToolBg, DrawCopyArrow),
            ["btnToggleDockablePanel"]  = (ToolBg, DrawDock),
            ["btnToggleToolbar"]        = (ToolBg, DrawFloatBar),
            ["btnExportJSON"]           = (ToolBg, dc => CenterText(dc, "{ }", 14)),
            ["btnToggleVosk"]           = (ToolBg, DrawMic),
            ["btnCreateSelectedItem"]   = (ToolBg, DrawPlus),
            ["btnAddClunasFromDwg"]     = (ToolBg, DrawPiles),
        };

        /// <summary>
        /// Returns a frozen BitmapSource icon for the given button name and pixel size,
        /// or null if the button name is not registered.
        /// </summary>
        public static BitmapSource? GetIcon(string buttonName, int size)
        {
            if (!Icons.TryGetValue(buttonName, out var entry))
                return null;

            var dv = new DrawingVisual();
            double s = size / 32.0;

            using (var dc = dv.RenderOpen())
            {
                if (Math.Abs(s - 1.0) > 0.001)
                    dc.PushTransform(new ScaleTransform(s, s));

                // Colored rounded-rect background
                dc.DrawRoundedRectangle(entry.Bg, null, new Rect(0, 0, 32, 32), 5, 5);

                // White foreground symbol
                entry.Draw(dc);

                if (Math.Abs(s - 1.0) > 0.001)
                    dc.Pop();
            }

            var bmp = new RenderTargetBitmap(size, size, 96, 96, PixelFormats.Pbgra32);
            bmp.Render(dv);
            bmp.Freeze();
            return bmp;
        }

        // =================================================================
        // Wall icons
        // =================================================================

        private static void DrawSync(DrawingContext dc)
        {
            // Two curved arrows forming a refresh/sync symbol
            var g = new StreamGeometry();
            using (var c = g.Open())
            {
                // Upper arc: right-to-left
                c.BeginFigure(new Point(20, 9), false, false);
                c.ArcTo(new Point(9, 15), new Size(9, 9), 0, false,
                    SweepDirection.Counterclockwise, true, true);
                // Arrowhead
                c.BeginFigure(new Point(9, 15), false, false);
                c.LineTo(new Point(9, 10), true, true);
                c.BeginFigure(new Point(9, 15), false, false);
                c.LineTo(new Point(14, 15), true, true);

                // Lower arc: left-to-right
                c.BeginFigure(new Point(12, 23), false, false);
                c.ArcTo(new Point(23, 17), new Size(9, 9), 0, false,
                    SweepDirection.Counterclockwise, true, true);
                // Arrowhead
                c.BeginFigure(new Point(23, 17), false, false);
                c.LineTo(new Point(23, 22), true, true);
                c.BeginFigure(new Point(23, 17), false, false);
                c.LineTo(new Point(18, 17), true, true);
            }
            g.Freeze();
            dc.DrawGeometry(null, Thick, g);
        }

        private static void DrawCopyRects(DrawingContext dc)
        {
            // Two offset rectangles (copy symbol)
            dc.DrawRectangle(null, Thick, new Rect(7, 10, 11, 15));
            dc.DrawRectangle(null, Thick, new Rect(14, 7, 11, 15));
        }

        private static void DrawWallFloor(DrawingContext dc)
        {
            // Vertical wall meeting horizontal floor
            dc.DrawRectangle(White, null, new Rect(12, 7, 5, 18));
            dc.DrawLine(Thick, new Point(7, 25), new Point(25, 25));
            // Down arrow indicating adjustment
            dc.DrawGeometry(null, Thin, Arrow(20, 12, 20, 22, 3));
        }

        private static void DrawWallHeight(DrawingContext dc)
        {
            // Wall with double-headed vertical arrow
            dc.DrawRectangle(White, null, new Rect(9, 7, 5, 18));
            dc.DrawLine(Thick, new Point(21, 7), new Point(21, 25));
            dc.DrawLine(Thick, new Point(18, 10), new Point(21, 7));
            dc.DrawLine(Thick, new Point(24, 10), new Point(21, 7));
            dc.DrawLine(Thick, new Point(18, 22), new Point(21, 25));
            dc.DrawLine(Thick, new Point(24, 22), new Point(21, 25));
        }

        private static void DrawCladding(DrawingContext dc)
        {
            // Wall with cladding layer
            dc.DrawRectangle(White, null, new Rect(10, 7, 5, 18));
            dc.DrawRectangle(null, Thick, new Rect(16, 7, 3, 18));
        }

        private static void DrawWindow(DrawingContext dc)
        {
            // Window: rectangle with cross dividers
            dc.DrawRectangle(null, Thick, new Rect(8, 8, 16, 16));
            dc.DrawLine(Thick, new Point(16, 8), new Point(16, 24));
            dc.DrawLine(Thick, new Point(8, 16), new Point(24, 16));
        }

        // =================================================================
        // Structure icons
        // =================================================================

        private static void DrawBeamDown(DrawingContext dc)
        {
            // Horizontal beam with downward arrow
            dc.DrawRectangle(White, null, new Rect(7, 9, 18, 4));
            dc.DrawGeometry(null, Thick, Arrow(16, 15, 16, 25, 3));
        }

        private static void DrawJoinGeom(DrawingContext dc)
        {
            // Two elements meeting at L-junction with join indicator
            dc.DrawRectangle(White, null, new Rect(8, 7, 4, 18));
            dc.DrawRectangle(White, null, new Rect(8, 21, 16, 4));
            dc.DrawEllipse(null, Thick, new Point(12, 21), 3, 3);
        }

        // =================================================================
        // Schedule icons
        // =================================================================

        private static void DrawSmallGrid(DrawingContext dc, double left)
        {
            double w = 14, h = 16;
            dc.DrawRectangle(null, Thin, new Rect(left, 8, w, h));
            dc.DrawLine(Thin, new Point(left, 13.3), new Point(left + w, 13.3));
            dc.DrawLine(Thin, new Point(left, 18.6), new Point(left + w, 18.6));
            dc.DrawLine(Thin, new Point(left + w / 2, 8), new Point(left + w / 2, 24));
        }

        private static void DrawExportGrid(DrawingContext dc)
        {
            DrawSmallGrid(dc, 6);
            // Right arrow (export)
            dc.DrawGeometry(null, Thick, Arrow(22, 16, 27, 16, 3));
        }

        private static void DrawEditGrid(DrawingContext dc)
        {
            DrawSmallGrid(dc, 6);
            // Pencil
            dc.DrawLine(Thick, new Point(23, 9), new Point(26, 25));
        }

        private static void DrawImportGrid(DrawingContext dc)
        {
            DrawSmallGrid(dc, 12);
            // Left arrow (import)
            dc.DrawGeometry(null, Thick, Arrow(10, 16, 5, 16, 3));
        }

        // =================================================================
        // BOQ icons
        // =================================================================

        private static void DrawGear(DrawingContext dc)
        {
            // Gear: circle with 8 radiating spokes
            dc.DrawEllipse(null, Thick, new Point(16, 16), 5, 5);
            dc.DrawLine(Thick, new Point(16, 6), new Point(16, 10));
            dc.DrawLine(Thick, new Point(16, 22), new Point(16, 26));
            dc.DrawLine(Thick, new Point(6, 16), new Point(10, 16));
            dc.DrawLine(Thick, new Point(22, 16), new Point(26, 16));
            dc.DrawLine(Thin, new Point(9, 9), new Point(12, 12));
            dc.DrawLine(Thin, new Point(20, 20), new Point(23, 23));
            dc.DrawLine(Thin, new Point(23, 9), new Point(20, 12));
            dc.DrawLine(Thin, new Point(9, 23), new Point(12, 20));
        }

        private static void DrawTable(DrawingContext dc)
        {
            // Table grid with rows and a column divider
            dc.DrawRectangle(null, Thick, new Rect(6, 7, 20, 18));
            dc.DrawLine(Thin, new Point(6, 12), new Point(26, 12));
            dc.DrawLine(Thin, new Point(6, 17), new Point(26, 17));
            dc.DrawLine(Thin, new Point(6, 22), new Point(26, 22));
            dc.DrawLine(Thin, new Point(13, 7), new Point(13, 25));
        }

        private static void DrawBolt(DrawingContext dc)
        {
            // Lightning bolt (autofill)
            var g = new StreamGeometry();
            using (var c = g.Open())
            {
                c.BeginFigure(new Point(18, 6), true, true);
                c.LineTo(new Point(11, 17), true, true);
                c.LineTo(new Point(16, 17), true, true);
                c.LineTo(new Point(13, 26), true, true);
                c.LineTo(new Point(22, 15), true, true);
                c.LineTo(new Point(17, 15), true, true);
            }
            g.Freeze();
            dc.DrawGeometry(White, null, g);
        }

        private static void DrawDoc(DrawingContext dc)
        {
            // Document with text lines (contract sections)
            dc.DrawRectangle(null, Thick, new Rect(8, 6, 16, 20));
            dc.DrawLine(Thin, new Point(11, 12), new Point(21, 12));
            dc.DrawLine(Thin, new Point(11, 16), new Point(21, 16));
            dc.DrawLine(Thin, new Point(11, 20), new Point(18, 20));
        }

        // =================================================================
        // Tool icons
        // =================================================================

        private static void DrawFunnel(DrawingContext dc)
        {
            // Funnel / filter shape
            var g = new StreamGeometry();
            using (var c = g.Open())
            {
                c.BeginFigure(new Point(6, 7), true, true);
                c.LineTo(new Point(26, 7), true, true);
                c.LineTo(new Point(18, 18), true, true);
                c.LineTo(new Point(18, 26), true, true);
                c.LineTo(new Point(14, 26), true, true);
                c.LineTo(new Point(14, 18), true, true);
            }
            g.Freeze();
            dc.DrawGeometry(White, null, g);
        }

        private static void DrawArea(DrawingContext dc)
        {
            // Square with "A" label (area / volume)
            dc.DrawRectangle(null, Thick, new Rect(8, 8, 14, 14));
            CenterText(dc, "A", 14);
        }

        private static void DrawHash(DrawingContext dc)
        {
            // # hash sign (codes)
            dc.DrawLine(Thick, new Point(13, 7), new Point(11, 25));
            dc.DrawLine(Thick, new Point(21, 7), new Point(19, 25));
            dc.DrawLine(Thick, new Point(7, 13), new Point(25, 13));
            dc.DrawLine(Thick, new Point(7, 19), new Point(25, 19));
        }

        private static void DrawCopyArrow(DrawingContext dc)
        {
            // Two rectangles with arrow between (copy from link)
            dc.DrawRectangle(null, Thick, new Rect(6, 11, 10, 13));
            dc.DrawRectangle(null, Thick, new Rect(16, 8, 10, 13));
            dc.DrawGeometry(null, Thin, Arrow(14, 17, 18, 17, 2));
        }

        private static void DrawDock(DrawingContext dc)
        {
            // Docked panel: window with left pane filled
            dc.DrawRectangle(null, Thick, new Rect(6, 7, 20, 18));
            dc.DrawRectangle(White, null, new Rect(6, 7, 8, 18));
            dc.DrawLine(Thick, new Point(14, 7), new Point(14, 25));
        }

        private static void DrawFloatBar(DrawingContext dc)
        {
            // Floating toolbar with title bar and small buttons
            dc.DrawRectangle(null, Thick, new Rect(6, 9, 20, 14));
            dc.DrawLine(Thick, new Point(6, 14), new Point(26, 14));
            dc.DrawRectangle(White, null, new Rect(9, 17, 4, 3));
            dc.DrawRectangle(White, null, new Rect(14, 17, 4, 3));
            dc.DrawRectangle(White, null, new Rect(19, 17, 4, 3));
        }

        private static void DrawMic(DrawingContext dc)
        {
            // Microphone with stand
            dc.DrawEllipse(White, null, new Point(16, 12), 4, 5);
            // U-shaped holder
            var g = new StreamGeometry();
            using (var c = g.Open())
            {
                c.BeginFigure(new Point(9, 13), false, false);
                c.ArcTo(new Point(23, 13), new Size(7, 8), 0, false,
                    SweepDirection.Clockwise, true, true);
            }
            g.Freeze();
            dc.DrawGeometry(null, Thin, g);
            // Stand
            dc.DrawLine(Thick, new Point(16, 21), new Point(16, 25));
            dc.DrawLine(Thick, new Point(12, 25), new Point(20, 25));
        }

        private static void DrawPlus(DrawingContext dc)
        {
            // Large + sign (create similar)
            dc.DrawLine(Thick, new Point(16, 8), new Point(16, 24));
            dc.DrawLine(Thick, new Point(8, 16), new Point(24, 16));
        }

        private static void DrawPiles(DrawingContext dc)
        {
            // Three circles in triangle layout (piles)
            dc.DrawEllipse(null, Thick, new Point(11, 12), 4, 4);
            dc.DrawEllipse(null, Thick, new Point(21, 12), 4, 4);
            dc.DrawEllipse(null, Thick, new Point(16, 21), 4, 4);
        }

        // =================================================================
        // Shared helpers
        // =================================================================

        private static StreamGeometry Arrow(double x1, double y1, double x2, double y2, double head)
        {
            var g = new StreamGeometry();
            using (var c = g.Open())
            {
                // Shaft
                c.BeginFigure(new Point(x1, y1), false, false);
                c.LineTo(new Point(x2, y2), true, true);

                // Arrowhead
                double dx = x2 - x1, dy = y2 - y1;
                double len = Math.Sqrt(dx * dx + dy * dy);
                if (len > 0)
                {
                    double ux = dx / len, uy = dy / len;
                    double px = -uy, py = ux;
                    double bx = x2 - ux * head, by = y2 - uy * head;

                    c.BeginFigure(new Point(bx + px * head, by + py * head), false, false);
                    c.LineTo(new Point(x2, y2), true, true);
                    c.LineTo(new Point(bx - px * head, by - py * head), true, true);
                }
            }
            g.Freeze();
            return g;
        }

        private static void CenterText(DrawingContext dc, string text, double fontSize)
        {
            var ft = new FormattedText(
                text,
                CultureInfo.InvariantCulture,
                FlowDirection.LeftToRight,
                new Typeface(new FontFamily("Segoe UI"), FontStyles.Normal, FontWeights.Bold, FontStretches.Normal),
                fontSize,
                White,
                1.0);
            dc.DrawText(ft, new Point((32 - ft.Width) / 2, (32 - ft.Height) / 2));
        }

        private static void Label(DrawingContext dc, string text, double x)
        {
            var ft = new FormattedText(
                text,
                CultureInfo.InvariantCulture,
                FlowDirection.LeftToRight,
                new Typeface(new FontFamily("Segoe UI"), FontStyles.Normal, FontWeights.Bold, FontStretches.Normal),
                10,
                White,
                1.0);
            dc.DrawText(ft, new Point(x, 12));
        }
    }
}

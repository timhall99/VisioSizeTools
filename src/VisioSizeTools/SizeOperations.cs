using System;
using System.Collections.Generic;
using System.Linq;
using Visio = Microsoft.Office.Interop.Visio;

namespace VisioSizeTools
{
    internal enum SizeMode
    {
        Tallest,
        Shortest,
        Widest,
        Narrowest,
        BothLarger,
        BothSmaller,
        FirstHeight,
        FirstWidth,
    }

    internal static class SizeOperations
    {
        private const short VisDrawingWindow = 1; // VisWinTypes.visDrawing

        /// <summary>
        /// Returns the selected shapes that can be resized: 2-D shapes whose width/height are not locked.
        /// 1-D shapes (connectors, lines) are ignored because their "height" is not meaningful.
        /// </summary>
        public static List<Visio.Shape> GetResizableSelection(Visio.Application app)
        {
            var result = new List<Visio.Shape>();
            Visio.Window window;
            try
            {
                window = app.ActiveWindow;
            }
            catch
            {
                return result;
            }

            if (window == null || window.Type != VisDrawingWindow)
                return result;

            Visio.Selection selection = window.Selection;
            if (selection == null)
                return result;

            int count = selection.Count;
            for (int i = 1; i <= count; i++)
            {
                Visio.Shape shape = selection[i];
                if (shape.OneD != 0)
                    continue;
                result.Add(shape);
            }
            return result;
        }

        public static int CountResizable(Visio.Application app)
        {
            try
            {
                return GetResizableSelection(app).Count;
            }
            catch
            {
                return 0;
            }
        }

        /// <summary>
        /// Applies the size operation. When <paramref name="alignToFirst"/> is true the shapes are then aligned to the
        /// first selected (primary) shape: bottoms for height modes, left edges for width modes, both for "Both" modes.
        /// Returns the number of shapes that could not be fully changed.
        /// </summary>
        public static int Apply(Visio.Application app, SizeMode mode, bool alignToFirst = false)
        {
            List<Visio.Shape> shapes = GetResizableSelection(app);
            if (shapes.Count < 2)
                return 0;

            var sizes = shapes.Select(s => (Width: Cell(s, "Width").ResultIU, Height: Cell(s, "Height").ResultIU)).ToList();

            double? targetWidth = null;
            double? targetHeight = null;
            switch (mode)
            {
                case SizeMode.Tallest:
                    targetHeight = sizes.Max(s => s.Height);
                    break;
                case SizeMode.Shortest:
                    targetHeight = sizes.Min(s => s.Height);
                    break;
                case SizeMode.Widest:
                    targetWidth = sizes.Max(s => s.Width);
                    break;
                case SizeMode.Narrowest:
                    targetWidth = sizes.Min(s => s.Width);
                    break;
                case SizeMode.BothLarger:
                    targetWidth = sizes.Max(s => s.Width);
                    targetHeight = sizes.Max(s => s.Height);
                    break;
                case SizeMode.BothSmaller:
                    targetWidth = sizes.Min(s => s.Width);
                    targetHeight = sizes.Min(s => s.Height);
                    break;
                case SizeMode.FirstHeight:
                    targetHeight = sizes[0].Height;
                    break;
                case SizeMode.FirstWidth:
                    targetWidth = sizes[0].Width;
                    break;
            }

            bool alignBottom = alignToFirst && targetHeight.HasValue;
            bool alignLeft = alignToFirst && targetWidth.HasValue;

            var failed = new bool[shapes.Count];

            // Anchor on the first selected shape's edges as they are *before* resizing, so its bottom
            // (or left) edge stays put and everything else lines up with it.
            double refLeft = 0, refBottom = 0;
            if (alignBottom || alignLeft)
                GetBounds(shapes[0], out refLeft, out refBottom);
            int undoScope = app.BeginUndoScope(UndoName(mode) + (alignToFirst ? " and Align" : ""));
            bool commit = false;
            try
            {
                for (int i = 0; i < shapes.Count; i++)
                {
                    bool ok = true;
                    if (targetWidth.HasValue)
                        ok &= TrySet(shapes[i], "Width", "LockWidth", targetWidth.Value);
                    if (targetHeight.HasValue)
                        ok &= TrySet(shapes[i], "Height", "LockHeight", targetHeight.Value);
                    failed[i] |= !ok;
                }

                if (alignBottom || alignLeft)
                {
                    // Shapes resize about their pin, so edges have moved; re-read each box and shift it
                    // (including the reference shape) back onto the anchor edge.
                    for (int i = 0; i < shapes.Count; i++)
                    {
                        GetBounds(shapes[i], out double left, out double bottom);
                        bool ok = true;
                        if (alignLeft)
                            ok &= TryMove(shapes[i], "PinX", "LockMoveX", refLeft - left);
                        if (alignBottom)
                            ok &= TryMove(shapes[i], "PinY", "LockMoveY", refBottom - bottom);
                        failed[i] |= !ok;
                    }
                }
                commit = true;
            }
            finally
            {
                app.EndUndoScope(undoScope, commit);
            }
            return failed.Count(f => f);
        }

        private const short VisBBoxUprightWH = 1;

        /// <summary>Left and bottom of the shape's upright width/height box in its parent's coordinates
        /// (the same box Visio's own Align uses, so rotated shapes line up the way users expect).</summary>
        private static void GetBounds(Visio.Shape shape, out double left, out double bottom)
        {
            shape.BoundingBox(VisBBoxUprightWH, out left, out bottom, out _, out _);
        }

        private static bool TryMove(Visio.Shape shape, string pinCellName, string lockCellName, double delta)
        {
            if (Math.Abs(delta) < 1e-9)
                return true;
            if (Cell(shape, lockCellName).ResultIU != 0)
                return false;
            try
            {
                Visio.Cell pin = Cell(shape, pinCellName);
                pin.ResultIU = pin.ResultIU + delta;
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static bool TrySet(Visio.Shape shape, string cellName, string lockCellName, double value)
        {
            Visio.Cell cell = Cell(shape, cellName);
            if (Math.Abs(cell.ResultIU - value) < 1e-9)
                return true;

            if (Cell(shape, lockCellName).ResultIU != 0)
                return false;

            try
            {
                cell.ResultIU = value; // respects GUARD(); guarded cells throw and are skipped
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static Visio.Cell Cell(Visio.Shape shape, string name) => shape.get_CellsU(name);

        public static string UndoName(SizeMode mode)
        {
            switch (mode)
            {
                case SizeMode.Tallest: return "Size to Tallest";
                case SizeMode.Shortest: return "Size to Shortest";
                case SizeMode.Widest: return "Size to Widest";
                case SizeMode.Narrowest: return "Size to Narrowest";
                case SizeMode.BothLarger: return "Size Both Larger";
                case SizeMode.BothSmaller: return "Size Both Smaller";
                case SizeMode.FirstHeight: return "Size to First: Height";
                case SizeMode.FirstWidth: return "Size to First: Width";
                default: return "Size";
            }
        }
    }
}

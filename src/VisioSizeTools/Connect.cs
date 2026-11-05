using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using Visio = Microsoft.Office.Interop.Visio;

namespace VisioSizeTools
{
    /// <summary>
    /// COM add-in entry point. Visio loads it via HKCU\Software\Microsoft\Visio\Addins\VisioSizeTools.Connect.
    /// Ribbon callbacks are late-bound through IDispatch, so the class keeps the default AutoDispatch interface.
    /// </summary>
    [ComVisible(true)]
    [Guid(ClassId)]
    [ProgId(ProgIdValue)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    public class Connect : IDTExtensibility2, Office.IRibbonExtensibility
    {
        public const string ClassId = "4A11EACB-6B73-4915-AD54-D92ECF41F340";
        public const string ProgIdValue = "VisioSizeTools.Connect";

        private static readonly string[] ControlIds =
        {
            "vstSizeMenu", "vstSizeTallest", "vstSizeShortest", "vstSizeWidest", "vstSizeNarrowest", "vstSizeBothLarger", "vstSizeBothSmaller",
        };

        private Visio.Application _app;
        private Office.IRibbonUI _ribbon;
        private bool _lastEnabled;

        #region IDTExtensibility2

        public void OnConnection(object application, int connectMode, object addInInst, IntPtr custom)
        {
            Log("OnConnection");
            _app = (Visio.Application)application;
            Log("cast ok");
            HookEvents();
            Log("events ok");
        }

        [System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.NoInlining)]
        private void HookEvents()
        {
            _app.SelectionChanged += OnSelectionChanged;
            _app.WindowActivated += OnWindowActivated;
        }

        public Connect()
        {
            Log("ctor");
        }

        public void OnDisconnection(int removeMode, IntPtr custom)
        {
            Log("OnDisconnection mode=" + removeMode);
            if (_app != null)
            {
                try
                {
                    _app.SelectionChanged -= OnSelectionChanged;
                    _app.WindowActivated -= OnWindowActivated;
                }
                catch
                {
                    // Visio may already be tearing down.
                }
            }
            _ribbon = null;
            _app = null;

            // Release every COM wrapper (Application, Shapes, ribbon) now rather than at some later GC,
            // so nothing this add-in touched can hold Visio open after its window closes.
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
            Log("OnDisconnection done");
        }

        public void OnAddInsUpdate(IntPtr custom) { }
        public void OnStartupComplete(IntPtr custom) { }
        public void OnBeginShutdown(IntPtr custom) => Log("OnBeginShutdown");

        #endregion

        #region IRibbonExtensibility + callbacks

        public string GetCustomUI(string ribbonId)
        {
            Log("GetCustomUI " + ribbonId);
            using (Stream s = Assembly.GetExecutingAssembly().GetManifestResourceStream("VisioSizeTools.Ribbon.xml"))
            using (var reader = new StreamReader(s))
                return reader.ReadToEnd();
        }

        public void OnRibbonLoad(Office.IRibbonUI ribbon)
        {
            Log("OnRibbonLoad");
            _ribbon = ribbon;
            _lastEnabled = IsEnabled();
        }

        public object GetImage(Office.IRibbonControl control) => Icons.Create(control.Id);

        public bool GetEnabled(Office.IRibbonControl control) => IsEnabled();

        public void OnAction(Office.IRibbonControl control)
        {
            SizeMode mode;
            switch (control.Id)
            {
                case "vstSizeTallest": mode = SizeMode.Tallest; break;
                case "vstSizeShortest": mode = SizeMode.Shortest; break;
                case "vstSizeWidest": mode = SizeMode.Widest; break;
                case "vstSizeNarrowest": mode = SizeMode.Narrowest; break;
                case "vstSizeBothLarger": mode = SizeMode.BothLarger; break;
                case "vstSizeBothSmaller": mode = SizeMode.BothSmaller; break;
                default: return;
            }

            try
            {
                // Shift held while clicking: also align to the first selected shape (bottom for height, left for width).
                bool align = (Control.ModifierKeys & Keys.Shift) == Keys.Shift;
                int skipped = SizeOperations.Apply(_app, mode, align);
                if (skipped > 0)
                {
                    MessageBox.Show(
                        $"{skipped} shape(s) could not be fully resized or moved because they are locked or protected.",
                        SizeOperations.UndoName(mode), MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, SizeOperations.UndoName(mode), MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        #endregion


        // Diagnostic trace (Debug builds only) -> %LOCALAPPDATA%/VisioSizeTools.log
        [System.Diagnostics.Conditional("DEBUG")]
        private static void Log(string message)
        {
            try { File.AppendAllText(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "VisioSizeTools.log"), DateTime.Now.ToString("HH:mm:ss.fff ") + message + Environment.NewLine); }
            catch { }
        }

        #region Selection tracking

        private bool IsEnabled() => _app != null && SizeOperations.CountResizable(_app) >= 2;

        private void OnSelectionChanged(Visio.Window window) => RefreshEnabledState();

        private void OnWindowActivated(Visio.Window window) => RefreshEnabledState();

        private void RefreshEnabledState()
        {
            if (_ribbon == null)
                return;

            bool enabled = IsEnabled();
            if (enabled == _lastEnabled)
                return;

            _lastEnabled = enabled;
            foreach (string id in ControlIds)
                _ribbon.InvalidateControl(id);
        }

        #endregion
    }
}

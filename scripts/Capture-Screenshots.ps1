<#
  Regenerates the README images in docs\images by driving a real Visio instance with the add-in loaded.

    ribbon.png               Home tab: Arrange group + the add-in's Size group      (screen grab)
    menu.png                 The Size menu dropped down                              (screen grab)
    before.png               The demo drawing before any operation                   (page export)
    <operation>.png          The demo drawing after each menu item is clicked        (page export)
    <operation>-align.png    ... with Shift held (align to the first selected shape)

  Every operation is performed by clicking the add-in's own ribbon menu items through UI Automation,
  so the images are produced by the shipped code path, not a re-implementation.

  Prerequisites: Visio desktop (2013+), and the add-in installed (the Setup .exe) or registered
  with scripts\Register.ps1. Visio cannot run on CI runners, so run this locally:

      powershell -ExecutionPolicy Bypass -File scripts\Capture-Screenshots.ps1

  A separate Visio instance is started and closed again; any Visio you already have open is not touched.
  Don't use the mouse/keyboard while it runs (about 30 seconds) - screen grabs need the window in front.
#>
param([string]$OutDir = (Join-Path $PSScriptRoot '..\docs\images'))

$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes, System.Drawing
Add-Type -Namespace VstCapture -Name Win32 -MemberDefinition @'
[DllImport("user32.dll")] public static extern bool SetProcessDpiAwarenessContext(IntPtr value);
[DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hWnd);
[DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
[DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
[DllImport("user32.dll")] public static extern void keybd_event(byte vk, byte scan, uint flags, UIntPtr extra);
[DllImport("user32.dll")] public static extern bool BringWindowToTop(IntPtr hWnd);
[DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, IntPtr processId);
[DllImport("user32.dll")] public static extern bool AttachThreadInput(uint idAttach, uint idAttachTo, bool fAttach);
[DllImport("kernel32.dll")] public static extern uint GetCurrentThreadId();
[DllImport("user32.dll", EntryPoint = "GetWindowThreadProcessId")] public static extern uint GetWindowProcessId(IntPtr hWnd, out uint processId);
'@
# A blocked out-of-process COM call can't be cancelled from PowerShell, so a timer thread ends the Visio
# instance this script started if a step overruns; the stuck call then fails and the step reports it.
Add-Type -TypeDefinition @'
using System.Diagnostics;
using System.Threading;
namespace VstCapture {
    public static class Watchdog {
        static Timer _timer;
        public static volatile bool Fired;
        public static void Arm(int processId, int milliseconds) {
            Disarm();
            Fired = false;
            _timer = new Timer(delegate {
                Fired = true;
                try { Process.GetProcessById(processId).Kill(); } catch { }
            }, null, milliseconds, Timeout.Infinite);
        }
        public static void Disarm() {
            if (_timer != null) { _timer.Dispose(); _timer = null; }
        }
    }
}
'@
# Per-monitor-v2 DPI awareness so UI Automation rectangles and screen pixels agree on scaled displays.
[void][VstCapture.Win32]::SetProcessDpiAwarenessContext([IntPtr]-4)

$UIA = [System.Windows.Automation.AutomationElement]
$VK_SHIFT = 0x10; $KEYUP = 0x2

New-Item -ItemType Directory -Force -Path $OutDir | Out-Null
$OutDir = (Resolve-Path $OutDir).Path

# ---------------------------------------------------------------- helpers

function Find-ByName($root, [string]$name, $controlType, [int]$timeoutMs = 5000) {
    $cond = New-Object System.Windows.Automation.AndCondition(
        (New-Object System.Windows.Automation.PropertyCondition($UIA::NameProperty, $name)),
        (New-Object System.Windows.Automation.PropertyCondition($UIA::ControlTypeProperty, $controlType)))
    $sw = [Diagnostics.Stopwatch]::StartNew()
    do {
        $el = $root.FindFirst([System.Windows.Automation.TreeScope]::Descendants, $cond)
        if ($el) { return $el }
        Start-Sleep -Milliseconds 200
    } while ($sw.ElapsedMilliseconds -lt $timeoutMs)
    throw "UI element not found: $name ($($controlType.ProgrammaticName))"
}

# Office hangs the dropped-down menu (a "Menu" element) and its items off the ribbon anchor itself.
function Open-SizeMenu($sizeMenu) {
    $sizeMenu.GetCurrentPattern([System.Windows.Automation.ExpandCollapsePattern]::Pattern).Expand()
    Find-ByName $sizeMenu 'Size' ([System.Windows.Automation.ControlType]::Menu)
}

function Union-Rect([System.Windows.Rect[]]$rects, [int]$pad) {
    $r = $rects[0]; foreach ($x in $rects) { $r.Union($x) }
    $r.Inflate($pad, $pad); $r
}

function Save-ScreenRect([System.Windows.Rect]$r, [string]$file) {
    $bmp = New-Object System.Drawing.Bitmap([int]$r.Width, [int]$r.Height)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.CopyFromScreen([int]$r.X, [int]$r.Y, 0, 0, $bmp.Size)
    $bmp.Save((Join-Path $OutDir $file), [System.Drawing.Imaging.ImageFormat]::Png)
    $g.Dispose(); $bmp.Dispose()
    Write-Host "  $file"
}

function Bring-ToFront([IntPtr]$hwnd) {
    # Borrow the current foreground window's input state so SetForegroundWindow is honoured.
    # (Don't use the usual "tap Alt" trick: in Office that enters ribbon KeyTips mode, which blocks COM calls.)
    $fg = [VstCapture.Win32]::GetForegroundWindow()
    if ($fg -eq $hwnd) { return }
    $fgThread = [VstCapture.Win32]::GetWindowThreadProcessId($fg, [IntPtr]::Zero)
    $me = [VstCapture.Win32]::GetCurrentThreadId()
    [void][VstCapture.Win32]::AttachThreadInput($me, $fgThread, $true)
    try {
        [void][VstCapture.Win32]::BringWindowToTop($hwnd)
        [void][VstCapture.Win32]::SetForegroundWindow($hwnd)
    } finally {
        [void][VstCapture.Win32]::AttachThreadInput($me, $fgThread, $false)
    }
    Start-Sleep -Milliseconds 500
    if ([VstCapture.Win32]::GetForegroundWindow() -ne $hwnd) { throw 'Could not bring the Visio window to the front.' }
}

# ---------------------------------------------------------------- demo drawing

$PageW = 9.0; $PageH = 5.0
# Three rectangles of different widths, heights and bottoms; "1" is selected first (the align reference).
$Demo = @(
    @{ Text = '1'; X1 = 1.0; Y1 = 1.5; X2 = 2.5; Y2 = 3.5; Fill = 'RGB(57,85,163)';   Font = 'RGB(255,255,255)' }
    @{ Text = '2'; X1 = 3.5; Y1 = 1.1; X2 = 6.0; Y2 = 2.3; Fill = 'RGB(205,216,240)'; Font = 'RGB(38,58,115)' }
    @{ Text = '3'; X1 = 7.0; Y1 = 1.8; X2 = 8.0; Y2 = 4.5; Fill = 'RGB(205,216,240)'; Font = 'RGB(38,58,115)' }
)

function Reset-Drawing($app, $page) {
    $win = $app.ActiveWindow
    $win.DeselectAll()
    while ($page.Shapes.Count -gt 0) { $page.Shapes.Item(1).Delete() }

    # Borderless white backdrop the size of the page, so every export has identical extents.
    $bg = $page.DrawRectangle(0, 0, $PageW, $PageH)
    $bg.CellsU('LinePattern').FormulaU = '0'
    $bg.CellsU('FillForegnd').FormulaU = 'RGB(255,255,255)'

    $shapes = foreach ($d in $Demo) {
        $s = $page.DrawRectangle($d.X1, $d.Y1, $d.X2, $d.Y2)
        $s.Text = $d.Text
        $s.CellsU('FillForegnd').FormulaU = $d.Fill
        $s.CellsU('LineColor').FormulaU = 'RGB(38,58,115)'
        $s.CellsU('LineWeight').FormulaU = '1.5 pt'
        $s.CellsU('Char.Size').FormulaU = '28 pt'
        $s.CellsU('Char.Color').FormulaU = $d.Font
        $s
    }
    # Drawing leaves the last new shape selected, so clear it first or "3" would become the primary selection.
    $win.DeselectAll()
    foreach ($s in $shapes) { $win.Select($s, 2) }   # 2 = visSelect, in order: "1" is the primary selection
    Start-Sleep -Milliseconds 400                    # let the add-in re-enable its menu
}

function Export-Page($page, [string]$file) {
    # Must be a plain [string]: Join-Path's PSObject-wrapped result makes Page.Export block forever.
    [string]$path = Join-Path $OutDir $file
    $page.Export($path)
    Write-Host "  $file"
}

$StepTimeoutSeconds = 10

# Runs one step under the watchdog. If Visio stops responding, the instance is closed and the step fails
# with a message naming it, instead of the whole script waiting forever.
function Invoke-Step([string]$name, [scriptblock]$body) {
    [VstCapture.Watchdog]::Arm($script:visioPid, $StepTimeoutSeconds * 1000)
    try {
        . $body
    } catch {
        if ([VstCapture.Watchdog]::Fired) {
            throw "Step '$name' got no response from Visio within $StepTimeoutSeconds s, so the Visio instance was closed."
        }
        throw
    } finally {
        [VstCapture.Watchdog]::Disarm()
    }
    if ([VstCapture.Watchdog]::Fired) {
        throw "Step '$name' got no response from Visio within $StepTimeoutSeconds s, so the Visio instance was closed."
    }
}

# ---------------------------------------------------------------- run

# Visio's Application.ProcessID isn't the Windows PID, so take the PID from the main window instead.
$existing = @(Get-Process VISIO -ErrorAction SilentlyContinue).Id
$app = New-Object -ComObject Visio.Application
$hwnd = [IntPtr]$app.WindowHandle32
$visioPid = [uint32]0
[void][VstCapture.Win32]::GetWindowProcessId($hwnd, [ref]$visioPid)
$visioPid = [int]$visioPid
$doc = $null
try {
    # The watchdog kills the instance on timeout, so never run against a Visio someone already had open.
    if ($existing -contains $visioPid) { throw 'Visio attached to an already running instance; close Visio and run again.' }

    Invoke-Step 'check add-in' {
        $addIn = $app.COMAddIns | Where-Object { $_.ProgId -eq 'VisioSizeTools.Connect' }
        if (-not $addIn -or -not $addIn.Connect) {
            throw 'The VisioSizeTools add-in is not loaded. Install it, or run scripts\Register.ps1 first.'
        }
    }

    Invoke-Step 'set up document' {
        $script:doc = $app.Documents.Add('')
        $script:page = $app.ActivePage
        $page.PageSheet.CellsU('PageWidth').FormulaU = "$PageW in"
        $page.PageSheet.CellsU('PageHeight').FormulaU = "$PageH in"
        # 2x the screen resolution keeps the exported images sharp on high-DPI displays and GitHub.
        $app.Settings.SetRasterExportResolution(3, 192, 192, 0)   # visRasterUseCustomResolution, px/in
        $app.Settings.RasterExportBackgroundColor = 0xFFFFFF

        [void][VstCapture.Win32]::ShowWindow($hwnd, 3)            # SW_MAXIMIZE
        $script:root = $UIA::FromHandle($hwnd)

        Reset-Drawing $app $page
        $app.ActiveWindow.ViewFit = 1                             # visFitPage
    }

    Write-Host 'Screen grabs'
    Invoke-Step 'ribbon screen grab' {
        Bring-ToFront $hwnd
        $sizeGroup = Find-ByName $root 'Size' ([System.Windows.Automation.ControlType]::Group)
        $arrangeGroup = Find-ByName $root 'Arrange' ([System.Windows.Automation.ControlType]::Group)
        $script:sizeMenu = Find-ByName $sizeGroup 'Size' ([System.Windows.Automation.ControlType]::MenuItem)
        $script:ribbonRect = Union-Rect @($arrangeGroup.Current.BoundingRectangle, $sizeGroup.Current.BoundingRectangle) 6
        Save-ScreenRect $ribbonRect 'ribbon.png'
    }
    Invoke-Step 'menu screen grab' {
        $popup = Open-SizeMenu $sizeMenu
        Start-Sleep -Milliseconds 400                             # menu open animation
        Save-ScreenRect (Union-Rect @($ribbonRect, $popup.Current.BoundingRectangle) 6) 'menu.png'
        $sizeMenu.GetCurrentPattern([System.Windows.Automation.ExpandCollapsePattern]::Pattern).Collapse()
    }

    Write-Host 'Before / after'
    Invoke-Step 'export before.png' { Export-Page $page 'before.png' }

    $operations = @(
        @{ Label = 'Size to Tallest';   File = 'tallest' }
        @{ Label = 'Size to Shortest';  File = 'shortest' }
        @{ Label = 'Size to Widest';    File = 'widest' }
        @{ Label = 'Size to Narrowest'; File = 'narrowest' }
        @{ Label = 'Size Both Larger';  File = 'both-larger' }
        @{ Label = 'Size Both Smaller'; File = 'both-smaller' }
        # Shift + height mode aligns bottoms to shape 1 (width/both modes would stack the shapes on top of each other).
        @{ Label = 'Size to Tallest';   File = 'tallest-align';  Shift = $true }
        @{ Label = 'Size to Shortest';  File = 'shortest-align'; Shift = $true }
    )
    foreach ($op in $operations) {
        Invoke-Step "$($op.File)" {
            Reset-Drawing $app $page
            Bring-ToFront $hwnd
            $popup = Open-SizeMenu $sizeMenu
            $item = Find-ByName $popup $op.Label ([System.Windows.Automation.ControlType]::MenuItem)
            if ($op.Shift) { [VstCapture.Win32]::keybd_event($VK_SHIFT, 0, 0, [UIntPtr]::Zero) }
            try {
                $item.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern).Invoke()
                Start-Sleep -Milliseconds 600                 # onAction runs on Visio's UI thread
            } finally {
                if ($op.Shift) { [VstCapture.Win32]::keybd_event($VK_SHIFT, 0, $KEYUP, [UIntPtr]::Zero) }
            }
            $app.ActiveWindow.DeselectAll()
            Export-Page $page "$($op.File).png"
        }
    }
    Write-Host "Done -> $OutDir"
}
finally {
    [VstCapture.Watchdog]::Disarm()
    [VstCapture.Win32]::keybd_event($VK_SHIFT, 0, $KEYUP, [UIntPtr]::Zero)
    # Close our document and instance; if the watchdog already ended Visio there is nothing to close.
    $proc = Get-Process -Id $visioPid -ErrorAction SilentlyContinue
    if ($proc) {
        try {
            if ($doc) { $doc.Saved = $true; $doc.Close() }
            $app.Quit()
        } catch {
            Stop-Process -Id $visioPid -Force -ErrorAction SilentlyContinue
        }
    }
}

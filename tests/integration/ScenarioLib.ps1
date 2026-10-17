# Helpers for Run-Scenarios.ps1 (dot-source after FeTest.ps1): real mouse/keyboard, window queries, screen sampling,
# geometry oracles computed independently of the add-in, and the PASS/FAIL bookkeeping.
. "$PSScriptRoot\FeTest.ps1"
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Windows.Forms

Add-Type -ReferencedAssemblies System.Drawing @"
using System; using System.Collections.Generic; using System.Drawing; using System.Drawing.Imaging;
using System.Runtime.InteropServices; using System.Text;
public static class FeWin2 {
  [StructLayout(LayoutKind.Sequential)] public struct RECT { public int Left, Top, Right, Bottom; }
  [StructLayout(LayoutKind.Sequential)] public struct POINT { public int X, Y; }
  [StructLayout(LayoutKind.Sequential)] public struct CURSORINFO { public int cbSize; public int flags; public IntPtr hCursor; public POINT pt; }
  [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
  [DllImport("user32.dll")] public static extern bool GetCursorPos(out POINT p);
  [DllImport("user32.dll")] public static extern bool GetCursorInfo(ref CURSORINFO ci);
  [DllImport("user32.dll")] public static extern IntPtr LoadCursor(IntPtr h, IntPtr id);
  [DllImport("user32.dll")] public static extern bool SetCursorPos(int x, int y);
  [DllImport("user32.dll")] public static extern void mouse_event(uint f, int dx, int dy, int d, UIntPtr e);
  [DllImport("user32.dll")] public static extern void keybd_event(byte vk, byte scan, uint flags, UIntPtr extra);
  [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out RECT r);
  [DllImport("user32.dll")] public static extern bool GetClientRect(IntPtr h, out RECT r);
  [DllImport("user32.dll")] public static extern bool SetWindowPos(IntPtr h, IntPtr after, int x, int y, int cx, int cy, uint flags);
  [DllImport("user32.dll")] public static extern IntPtr GetWindow(IntPtr h, uint cmd);
  [DllImport("user32.dll")] public static extern IntPtr GetAncestor(IntPtr h, uint flags);
  [DllImport("user32.dll")] public static extern bool IsWindow(IntPtr h);
  [DllImport("user32.dll")] public static extern bool IsIconic(IntPtr h);
  [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetClassName(IntPtr h, StringBuilder sb, int max);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetWindowText(IntPtr h, StringBuilder sb, int max);
  [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
  [DllImport("user32.dll")] public static extern IntPtr WindowFromPoint(POINT p);
  [DllImport("user32.dll")] public static extern IntPtr GetParent(IntPtr h);
  [DllImport("user32.dll")] public static extern int GetSystemMetrics(int i);
  [DllImport("user32.dll")] public static extern IntPtr GetWindowLongPtr(IntPtr h, int idx);
  [DllImport("user32.dll")] public static extern uint GetDpiForWindow(IntPtr h);
  [DllImport("user32.dll")] public static extern bool PostMessage(IntPtr h, uint msg, IntPtr w, IntPtr l);
  public delegate bool EnumProc(IntPtr h, IntPtr l);
  [DllImport("user32.dll")] public static extern bool EnumWindows(EnumProc cb, IntPtr l);
  [DllImport("user32.dll")] public static extern bool EnumChildWindows(IntPtr p, EnumProc cb, IntPtr l);

  [StructLayout(LayoutKind.Sequential)] public struct GUITHREADINFO { public int cbSize; public int flags; public IntPtr hwndActive; public IntPtr hwndFocus; public IntPtr hwndCapture; public IntPtr hwndMenuOwner; public IntPtr hwndMoveSize; public IntPtr hwndCaret; public RECT rcCaret; }
  [DllImport("user32.dll")] public static extern bool GetGUIThreadInfo(uint tid, ref GUITHREADINFO info);
  // "active=0x..[class] focus=0x..[class]" of the thread owning window h (the Excel UI thread).
  public static string ThreadFocus(IntPtr anyWindowOfThread) {
    uint pid; uint tid = GetWindowThreadProcessId(anyWindowOfThread, out pid);
    var gi = new GUITHREADINFO(); gi.cbSize = Marshal.SizeOf(typeof(GUITHREADINFO));
    GetGUIThreadInfo(tid, ref gi);
    return "active=0x" + gi.hwndActive.ToInt64().ToString("X") + "[" + ClassOf(gi.hwndActive) + "] focus=0x" + gi.hwndFocus.ToInt64().ToString("X") + "[" + ClassOf(gi.hwndFocus) + "]";
  }
  public static string ClassOf(IntPtr h) { var sb = new StringBuilder(256); GetClassName(h, sb, 256); return sb.ToString(); }
  public static string TitleOf(IntPtr h) { var sb = new StringBuilder(512); GetWindowText(h, sb, 512); return sb.ToString(); }
  public static int[] Rect(IntPtr h) { RECT r; if (!GetWindowRect(h, out r)) return null; return new[] { r.Left, r.Top, r.Right, r.Bottom }; }
  public static int[] CursorPos() { POINT p; GetCursorPos(out p); return new[] { p.X, p.Y }; }
  public static IntPtr CurrentCursor() { var ci = new CURSORINFO(); ci.cbSize = Marshal.SizeOf(typeof(CURSORINFO)); GetCursorInfo(ref ci); return ci.hCursor; }
  public static IntPtr SystemCursor(int id) { return LoadCursor(IntPtr.Zero, (IntPtr)id); }
  public static IntPtr TopWindowAt(int x, int y) { POINT p; p.X = x; p.Y = y; var h = WindowFromPoint(p); return h == IntPtr.Zero ? h : GetAncestor(h, 2); }

  // Visible top-level windows of a process: "hwnd|class|title|left,top,right,bottom".
  public static List<string> TopLevel(uint pid) {
    var list = new List<string>();
    EnumWindows((h, l) => { uint p; GetWindowThreadProcessId(h, out p);
      if (p == pid && IsWindowVisible(h)) { var r = Rect(h); list.Add(h.ToInt64() + "|" + ClassOf(h) + "|" + TitleOf(h) + "|" + string.Join(",", r)); }
      return true; }, IntPtr.Zero);
    return list;
  }
  public static List<string> ChildrenOf(IntPtr parent) {
    var list = new List<string>();
    EnumChildWindows(parent, (h, l) => { var r = Rect(h); list.Add(h.ToInt64() + "|" + ClassOf(h) + "|" + string.Join(",", r) + "|" + IsWindowVisible(h)); return true; }, IntPtr.Zero);
    return list;
  }

  // Real input. Absolute move over the whole virtual desktop so Windows sees a genuine mouse move.
  public static void MoveTo(int x, int y) {
    SetCursorPos(x, y);
    int vx = GetSystemMetrics(76), vy = GetSystemMetrics(77), vw = GetSystemMetrics(78), vh = GetSystemMetrics(79);
    int ax = (int)Math.Ceiling((x - vx + 0.5) * 65536.0 / vw), ay = (int)Math.Ceiling((y - vy + 0.5) * 65536.0 / vh);
    mouse_event(0x0001 | 0x8000 | 0x4000, ax, ay, 0, UIntPtr.Zero);   // MOVE | ABSOLUTE | VIRTUALDESK
  }
  public static void Down(bool right) { mouse_event(right ? 8u : 2u, 0, 0, 0, UIntPtr.Zero); }
  public static void Up(bool right) { mouse_event(right ? 16u : 4u, 0, 0, 0, UIntPtr.Zero); }
  public static void Wheel(int delta) { mouse_event(0x0800, 0, 0, delta, UIntPtr.Zero); }
  public static void KeyDown(byte vk) { keybd_event(vk, 0, 0, UIntPtr.Zero); }
  public static void KeyUp(byte vk) { keybd_event(vk, 0, 2, UIntPtr.Zero); }

  // Screen sampling ------------------------------------------------------------------------------------------------
  public static Bitmap Capture(int x, int y, int w, int h) {
    var bmp = new Bitmap(w, h, PixelFormat.Format32bppArgb);
    using (var g = Graphics.FromImage(bmp)) g.CopyFromScreen(x, y, 0, 0, bmp.Size);
    return bmp;
  }
  // Number of pixels within 'tol' (per channel) of the colour, inside rect (x,y,w,h) of the bitmap.
  public static int CountColor(Bitmap bmp, int x, int y, int w, int h, int rgb, int tol) {
    int r0 = (rgb >> 16) & 255, g0 = (rgb >> 8) & 255, b0 = rgb & 255, n = 0;
    x = Math.Max(0, x); y = Math.Max(0, y); w = Math.Min(w, bmp.Width - x); h = Math.Min(h, bmp.Height - y);
    for (int j = y; j < y + h; j++) for (int i = x; i < x + w; i++) {
      var c = bmp.GetPixel(i, j);
      if (Math.Abs(c.R - r0) <= tol && Math.Abs(c.G - g0) <= tol && Math.Abs(c.B - b0) <= tol) n++;
    }
    return n;
  }
  // ---- bitmap differences (overlay on vs. off) ----
  static byte[] Bits(Bitmap b, out int stride) {
    var d = b.LockBits(new Rectangle(0, 0, b.Width, b.Height), ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb);
    stride = d.Stride; var buf = new byte[d.Stride * b.Height]; Marshal.Copy(d.Scan0, buf, 0, buf.Length); b.UnlockBits(d); return buf;
  }
  // Changed pixels between two same-size bitmaps: returns int[] { count, minX, minY, maxX, maxY }.
  public static int[] DiffStats(Bitmap a, Bitmap b) {
    int sa, sb; var A = Bits(a, out sa); var B = Bits(b, out sb);
    int n = 0, x0 = int.MaxValue, y0 = int.MaxValue, x1 = -1, y1 = -1;
    for (int y = 0; y < a.Height; y++) { int ra = y * sa, rb = y * sb;
      for (int x = 0; x < a.Width; x++) { int ia = ra + x * 4, ib = rb + x * 4;
        if (A[ia] != B[ib] || A[ia + 1] != B[ib + 1] || A[ia + 2] != B[ib + 2]) { n++; if (x < x0) x0 = x; if (x > x1) x1 = x; if (y < y0) y0 = y; if (y > y1) y1 = y; } } }
    return new[] { n, x0, y0, x1, y1 };
  }
  // Changed pixels NOT inside any of the rects ({l,t,r,b} flattened: l,t,r,b,l,t,r,b...). Returns { outsideCount, sampleX, sampleY }.
  public static int[] DiffOutside(Bitmap a, Bitmap b, int[] rects) {
    int sa, sb; var A = Bits(a, out sa); var B = Bits(b, out sb);
    int n = 0, sx = -1, sy = -1;
    for (int y = 0; y < a.Height; y++) { int ra = y * sa, rb = y * sb;
      for (int x = 0; x < a.Width; x++) { int ia = ra + x * 4, ib = rb + x * 4;
        if (A[ia] != B[ib] || A[ia + 1] != B[ib + 1] || A[ia + 2] != B[ib + 2]) {
          bool inside = false;
          for (int k = 0; k + 3 < rects.Length; k += 4) if (x >= rects[k] && x < rects[k + 2] && y >= rects[k + 1] && y < rects[k + 3]) { inside = true; break; }
          if (!inside) { if (n == 0) { sx = x; sy = y; } n++; } } } }
    return new[] { n, sx, sy };
  }
  // Changed pixels inside rect {l,t,r,b}.
  public static int DiffInside(Bitmap a, Bitmap b, int l, int t, int r, int bt) {
    int sa, sb; var A = Bits(a, out sa); var B = Bits(b, out sb); int n = 0;
    l = Math.Max(0, l); t = Math.Max(0, t); r = Math.Min(a.Width, r); bt = Math.Min(a.Height, bt);
    for (int y = t; y < bt; y++) for (int x = l; x < r; x++) { int ia = y * sa + x * 4, ib = y * sb + x * 4;
      if (A[ia] != B[ib] || A[ia + 1] != B[ib + 1] || A[ia + 2] != B[ib + 2]) n++; }
    return n;
  }
  // Changed pixels inside rect: { count, minX, minY, maxX, maxY }.
  public static int[] DiffBox(Bitmap a, Bitmap b, int l, int t, int r, int bt) {
    int sa, sb; var A = Bits(a, out sa); var B = Bits(b, out sb); int n = 0, x0 = int.MaxValue, y0 = int.MaxValue, x1 = -1, y1 = -1;
    l = Math.Max(0, l); t = Math.Max(0, t); r = Math.Min(a.Width, r); bt = Math.Min(a.Height, bt);
    for (int y = t; y < bt; y++) for (int x = l; x < r; x++) { int ia = y * sa + x * 4, ib = y * sb + x * 4;
      if (A[ia] != B[ib] || A[ia + 1] != B[ib + 1] || A[ia + 2] != B[ib + 2]) { n++; if (x < x0) x0 = x; if (x > x1) x1 = x; if (y < y0) y0 = y; if (y > y1) y1 = y; } }
    return new[] { n, x0, y0, x1, y1 };
  }
  [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr h, int cmd);
  public static int PixelRgb(Bitmap bmp, int x, int y) { var c = bmp.GetPixel(x, y); return (c.R << 16) | (c.G << 8) | c.B; }
  // The most "saturated/dark" pixel in a rect (darkest by luminance) — the ink colour of text drawn on a light background.
  public static int DarkestRgb(Bitmap bmp, int x, int y, int w, int h) {
    int best = 0xFFFFFF; double bl = 1e9;
    x = Math.Max(0, x); y = Math.Max(0, y); w = Math.Min(w, bmp.Width - x); h = Math.Min(h, bmp.Height - y);
    for (int j = y; j < y + h; j++) for (int i = x; i < x + w; i++) {
      var c = bmp.GetPixel(i, j); double l = 0.3 * c.R + 0.59 * c.G + 0.11 * c.B;
      if (l < bl) { bl = l; best = (c.R << 16) | (c.G << 8) | c.B; }
    }
    return best;
  }
}
"@

# --- constants -------------------------------------------------------------------------------------------------------
$script:VK = @{ Esc = 0x1B; Enter = 0x0D; F2 = 0x71; Ctrl = 0x11; Shift = 0x10; Alt = 0x12; H = 0x48; Tab = 0x09; Down = 0x28; Up = 0x26; Space = 0x20; Home = 0x24; Right = 0x27; Left = 0x25; F5 = 0x74; G = 0x47 }
$script:ExpectedPalette = @{ Blue = 0x0000FF; Green = 0x008000; Purple = 0x9900CC; Maroon = 0x800000; LightGreen = 0x00CC33; Orange = 0xFF6600; Magenta = 0xCC0099 }
$script:Gray = 0xDCDCDC
$script:Gold = 0xFFD700

# --- result bookkeeping ----------------------------------------------------------------------------------------------
$script:Results = [Collections.Generic.List[object]]::new()
$script:Cur = $null

function Begin-Scenario([string]$Id, [string]$Title) {
    $script:Cur = [pscustomobject]@{ id = $Id; title = $Title; checks = [Collections.Generic.List[object]]::new(); notes = [Collections.Generic.List[string]]::new(); error = $null }
    Write-Host ""
    Write-Host "=== $Id $Title" -ForegroundColor Cyan
}

# Level 'hard' failure => FAIL; 'soft' failure => PARTIAL.
function Check([string]$Name, [bool]$Ok, $Observed = $null, [ValidateSet('hard', 'soft')][string]$Level = 'hard') {
    $obs = if ($null -eq $Observed) { '' } elseif ($Observed -is [string]) { $Observed } else { ($Observed | ConvertTo-Json -Compress -Depth 4) }
    $script:Cur.checks.Add([pscustomobject]@{ name = $Name; ok = $Ok; level = $Level; observed = $obs })
    $tag = if ($Ok) { 'ok  ' } elseif ($Level -eq 'soft') { 'WARN' } else { 'FAIL' }
    $color = if ($Ok) { 'Green' } elseif ($Level -eq 'soft') { 'Yellow' } else { 'Red' }
    Write-Host ("  [{0}] {1}  {2}" -f $tag, $Name, $obs) -ForegroundColor $color
}

function Note([string]$Text) { $script:Cur.notes.Add($Text); Write-Host "  note: $Text" -ForegroundColor DarkGray }

function End-Scenario([string]$ForceStatus) {
    $c = $script:Cur
    $status = 'PASS'
    if ($c.error) { $status = 'FAIL' }
    elseif (@($c.checks | Where-Object { -not $_.ok -and $_.level -eq 'hard' }).Count -gt 0) { $status = 'FAIL' }
    elseif (@($c.checks | Where-Object { -not $_.ok }).Count -gt 0) { $status = 'PARTIAL' }
    if ($ForceStatus) { $status = $ForceStatus }
    $c | Add-Member -NotePropertyName status -NotePropertyValue $status -Force
    $script:Results.Add($c)
    $color = switch ($status) { 'PASS' { 'Green' } 'PARTIAL' { 'Yellow' } default { 'Red' } }
    Write-Host ("=== {0} {1}" -f $c.id, $status) -ForegroundColor $color
}

# --- generic waiting -------------------------------------------------------------------------------------------------
function Wait-Until([scriptblock]$Condition, [int]$TimeoutMs = 5000, [int]$PollMs = 50) {
    $sw = [Diagnostics.Stopwatch]::StartNew()
    do {
        $v = $false
        try { $v = [bool](& $Condition) } catch { }
        if ($v) { return $true }
        Start-Sleep -Milliseconds $PollMs
    } while ($sw.ElapsedMilliseconds -lt $TimeoutMs)
    $false
}

# --- window queries --------------------------------------------------------------------------------------------------
function Get-ExcelHwnd($fe) { [IntPtr]$fe.Excel.Hwnd }
function Get-ExplorerHwnd($fe) { [IntPtr][int64](Get-FeState $fe).explorerHwnd }
function Get-FgHwnd { [FeWin2]::GetForegroundWindow() }
function Test-ExplorerForeground($fe, $state) { [FeWin2]::GetForegroundWindow().ToInt64() -eq [int64]$state.explorerHwnd }
function Get-Rect([IntPtr]$h) { [FeWin2]::Rect($h) }
function Describe-Fg { $h = [FeWin2]::GetForegroundWindow(); "0x{0:X} [{1}] '{2}'" -f $h.ToInt64(), [FeWin2]::ClassOf($h), [FeWin2]::TitleOf($h) }

# --- mouse -----------------------------------------------------------------------------------------------------------
function Move-Mouse([int]$X, [int]$Y) { [FeWin2]::MoveTo($X, $Y); Start-Sleep -Milliseconds 40 }
function Click-At([int]$X, [int]$Y, [switch]$Double, [switch]$Right, [int]$GapMs = 90) {
    Move-Mouse $X $Y
    [FeWin2]::Down($Right.IsPresent); Start-Sleep -Milliseconds 25; [FeWin2]::Up($Right.IsPresent)
    if ($Double) { Start-Sleep -Milliseconds $GapMs; [FeWin2]::Down($false); Start-Sleep -Milliseconds 25; [FeWin2]::Up($false) }
}
function Press-Key([int]$Vk, [int[]]$Modifiers = @()) {
    foreach ($m in $Modifiers) { [FeWin2]::KeyDown([byte]$m) }
    [FeWin2]::KeyDown([byte]$Vk); Start-Sleep -Milliseconds 40; [FeWin2]::KeyUp([byte]$Vk)
    foreach ($m in $Modifiers) { [FeWin2]::KeyUp([byte]$m) }
    Start-Sleep -Milliseconds 80
}
function Get-CursorKind {
    $c = [FeWin2]::CurrentCursor()
    if ($c -eq [FeWin2]::SystemCursor(32649)) { return 'hand' }
    if ($c -eq [FeWin2]::SystemCursor(32512)) { return 'arrow' }
    if ($c -eq [FeWin2]::SystemCursor(32514)) { return 'wait' }
    if ($c -eq [FeWin2]::SystemCursor(32650)) { return 'appstarting' }
    if ($c -eq [FeWin2]::SystemCursor(32513)) { return 'ibeam' }
    "other(0x{0:X})" -f $c.ToInt64()
}

# --- state helpers ---------------------------------------------------------------------------------------------------
# Opens the explorer and waits until the root's text is displayed.
function Open-Explorer($fe, [string]$Sheet, [string]$Address, [switch]$KeepOpen) {
    if (-not $KeepOpen) { Close-Explorer $fe }
    if ($Sheet) { Select-FeCell $fe $Sheet $Address }
    $r = Invoke-Fe $fe 'FE_Test_Open'
    $s = Wait-FeState $fe { $_.isOpen -and $_.view -and $_.session.root -and -not $_.view.isLoading } 6000
    Start-Sleep -Milliseconds 250   # let the overlay's first tick land
    Get-FeState $fe
}

function Wait-Loaded($fe, [string]$TextLike, [int]$TimeoutMs = 5000) {
    Wait-FeState $fe { $_.isOpen -and $_.view -and -not $_.view.isLoading -and $_.view.displayText -like $TextLike } $TimeoutMs
}

# Char index of a token within the displayed text (see Find-FeIndex).
function Ix($state, [string]$Needle, [int]$Occurrence = 1, [int]$Offset = 0) { Find-FeIndex $state.view.displayText $Needle $Occurrence $Offset }

function Char-Point($state, [int]$Index) {
    $p = $state.charPoints."$Index"
    if (-not $p) { throw "No charPoint for index $Index" }
    [pscustomobject]@{ X = [int]$p[0]; Y = [int]$p[1] }
}

function Hover-Char($fe, [int]$Index) {
    Invoke-Fe $fe 'FE_Test_Hover' $Index | Out-Null
}

function Wait-Tooltip($fe, [int]$TimeoutMs = 2500) {
    Wait-FeState $fe { $_.view.tooltipVisible } $TimeoutMs
}

# Substring highlighted by the hover (null when nothing).
# ConvertFrom-Json turns an empty JSON array into $null: these helpers always return real arrays.
function Get-Overlay($state) { , @($state.overlay | Where-Object { $null -ne $_ }) }
function Get-Stack($state) { , @($state.session.stack | Where-Object { $null -ne $_ }) }
function Get-HoverText($state) {
    if ($null -eq $state.view.hoverStart) { return $null }
    $state.view.displayText.Substring([int]$state.view.hoverStart, [int]$state.view.hoverLength)
}
function Get-GoldText($state) {
    if ($null -eq $state.view.goldStart) { return $null }
    $state.view.displayText.Substring([int]$state.view.goldStart, [int]$state.view.goldLength)
}

function Close-Explorer($fe) {
    try { Invoke-Fe $fe 'FE_Test_Close' | Out-Null } catch { }
    Start-Sleep -Milliseconds 200
}

# --- Excel geometry oracle (independent of the add-in) -----------------------------------------------------------------
# Pane.PointsToScreenPixelsX/Y take whole points; interpolate fractional ones.
function Convert-PointsToPx($pane, [string]$Axis, [double]$Points) {
    $fl = [math]::Floor($Points); $frac = $Points - $fl
    if ($Axis -eq 'X') { $p0 = $pane.PointsToScreenPixelsX([int]$fl) } else { $p0 = $pane.PointsToScreenPixelsY([int]$fl) }
    if ($frac -lt 1e-6) { return [int]$p0 }
    if ($Axis -eq 'X') { $p1 = $pane.PointsToScreenPixelsX([int]$fl + 1) } else { $p1 = $pane.PointsToScreenPixelsY([int]$fl + 1) }
    [int][math]::Round($p0 + $frac * ($p1 - $p0), [MidpointRounding]::AwayFromZero)
}

# Screen rect [l,t,r,b] of a range through a pane (default: the active pane).
function Get-RangeScreenRect($fe, $Range, $Pane = $null) {
    if (-not $Pane) { $Pane = $fe.Excel.ActiveWindow.ActivePane }
    $l = [double]$Range.Left; $t = [double]$Range.Top; $w = [double]$Range.Width; $h = [double]$Range.Height
    @((Convert-PointsToPx $Pane 'X' $l), (Convert-PointsToPx $Pane 'Y' $t), (Convert-PointsToPx $Pane 'X' ($l + $w)), (Convert-PointsToPx $Pane 'Y' ($t + $h)))
}

function Get-CellCenter($fe, $Range, $Pane = $null) {
    $r = Get-RangeScreenRect $fe $Range $Pane
    [pscustomobject]@{ X = [int](($r[0] + $r[2]) / 2); Y = [int](($r[1] + $r[3]) / 2) }
}

function Rect-Near($a, $b, [int]$Tol = 2) {
    for ($i = 0; $i -lt 4; $i++) { if ([math]::Abs($a[$i] - $b[$i]) -gt $Tol) { return $false } }
    $true
}
function Overlay-Rect($o) { @([int]$o.left, [int]$o.top, [int]$o.right, [int]$o.bottom) }

# --- screen sampling -------------------------------------------------------------------------------------------------
function Save-Crop($fe, [string]$Name, [int[]]$Rect) { Save-FeShot $fe $Name $Rect }

# Rectangle (x,y,w,h in screen pixels) covering chars [start, start+len) of the displayed formula.
function Get-SpanRect($state, [int]$Start, [int]$Length) {
    $pts = foreach ($i in $Start..($Start + $Length - 1)) { $p = $state.charPoints."$i"; if ($p) { , @([int]$p[0], [int]$p[1]) } }
    $xs = $pts | ForEach-Object { $_[0] }; $ys = $pts | ForEach-Object { $_[1] }
    $minX = ($xs | Measure-Object -Minimum).Minimum; $maxX = ($xs | Measure-Object -Maximum).Maximum
    $minY = ($ys | Measure-Object -Minimum).Minimum; $maxY = ($ys | Measure-Object -Maximum).Maximum
    # charPoints are char centres; widen by roughly half a char (Calibri 11pt at the window's scale).
    $halfW = [int](4 * [math]::Max(1.0, [double]$state.view.pixelsPerDip))
    [pscustomobject]@{ X = $minX - $halfW; Y = $minY - 8; W = ($maxX - $minX) + 2 * $halfW; H = ($maxY - $minY) + 16 }
}

# Dominant "ink" colour of a char span of the explorer text (darkest pixel), from a fresh screen capture.
function Get-SpanInk($state, [int]$Start, [int]$Length) {
    $r = Get-SpanRect $state $Start $Length
    $bmp = [FeWin2]::Capture($r.X, $r.Y, $r.W, $r.H)
    try { [FeWin2]::DarkestRgb($bmp, 0, 0, $r.W, $r.H) } finally { $bmp.Dispose() }
}
function Count-SpanColor($state, [int]$Start, [int]$Length, [int]$Rgb, [int]$Tol = 8) {
    $r = Get-SpanRect $state $Start $Length
    $bmp = [FeWin2]::Capture($r.X, $r.Y, $r.W, $r.H)
    try { [FeWin2]::CountColor($bmp, 0, 0, $r.W, $r.H, $Rgb, $Tol) } finally { $bmp.Dispose() }
}
function Hex([int]$v) { '#{0:X6}' -f $v }
function Color-Dist([int]$a, [int]$b) {
    $d = 0
    foreach ($s in 16, 8, 0) { $d = [math]::Max($d, [math]::Abs((($a -shr $s) -band 255) - (($b -shr $s) -band 255))) }
    $d
}

# --- real settings / log (outside the MSIX virtualization) ----------------------------------------------------------------
function Backup-RealSettings([string]$OutDir) {
    $dst = Join-Path $OutDir 'settings.backup.json'
    if (Test-Path $dst) { Remove-Item $dst -Force }
    Invoke-Unpackaged "cmd.exe /c if exist `"%LOCALAPPDATA%\FormulaExplorer\settings.json`" copy /y `"%LOCALAPPDATA%\FormulaExplorer\settings.json`" `"$dst`""
}
function Clear-RealSettings { Invoke-Unpackaged "cmd.exe /c del /q `"%LOCALAPPDATA%\FormulaExplorer\settings.json`"" }
function Restore-RealSettings([string]$OutDir) {
    $src = Join-Path $OutDir 'settings.backup.json'
    if (Test-Path $src) { Invoke-Unpackaged "cmd.exe /c copy /y `"$src`" `"%LOCALAPPDATA%\FormulaExplorer\settings.json`"" }
    else { Clear-RealSettings }
}
function Get-RealSettingsText([string]$OutDir) {
    $dst = Join-Path $OutDir 'settings.current.json'
    if (Test-Path $dst) { Remove-Item $dst -Force }
    Invoke-Unpackaged "cmd.exe /c if exist `"%LOCALAPPDATA%\FormulaExplorer\settings.json`" copy /y `"%LOCALAPPDATA%\FormulaExplorer\settings.json`" `"$dst`""
    if (Test-Path $dst) { Get-Content $dst -Raw } else { $null }
}
function Get-RealLogLines([string]$OutDir) {
    $dst = Join-Path $OutDir 'real-log.txt'
    if (Test-Path $dst) { Remove-Item $dst -Force }
    Copy-FeRealLog $dst
    if (Test-Path $dst) { , @(Get-Content $dst -Encoding UTF8 | ForEach-Object { [string]$_ }) } else { , @() }
}
# The developer's real add-in log (diagnostics of real use) must survive a run: copy log.txt and log.txt.old to
# the folder real-log-backups next to OutDir (a time-stamped pair, outside OutDir so that -Clean keeps it) BEFORE
# Clear-RealLog empties them. Returns the backup folder (a file is only written when its log exists).
function Backup-RealLog([string]$OutDir) {
    $dir = Join-Path (Split-Path -Parent $OutDir) 'real-log-backups'
    New-Item -ItemType Directory -Force $dir | Out-Null
    $stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    foreach ($name in 'log.txt', 'log.txt.old') {
        $dst = Join-Path $dir "$stamp-$name"
        Invoke-Unpackaged "cmd.exe /c if exist `"%LOCALAPPDATA%\FormulaExplorer\$name`" copy /y `"%LOCALAPPDATA%\FormulaExplorer\$name`" `"$dst`""
    }
    $dir
}
function Clear-RealLog { Invoke-Unpackaged "cmd.exe /c del /q `"%LOCALAPPDATA%\FormulaExplorer\log.txt`" `"%LOCALAPPDATA%\FormulaExplorer\log.txt.old`"" }

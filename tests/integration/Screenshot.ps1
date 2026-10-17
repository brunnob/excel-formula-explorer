# Captures the full virtual screen (physical pixels) to a PNG. Usage: Screenshot.ps1 -Path out.png [-Rect x,y,w,h]
param([Parameter(Mandatory)] [string]$Path, [int[]]$Rect)
Add-Type -AssemblyName System.Drawing
Add-Type @"
using System; using System.Runtime.InteropServices;
public static class DpiUtil {
  [DllImport("user32.dll")] public static extern bool SetProcessDpiAwarenessContext(IntPtr v);
  [DllImport("user32.dll")] public static extern int GetSystemMetrics(int i);
}
"@
[DpiUtil]::SetProcessDpiAwarenessContext([IntPtr]-4) | Out-Null   # PER_MONITOR_AWARE_V2
if ($Rect) { $x,$y,$w,$h = $Rect } else {
  $x = [DpiUtil]::GetSystemMetrics(76); $y = [DpiUtil]::GetSystemMetrics(77)
  $w = [DpiUtil]::GetSystemMetrics(78); $h = [DpiUtil]::GetSystemMetrics(79) }
$bmp = New-Object System.Drawing.Bitmap $w, $h
$g = [System.Drawing.Graphics]::FromImage($bmp)
$g.CopyFromScreen($x, $y, 0, 0, $bmp.Size)
$bmp.Save($Path, [System.Drawing.Imaging.ImageFormat]::Png)
$g.Dispose(); $bmp.Dispose()
"$Path ${w}x${h}"

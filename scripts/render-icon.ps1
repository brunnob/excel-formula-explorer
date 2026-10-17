# Renders the Formula Explorer ribbon icon (an original blue globe) to PNG with WPF.
# Output: src/FormulaExplorer.AddIn/Ribbon/globe32.png (large button), globe16.png (small), globe64.png (high-DPI source).
# Run with Windows PowerShell in STA mode:  powershell -STA -File scripts\render-icon.ps1
[CmdletBinding()]
param(
    [string]$OutDir
)
$ErrorActionPreference = 'Stop'
if (-not $OutDir) { $OutDir = Join-Path $PSScriptRoot '..\src\FormulaExplorer.AddIn\Ribbon' }
Add-Type -AssemblyName PresentationCore, PresentationFramework, WindowsBase

function New-Color([string]$hex) { [System.Windows.Media.ColorConverter]::ConvertFromString($hex) }

function New-Pen([string]$hex, [double]$width, [double]$alpha = 1.0) {
    $c = New-Color $hex
    $c.A = [byte][math]::Round(255 * $alpha)
    $brush = New-Object System.Windows.Media.SolidColorBrush $c
    $brush.Freeze()
    $pen = New-Object System.Windows.Media.Pen $brush, $width
    $pen.StartLineCap = 'Round'; $pen.EndLineCap = 'Round'
    $pen.Freeze()
    $pen
}

function New-Arc([double]$cx, [double]$cy, [double]$rx, [double]$ry, [double]$fromDeg, [double]$toDeg) {
    # Elliptical arc as a StreamGeometry (angles measured clockwise from +X, screen coordinates).
    $a0 = $fromDeg * [math]::PI / 180; $a1 = $toDeg * [math]::PI / 180
    $g = New-Object System.Windows.Media.StreamGeometry
    $ctx = $g.Open()
    $p0 = New-Object System.Windows.Point ($cx + $rx * [math]::Cos($a0)), ($cy + $ry * [math]::Sin($a0))
    $p1 = New-Object System.Windows.Point ($cx + $rx * [math]::Cos($a1)), ($cy + $ry * [math]::Sin($a1))
    $ctx.BeginFigure($p0, $false, $false)
    $size = New-Object System.Windows.Size $rx, $ry
    $large = ([math]::Abs($toDeg - $fromDeg) -gt 180)
    $ctx.ArcTo($p1, $size, 0, $large, [System.Windows.Media.SweepDirection]::Clockwise, $true, $false)
    $ctx.Close()
    $g.Freeze()
    $g
}

function Draw-Globe([System.Windows.Media.DrawingContext]$dc, [bool]$detailed) {
    # Design space: 32 x 32 units. Sphere centre (16,16), radius 14.2.
    # The radius is a whole number of pixels at 32/16/64 px so the silhouette lands on the pixel grid.
    $cx = 16.0; $cy = 16.0; $r = 14.0
    $rimWidth = if ($detailed) { 1.0 } else { 2.0 }       # always about one device pixel
    $center = New-Object System.Windows.Point $cx, $cy

    # Soft ground shadow under the sphere (very subtle, keeps the alpha edge clean).
    $shadow = New-Object System.Windows.Media.RadialGradientBrush
    $shadow.GradientStops.Add((New-Object System.Windows.Media.GradientStop (New-Color '#40001A40'), 0.0))
    $shadow.GradientStops.Add((New-Object System.Windows.Media.GradientStop (New-Color '#00001A40'), 1.0))
    $shadow.Freeze()
    $dc.DrawEllipse($shadow, $null, (New-Object System.Windows.Point 16, 30.2), 9.5, 1.5)

    # Sphere body: radial gradient with the light coming from the upper left.
    $body = New-Object System.Windows.Media.RadialGradientBrush
    $body.GradientOrigin = New-Object System.Windows.Point 0.34, 0.27
    $body.Center = New-Object System.Windows.Point 0.45, 0.42
    $body.RadiusX = 0.68; $body.RadiusY = 0.68
    $body.GradientStops.Add((New-Object System.Windows.Media.GradientStop (New-Color '#B4E0FF'), 0.0))
    $body.GradientStops.Add((New-Object System.Windows.Media.GradientStop (New-Color '#4AA8F5'), 0.32))
    $body.GradientStops.Add((New-Object System.Windows.Media.GradientStop (New-Color '#1A70D0'), 0.68))
    $body.GradientStops.Add((New-Object System.Windows.Media.GradientStop (New-Color '#0B3E92'), 1.0))
    $body.Freeze()
    $dc.DrawEllipse($body, $null, $center, $r, $r)

    # Meridians and parallels, clipped to the sphere.
    $clip = New-Object System.Windows.Media.EllipseGeometry $center, ($r - 0.2), ($r - 0.2)
    $clip.Freeze()
    $dc.PushClip($clip)
    if ($detailed) {
        $grid = New-Pen '#FFFFFF' 1.0 0.72
        $dc.DrawEllipse($null, $grid, $center, ($r * 0.46), $r)           # inner meridian
        $dc.DrawEllipse($null, $grid, $center, ($r * 0.84), $r)           # outer meridian
        # Central meridian on a pixel centre (half a unit off) so it stays one crisp pixel wide.
        $dc.DrawLine($grid, (New-Object System.Windows.Point ($cx - 0.5), ($cy - $r)), (New-Object System.Windows.Point ($cx - 0.5), ($cy + $r)))
        # Parallels seen from slightly above: front half of each circle, bending downwards.
        foreach ($lat in 0, 38, -38) {
            $y = $cy - $r * [math]::Sin($lat * [math]::PI / 180)
            $pr = $r * [math]::Cos($lat * [math]::PI / 180)
            $dc.DrawGeometry($null, $grid, (New-Arc $cx $y $pr ($pr * 0.30) 0 180))
        }
    } else {
        # 16 px variant: one-pixel lines (2 units), a single meridian ellipse, the equator and the central meridian.
        $grid = New-Pen '#FFFFFF' 2.0 0.78
        $dc.DrawEllipse($null, $grid, $center, ($r * 0.52), $r)
        $dc.DrawGeometry($null, $grid, (New-Arc $cx ($cy + 1) $r ($r * 0.24) 0 180))
        $dc.DrawLine($grid, (New-Object System.Windows.Point ($cx - 1), ($cy - $r)), (New-Object System.Windows.Point ($cx - 1), ($cy + $r)))
    }
    $dc.Pop()

    # Dark rim.
    $dc.DrawEllipse($null, (New-Pen '#07306F' $rimWidth 0.95), $center, ($r - $rimWidth / 2), ($r - $rimWidth / 2))

    # Glossy highlight on the upper left.
    $gloss = New-Object System.Windows.Media.LinearGradientBrush
    $gloss.StartPoint = New-Object System.Windows.Point 0.2, 0
    $gloss.EndPoint = New-Object System.Windows.Point 0.6, 1
    $gloss.GradientStops.Add((New-Object System.Windows.Media.GradientStop (New-Color '#B0FFFFFF'), 0.0))
    $gloss.GradientStops.Add((New-Object System.Windows.Media.GradientStop (New-Color '#00FFFFFF'), 1.0))
    $gloss.Freeze()
    $dc.PushTransform((New-Object System.Windows.Media.RotateTransform -35, 11.2, 8.6))
    $dc.DrawEllipse($gloss, $null, (New-Object System.Windows.Point 11.2, 8.6), 6.4, 3.4)
    $dc.Pop()
}

function Save-Globe([int]$px, [bool]$detailed, [string]$path) {
    $dv = New-Object System.Windows.Media.DrawingVisual
    [System.Windows.Media.RenderOptions]::SetEdgeMode($dv, 'Unspecified')
    $dc = $dv.RenderOpen()
    $dc.PushTransform((New-Object System.Windows.Media.ScaleTransform ($px / 32.0), ($px / 32.0)))
    Draw-Globe $dc $detailed
    $dc.Pop()
    $dc.Close()
    $rtb = New-Object System.Windows.Media.Imaging.RenderTargetBitmap $px, $px, 96, 96, ([System.Windows.Media.PixelFormats]::Pbgra32)
    $rtb.Render($dv)
    $enc = New-Object System.Windows.Media.Imaging.PngBitmapEncoder
    $enc.Frames.Add([System.Windows.Media.Imaging.BitmapFrame]::Create($rtb))
    $full = [System.IO.Path]::GetFullPath($path)
    $fs = [System.IO.File]::Create($full)
    try { $enc.Save($fs) } finally { $fs.Dispose() }
    "$full ($px x $px)"
}

New-Item -ItemType Directory -Force $OutDir | Out-Null
Save-Globe 32 $true (Join-Path $OutDir 'globe32.png')
Save-Globe 16 $false (Join-Path $OutDir 'globe16.png')
Save-Globe 64 $true (Join-Path $OutDir 'globe64.png')

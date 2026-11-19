function Initialize-StaWindowCapture {
    if ('StaCli.WindowCapture' -as [type]) { return }
    Add-Type -AssemblyName System.Drawing.Common
    Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
using System.Text;

namespace StaCli
{
    public static class WindowCapture
    {
        [StructLayout(LayoutKind.Sequential)]
        public struct Rect { public int Left, Top, Right, Bottom; }

        [DllImport("user32.dll")]
        public static extern bool GetWindowRect(IntPtr window, out Rect bounds);
        [DllImport("user32.dll")]
        public static extern bool IsWindowVisible(IntPtr window);
        [DllImport("user32.dll")]
        public static extern bool IsIconic(IntPtr window);
        [DllImport("user32.dll")]
        public static extern uint GetWindowThreadProcessId(IntPtr window, out uint processId);
        [DllImport("user32.dll")]
        public static extern bool PrintWindow(IntPtr window, IntPtr deviceContext, uint flags);

        [DllImport("user32.dll")]
        private static extern IntPtr OpenInputDesktop(uint flags, bool inherit, uint access);
        [DllImport("user32.dll")]
        private static extern bool CloseDesktop(IntPtr desktop);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern bool GetUserObjectInformation(
            IntPtr handle, int index, StringBuilder result, uint bytes, out uint required);

        public static bool HasInteractiveInputDesktop()
        {
            if (!Environment.UserInteractive) return false;
            IntPtr input = OpenInputDesktop(0, false, 1);
            if (input == IntPtr.Zero) return false;
            try
            {
                var name = new StringBuilder(256);
                return GetUserObjectInformation(input, 2, name, 512, out _) &&
                    string.Equals(name.ToString(), "Default", StringComparison.OrdinalIgnoreCase);
            }
            finally { CloseDesktop(input); }
        }
    }
}
"@
}

function Get-StaWindowBounds([IntPtr]$Handle, [int]$ExpectedPid) {
    $owner = [uint32]0
    $thread = [StaCli.WindowCapture]::GetWindowThreadProcessId($Handle, [ref]$owner)
    $bounds = [StaCli.WindowCapture+Rect]::new()
    if ($thread -eq 0 -or $owner -ne $ExpectedPid -or
        -not [StaCli.WindowCapture]::IsWindowVisible($Handle) -or
        [StaCli.WindowCapture]::IsIconic($Handle) -or
        -not [StaCli.WindowCapture]::GetWindowRect($Handle, [ref]$bounds)) {
        return
    }
    $bounds
}

function Test-StaActiveDocument($App, [string]$ActiveProperty, [IntPtr]$ExpectedWindow) {
    $active = $App.$ActiveProperty
    if ($null -eq $active) { return $false }
    $window = $null
    try {
        $window = $active.Window
        $null -ne $window -and [IntPtr]$window.Handle -eq $ExpectedWindow
    } finally {
        Release-StaCom $window
        Release-StaCom $active
    }
}

function Get-StaFrameSignature($Bitmap, $Region) {
    # Sample the document area, leaving out most of its window chrome.
    $colors = [Collections.Generic.HashSet[int]]::new()
    $samples = [Collections.Generic.List[int]]::new()
    $left = $Region.Left + 12
    $top = $Region.Top + 32
    $width = $Region.Width - 24
    $height = $Region.Height - 44
    $base = $Bitmap.GetPixel($left, $top).ToArgb()
    $different = 0
    for ($y = 0; $y -lt 54; $y++) {
        for ($x = 0; $x -lt 72; $x++) {
            $color = $Bitmap.GetPixel($left + [int]($x * $width / 72), $top + [int]($y * $height / 54)).ToArgb()
            [void]$colors.Add($color)
            $samples.Add($color)
            if ($color -ne $base) { $different++ }
        }
    }
    if ($colors.Count -lt 3 -or $different -lt 8) { return }
    $samples -join ','
}

function Save-StaWindowCapture($Session, $Options, [string]$Stage) {
    Initialize-StaWindowCapture
    Invoke-StaCall 'Capture' 'Interactive desktop' {
        if (-not [StaCli.WindowCapture]::HasInteractiveInputDesktop()) {
            throw 'Screenshot requires an unlocked interactive Windows desktop.'
        }
    }
    $collection, $activeProperty = switch ([IO.Path]::GetExtension($Options.InputPath)) {
        '.sta' { 'Spreadsheets'; 'ActiveSpreadsheet' }
        '.stg' { 'Graphs'; 'ActiveGraph' }
        '.stw' { 'Workbooks'; 'ActiveWorkbook' }
    }
    $doc = Open-StaDocument $Session $collection $Options.InputPath
    $window = $null
    try {
        Invoke-StaCall 'Capture' 'Document.Activate' {
            $Session.App.Visible = $true
            $doc.Visible = $true
            $doc.Activate()
        }
        $window = $doc.Window
        $mainHandle = [IntPtr]$Session.App.Handle
        $docHandle = [IntPtr]$window.Handle
        $previous = $null
        $saved = $false
        $clock = [Diagnostics.Stopwatch]::StartNew()
        do {
            $main = Get-StaWindowBounds $mainHandle $Session.Pid
            $target = Get-StaWindowBounds $docHandle $Session.Pid
            if ($main -and $target) {
                $width = $main.Right - $main.Left
                $height = $main.Bottom - $main.Top
                $bounds = [Drawing.Rectangle]::new(0, 0, $width, $height)
                $region = [Drawing.Rectangle]::new(
                    $target.Left - $main.Left, $target.Top - $main.Top,
                    $target.Right - $target.Left, $target.Bottom - $target.Top)
                $region = [Drawing.Rectangle]::Intersect($bounds, $region)
                if ($region.Width -ge 160 -and $region.Height -ge 100 -and
                    (Test-StaActiveDocument $Session.App $activeProperty $docHandle)) {
                    $bitmap = [Drawing.Bitmap]::new($width, $height)
                    try {
                        $copied = Invoke-StaCall 'Capture' 'PrintWindow' {
                            $graphics = [Drawing.Graphics]::FromImage($bitmap)
                            try {
                                $dc = $graphics.GetHdc()
                                try { [StaCli.WindowCapture]::PrintWindow($mainHandle, $dc, 0) }
                                finally { $graphics.ReleaseHdc($dc) }
                            } finally { $graphics.Dispose() }
                        }
                        if ($copied) {
                            $signature = Get-StaFrameSignature $bitmap $region
                            if ($signature -and $signature -ceq $previous -and
                                (Test-StaActiveDocument $Session.App $activeProperty $docHandle)) {
                                Invoke-StaCall 'Capture' 'Save PNG' {
                                    if ($Options['CropToDocument']) {
                                        $crop = $bitmap.Clone($region, $bitmap.PixelFormat)
                                        try { $crop.Save($Stage, [Drawing.Imaging.ImageFormat]::Png) }
                                        finally { $crop.Dispose() }
                                    } else {
                                        $bitmap.Save($Stage, [Drawing.Imaging.ImageFormat]::Png)
                                    }
                                }
                                $saved = $true
                            }
                            $previous = $signature
                        } else { $previous = $null }
                    } finally { $bitmap.Dispose() }
                } else { $previous = $null }
            } else { $previous = $null }
            if (-not $saved) { Start-Sleep -Milliseconds 160 }
        } until ($saved -or $clock.Elapsed.TotalSeconds -ge $Options.CaptureTimeoutSeconds)
        if (-not $saved) {
            $staState.Failure = New-StaError 'Capture' 'Wait for stable frame' 'The active document did not produce a stable, nonblank frame.'
            throw $staState.Failure.message
        }
    } finally {
        Release-StaCom $window
        Close-StaDocument $doc -LeaveForQuit
    }
}

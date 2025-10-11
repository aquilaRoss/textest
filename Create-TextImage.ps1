param (
    [string]$Text,
    [int]$Width,
    [int]$Height,
    [string]$Filename
)

Add-Type -AssemblyName System.Drawing

try {
    # --- Resolve target path safely ---
    # If user provided an absolute path, use it. If not, resolve relative to the script's folder.
    $isRooted = [System.IO.Path]::IsPathRooted($Filename)
    if (-not $isRooted) {
        # Script folder (works even when executed from other folders)
        $scriptPath = $MyInvocation.MyCommand.Path
        if ([string]::IsNullOrEmpty($scriptPath)) {
            # fallback if running interactively (shouldn't happen when running script file)
            $scriptDir = Get-Location
        } else {
            $scriptDir = Split-Path -Parent $scriptPath
        }
        $fullPath = [System.IO.Path]::GetFullPath((Join-Path $scriptDir $Filename))
    } else {
        $fullPath = [System.IO.Path]::GetFullPath($Filename)
    }

    Write-Host "Saving image to: $fullPath"

    # Ensure target folder exists
    $folder = Split-Path $fullPath
    if (-not (Test-Path $folder)) {
        New-Item -ItemType Directory -Force -Path $folder | Out-Null
    }

    # --- Create drawing surface ---
    $bmp = New-Object System.Drawing.Bitmap($Width, $Height)
    $graphics = [System.Drawing.Graphics]::FromImage($bmp)
    $graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $graphics.Clear([System.Drawing.Color]::White)

    # Draw black border
    $pen = New-Object System.Drawing.Pen([System.Drawing.Color]::Black)
    $graphics.DrawRectangle($pen, 0, 0, $Width - 1, $Height - 1)

    # Font and brush (explicit constructor)
    $font = New-Object System.Drawing.Font("Arial", 20, [System.Drawing.FontStyle]::Bold)
    $brush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::Black)

    # Measure and center text
    $textSize = $graphics.MeasureString($Text, $font)
    $x = [Math]::Max(0, ($Width - $textSize.Width) / 2)
    $y = [Math]::Max(0, ($Height - $textSize.Height) / 2)

    # Draw text
    $graphics.DrawString($Text, $font, $brush, $x, $y)

    # Dispose drawing objects BEFORE saving
    $brush.Dispose()
    $pen.Dispose()
    $font.Dispose()
    $graphics.Dispose()

    # If a file exists, remove it first (avoid locked/readonly edge cases)
    if (Test-Path $fullPath) {
        Remove-Item $fullPath -Force -ErrorAction SilentlyContinue
    }

    # --- Save using a FileStream (reliable with Dropbox/OneDrive/locked folders) ---
    $fs = $null
    try {
        $fs = [System.IO.File]::Open($fullPath,
            [System.IO.FileMode]::Create,
            [System.IO.FileAccess]::Write,
            [System.IO.FileShare]::None)

        $bmp.Save($fs, [System.Drawing.Imaging.ImageFormat]::Png)
    }
    finally {
        if ($fs -ne $null) {
            $fs.Close()
            $fs.Dispose()
        }
    }

    # Dispose bitmap last
    $bmp.Dispose()

    Write-Host "? Image saved successfully to: $fullPath"
}
catch {
    Write-Host "? Error: $($_.Exception.Message)"
    throw
}

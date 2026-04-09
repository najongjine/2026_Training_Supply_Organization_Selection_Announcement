param(
    [string]$InputDir = ".",
    [string]$OutputPath = "ocr_result.txt"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

Add-Type -AssemblyName System.Runtime.WindowsRuntime
Add-Type -AssemblyName System.Drawing
$null = [Windows.Media.Ocr.OcrEngine, Windows.Foundation, ContentType = WindowsRuntime]
$null = [Windows.Graphics.Imaging.SoftwareBitmap, Windows.Foundation, ContentType = WindowsRuntime]
$null = [Windows.Storage.StorageFile, Windows.Storage, ContentType = WindowsRuntime]

function Await-AsyncResult {
    param(
        [Parameter(Mandatory = $true)]
        $Operation,
        [Parameter(Mandatory = $false)]
        [Type]$ResultType
    )

    if ($ResultType) {
        $asTaskMethod = [System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object {
            $_.Name -eq "AsTask" -and
            $_.IsGenericMethodDefinition -and
            $_.GetParameters().Count -eq 1 -and
            $_.GetGenericArguments().Count -eq 1
        } | Select-Object -First 1

        $task = $asTaskMethod.MakeGenericMethod($ResultType).Invoke($null, @($Operation))
    }
    else {
        $task = [System.WindowsRuntimeSystemExtensions]::AsTask([Windows.Foundation.IAsyncAction]$Operation)
    }

    $task.Wait()
    if ($task.Exception) {
        throw $task.Exception.InnerException
    }
    if ($task.GetType().IsGenericType) {
        return $task.Result
    }
    return $null
}

function Get-Recognizer {
    $null = [Windows.Globalization.Language, Windows.Globalization, ContentType = WindowsRuntime]
    $lang = [Windows.Globalization.Language]::new("ko")
    $engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromLanguage($lang)
    if (-not $engine) {
        throw "Korean OCR recognizer is not available on this system."
    }
    return $engine
}

function Save-RotatedImage {
    param(
        [Parameter(Mandatory = $true)]
        [string]$SourcePath,
        [Parameter(Mandatory = $true)]
        [int]$Angle,
        [Parameter(Mandatory = $true)]
        [string]$TargetPath
    )

    $bitmap = [System.Drawing.Bitmap]::new($SourcePath)
    try {
        switch ($Angle) {
            90 { $bitmap.RotateFlip([System.Drawing.RotateFlipType]::Rotate90FlipNone) }
            180 { $bitmap.RotateFlip([System.Drawing.RotateFlipType]::Rotate180FlipNone) }
            270 { $bitmap.RotateFlip([System.Drawing.RotateFlipType]::Rotate270FlipNone) }
            default { }
        }
        $bitmap.Save($TargetPath, [System.Drawing.Imaging.ImageFormat]::Png)
    }
    finally {
        $bitmap.Dispose()
    }
}

function Get-OcrResult {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ImagePath,
        [Parameter(Mandatory = $true)]
        $Engine
    )

    $storageFile = Await-AsyncResult ([Windows.Storage.StorageFile]::GetFileFromPathAsync($ImagePath)) ([Windows.Storage.StorageFile])
    $stream = Await-AsyncResult ($storageFile.OpenAsync([Windows.Storage.FileAccessMode]::Read)) ([Windows.Storage.Streams.IRandomAccessStream])
    try {
        $decoder = Await-AsyncResult ([Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($stream)) ([Windows.Graphics.Imaging.BitmapDecoder])
        $softwareBitmap = Await-AsyncResult ($decoder.GetSoftwareBitmapAsync()) ([Windows.Graphics.Imaging.SoftwareBitmap])
        if ($softwareBitmap.BitmapPixelFormat -ne [Windows.Graphics.Imaging.BitmapPixelFormat]::Bgra8) {
            $converted = [Windows.Graphics.Imaging.SoftwareBitmap]::Convert(
                $softwareBitmap,
                [Windows.Graphics.Imaging.BitmapPixelFormat]::Bgra8
            )
            $softwareBitmap.Dispose()
            $softwareBitmap = $converted
        }

        try {
            return Await-AsyncResult ($Engine.RecognizeAsync($softwareBitmap)) ([Windows.Media.Ocr.OcrResult])
        }
        finally {
            $softwareBitmap.Dispose()
        }
    }
    finally {
        $stream.Dispose()
    }
}

function Get-TextScore {
    param(
        [string]$Text
    )

    if ([string]::IsNullOrWhiteSpace($Text)) {
        return 0
    }

    $hangul = ([regex]::Matches($Text, "[\uAC00-\uD7A3]")).Count
    $latin = ([regex]::Matches($Text, "[A-Za-z0-9]")).Count
    $lines = (($Text -split "`r?`n") | Where-Object { $_.Trim().Length -gt 0 }).Count
    return ($hangul * 5) + $latin + ($lines * 3) + $Text.Trim().Length
}

function Format-OcrLines {
    param(
        [Parameter(Mandatory = $true)]
        $OcrResult
    )

    $formattedLines = New-Object System.Collections.Generic.List[string]
    $previousBottom = $null

    foreach ($line in $OcrResult.Lines) {
        $words = @($line.Words)
        if ($words.Count -eq 0) {
            continue
        }

        $sortedWords = $words | Sort-Object { $_.BoundingRect.X }
        $tops = @($sortedWords | ForEach-Object { $_.BoundingRect.Y })
        $bottoms = @($sortedWords | ForEach-Object { $_.BoundingRect.Y + $_.BoundingRect.Height })
        $lefts = @($sortedWords | ForEach-Object { $_.BoundingRect.X })
        $rights = @($sortedWords | ForEach-Object { $_.BoundingRect.X + $_.BoundingRect.Width })
        $lineTop = ($tops | Measure-Object -Minimum).Minimum
        $lineBottom = ($bottoms | Measure-Object -Maximum).Maximum
        $lineLeft = ($lefts | Measure-Object -Minimum).Minimum
        $lineRight = ($rights | Measure-Object -Maximum).Maximum
        $lineHeight = [Math]::Max($lineBottom - $lineTop, 1)

        if ($previousBottom -ne $null) {
            $verticalGap = $lineTop - $previousBottom
            if ($verticalGap -gt ($lineHeight * 1.2)) {
                $formattedLines.Add("")
            }
        }

        $avgCharWidth = 12.0
        if ($line.Text.Length -gt 0) {
            $avgCharWidth = [Math]::Max(($lineRight - $lineLeft) / [Math]::Max($line.Text.Length, 1), 7.0)
        }

        $builder = New-Object System.Text.StringBuilder
        $lastRight = $null
        foreach ($word in $sortedWords) {
            if ($lastRight -ne $null) {
                $gap = $word.BoundingRect.X - $lastRight
                if ($gap -gt ($avgCharWidth * 2.2)) {
                    $tabCount = [Math]::Max([Math]::Floor($gap / ($avgCharWidth * 6)), 1)
                    [void]$builder.Append(("`t" * $tabCount))
                }
                elseif ($builder.Length -gt 0) {
                    [void]$builder.Append(" ")
                }
            }
            [void]$builder.Append($word.Text)
            $lastRight = $word.BoundingRect.X + $word.BoundingRect.Width
        }

        $formattedLines.Add($builder.ToString().TrimEnd())
        $previousBottom = $lineBottom
    }

    return ($formattedLines -join [Environment]::NewLine).Trim()
}

function Get-BestOrientationResult {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ImagePath,
        [Parameter(Mandatory = $true)]
        $Engine,
        [Parameter(Mandatory = $true)]
        [string]$TempDir
    )

    $best = $null
    foreach ($angle in 0, 90, 180, 270) {
        $tempPath = Join-Path $TempDir ("{0}_{1}.png" -f [IO.Path]::GetFileNameWithoutExtension($ImagePath), $angle)
        Save-RotatedImage -SourcePath $ImagePath -Angle $angle -TargetPath $tempPath

        $ocr = Get-OcrResult -ImagePath $tempPath -Engine $Engine
        $formattedText = Format-OcrLines -OcrResult $ocr
        $score = Get-TextScore -Text $formattedText

        $candidate = [PSCustomObject]@{
            Angle = $angle
            Score = $score
            Text = $formattedText
        }

        if (-not $best -or $candidate.Score -gt $best.Score) {
            $best = $candidate
        }
    }

    return $best
}

$engine = Get-Recognizer
$resolvedInputDir = (Resolve-Path $InputDir).Path
$resolvedOutputPath = [IO.Path]::GetFullPath((Join-Path $resolvedInputDir $OutputPath))
$tempDir = Join-Path $resolvedInputDir ".ocr_temp"

if (-not (Test-Path $tempDir)) {
    New-Item -ItemType Directory -Path $tempDir | Out-Null
}

try {
    $images = Get-ChildItem -Path $resolvedInputDir -File | Where-Object {
        $_.Extension.ToLowerInvariant() -in @(".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff", ".webp")
    } | Sort-Object Name

    if (-not $images) {
        throw "No images were found in $resolvedInputDir."
    }

    $outputLines = New-Object System.Collections.Generic.List[string]
    $outputLines.Add("OCR result generated on $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")")
    $outputLines.Add("Source directory: $resolvedInputDir")
    $outputLines.Add("")

    foreach ($image in $images) {
        Write-Host "Processing $($image.Name)..."
        $best = Get-BestOrientationResult -ImagePath $image.FullName -Engine $engine -TempDir $tempDir

        $outputLines.Add(("=" * 80))
        $outputLines.Add("FILE: $($image.Name)")
        $outputLines.Add("BEST_ROTATION: $($best.Angle) degrees clockwise")
        $outputLines.Add(("=" * 80))
        if ([string]::IsNullOrWhiteSpace($best.Text)) {
            $outputLines.Add("[No OCR text detected]")
        }
        else {
            $outputLines.Add($best.Text)
        }
        $outputLines.Add("")
    }

    [IO.File]::WriteAllText($resolvedOutputPath, ($outputLines -join [Environment]::NewLine), [Text.Encoding]::UTF8)
    Write-Host "Saved OCR output to $resolvedOutputPath"
}
finally {
    if (Test-Path $tempDir) {
        Remove-Item -Path $tempDir -Recurse -Force
    }
}

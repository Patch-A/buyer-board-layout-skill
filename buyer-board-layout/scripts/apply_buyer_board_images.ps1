$ErrorActionPreference = "Stop"

param(
    [Parameter(Mandatory = $true)][string]$InputPpt,
    [Parameter(Mandatory = $true)][string]$ConfigJson,
    [Parameter(Mandatory = $true)][string]$OutputPpt,
    [Parameter(Mandatory = $true)][string]$PreviewDir
)

function Get-ShapeAtPosition {
    param(
        $Slide,
        [double]$Left,
        [double]$Top,
        [int[]]$AllowedTypes = @(13, 28)
    )

    foreach ($shape in $Slide.Shapes) {
        if ($AllowedTypes -notcontains $shape.Type) {
            continue
        }
        if ([Math]::Abs($shape.Left - $Left) -lt 2 -and [Math]::Abs($shape.Top - $Top) -lt 2) {
            return $shape
        }
    }

    throw "Target shape not found at expected position."
}

function Fit-PictureIntoBox {
    param(
        $Shape,
        [double]$BoxLeft,
        [double]$BoxTop,
        [double]$BoxWidth,
        [double]$BoxHeight
    )

    $ratio = [Math]::Min($BoxWidth / $Shape.Width, $BoxHeight / $Shape.Height)
    $Shape.LockAspectRatio = -1
    $Shape.Width = $Shape.Width * $ratio
    $Shape.Height = $Shape.Height * $ratio
    $Shape.Left = $BoxLeft + (($BoxWidth - $Shape.Width) / 2)
    $Shape.Top = $BoxTop + (($BoxHeight - $Shape.Height) / 2)
}

function Fill-PictureIntoBox {
    param(
        $Shape,
        [double]$BoxLeft,
        [double]$BoxTop,
        [double]$BoxWidth,
        [double]$BoxHeight
    )

    $ratio = [Math]::Max($BoxWidth / $Shape.Width, $BoxHeight / $Shape.Height)
    $Shape.LockAspectRatio = -1
    $Shape.Width = $Shape.Width * $ratio
    $Shape.Height = $Shape.Height * $ratio
    $Shape.Left = $BoxLeft + (($BoxWidth - $Shape.Width) / 2)
    $Shape.Top = $BoxTop + (($BoxHeight - $Shape.Height) / 2)
}

function Replace-PictureShape {
    param(
        $Slide,
        $Target,
        [string]$ImagePath,
        [bool]$FillBox = $false
    )

    if (-not (Test-Path -LiteralPath $ImagePath)) {
        throw "Missing image asset: $ImagePath"
    }

    $left = $Target.Left
    $top = $Target.Top
    $width = $Target.Width
    $height = $Target.Height
    $zOrder = $Target.ZOrderPosition

    $Target.Delete()
    $newShape = $Slide.Shapes.AddPicture($ImagePath, $false, $true, $left, $top, -1, -1)
    if ($FillBox) {
        Fill-PictureIntoBox -Shape $newShape -BoxLeft $left -BoxTop $top -BoxWidth $width -BoxHeight $height
    }
    else {
        Fit-PictureIntoBox -Shape $newShape -BoxLeft $left -BoxTop $top -BoxWidth $width -BoxHeight $height
    }

    while ($newShape.ZOrderPosition -lt $zOrder) {
        $newShape.ZOrder(0)
    }
}

function Add-LogoPicture {
    param(
        $Slide,
        [string]$ImagePath,
        [double]$Left,
        [double]$Top,
        [double]$Width,
        [double]$Height
    )

    if (-not (Test-Path -LiteralPath $ImagePath)) {
        throw "Missing image asset: $ImagePath"
    }

    $shape = $Slide.Shapes.AddPicture($ImagePath, $false, $true, $Left, $Top, -1, -1)
    Fit-PictureIntoBox -Shape $shape -BoxLeft $Left -BoxTop $Top -BoxWidth $Width -BoxHeight $Height
}

function Remove-HeaderArtifacts {
    param(
        $Slide
    )

    $toDelete = @()
    foreach ($shape in $Slide.Shapes) {
        if (
            $shape.Left -ge 55 -and
            $shape.Left -le 180 -and
            $shape.Top -ge 100 -and
            $shape.Top -le 155 -and
            ($shape.Type -eq 28 -or $shape.Type -eq 13)
        ) {
            $toDelete += $shape
        }
    }

    foreach ($shape in $toDelete) {
        $shape.Delete()
    }
}

if (-not (Test-Path -LiteralPath $InputPpt)) {
    throw "Input PPT not found."
}

if (-not (Test-Path -LiteralPath $ConfigJson)) {
    throw "Config JSON not found."
}

$slideAssets = Get-Content -Raw -LiteralPath $ConfigJson | ConvertFrom-Json
New-Item -ItemType Directory -Force -Path $PreviewDir | Out-Null

$powerPoint = New-Object -ComObject PowerPoint.Application
$powerPoint.Visible = -1

try {
    $presentation = $powerPoint.Presentations.Open($InputPpt, $false, $false, $false)

    foreach ($item in $slideAssets) {
        $slide = $presentation.Slides.Item([int]$item.SlideIndex)
        $siteTarget = Get-ShapeAtPosition -Slide $slide -Left ([double]$item.SiteLeft) -Top ([double]$item.SiteTop) -AllowedTypes @(13)
        Replace-PictureShape -Slide $slide -Target $siteTarget -ImagePath $item.SitePath -FillBox $true

        if ($null -ne $item.AddLogoLeft) {
            Remove-HeaderArtifacts -Slide $slide
            Add-LogoPicture `
                -Slide $slide `
                -ImagePath $item.LogoPath `
                -Left ([double]$item.AddLogoLeft) `
                -Top ([double]$item.AddLogoTop) `
                -Width ([double]$item.AddLogoWidth) `
                -Height ([double]$item.AddLogoHeight)
        }
        else {
            $logoTarget = Get-ShapeAtPosition -Slide $slide -Left ([double]$item.LogoLeft) -Top ([double]$item.LogoTop) -AllowedTypes @(13)
            Replace-PictureShape -Slide $slide -Target $logoTarget -ImagePath $item.LogoPath
        }
    }

    $presentation.SaveAs($OutputPpt)
    $presentation.Export($PreviewDir, "PNG")
    $presentation.Close()
}
finally {
    $powerPoint.Quit()
}

Write-Output $OutputPpt
Write-Output $PreviewDir

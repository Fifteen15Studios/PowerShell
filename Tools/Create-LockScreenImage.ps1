###############################################################################
#
# This script takes a background image, and adds an overlay on the image. It
#   can overlay an image (like a QR code) in the top left of the screen, the
#   computer name and an additional message in the top right of the screen,
#   and a "message of the day" (MOTD) at the bottom center of the screen.
#
# Originally written by reddit user Shamalamadindong
#
###############################################################################

# Define paths
$BlankSrcPath = "PATH\OrigWallpaper.jpg" # Your blank wallpaper
$OverlaySrcPath = "PATH\OverlayImage.jpg" # Image that has to be overlaid on it.
$MOTDSrcPath = "PATH\MOTD.txt" # MOTD message
$FinalDestPath = "PATH\lockscreenbg.bmp" # Final result file, save to somewhere and point your GPO lockscreen setting to that location

# Message to put below computer name
$SubMessage = "Help Desk:`r`n<Email>`r`n<Phone>`r`nAfter Hours: <After Hours Phone>"

function Write-Text {
    [CmdletBinding()]
    param (
        [Parameter()][String] $Text=" ", # Defaults to a space to prevent errors
        [Parameter()][String] $FontType="Segoe UI",
        [Parameter()][String] $Color="#ffffff",
        [Parameter()][string] $TextSize="60",
        [Parameter(Mandatory=$true)][ValidateSet("Near", "Center", "Far")][String] $HorAlign,
        [Parameter(Mandatory=$true)][ValidateSet("Near", "Center", "Far")][String] $VerAlign
    )
 
    # Rectangle
    $rect = [System.Drawing.RectangleF]::FromLTRB(0, 0, $BlanksrcImg.Width, $BlanksrcImg.Height)
 
    # Style text
    $Brush = New-Object Drawing.SolidBrush([System.Drawing.ColorTranslator]::FromHtml($Color))
 
    # Define font
    $FontFN = new-object System.Drawing.Font($FontType, $TextSize)
 
    # Set location
    $format = [System.Drawing.StringFormat]::GenericDefault
    $format.Alignment = [System.Drawing.StringAlignment]::$HorAlign
    $format.LineAlignment = [System.Drawing.StringAlignment]::$VerAlign
 
    # Draw text
    $Image.DrawString($Text, $FontFN, $Brush, $Rect, $format)
}
 
Function Add-Overlay {
    # Get overlay source image
    $OverlaysrcImg = [System.Drawing.Image]::FromFile($OverlaySrcPath)
    
    # create graphic object from source image
    $graphics = [system.drawing.Graphics]::FromImage($BlanksrcImg)

    # Make new image 1/10th the size of the background
    $Height = $BlanksrcImg.height / 5
    $Width = $BlanksrcImg.height / 5

    # Use these if you want the image in bottom right instead of top left
    $X = $BlanksrcImg.width - $Width - 10
    $Y = $BlanksrcImg.height - $Height - 10

    # draw overlay image onto blank image. Use "$X, $Y" instead of "10, 10" for bottom right placement
    $graphics.DrawImage($OverlaysrcImg, 10, 10, $Width, $Height)
 
    # Clean up
    $OverlaysrcImg.Dispose()
    $graphics.Dispose()
}
 
Function Write-Lockscreenbg {
    
    # If image path insn't valid, notify and exit
    if(!(Test-Path $BlankSrcPath)) {
        Write-Host "Invalid image path!" -ForegroundColor Red
        break script
    }

    # Get blank source image
    $BlanksrcImg = [System.Drawing.Image]::FromFile($BlankSrcPath)
 
    # Create a bitmap as $destPath
    $outputIImg = new-object System.Drawing.Bitmap([int]($BlanksrcImg.width)),([int]($BlanksrcImg.height))
 
    # Intialize Graphics
    $Image = [System.Drawing.Graphics]::FromImage($outputIImg)
    $Image.SmoothingMode = "AntiAlias"
 
    # Add overlay - if path is valid
    if(Test-Path $OverlaySrcPath)
    {Add-Overlay}
 
    $Rectangle = New-Object Drawing.Rectangle 0, 0, $BlanksrcImg.Width, $BlanksrcImg.Height
    $Image.DrawImage($BlanksrcImg, $Rectangle, 0, 0, $BlanksrcImg.Width, $BlanksrcImg.Height, ([Drawing.GraphicsUnit]::Pixel))
 
    # Write-Text - Feel free to change parameters such as colors and sizes to fit your environment
    Write-Text -Text "$env:COMPUTERNAME`r`n$SubMessage" -TextSize "20" -FontType "Segoe UI" -Color "#ffffff" -HorAlign "Far" -VerAlign "Near"
    # Write Message of the day - If path is valid
    if(Test-Path $MOTDSrcPath)
    {
        Write-Text -Text (Get-Content $MOTDSrcPath) -TextSize "30" -FontType "Segoe UI" -Color "#ffffff" -HorAlign "Center" -VerAlign "Far"
    }
 
    # save altered image
    $outputIImg.Save($FinalDestPath)
    # clean up
    $outputIImg.Dispose()
    $Image.Dispose()
    $BlanksrcImg.Dispose()
}
 
Write-Lockscreenbg

cls

function ConvertToGrayScale 
{ 
    # Converts a color bitmap image to grayscale 
    # 
    # $image is of type System.Drawing.Bitmap 
    # and can be one of System.Drawing.Imaging.ImageFormat 
    # I have tested it with .bmp, .png, .gif and .jpg files 
 
    param($image) 
     
    $grayImage = New-Object System.Drawing.Bitmap( 
        $image.Width, $image.Height) 
    $g = [System.Drawing.Graphics]::FromImage($grayImage) 
         
    $matrix = New-Object System.Drawing.Imaging.Colormatrix 
     
    $matrix.Matrix00 = 0.3 
    $matrix.Matrix01 = 0.3 
    $matrix.Matrix02 = 0.3 
    $matrix.Matrix03 = 0.0 
    $matrix.Matrix04 = 0.0 
     
    $matrix.Matrix10 = 0.59 
    $matrix.Matrix11 = 0.59 
    $matrix.Matrix12 = 0.59 
    $matrix.Matrix13 = 0.0 
    $matrix.Matrix14 = 0.0 
     
    $matrix.Matrix20 = 0.11 
    $matrix.Matrix21 = 0.11 
    $matrix.Matrix22 = 0.11 
    $matrix.Matrix23 = 0.0 
    $matrix.Matrix24 = 0.0 
     
    $matrix.Matrix30 = 0.0 
    $matrix.Matrix31 = 0.0 
    $matrix.Matrix32 = 0.0 
    $matrix.Matrix33 = 1.0 
    $matrix.Matrix34 = 0.0 
     
    $matrix.Matrix40 = 0.0 
    $matrix.Matrix41 = 0.0 
    $matrix.Matrix42 = 0.0 
    $matrix.Matrix43 = 0.0 
    $matrix.Matrix44 = 1.0 
     
    $attributes = New-Object System.Drawing.Imaging.ImageAttributes 
    $attributes.SetColorMatrix($matrix) 
     
    $imageRectangle = New-Object System.Drawing.Rectangle( 
        0, 0, $image.Width, $image.Height) 
    $g.DrawImage($image, $imageRectangle, 0, 0,  
        $image.Width, $image.Height,  
        [System.Drawing.GraphicsUnit]::Pixel, $attributes) 
     
    $g.Dispose() 
     
    $grayImage 
}

function Crop([System.IO.FileInfo]$pngFile)
{
    $rectangle = New-Object System.Drawing.Rectangle(104, 64, 170, 25)

    $img = New-Object System.Drawing.Bitmap($png.FullName);
    $crop = $img.Clone($rectangle, $img.PixelFormat);
    $img.Dispose();

    return $crop;
}

function Ocr ([System.Drawing.Bitmap]$image)
{
    Unblock-File ".\tessnet2_64.dll"
    [void][System.Reflection.Assembly]::LoadFrom(".\tessnet2_64.dll")

    $ocr = New-Object tessnet2.Tesseract
    $void = $ocr.Init(".\tessdata", "eng", $FALSE)
    $results = $ocr.DoOCR($crop, [System.Drawing.Rectangle]::Empty) | Where { $_.Confidence -le 160 } | Sort Confidence | Select -First 1
    $ocr.Dispose()

    if ($results.count -gt 0)
    {
        return $results[0].Text
    }
    else
    {
        return $null
    }
}

function Dialog ([System.Drawing.Bitmap]$crop, [System.String]$user, [System.IO.FileInfo]$fileInfo)
{
    [System.Windows.Forms.Application]::EnableVisualStyles();
    $objForm = New-Object System.Windows.Forms.Form 
    $objForm.Text = "Hearthstone Bot Report Form"
    $objForm.Size = New-Object System.Drawing.Size(500,200) 
    $objForm.StartPosition = "CenterScreen"

    $objForm.KeyPreview = $True
    $objForm.Add_KeyDown({if ($_.KeyCode -eq "Enter") 
        {$specifiedUser=$objUserTextBox.Text;$specifiedDate=$objDateTextBox.Text;$objForm.Close()}})
    $objForm.Add_KeyDown({if ($_.KeyCode -eq "Escape") 
        {$objForm.Close()}})

    $objLabel = New-Object System.Windows.Forms.Label
    $objLabel.Location = New-Object System.Drawing.Size(75,17) 
    $objLabel.Size = New-Object System.Drawing.Size(90,20)
    $objLabel.Text = "User:"
    $objForm.Controls.Add($objLabel)

    $pictureBox = new-object Windows.Forms.PictureBox
    $pictureBox.Location = New-Object System.Drawing.Size(75,37)
    $pictureBox.Width = $crop.Size.Width;
    $pictureBox.Height = $crop.Size.Height;
    $pictureBox.Image = $crop;
    $objform.controls.add($pictureBox)

    $objUserTextBox = New-Object System.Windows.Forms.TextBox
    $objUserTextBox.Location = New-Object System.Drawing.Size(247,40)
    $objUserTextBox.Size = New-Object System.Drawing.Size(125,20)
    $objUserTextBox.Text = $user
    $objForm.Controls.Add($objUserTextBox)

    $objLabel = New-Object System.Windows.Forms.Label
    $objLabel.Location = New-Object System.Drawing.Size(75,67) 
    $objLabel.Size = New-Object System.Drawing.Size(90,20)
    $objLabel.Text = "Date:"
    $objForm.Controls.Add($objLabel)

    $objDateTextBox = New-Object System.Windows.Forms.TextBox
    $objDateTextBox.Location = New-Object System.Drawing.Size(247,87)
    $objDateTextBox.Size = New-Object System.Drawing.Size(125,20)
    $objDateTextBox.Text = $fileInfo.LastWriteTimeUtc.ToString()
    $objForm.Controls.Add($objDateTextBox)

    $OKButton = New-Object System.Windows.Forms.Button
    $OKButton.Location = New-Object System.Drawing.Size(75,120)
    $OKButton.Size = New-Object System.Drawing.Size(75,23)
    $OKButton.Text = "OK"
    $OKButton.Add_Click({$specifiedUser=$objUserTextBox.Text;$specifiedDate=$objDateTextBox.Text;$objForm.Close()})
    $objForm.Controls.Add($OKButton)
    
    $CancelButton = New-Object System.Windows.Forms.Button
    $CancelButton.Location = New-Object System.Drawing.Size(150,120)
    $CancelButton.Size = New-Object System.Drawing.Size(75,23)
    $CancelButton.Text = "Cancel"
    $CancelButton.Add_Click({$objForm.Close()})
    $objForm.Controls.Add($CancelButton)

    $objForm.Topmost = $True

    $objForm.Add_Shown({$objForm.Activate()})
    [void] $objForm.ShowDialog()
}

$pngs = get-childitem $PSScriptRoot\*.* -Include *.png, *.jpg

foreach($png in $pngs)
{
    $crop = Crop $png
    $user = Ocr $crop

    $botReport = Dialog $crop $user $png

    $crop.Dispose()
}
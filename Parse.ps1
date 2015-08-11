Function Get-FileName {   
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = "Comma Separated files (*.csv)| *.csv"
    $OpenFileDialog.ShowDialog() | Out-Null
    $OpenFileDialog.filename
}

$Input  = Get-FileName
$Output = ".\Output.csv"

if (Test-Path $Output) {
    echo "$Output already exists and will be overwritten (Ctlr-C to abort)"
    pause
}

echo "Testing if nuget is available"
Try {
    gcm nuget 2> $null
    $nuget = nuget
} Catch {
    #we'll use local nuget instead
    $nuget = ".\nuget.exe"
    if (!(Test-Path $nuget))
    {
        #download nuget
        echo "Downloading nuget to this directory"
        iwr https://www.nuget.org/nuget.exe -OutFile $nuget
    }
}
echo "Installing HtmlAgilityPack version 1.4.9"
& $nuget install HtmlAgilityPack -Version 1.4.9
add-type -Path "HtmlAgilityPack.1.4.9\lib\Net40\HtmlAgilityPack.dll"
$doc = New-Object HtmlAgilityPack.HtmlDocument 

echo "Parsing $Input"
ConvertFrom-Csv (gc $Input)| % {
    
    $doc.LoadHTML($_.Description) > $null
    [PSCustomObject]@{ 
        "SKU" = $_."SKU*";
        "Description" = $doc.DocumentNode.SelectNodes("//div[@class='desc']").OuterHtml -Join " "
    }
} | Export-Csv $Output -NoTypeInformation
echo "Done"
#Script to parse csv file

$OutputFile = ".\Output.csv"

Function Get-FileName {
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = "Comma Separated files (*.csv)| *.csv"
    $OpenFileDialog.ShowDialog() | Out-Null
    $OpenFileDialog.filename
}

Function Test-Command( [string] $CommandName )
{
    (Get-Command $CommandName -ErrorAction SilentlyContinue -WarningAction SilentlyContinue) -ne $null
}
$InputFile  = Get-FileName

if (Test-Path $OutputFile) {
    echo "$OutputFile already exists and will be overwritten (Ctlr-C to abort)"
    pause
}

echo "Testing if nuget is available"
if (!(Test-Path ".\nuget.exe"))
{
    #it's not local, is it on path?
    $nuget = "nuget"
    if (!(Test-Command("nuget"))) {
        $nuget = ".\nuget.exe"
        echo ">>Downloading nuget to this directory"
        iwr https://www.nuget.org/nuget.exe -OutFile $nuget
    } else {
        echo ">>nuget was found globally"
    }
} else {
    echo ">>nuget was found locally"
}
echo "Installing HtmlAgilityPack version 1.4.9"
& $nuget install HtmlAgilityPack -Version 1.4.9
add-type -Path "HtmlAgilityPack.1.4.9\lib\Net40\HtmlAgilityPack.dll"
$doc = New-Object HtmlAgilityPack.HtmlDocument 

echo "Parsing $InputFile"
ConvertFrom-Csv (gc $InputFile)| % {
    
    $doc.LoadHTML($_.Description) > $null
    [PSCustomObject]@{ 
        "SKU" = $_."SKU*";
        "Description" = $doc.DocumentNode.SelectNodes("//div[@class='desc']").OuterHtml -Join " "
    }
} | Export-Csv $OutputFile -NoTypeInformation
echo "Done"

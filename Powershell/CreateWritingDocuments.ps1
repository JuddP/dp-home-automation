param(
    [Parameter(Position=0,mandatory=$true)]
    [string] $childName,
    [Parameter(Position=1,mandatory=$true)]
    [string] $date,
    [Parameter(Position=2,mandatory=$true)]
    [string] $handDirection,
    [string] $templateFile="Numbers and Symbols Template.docx",
    [string] $subject = "Writing",
    [string] $saveDir= "c:/homework",
    [string] $templateDir = "../Templates",
    [bool] $overWrite = $true
)

Write-Host "Creating the Writing Templates for Name(s): $childName"

[Reflection.Assembly]::LoadWithPartialName("Microsoft.Office.Interop.Word") | Out-Null

$names = $childName.Split(",")

# Open the template via Word
Write-Host "==============================="
Write-Host "> Opening Word Application (hidden) ..."
$objWord = New-Object -comobject Word.Application  
$saveFormat = [Microsoft.Office.Interop.Word.WdSaveFormat]::wdFormatDocument
$objWord.Visible = $false

 

foreach($name in $names){
    Write-Host "----------------------------"
    Write-Host ">> Processing Template for child: $name"

    # We need to make sure that we create the necessary working directory.
    $dir = "$saveDir/$name/$subject"
    if (!(Test-Path -Path $dir)){
        Write-Host ">> Creating save directory at: $dir"
        md $dir -Force | Out-Null
    }

    Write-Host ">> Opening Template file: $PSScriptRoot/$templateDir/$subject/$templateFile"
    $objDoc = $objWord.Documents.Open("$PSScriptRoot/$templateDir/$subject/$templateFile") 
    $objSelection = $objWord.Selection

    $section = $objDoc.Sections.item(1);
    $header = $Section.Headers.Item(1)
    
    Write-Host ">> Replacing contents..."
    $find = $header.Range.find
    $find.Execute("%%name%%", $False, $true, $false, $false, $false, $true, 1, $false, $name , 2)
    
    #$header = $section.Headers.Item(2);
    $find.Execute("%%Hand%%", $False, $true, $false, $false, $false, $true, 1, $false, $handDirection , 2)
    
    #$header = $section.Headers.Item(3);
    $find.Execute("%%Data%%", $False, $true, $false, $false, $false, $true, 1, $false, $date , 2)

    Write-Host ">> Saving out file to child's directory..."
    $saveFile = "$dir/$name-$handDirection-$subject-$date"
    $objDoc.SaveAs([ref][system.object]$saveFile, [ref]$saveFormat)
    $objDoc.Close()

    Write-Host ">> Work for child completed."
    Write-Host "----------------------------"
}

$objWord.quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($objWord) | Out-Null
Write-Host "> Closing Word Application (hidden) ..."
Write-Host "==============================="

Write-Host "Writing Templates Created."
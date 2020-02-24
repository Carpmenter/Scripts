<### USAGE INFO ###

<> 
<> INPUT FILE MUST BE IN Desktop FOLDER
<> PARAMETERS: .\MappingFileBuilder.ps1 <FileName> <TenantName> -> .\MappingFileBuilder.ps1 Interbrand-ExchMapping.csv Interbrand
 
##################>

Function Upload-InputFile {

    $openFileDialog = New-Object windows.forms.openfiledialog   
    $openFileDialog.initialDirectory = [Environment]::GetFolderPath("Desktop")  
    $openFileDialog.title = "Select an Input file to create Mappings"   
    $openFileDialog.filter = "All files (*.*)| *.*"
    $openFileDialog.ShowHelp = $True   
    Write-Host "Select Input File... (see FileOpen Dialog)" -ForegroundColor Green  
    $result = $openFileDialog.ShowDialog() # in ISE you may have to alt-tab or minimize ISE to see dialog box  

     if($result -eq "OK"){    
        Write-Host "Selected Downloaded Settings File:"  -ForegroundColor Green $OpenFileDialog.filename   
        Write-Host "Input File Imported!" -ForegroundColor Green 
    } 
    else { Write-Host "Input File Cancelled!" -ForegroundColor Yellow} 

    return $OpenFileDialog.filename

}

$InputFile = Upload-InputFile

# Importing input file
Try {
    $userdata = Import-Csv $InputFile
} 
Catch {
    Write-Host $_
    exit 0
}

$TenantName = $args[0]
if (!$args[0])
{ 
    $TenantName = Read-Host -Prompt 'Enter a prefix for final mapping file names'
}

$ColumnHeaders = ($userdata[0].psobject.Properties).name
$ExportPath = [Environment]::GetFolderPath("Desktop")

# Creates 4 mapping files plus wave file in cutover format
Function Create-CutoverMappings
{
    $WaveExport = $ExportPath + "\" + $TenantName + "-WaveFile.csv"
    $ArchiveWaveExport = $ExportPath + "\" + $TenantName + "-ArchiveWaveFile.csv"
    $OldNewSourceExport = $ExportPath + "\" + $TenantName + "-OldNewSource.csv"
    $OldNewTargetExport = $ExportPath + "\" + $TenantName + "-OldNewTarget.csv"
    $UserMailboxPreExport = $ExportPath + "\" + $TenantName + "-UserMailboxPre.csv"
    $UserMailboxFinalExport = $ExportPath + "\" + $TenantName + "-UserMailboxFinal.csv"

    
    ForEach ($user in $userdata){

    $user.($ColumnHeaders[0]) + ";0;" + $user.($ColumnHeaders[2]) + ";0" | Out-file $WaveExport -Append
    $user.($ColumnHeaders[0]) + ";1;" + $user.($ColumnHeaders[2]) + ";1" | Out-file $ArchiveWaveExport -Append
    $user.($ColumnHeaders[0]) + ";" + $user.($ColumnHeaders[1]) + ";3" | Out-file $OldNewSourceExport -Append
    $user.($ColumnHeaders[2]) + ";" + $user.($ColumnHeaders[3]) + ";3" | Out-file $OldNewTargetExport -Append
    $user.($ColumnHeaders[0]) + ";" + $user.($ColumnHeaders[2]) + ";0" | Out-file $UserMailboxPreExport -Append
    $user.($ColumnHeaders[1]) + ";" + $user.($ColumnHeaders[3]) + ";0" | Out-file $UserMailboxFinalExport -Append

    }
}

# Creates 1 Wave file and 1 Mapping file
Function Create-RegularMappings
{
    $WaveExport = $ExportPath + "\" + $TenantName + "-WaveFile.csv"
    $ArchiveWaveExport = $ExportPath + "\" + $TenantName + "-ArchiveWaveFile.csv"
    $MappingExport = $ExportPath + "\" + $TenantName + "-MappingFile.csv"

    ForEach ($user in $userdata){

    $user.($ColumnHeaders[0]) + ";0;" + $user.($ColumnHeaders[1]) + ";0" | Out-file $WaveExport -Append
    $user.($ColumnHeaders[0]) + ";1;" + $user.($ColumnHeaders[1]) + ";1" | Out-file $ArchiveWaveExport -Append
    $user.($ColumnHeaders[0]) + ";" + $user.($ColumnHeaders[1]) + ";0" | Out-file $MappingExport -Append
    }
}


if ($ColumnHeaders.Count -eq 2)
{ 
    Create-RegularMappings
    Write-Host "Primary/Secondary Wave files and a UserMailbox Mapping file have been created!"
} 
elseif ($ColumnHeaders.Count -eq 4)
{
    Create-CutoverMappings
    Write-Host "All Cutover Mappings and Wave files have been created!"
} 
else
{
    Write-Host "CSV format not supported!"
    exit 0
}

<# 
Handle other input formats and archive option for waves
Add logic to limit wave files to 400 lines

#>
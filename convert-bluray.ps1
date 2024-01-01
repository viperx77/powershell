$ErrorActionPreference = "Stop"
$DebugPreference = "Continue"

# This powershell script requires the following tools:
# - makemkvcon64
# - https://github.com/lisamelton/video_transcoding

$encoder = "nvenc_h265"
$makemkv = """C:\Program Files (x86)\MakeMKV\makemkvcon64.exe"""

# Uses the $sourceFile as the input to makemkvcon64.exe 
# and creates all the mkv files in the $outputFolder
function MakeMkv($sourceFile, $outputFolder) {
  if (Test-Path $outputFolder) {
    Write-Debug "output folder exists '$($outputFolder)', skipping creation of mkv"
    return
  }
  if (-not (Test-Path -Path $outputFolder -PathType Container)) {
    Write-Debug "create folder '$($outputFolder)'"
    New-Item $outputFolder -ItemType Directory -Force | Out-Null
  }
  $cmd = "echo yes | $($makemkv) -r  mkv ""$($sourceFile)"" all ""$($outputFolder)"""
  Write-Debug $cmd
  & cmd.exe /c $cmd
}

function TranscodeFile($encoder, $source, $output) {
  if (Test-Path $output) {
    Write-Debug "File exists $($output), skipping."
    return;
  }
  Write-Debug "transcode-video --encoder $($encoder) --crop detect --fallback-crop minimal $($source) -o $($output) "
  & "transcode-video" "--encoder" $encoder "--crop" "detect" "--fallback-crop" "minimal" $source "-o" $output
}

function TranscodeMkvFolder($folder, $movieFolder, $movieExtrasFolder) {
  $title = $folder.Name
  # Assume largest file is the movie
  $movieFile = Get-ChildItem -Path $folder.FullName -File | Sort-Object -Property Length -Descending | Select-Object -First 1
  if ($movieFile) {
    $output = "$($movieFolder)\$($title).mkv"
    TranscodeFile $encoder $movieFile.FullName $output 
    # All other files are the extras
    $extraFiles = Get-ChildItem -Path $folder -File | Sort-Object -Property Length -Descending | Select-Object -Skip 1
    if ($extraFiles -and $extraFiles.Count -ne 1 -and $extraFiles[0].FullName -ne $movieFile.FullName) {
      $outputFolder = "$($movieExtrasFolder)\$($title)"
      if (-not (Test-Path -Path $outputFolder -PathType Container)) {
        Write-Debug "create folder '$($outputFolder)'"
        New-Item $outputFolder -ItemType Directory -Force | Out-Null
      }
      $i = 0
      $extraFiles | ForEach-Object {
        $i += 1
        $output = "$($outputFolder)\$($title) (extra $($i)).mkv"
        TranscodeFile $encoder $_.FullName $output
      }
    }
  }
}

function SkipSourceFolder($folder) {
  # Skip source folder that have an underscore as the first character of the folder name so
  # they can be manually renamed  to be skipped. This allows the script to be srestarted in case of 
  # bad files or the script hanging up.
  return ($folder.Substring(0, 1) -eq "_")
}

# This is the root location of the source rips
$sourceRoot = "I:\"
# This is the root folder to create folder per rip containing all the mkvs (output from makemkv)
$rootMkvFolder = "F:\Carl\mkv"
# This is the root folder to output all the main transcoded movie files
$movieFolder = "F:\Carl\movies"
# This is the root folder to create a folder per movie with all the extra transcoded files
$movieExtrasFolder = "F:\Carl\movie-extras"
# Should the makemkv output files nto be deleted
$keepMakeMkvOutput = $false

$brFiles =  Get-ChildItem $sourceRoot -Recurse -Filter 'index.bdmv' | Where-Object { $_.DirectoryName -notmatch 'BACKUP' }
$isoFiles = Get-ChildItem $sourceRoot -Recurse -Filter '*.iso'
 $tsFiles = Get-ChildItem $sourceRoot -Recurse -Filter 'VIDEO_TS.IFO'
$allFiles = $brFiles + $isoFiles + $tsFiles

$allFiles | ForEach-Object { 
  $filePath = $_.FullName
  # Skip source folder that have an underscore as the first character of the folder name
  # to allow the script to be cancelled and it will pick up where it left off
  $pathParts = $filePath -split '\\' | Where-Object { $_ }
  $mkvsFolder = $pathParts[1]
  if (SkipSourceFolder $mkvsFolder) {
    Write-Debug "Skipping $($filePath)"
    return;
  }
  $outputFolder = "$($rootMkvFolder)\$($mkvsFolder)"  
  MakeMkv $filePath $outputFolder
  $outputFolderObject = Get-Item -Path $outputFolder
  TranscodeMkvFolder $outputFolderObject $movieFolder $movieExtrasFolder
  if (-not $keepMakeMkvOutput) {
    Write-Debug "delete converted source mkv folder '$($outputFolder)'"
    Remove-Item -Path $outputFolder -Recurse -Force 
  }
  Write-Debug "rename '$($pathParts[0])\$($mkvsFolder)' to '$($pathParts[0])\_$($mkvsFolder)'"
  Rename-Item -Path "$($pathParts[0])\$($mkvsFolder)" -NewName "$($pathParts[0])\_$($mkvsFolder)"
}

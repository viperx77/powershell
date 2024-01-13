$ErrorActionPreference = "Stop"
$DebugPreference = "Continue"

# This powershell script requires the following tools:
# - makemkvcon64
# - https://github.com/lisamelton/video_transcoding

$encoder = "nvenc_h265"
$makemkv = "C:\Program Files (x86)\MakeMKV\makemkvcon64.exe"

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

  $titleNums = GetTitlesToConvert $sourceFile
  $titleNums | % {
    $cmd = "echo yes | ""$($makemkv)"" -r  mkv ""$($sourceFile)"" $_ ""$($outputFolder)"""
    Write-Debug $cmd
    & cmd.exe /c $cmd
  }  
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

function GetSectionLines($lines, $section, $num, $num2) {
  if (-not $num2) {
    $lines | ? { $_ -match "^$($section):$($num),.*" }
  } else {
    $lines | ? { $_ -match "^$($section):$($num),$($num2),.*" }
  }
}

function FilterNum($lines, $section, $num, $num2) {
  $lines | ? { $_ -match "^$($section):$($num),$($num2),.*" }
}

function CINFO($line) {
  $parts = $line.Split(',')
  $parts[2].Replace('"', '')  
}

function TINFO($line) {
  if (-not $line) {
    return ""
  }
  $parts = $line.Split(',')
  $parts[3].Replace('"', '')  
}

function GetTitlesToConvert($sourceFile) {

  $lines = (& $makemkv "-r" "info" ""$sourceFile"")
  
  $t = CINFO (GetSectionLines $lines "CINFO" "2")
  $f = CINFO (GetSectionLines $lines "CINFO" "32")
  
  $CHAPTERS=8
  $SIZE = 11
  $LANGUAGE=29
  
  $titleNum = 0
  $tinfo = GetSectionLines $lines "TINFO" $titleNum
  $titles = @()
  
  while ($tinfo) {
    $c = TINFO (FilterNum $tinfo "TINFO" $titleNum $CHAPTERS)
    $s = TINFO (FilterNum $tinfo "TINFO" $titleNum $SIZE)
    $l = TINFO (FilterNum $tinfo "TINFO" $titleNum $LANGUAGE)
    Write-Debug "title: $titleNum, chapters: $c, size: $s, language: $l"
    $titles += [PSCustomObject]@{
      "track" = $titleNum
      "title" = $t
      "chapters" = $c
      "size" = [double]$s
      "language" = $l
    }
    $titleNum += 1
    $tinfo = GetSectionLines $lines "TINFO" $titleNum
  }
  
  $onlyEnglishTitles = $titles | Where-Object { $_.language -eq "English" }
  
  $sortedTitles = $onlyEnglishTitles | Sort-Object -Property "size" -Descending
  $largestTitle = $sortedTitles | Select-Object -First 1
  $nonRepeatingTitles = $titles | Where-Object { $_.size -ne $largestTitle.size }
  
  $tout = @()
  $tout += $largestTitle.track 
  $tout += ($nonRepeatingTitles | % { $_.track })
  $tout
}

function SkipSourceFolder($folder) {
  # Skip source folder that have an underscore as the first character of the folder name so
  # they can be manually renamed  to be skipped. This allows the script to be srestarted in case of 
  # bad files or the script hanging up.
  return ($folder.Substring(0, 1) -eq "_")
}

# This is the root location of the source rips
$sourceRoot = "H:\"
# This is the root folder to create folder per rip containing all the mkvs (output from makemkv)
$rootMkvFolder = "f:\carl\mkv"
# This is the root folder to output all the main transcoded movie files
$movieFolder = "i:\movies"
# This is the root folder to create a folder per movie with all the extra transcoded files
$movieExtrasFolder = "i:\movie-extras"
# Should the makemkv output files nto be deleted
$keepMakeMkvOutput = $false

# pause for this amount number of seconds  to ask if should cancel
$timeoutSeconds = 5

$brFiles = Get-ChildItem $sourceRoot -Recurse -Filter 'index.bdmv' | Where-Object { $_.DirectoryName -notmatch 'BACKUP' }
$isoFiles = Get-ChildItem $sourceRoot -Recurse -Filter '*.iso'
$tsFiles = Get-ChildItem $sourceRoot -Recurse -Filter 'VIDEO_TS.IFO'
$allFiles = [array]$brFiles + [array]$isoFiles + [array]$tsFiles

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

  if ($timeoutSeconds -gt 0) {
    Write-Host "Stop processing?"
    $scriptBlock = {
      $response = Read-Host "Stop processing?"
      return $response
    }
    $job = Start-Job -ScriptBlock $scriptBlock
    Start-Sleep -Seconds $timeoutSeconds
    if (Get-Job -Id $job.Id -ErrorAction SilentlyContinue) {
      Write-Debug "Timed out, continuing..."
      Stop-Job -Id $job.Id
      Remove-Job -Id $job.Id
    }
    else {
      $result = Receive-Job -Id $job.Id
      Remove-Job -Id $job.Id
      if ($result -eq "Y" -or $result -eq "y") {
        exit
      }      
    }  
  }
}

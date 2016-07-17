<#==================================================================================
conv2mp4 - https://github.com/BrianDMG/conv2mp4-ps v1.3

This Powershell script will recursively search through a defined file path and
convert all MKV, AVI, FLV, and MPEG files to MP4 using ffmpeg + audio to AAC. If it
detects a conversion failure, it will re-encode the file with Handbrake.
It then refreshes a Plex library, and upon conversion success deletes the source 
(original) file. The purpose of this script is to reduce transcodes from Plex.
=====================================================================================

This script requires ffmpeg and Handbrake. You can download them at the URLs below.
ffmpeg : https://ffmpeg.org/download.html
handbrakecli : https://handbrake.fr/downloads.php

-------------------------------------------------------------------------------------
User-specific variables
-------------------------------------------------------------------------------------
$mediaPath = the path to the media you want to convert
NOTE: to use mapped drive, you must run "net use z: \\server\share /persistent:yes" as the user you're going to run the script as (generally Administrator) prior to running the script
$fileTypes = the extensions of the files you want to convert in the format "*.ex1", "*.ex2" 
$log = path you want the log file to save to. defaults to your desktop.
$plexIP = the IP address and port of your Plex server (for the purpose of refreshing its library)
$plexToken = your Plex server's token (for the purpose of refreshing its library). 
-Plex server token - See https://support.plex.tv/hc/en-us/articles/204059436-Finding-your-account-token-X-Plex-Token
--Plex server token is also easy to retrive with Couchpotato or SickRage 
$ffmpeg = path to ffmpeg.exe
$handbrake = path to HandBrakeCLI.exe
$filebot = path to filebot.exe#>
$mediaPath = "Path to media to be converted"
$fileTypes = "*.mkv", "*.avi", "*.flv", "*.mpeg", "*.wmv"
$log = "C:\Users\$env:username\Desktop\conv2mp4-ps.log"
$plexIP = '127.0.0.1:32400'
$plexToken = ''
$ffmpeg = "C:\ffmpeg\bin\ffmpeg.exe"
$handbrake = "C:\Program Files\HandBrake\HandBrakeCLI.exe"
$fileBot = "C:\Program Files\FileBot\filebot.exe"
$tvShowPattern1 = "s[0-1][0-9]"
$tvShowPattern2	= "[0-9]x[0-9][0-9]"
$langPattern = "hun"
$movieTarget = "Path where you want your movies"
$langMovieTarget = "Path where you want your different language movies"
$tvShowTarget = "Path where you want your TV SHows"
$langTVShowTarget = "Path where you want your different langue Tv Shows"

<#----------------------------------------------------------------------------------
	Rename Files using Filebot,
	looks for patterns in the name to determine if it's a TV Show or a Movie
----------------------------------------------------------------------------------#>
function RenameFiles {
	
	$mPath = Get-Item -Path $mediaPath
	$fileList = Get-ChildItem "$($mPath.FullName)\*" -i $fileTypes -recurse
	
	ForEach ($file in $fileList)
	{
		$fbarg1 = "-rename"
		$fbarg2 = $file.FullName
		$fbarg3 = "--db"
		$fbarg4 = ""
		$fbarg5 = "-non-strict"
		
		if($file.FullName -like "*$tvShowPattern1*" -Or $file.FullName -like "*$tvShowPattern2*"){
			$fbarg4 = "TheTVDB"
		}else {	
			$fbarg4 = "TheMovieDB"
		}
		
		$fbargs = @($fbarg1, $fbarg2, $fbarg3, $fbarg4, $fbarg5)
		$fbcmd = &$filebot $fbargs

		$fbcmd
	}
}

function GetFileDifference([string]$oldFile,[string]$newFile){
	$LengthColumn = 27
	$shell = New-Object -ComObject Shell.Application 
	$folder = $folder = Split-Path $oldFile
	$shellfolder = $shell.Namespace($folder)
	$old = Split-Path $oldFile -Leaf
	$new = Split-Path $newFile -Leaf
	$fileOld = $shellfolder.ParseName($old)
	$oldLength = $shellfolder.GetDetailsOf($fileOld, $LengthColumn)
	$shellfolder = $shell.Namespace($file.DirectoryName)
	$fileNew = $shellfolder.ParseName($new)
	$newLength = $shellfolder.GetDetailsOf($fileNew, $LengthColumn)
	
	$Difference = ([timespan]$oldLength).Subtract([timespan]$newLength).Duration().TotalSeconds
	
	return $Difference
}

function ffmpegFallback([string]$oldFile,[string]$newFile){
	
	Remove-Item $newFile -Force
	Write-Output "$($time.Invoke()) EXCEPTION: New file is over 25% smaller ($($diffErr)MB). $newFile deleted." | Out-File $log -Append
	Write-Output "$($time.Invoke()) FAILOVER: Re-encoding $oldFile with Handbrake." | Out-File $log -Append
		
	<#----------------------------------------------------------------------------------
	ffmpeg variables
	----------------------------------------------------------------------------------#>
	$ffarg1 = "-n"
	$ffarg2 = "-fflags"
	$ffarg3 = "+genpts"
	$ffarg4 = "-i"
	$ffarg5 = "$oldFile"
	$ffarg6 = "-map"
	$ffarg7 = "0"
	$ffarg8 = "-vcodec"
	$ffarg9 = "copy"
	$ffarg13 = "-scodec"
	$ffarg14 = "mov_text"
	$ffarg10 = "-acodec"
	$ffarg11 = "aac"
	$ffarg12 = "$newFile"

	$ffargs = @($ffarg1, $ffarg2, $ffarg3, $ffarg4, $ffarg5, $ffarg6, $ffarg7, $ffarg8, $ffarg9, $ffarg13, $ffarg14, $ffarg10, $ffarg11, $ffarg12)
	$ffcmd = &$ffmpeg $ffargs
	
	<#----------------------------------------------------------------------------------
	Begin ffmpeg conversion (lossless)
	-----------------------------------------------------------------------------------#>
	$ffcmd
	Write-Output "$($time.Invoke()) ffmpeg completed" | Out-File $log -Append	

}

function MoveFileToPlex([string]$newFile){
	<#----------------------------------------------------------------------------------
	Check if it's a TV Show or movie
	Then check if it's Hungarian or not.
	----------------------------------------------------------------------------------#>
	$fileTarget = ""
	$fileName = Split-Path $newFile -Leaf
	Write-Host $newFile
	if($newFile -like "*$tvShowPattern1*" -Or $newFile -like "*$tvShowPattern2*") {
		if($newFile -like "*$langPattern*"){
			$fileTarget = $langTVShowTarget + "\"
		}else{
			$fileTarget = $tvShowTarget + "\"
		}
		$showFolder = $filename.split("-").trim()
		$showFolder = $showFolder[0].Trim()
		$fileTarget = $fileTarget + $showFolder
		Write-Host $fileTarget
		if( -Not(Test-Path $fileTarget)){
			New-Item -ItemType directory -Path $fileTarget
		}
	} else {
		if($newFile -like "*$langPattern*"){
			$fileTarget = $movieTarget + "\"
		}else{
			$fileTarget = $langMovieTarget + "\"
		}
	}
	$fileTarget = $fileTarget + "\" + $fileName
	Move-Item $newFile $fileTarget
}


function ConvertWithHandbrake {	
	<#----------------------------------------------------------------------------------
	Variables
	----------------------------------------------------------------------------------#>
	$mPath = Get-Item -Path $mediaPath
	$fileList = Get-ChildItem "$($mPath.FullName)\*" -i $fileTypes -recurse
	$num = $fileList | measure
	$fileCount = $num.count
	$time = {Get-Date -format "MM/dd/yy HH:mm:ss"}
		
	<#----------------------------------------------------------------------------------
	Begin search loop 
	----------------------------------------------------------------------------------#>
	$i = 0;
	ForEach ($file in $fileList)
	{
		$i++;
		$oldFile = $file.DirectoryName + "\" + $file.BaseName + $file.Extension;
		$newFile = $file.DirectoryName + "\" + $file.BaseName + ".mp4";
		$plexURL = "http://$plexIP/library/sections/all/refresh?X-Plex-Token=$plexToken"
		$progress = ($i / $fileCount) * 100
		$progress = [Math]::Round($progress,2)
		
		<#----------------------------------------------------------------------------------
		Do work
		----------------------------------------------------------------------------------#>
		Clear-Host
		Write-Output "------------------------------------------------------------------------------------" | Out-File $log -Append
		Write-Output "$($time.Invoke()) Processing - $oldFile" | Out-File $log -Append
		Write-Output "$($time.Invoke()) File $i of $fileCount - $progress%" | Out-File $log -Append
		
		<#----------------------------------------------------------------------------------
		Refresh Plex libraries in Chrome
		-----------------------------------------------------------------------------------#>
		Invoke-WebRequest $plexURL 
		Write-Output "$($time.Invoke()) Plex library refreshed" | Out-File $log -Append 
		
		# Begin Handbrake encode (lossy)
		# Handbrake CLI: https://trac.handbrake.fr/wiki/CLIGuide#presets
		
		<#----------------------------------------------------------------------------------
		HandbrakeCLI variables
		----------------------------------------------------------------------------------#>
		$hbarg1 = "-i"
		$hbarg2 = "$oldFile"
		$hbarg3 = "-o"
		$hbarg4 = "$newFile"
		$hbarg5 = "-f"
		$hbarg6 = "mp4"
		$hbarg7 = "--preset=High Profile"
		$hbarg8 = "--verbose=1"
		$hbarg9 = "-s"
		$hbarg10 = "1,2,3,4"
		$hbarg11 = "-a"
		$hbarg12 = "1,2,3,4"
		$hbargs = @($hbarg1, $hbarg2, $hbarg3, $hbarg4, $hbarg5, $hbarg6, $hbarg7, $hbarg8, $hbarg9, $hbarg10, $hbarg11, $hbarg12)
		$hbcmd = &$handbrake $hbargs

		$hbcmd
		Write-Output "$($time.Invoke()) Handbrake finished." | Out-File $log -Append
		
		<#----------------------------------------------------------------------------------
		Check Difference in file length
		if it's greater than 5 seconds, fallback to ffmpeg
		----------------------------------------------------------------------------------#>
		$fileDifference = GetFileDifference $oldFile $newFile
		
		if($fileDifference -gt 5 -Or $fileDifference -lt -5){
			ffmpegFallback $oldFile $newFile
		}else{
			Remove-Item $oldFile -Force
			Write-Output "$($time.Invoke()) Same file duration $oldFile deleted." | Out-File $log -Append
		}
		
		MoveFileToPlex $newFile
	}
}

<#----------------------------------------------------------------------------------
Rename Files
----------------------------------------------------------------------------------#>
RenameFiles
ConvertWithHandbrake


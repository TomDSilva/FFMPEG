###############################################################################################################################
#                                                                                                                             #
#  A PowerShell script to provide a CLI menu frontend to the popular FFMPEG tool with a selection of useful options.          #
#                                                                                                                             #
#  DISCLAIMER: THIS CODE IS PROVIDED FREE OF CHARGE. UNDER NO CIRCUMSTANCES SHALL I HAVE ANY LIABILITY TO YOU FOR ANY LOSS    #
#  OR DAMAGE OF ANY KIND INCURRED AS A RESULT OF THE USE OF THIS CODE. YOUR USE OF THIS CODE IS SOLELY AT YOUR OWN RISK.      #
#                                                                                                                             #
#  By Tom D'Silva 2019 - https://github.com/TomDSilva                                                                         #
#                                                                                                                             #
###############################################################################################################################

###############################################################################################################################
### Version History                                                                                                         ###
###############################################################################################################################
# 1.0  :          : First release                                                                                             #
# 1.1  :          : Major rewrite, added new option (join), moved file selection into function, filtered files it displays    #
#                   for choice.                                                                                               #
# 1.2  :          : Fixed file display bugs.                                                                                  #
# 1.3  :          : Fixed bug where it wouldnt copy all audio streams.                                                        #
# 1.4  :          : Added option to change volume level.                                                                      #
# 1.5  : 22/08/20 : Moved start time to beginning of command as this saves time parsing the input file.                       #
# 1.6  : 23/08/20 : Added 6th option to cut and re-encode in the case that the normal cut method results in frozen frames at  #
#                   the start.                                                                                                #
# 1.7  : 27/04/21 : Now has the option to remove video and leave audio                                                        #
# 1.8  : 19/09/21 : Added option to reverse a video.                                                                          #
# 1.9  : 27/05/22 : Added option to convert WAV to FLAC.                                                                      #
#                   Tidied up script.                                                                                         #
#                   Now asks to open file and then delete                                                                     #
# 1.10 : 13/03/23 : Added handling for start time variables to be automatically set to the start if null.                     #
# 1.11 : 26/04/23 : Changed banner to include warning and author information.                                                 #
#                   Added a new function to check if a variable exists.                                                       #
#                   Better checks for existing variables.                                                                     #
#                   Variables laid out better using script scope.                                                             #
#                   File selection will now only accept integers and only those for files that it has checked already exist.  #
#                   Now checks if the edited files output location exists, if it doesn't then it creates the folder.          #
#                   Checks if ffmpeg.exe exists in the right location, if not then exit with an error.                        #
#                   Reformatted ffmpeg commands so they adhere to best practices.                                             #
#                   Fixed bug where option 10 was hardcoded to a set location.                                                #
#                   Tidied up version history formatting.                                                                     #
#                                                                                                                             #
# Possible future changes:                                                                                                    #
# Join more then 2 files.                                                                                                     #
# Show progress bar.                                                                                                          #
# Make options more dynamic.                                                                                                  #
# Splice in another audio track                                                                                               #
###############################################################################################################################

###############################################################################################################################
### Script Location Checker                                                                                                 ###
###############################################################################################################################

# Get the full path of this script:
$scriptPath = $MyInvocation.MyCommand.Path
# Remove the "file" part of the path so that only the directory path remains:
$scriptPath = Split-Path $scriptPath
# Change location to where the script is being run:
Set-Location $scriptPath

###############################################################################################################################
### End Of Script Location Checker                                                                                          ###
###############################################################################################################################

###############################################################################################################################
### Functions                                                                                                               ###
###############################################################################################################################

Function VariableExists ($variable) {
    if (Get-Variable $variable -ErrorAction SilentlyContinue) {
        return $true
    }
    else {
        return $false
    }
}

# Main menu function that is used with the Show-SelectionMenu:
Function Show-Menu {
    param (
        [string]$title = 'FFMPEG'
    )
    Clear-Host
    Write-Host "===================== $title ====================="
    Write-Host ""
    Write-Host "By Tom D'Silva 2019 - https://github.com/TomDSilva"
    Write-Host ""
    Write-Host "DISCLAIMER: THIS CODE IS PROVIDED FREE OF CHARGE. UNDER NO CIRCUMSTANCES SHALL I HAVE ANY LIABILITY TO YOU FOR ANY LOSS"
    Write-Host "OR DAMAGE OF ANY KIND INCURRED AS A RESULT OF THE USE OF THIS CODE. YOUR USE OF THIS CODE IS SOLELY AT YOUR OWN RISK."
    Write-Host ""
    Write-Host "Make your selection"
    Write-Host "1  -  Trim"
    Write-Host "2  -  Remove Audio"
    Write-Host "3  -  Remove Audio & Trim"
    Write-Host "4  -  Remove Video"
    Write-Host "5  -  Join 2 Videos (Must Be Same Codecs)"
    Write-Host "6  -  Change Volume"
    Write-Host "7  -  Cut And Convert To MP4 (THIS RE-ENCODES), if the normal cut method had freezing issues at the start then this may work better with the source file"
    Write-Host "8  -  Reverse A Video (THIS USES MASSIVE AMOUNTS OF RAM!!!)"
    Write-Host "9  -  Convert ALL WAV files to FLAC"
    Write-Host "10 -  Convert A Single WAV file to FLAC"
    Write-Host "Q  -  Submit 'Q' to quit"
}

Function Show-SelectionMenu ($selection, [INT]$numberOfSelections) {

    "You chose option #$selection"
    if ($numberOfSelections -NE 0) {
        $filesPath = $scriptPath + "\*"
        $arrayFiles = Get-ChildItem -Path $filesPath -Attributes !Directory+!System -Include '*.mkv', '*.flv', '*.mp4', '*.m4a', '*.m4v', '*.f4v', '*.f4a', '*.m4b', '*.m4r', '*.f4b', '*.mov', '*.3gp', '*.3gp2', '*.3g2', '*.3gpp', '*.3gpp2', '*.ogg', '*.oga', '*.ogv', '*.ogx', '*.wmv', '*.wma', '*.asf', '*.VOB'
        $menu = @{ }

        For ($i = 1; $i -le $arrayFiles.count; $i++) {
            Write-Host "$i. $($arrayFiles[$i-1].name),$($arrayFiles[$i-1].status)" 
            $menu.Add($i, ($arrayFiles[$i - 1].name))
        }

        If ($numberOfSelections -EQ 1) {
            $i = 0
            do {
                if ($i -ge 1) {
                    Write-Warning "ERROR - Not a valid selection, please try again"
                    Write-Host ($menu | Out-String)
                }
                [INT]$ans = Read-Host -Prompt "Please select a file by number:"
                $i++
            } until ($ans -in $menu.Keys)
            if (VariableExists 'ans1') {
                Remove-Variable 'ans1'
            }
            New-Variable -Name 'ans1' -Value $menu.Item($ans) -Scope 'Script'
            if (VariableExists 'outputFile') {
                Remove-Variable 'outputFile'
            }
            New-Variable -Name 'outputFile' -Value "$scriptPath\Edited Files\$ans1" -Scope 'Script'
        }

        If ($numberOfSelections -EQ 2) {
            $i = 0
            do {
                if ($i -ge 1) {
                    Write-Warning "ERROR - Not a valid selection, please try again"
                    Write-Host ($menu | Out-String)
                }
                [INT]$ans = Read-Host "Please select the first file by number:"
                $i++
            } until ($ans -in $menu.Keys)
            if (VariableExists 'ans1') {
                Remove-Variable 'ans1'
            }
            New-Variable -Name 'ans1' -Value $menu.Item($ans) -Scope 'Script'

            $i = 0
            do {
                if ($i -ge 1) {
                    Write-Warning "ERROR - Not a valid selection, please try again"
                    Write-Host ($menu | Out-String)
                }
                [INT]$ans = Read-Host "Please select the second file by number:"
                $i++
            } until ($ans -in $menu.Keys)
            if (VariableExists 'ans2') {
                Remove-Variable 'ans2'
            }
            New-Variable -Name 'ans2' -Value $menu.Item($ans) -Scope 'Script'

            if (VariableExists 'outputFile') {
                Remove-Variable 'outputFile'
            }
            New-Variable -Name 'outputFile' -Value "$scriptPath\Edited Files\JOINED $ans1 - $ans2" -Scope 'Script'
        }
    }
}

Function HandleFile ($file) {
    $selection = Read-Host "Do you want to open the file? [y/N]"
    if ($selection -eq 'y') {
        Invoke-Item $file
        $selection = Read-Host "Delete? [y/N]"
        if ($selection -eq 'y') {
            try {
                Remove-Item $file
            }
            catch {
                Write-Warning '---'
                Write-Warning "OOPS! I COULDNT DELETE THE FOLLOWING FILE:"
                Write-Warning $file
                Write-Warning "The error message was:"
                Write-Warning $error[0]
                Write-Warning '---'
                Pause
            }
        }
    }
}

###############################################################################################################################
### End Of Functions                                                                                                        ###
###############################################################################################################################

# If the "Edited Files" folder doesnt exist, then create it:
if (!(Test-Path "$scriptPath\Edited Files\")) {
    New-Item -ItemType "directory" -Path "$scriptPath\Edited Files\"
}

# If ffmpeg.exe is not present in the correct location then exit with error:
if (!(Test-Path "$scriptPath\bin\ffmpeg.exe")) {
    Write-Error "ffmpeg.exe not found in bin folder"
    exit 1
}

do {
    Show-Menu 
    $selection = Read-Host "Please make a selection"
    switch ($selection) {
        '1' {
            Show-SelectionMenu "1  -  Trim" 1

            do {
                # Prompts for the Start Time
                $startTime = Read-Host -Prompt "Start Time? (HH:MM:SS or HH:MM:SS.SSS) OR just press enter to begin at the start of file."
                # Prompts for the Stop Time
                $stopTime = Read-Host -Prompt "Stop Time? (HH:MM:SS or HH:MM:SS.SSS) OR just press enter to skip to end of file."

                if ([string]::IsNullOrWhiteSpace($startTime) -and [string]::IsNullOrWhiteSpace($stopTime)) {
                    Write-Output "-------------------------------------------"
                    Write-Output "Start time AND stop time are both null!!!"
                    Write-Output "Please set at least 1 of these to continue."
                    Write-Output "-------------------------------------------"
                }
            } until (![string]::IsNullOrWhiteSpace($startTime) -or ![string]::IsNullOrWhiteSpace($stopTime))

            if (![string]::IsNullOrWhiteSpace($startTime) -and ![string]::IsNullOrWhiteSpace($stopTime)) {
                .\bin\ffmpeg.exe -i "$ans1" -ss $startTime -to $stopTime -acodec copy -vcodec copy -map 0 $outputFile
            }
            elseif ([string]::IsNullOrWhiteSpace($startTime) -and ![string]::IsNullOrWhiteSpace($stopTime)) {
                $startTime = '00:00:00'
                .\bin\ffmpeg.exe -i "$ans1" -ss $startTime -to $stopTime -acodec copy -vcodec copy $outputFile
            }
            elseif (![string]::IsNullOrWhiteSpace($startTime) -and [string]::IsNullOrWhiteSpace($stopTime)) {
                .\bin\ffmpeg.exe -i "$ans1" -ss $startTime -acodec copy -vcodec copy $outputFile
            }

            HandleFile $outputFile
        }

        '2' {
            Show-SelectionMenu "2  -  Remove Audio" 1

            .\bin\ffmpeg.exe -loglevel quiet -i "$ans1" -vcodec copy -an $outputFile

            HandleFile $outputFile
        }

        '3' {
            Show-SelectionMenu "3  -  Remove Audio & Trim" 1

            do {
                # Prompts for the Start Time
                $startTime = Read-Host -Prompt "Start Time? (HH:MM:SS or HH:MM:SS.SSS) OR just press enter to begin at the start of file."
                # Prompts for the Stop Time
                $stopTime = Read-Host -Prompt "Stop Time? (HH:MM:SS or HH:MM:SS.SSS) OR just press enter to skip to end of file."

                if ([string]::IsNullOrWhiteSpace($startTime) -and [string]::IsNullOrWhiteSpace($stopTime)) {
                    Write-Output "-------------------------------------------"
                    Write-Output "Start time AND stop time are both null!!!"
                    Write-Output "Please set at least 1 of these to continue."
                    Write-Output "-------------------------------------------"
                }
            } until (![string]::IsNullOrWhiteSpace($startTime) -or ![string]::IsNullOrWhiteSpace($stopTime))

            if (![string]::IsNullOrWhiteSpace($startTime) -and ![string]::IsNullOrWhiteSpace($stopTime)) {
                .\bin\ffmpeg.exe -i "$ans1" -ss $startTime -to $stopTime -vcodec copy -map 0 -an $outputFile
            }
            elseif ([string]::IsNullOrWhiteSpace($startTime) -and ![string]::IsNullOrWhiteSpace($stopTime)) {
                $startTime = '00:00:00'
                .\bin\ffmpeg.exe -i "$ans1" -ss $startTime -to $stopTime -vcodec copy -map 0 -an $outputFile
            }
            elseif (![string]::IsNullOrWhiteSpace($startTime) -and [string]::IsNullOrWhiteSpace($stopTime)) {
                .\bin\ffmpeg.exe -i "$ans1" -ss $startTime -vcodec copy -map 0 -an $outputFile
            }

            HandleFile $outputFile
        }

        '4' {
            Show-SelectionMenu "4  -  Remove Video" 1

            .\bin\ffmpeg.exe -loglevel quiet -i "$ans1" -acodec copy -vn $outputFile

            HandleFile $outputFile
        }

        '5' {
            Show-SelectionMenu "5  -  Join 2 Videos (Must Be Same Codecs)" 2

            New-Item "$scriptPath\Input.txt"
            Add-Content -Path "$scriptPath\Input.txt" -Value "file '$ans1'"
            Add-Content -Path "$scriptPath\Input.txt" -Value "file '$ans2'"

            .\bin\ffmpeg.exe -loglevel quiet -f concat -safe 0 -i Input.txt -acodec copy -vcodec copy -map 0 $outputFile

            Remove-Item "$scriptPath\Input.txt"

            HandleFile $outputFile
        }

        '6' {
            Show-SelectionMenu "6  -  Change Volume" 1

            [string]$volume = Read-Host "Input dB change (use - to indicate a reduction)"

            $volume = $volume + 'dB'

            .\bin\ffmpeg.exe -loglevel quiet -i "$ans1" -vcodec copy -map 0 -af "volume=$volume" $outputFile

            HandleFile $outputFile
        }
        
        '7' {
            Show-SelectionMenu "7  -  Cut and convert to MP4 (THIS RE-ENCODES)" 1
            
            do {
                # Prompts for the Start Time
                $startTime = Read-Host -Prompt "Start Time? (HH:MM:SS or HH:MM:SS.SSS) OR just press enter to begin at the start of file."
                # Prompts for the Stop Time
                $stopTime = Read-Host -Prompt "Stop Time? (HH:MM:SS or HH:MM:SS.SSS) OR just press enter to skip to end of file."

                if ([string]::IsNullOrWhiteSpace($startTime) -and [string]::IsNullOrWhiteSpace($stopTime)) {
                    Write-Output "-------------------------------------------"
                    Write-Output "Start time AND stop time are both null!!!"
                    Write-Output "Please set at least 1 of these to continue."
                    Write-Output "-------------------------------------------"
                }
            } until (![string]::IsNullOrWhiteSpace($startTime) -or ![string]::IsNullOrWhiteSpace($stopTime))

            if (![string]::IsNullOrWhiteSpace($startTime) -and ![string]::IsNullOrWhiteSpace($stopTime)) {
                .\bin\ffmpeg.exe -i "$ans1" -ss $startTime -to $stopTime -acodec copy -map 0 $outputFile
            }
            elseif ([string]::IsNullOrWhiteSpace($startTime) -and ![string]::IsNullOrWhiteSpace($stopTime)) {
                $startTime = '00:00:00'
                .\bin\ffmpeg.exe -i "$ans1" -ss $startTime -to $stopTime -acodec copy -map 0 $outputFile
            }
            elseif (![string]::IsNullOrWhiteSpace($startTime) -and [string]::IsNullOrWhiteSpace($stopTime)) {
                .\bin\ffmpeg.exe -i "$ans1" -ss $startTime -acodec copy -map 0 $outputFile
            }

            HandleFile $outputFile
        }

        '8' {
            Show-SelectionMenu "8  -  Reverse a Video (THIS USES MASSIVE AMOUNTS OF RAM!!!)" 1

            .\bin\ffmpeg.exe -i "$ans1" -vf reverse -af areverse $outputFile

            HandleFile $outputFile
        }

        '9' {
            Show-SelectionMenu "9  -  Convert ALL WAV files to FLAC" 0

            $filesPath = $scriptPath + "\*"
            $arrayFiles = Get-ChildItem -Path $filesPath -Attributes !Directory+!System -Include '*.wav'

            foreach ($file in $arrayFiles) {
                .\bin\ffmpeg.exe -i $file ""$scriptPath'\Edited Files\'$($file.BaseName).FLAC""
            }
        }

        '10' {
            Show-SelectionMenu "10  -  Convert A Single WAV file to FLAC" 0

            $filesPath = $scriptPath + "\*"
            $arrayFiles = Get-ChildItem -Path $filesPath -Attributes !Directory+!System -Include '*.wav'
            $menu = @{ }
            
            For ($i = 1; $i -le $arrayFiles.count; $i++) {
                Write-Host "$i. $($arrayFiles[$i-1].name),$($arrayFiles[$i-1].status)" 
                $menu.Add($i, ($arrayFiles[$i - 1].name))
            }

            [INT]$ans = Read-Host -Prompt "Please select a file:"
            [STRING]$ans1 = $menu.Item($ans)

            .\bin\ffmpeg.exe -i $ans1 ""$scriptPath'\Edited Files\'$($ans1.Substring(0,$ans1.Length-4)).FLAC""
        }

    }

}
Until ($selection -eq 'q')

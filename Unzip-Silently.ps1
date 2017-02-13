<#
Unzip-Silently.ps1
#>


[CmdletBinding()]
Param (
    [Parameter(ValueFromPipeline=$true,
    ValueFromPipelineByPropertyName=$true,
    Mandatory=$true,
      HelpMessage="`r`nFilePath: Which zip file would you like to unzip? `r`n`r`nPlease enter a valid system filename ('FullPath'), which preferably includes the path to the zip file as well (a full path name of a file such as C:\Windows\archive.zip). `r`n`r`nNotes:`r`n`t- If no path is defined, the current directory gets searched for the filename. `r`n`t- If the full filename or the directory name includes space characters, `r`n`t   please enclose the whole inputted string in quotation marks (single or double). `r`n`t- To stop entering new values, please press [Enter] at an empty input row (and the script will run). `r`n`t- To exit this script, please press [Ctrl] + C`r`n")]
    [Alias("FilenameWithPathName","FullPath","Source","File","ZipFile","Zip")]
    [string[]]$FilePath,
    [Alias("OutputFolder")]
    [string]$Output,
    [Alias("IncludeZipFilesInTheFolderDefinedWithFilepathParameter","IncludesFolders","Folders","Folder")]
    [switch]$Include,
    [switch]$Recurse,
    [Alias("DeleteZip","DeleteOriginal","Delete")]
    [switch]$Purge
)




Begin {


    # Set the common parameters
    $computer = $env:COMPUTERNAME
    $ErrorActionPreference = "Stop"
    $start_time = Get-Date
    $empty_line = ""
    $files = @()
    $skipped = @()
    $zip_files = @()
    $skipped_path_names = @()
    $num_invalid_paths = 0


    # A function to calculate hash values in PowerShell versions 2 and 3
    # Requires .NET Framework v3.5
    # Example: dir C:\Temp | Check-FileHash
    # Example: Check-FileHash C:\Windows\explorer.exe -Algorithm SHA1
    # Source: http://poshcode.org/2154
    # Credit: Lee Holmes: "Windows PowerShell Cookbook (O'Reilly)" (Get-FileHash script) http://www.leeholmes.com/guide

    Function Check-FileHash {

        param(
        $hash_path,
        [ValidateSet("MD5","SHA1","SHA256","SHA384","SHA512","MACTripleDES","RIPEMD160")]
        $Algorithm = "SHA256"
        )

        $hash_files = @()

        # Create the hash value calculator
        # Source: http://stackoverflow.com/questions/21252824/how-do-i-get-powershell-4-cmdlets-such-as-test-netconnection-to-work-on-windows
        # Source: https://msdn.microsoft.com/en-us/library/system.security.cryptography.sha256cryptoserviceprovider(v=vs.110).aspx
        # Source: https://msdn.microsoft.com/en-us/library/system.security.cryptography.md5cryptoserviceprovider(v=vs.110).aspx
        # Source: https://msdn.microsoft.com/en-us/library/system.security.cryptography(v=vs.110).aspx
        # Source: https://msdn.microsoft.com/en-us/library/system.security.cryptography.mactripledes(v=vs.110).aspx
        # Source: https://msdn.microsoft.com/en-us/library/system.security.cryptography.ripemd160(v=vs.110).aspx
        # Credit: Twon of An: "Get the SHA1,SHA256,SHA384,SHA512,MD5 or RIPEMD160 hash of a file" https://community.spiceworks.com/scripts/show/2263-get-the-sha1-sha256-sha384-sha512-md5-or-ripemd160-hash-of-a-file
        If (($Algorithm -eq "MD5") -or ($Algorithm -like "SHA*")) {
            $typename = [string]"System.Security.Cryptography." + $Algorithm + "CryptoServiceProvider"
            $hasher = New-Object -TypeName $typename
        } ElseIf ($Algorithm -eq "MACTripleDES") {
            $hasher = New-Object -TypeName System.Security.Cryptography.MACTripleDES
        } ElseIf ($Algorithm -eq "RIPEMD160") {
            $hasher = [System.Security.Cryptography.HashAlgorithm]::Create("RIPEMD160")
        } Else {
            $continue = $true
        } # Else

                    # If a file name is specified, add that to the list of files to process
                    If ($hash_path) {
                        $hash_files += $hash_path
                    } Else {
                        # Take the files that are piped into the script
                        $hash_files += @($input | ForEach-Object { $_.FullName })
                    } # Else (If $hash_path)

        ForEach ($hash_file in $hash_files) {

                # Source: http://go.microsoft.com/fwlink/?LinkID=113418
                If ((Test-Path $hash_file -PathType Leaf) -eq $false) {

                    # Skip the item ($hash_file) if it is not a file (return to top of the program loop (ForEach $hash_file)
                    Continue

                } Else {

                    # Convert the item ($hash_file) to a fully-qualified path
                    $path_to_a_file = (Resolve-Path -Path $hash_file).Path

                    # Calculate the hash of the file regardless whether it is opened in another program or not
                    # Source: http://stackoverflow.com/questions/21252824/how-do-i-get-powershell-4-cmdlets-such-as-test-netconnection-to-work-on-windows
                    # Credit: Gisli: "Unable to read an open file with binary reader" http://stackoverflow.com/questions/8711564/unable-to-read-an-open-file-with-binary-reader
                    $source_file = [System.IO.File]::Open("$path_to_a_file",[System.IO.Filemode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
                    $hash = [System.BitConverter]::ToString($hasher.ComputeHash($source_file)) -replace "-",""
                    $source_file.Close()

                            # Return a custom object with the important details from the hashing
                            # Source: https://msdn.microsoft.com/en-us/library/system.io.path_methods(v=vs.110).aspx
                            $result = New-Object PsObject -Property @{
                                    'FileName'                      = ([System.IO.Path]::GetFileName($hash_file));
                                    'Path'                          = ([System.IO.Path]::GetFullPath($hash_file));
                                    'Directory'                     = ([System.IO.Path]::GetDirectoryName($hash_file));
                                  #  'Extension'                     = ([System.IO.Path]::GetExtension($hash_file));
                                  #  'FileNameWithoutExtension'      = ([System.IO.Path]::GetFileNameWithoutExtension($hash_file));
                                    'Algorithm'                     = $Algorithm;
                                    'Hash'                          = $hash
                            } # New-Object
                } # Else (Test-Path $hash_file)
            $result
        } # ForEach ($hash_file)
    } # function




    If ($Output) {

        # Test if the Output-path ("OutputFolder") exists
        If ((Test-Path $Output) -eq $false) {

            # Display an error message in console
            $empty_line | Out-String
            Write-Warning "The -Output folder value '$Output' doesn't seem to be a valid file system path."
            $empty_line | Out-String
            Write-Verbose "Please consider checking that the Output ('OutputFolder') location '$Output', where the unzipped files are ought to be written, was typed correctly and that it is a valid file system path, which points to a directory. If the path name includes space characters, please enclose the path name in quotation marks (single or double)." -verbose
            $empty_line | Out-String
            $skip_text = "Couldn't find the -Output folder '$Output'..."
            Write-Output $skip_text
            $empty_line | Out-String
            Exit

        } Else {
            # Test if the Output-path is a folder
            # Source: http://go.microsoft.com/fwlink/?LinkID=113418
            If ((Test-Path $Output -PathType Container) -eq $true) {
                # Resolve the Output-path ("ReportPath") (if the Output-path is specified as relative)
                $real_output_path = (Resolve-Path -Path $Output).Path
            } Else {
                $empty_line | Out-String
                $exit_text = "It seems that the -Output value of '$Output' points to a file (while a path to a folder, where the unzipped files are ought to be written, was the expected value for the -Output parameter...)"
                Write-Output $exit_text
                $empty_line | Out-String
                Exit
            } # Else (If Test-Path $Output -PathType)
        } # Else (If Test-Path $Output)
    } Else {
        $continue = $true
    } # Else (If $Output)




    # If a zip-file or a folder is specified, add those to the list of files to process
    # Source: http://poshcode.org/2154
    # Credit: Lee Holmes: "Windows PowerShell Cookbook (O'Reilly)" (Get-FileHash script) http://www.leeholmes.com/guide
    If ($FilePath) {

        ForEach ($path in $FilePath) {

            # Test if the path exists
            If ((Test-Path $path) -eq $false) {

                $invalid_filepath_was_found = $true

                # Increment the error counter
                $num_invalid_paths++

                # Display an error message in console
                $empty_line | Out-String
                Write-Warning "The zip-file path '$path' doesn't seem to be a valid FullPath or FilePath value."
                $empty_line | Out-String
                Write-Verbose "Please consider checking that the full zip filename with the path name (the '-FilePath' variable value) '$path' was typed correctly and that it includes the path to the zip-file as well. If the full filename or the directory name includes space characters, please enclose the whole string in quotation marks (single or double)." -verbose
                $empty_line | Out-String
                $skip_text = "Skipping '$path' from the files to be processed."
                Write-Output $skip_text

                    # Add the file candidate as an object (with properties) to a collection of skipped paths
                    $skipped += $obj_skipped = New-Object -TypeName PSCustomObject -Property @{

                                'Skipped FilePath Values'   = $path
                                'Owner'                     = ""
                                'Created on'                = ""
                                'Last Updated'              = ""
                                'Size'                      = "-"
                                'Error'                     = "The path was not found on $computer."
                                'raw_size'                  = 0

                        } # New-Object

                # Add the file candidate to a list of failed filenames
                $skipped_path_names += $path

                # Return to top of the program loop (ForEach $path) and skip just this iteration of the loop.
                Continue

            } Else {

                # Test the item ($path) to see whether it's a file or folder
                # Source: http://go.microsoft.com/fwlink/?LinkID=113418
                If ((Test-Path $path -PathType Leaf) -eq $false) {

                    $a_valid_path_to_a_folder = $true

                    If ($Include) {
                        # Search for zip files from the folder according to the user-set recurse option
                        # Source: http://poshcode.org/5668
                        $paths = Get-ChildItem '*.zip' -Path "$path" -Recurse:$Recurse -ErrorAction SilentlyContinue -Force
                        $files += @($paths | Foreach-Object { $_.FullName })
                    } Else {
                        # Skip the item ($path) if it is not a file (return to top of the program loop (ForEach $path)
                        $empty_line | Out-String
                        $exit_text = "It seems that the -FilePath value of '$path' points to a folder. Please consider adding the -Include parameter to the launching command in order to automatically extract the contents of all the zip files, which are located in the '$path' directory."
                        Write-Output $exit_text
                        Continue
                    } # Else (If $Recurse)

                } Else {

                    # Resolve the path of a file (if path is specified as relative)
                    $full_path = (Resolve-Path -Path $path).Path
                    $files += $full_path

                } # Else (If Test-Path $path -PathType)
            } # Else (Test-Path $path -eq)

        } # ForEach $path


    } Else {
        # Take the files that are piped into the script
        $files += @($input | Foreach-Object { $_.FullName })
    } # Else (If $FilePath)

} # Begin




Process {

    $empty_line | Out-String
    $timestamp = Get-Date -Format HH:mm:ss

    # Process each file
    # Source: https://msdn.microsoft.com/en-us/powershell/reference/5.1/microsoft.powershell.utility/new-object
    # Source: http://www.computerperformance.co.uk/powershell/powershell_com_shell.htm
    # Source: http://poshcode.org/5905
    If ($files.Count -ge 1) {

        # Initiate the unzipping session
        $shell = New-Object -ComObject "Shell.Application"
        $initiation_text = "$timestamp - Initiating the unzipping session..."
        Write-Output $initiation_text

        # Try to process one instance only once
        $unique_files = $files | select -Unique

        # Set the progress bar variables ($id denominates different progress bars, if more than one is being displayed)
        $activity           = "Extracting Zip Files"
        $status             = "Status"
        $task               = "Setting Initial Variables"
        $total_steps        = $unique_files.Count
        $task_number        = 0
        $id                 = 1

                            # Start the progress bar if there is more than one unique file to process
                            If (($unique_files.Count) -gt 1) {
                                Write-Progress -Id $id -Activity $activity -Status $status -CurrentOperation $task -PercentComplete ((0.000002 / $total_steps) * 100)
                            } # If ($unique_files.Count)


        ForEach ($file in $unique_files) {

                # Increment the step counter
                $task_number++

                            # Update the progress bar if there is more than one unique file to process
                            $activity = "Extracting Zip Files $task_number/$total_steps"
                            If (($unique_files.Count) -gt 1) {
                                Write-Progress -Id $id -Activity $activity -Status $status -CurrentOperation $file -PercentComplete (($task_number / $total_steps) * 100)
                            } # If ($unique_files.Count)

                # Determine the default output folder
                # Source: https://msdn.microsoft.com/en-us/library/system.io.path_methods(v=vs.110).aspx
                $new_folder_name = [System.IO.Path]::GetFileNameWithoutExtension($file)
                $existing_directory_path = [System.IO.Path]::GetDirectoryName($file)
                $filename = ([System.IO.Path]::GetFileName($file))
                $full_path = ([System.IO.Path]::GetFullPath($file))
                $extension = ([System.IO.Path]::GetExtension($file))

                            If ($Output) {
                                        $output_candidate = "$real_output_path\$new_folder_name"
                                        If ($output_candidate.Contains("\\")) {
                                            $output_folder = $output_candidate.Replace("\\","\")
                                        } Else {
                                            $output_folder = $output_candidate
                                        } # Else (If $output_candidate.Contains)
                            } Else {
                                $output_folder = "$existing_directory_path\$new_folder_name"
                            } # Else

                # Create an unique output folder with a generic directory name regardless whether or not a similar item already exists
                # Source: https://msdn.microsoft.com/en-us/library/system.io.directory_methods(v=vs.110).aspx
                # Source: http://www.winserverhelp.com/2010/04/powershell-tutorial-loops-for-foreach-while-do-while-do-until/2/
                If ((Test-Path $output_folder) -eq $false) {
                    New-Item -ItemType Directory -Path "$output_folder" -Force | Out-Null
                } Else {
                    $empty_line | Out-String
                    Write-Warning "The default output folder '$output_folder' already exists."

                            $i = 2
                            do { $candidate_folder = [string]$output_folder + '_' + $i ; $i++ }
                            until ([System.IO.Directory]::Exists($candidate_folder) -eq $false)

                    $output_folder = $candidate_folder
                    $empty_line | Out-String
                    Write-Verbose "The output folder changed to $candidate_folder." -Verbose
                    New-Item -ItemType Directory -Path "$output_folder" -Force | Out-Null
                } # Else (If Test-Path $output_folder)


                            $zip_files += $obj_zip = New-Object -TypeName PSCustomObject -Property @{
                                        'File'              = $filename
                                        'Directory'         = $existing_directory_path
                                        'Output Folder'     = $output_folder
                                        'Output_Folder'     = $output_folder
                                        'Extension'         = $extension
                                        'New Folder'        = $new_folder_name
                                        'New_Folder'        = $new_folder_name
                                        'Full Path'         = $full_path
                                        'Full_Path'         = $full_path
                                        'MD5'               = If (($PSVersionTable.PSVersion).Major -ge 4) {
                                                                    Get-FileHash $file -Algorithm MD5 | Select-Object -ExpandProperty Hash
                                                                } Else {
                                                                    Check-FileHash $file -Algorithm MD5 | Select-Object -ExpandProperty Hash
                                                                } # Else (If $PSVersionTable.PSVersion)
                                        'SHA256'            = If (($PSVersionTable.PSVersion).Major -ge 4) {
                                                                    Get-FileHash $file -Algorithm SHA256 | Select-Object -ExpandProperty Hash
                                                                } Else {
                                                                    Check-FileHash $file -Algorithm SHA256 | Select-Object -ExpandProperty Hash
                                                                } # Else (If $PSVersionTable.PSVersion)
                                } # New-Object

                # Unzip with PowerShell
                # CopyHere vOptions:    4    = Do not display a progress dialog box.
                #                       8    = Give the file being operated on a new name in a move, copy, or rename operation if a file with the target name already exists.
                #                       16   = Respond with "Yes to All" for any dialog box that is displayed.
                #                       256  = Display a progress dialog box but do not show the file names.
                #                       512  = Do not confirm the creation of a new directory if the operation requires one to be created.
                #                       1024 = Do not display a user interface if an error occurs.
                # $unzip_destination_folder.CopyHere($zip_package_obj.Items(),512)
                # Source: https://msdn.microsoft.com/en-us/library/windows/desktop/bb787866(v=vs.85).aspx
                # Source: http://serverfault.com/questions/18872/how-to-zip-unzip-files-in-powershell#201604
                # Source: https://gist.github.com/tcotav/6058400
                # Source: http://poshcode.org/5668
                $zip_package_obj = $shell.NameSpace("$file")
                $unzip_destination_folder = $shell.NameSpace("$output_folder")
                $unzip_destination_folder.CopyHere($zip_package_obj.Items())
                Start-Sleep -s 1
                $list = Get-ChildItem $output_folder


                # Delete the original zip-file if set to do so with the -Purge parameter
                If ($Purge) {
                    If (($list.Count) -ge 1) {
                        Remove-Item $file -Force
                    } Else {
                        Write-Verbose "$file - Something went wrong with unzip procedure."
                    } # Else
                } Else {
                    $continue = $true
                } # Else (If $Purge)

        } # ForEach $file

                            # Close the progress bar if it has been opened
                            If (($unique_files.Count) -gt 1) {
                                $task_number = $unique_files.Count
                                $task = "Finished extracting zip files."
                                Write-Progress -Id $id -Activity $activity -Status $status -CurrentOperation $task -PercentComplete (($task_number / $total_steps) * 100) -Completed
                            } # If ($unique_files.Count)

    } Else {
        $continue = $true
    } # Else (If $files.Count)
} # Process




End {
                # Do the background work for natural language
                If ($zip_files.Count -gt 1) { $item_text = "zip-files"; $verb = "were" } Else { $item_text = "zip-file"; $verb = "was" }
                $unique_folders = "$(($zip_files | select -ExpandProperty Directory -Unique) -join ', ')"
                $empty_line | Out-String

                # Write the operational stats in console
                If ($skipped_path_names.Count -eq 0) {
                            If ($zip_files.Count -eq 0) {
                                $stats_text = "Didn't unzip any files."
                            } ElseIf ($zip_files.Count -le 4) {
                                $stats_text = "Total $($zip_files.Count) $item_text processed at $unique_folders."
                            } Else {
                                $stats_text = "Total $($zip_files.Count) $item_text processed."
                            } # Else (If $zip_files.Count)
                    Write-Output $stats_text

                } Else {
                    # Display the skipped path names and write the operational stats in console
                    $skipped.PSObject.TypeNames.Insert(0,"Skipped FilePath Names")
                    $skipped_selection = $skipped | Select-Object 'Skipped FilePath Values','Size','Error' | Sort-Object 'Skipped FilePath Values'
                    $skipped_selection | Format-Table -auto
                            If ($num_invalid_paths -gt 1) {
                                If ($zip_files.Count -eq 0) {
                                    $stats_text = "There were $num_invalid_paths skipped FilePath values. Didn't unzip any files."
                                } ElseIf ($zip_files.Count -le 4) {
                                    $stats_text = "Total $($zip_files.Count) $item_text processed at $unique_folders. There were $num_invalid_paths skipped FilePath values."
                                } Else {
                                    $stats_text = "Total $($zip_files.Count) $item_text processed. There were $num_invalid_paths skipped FilePath values."
                                } # Else (If $zip_files.Count)
                            } Else {
                                If ($zip_files.Count -eq 0) {
                                    $stats_text = "One FilePath value was skipped. Didn't unzip any files."
                                } ElseIf ($zip_files.Count -le 4) {
                                    $stats_text = "Total $($zip_files.Count) $item_text processed at $unique_folders. One FilePath value was skipped."
                                } Else {
                                    $stats_text = "Total $($zip_files.Count) $item_text processed. One FilePath value was skipped."
                                } # Else (If $zip_files.Count)
                            } # Else (If $num_invalid_paths)
                    Write-Output $stats_text

                } # Else (If $skipped_path_names.Count)


    If ($zip_files.Count -ge 1) {


        # Write the hash values in console
        $zip_files.PSObject.TypeNames.Insert(0,"Unzipped Files")
        $zip_files_selection = $zip_files | select "File","Full Path",'Output Folder','MD5','SHA256'
        Write-Output $zip_files_selection | Format-List


        # Find out how long the script took to complete
        $end_time = Get-Date
        $runtime = ($end_time) - ($start_time)

            If ($runtime.Days -ge 2) {
                $runtime_result = [string]$runtime.Days + ' days ' + $runtime.Hours + ' h ' + $runtime.Minutes + ' min'
            } ElseIf ($runtime.Days -gt 0) {
                $runtime_result = [string]$runtime.Days + ' day ' + $runtime.Hours + ' h ' + $runtime.Minutes + ' min'
            } ElseIf ($runtime.Hours -gt 0) {
                $runtime_result = [string]$runtime.Hours + ' h ' + $runtime.Minutes + ' min'
            } ElseIf ($runtime.Minutes -gt 0) {
                $runtime_result = [string]$runtime.Minutes + ' min ' + $runtime.Seconds + ' sec'
            } ElseIf ($runtime.Seconds -gt 0) {
                $runtime_result = [string]$runtime.Seconds + ' sec'
            } ElseIf ($runtime.Milliseconds -gt 1) {
                $runtime_result = [string]$runtime.Milliseconds + ' milliseconds'
            } ElseIf ($runtime.Milliseconds -eq 1) {
                $runtime_result = [string]$runtime.Milliseconds + ' millisecond'
            } ElseIf (($runtime.Milliseconds -gt 0) -and ($runtime.Milliseconds -lt 1)) {
                $runtime_result = [string]$runtime.Milliseconds + ' milliseconds'
            } Else {
                $runtime_result = [string]''
            } # Else (If)

                If ($runtime_result.Contains(" 0 h")) {
                    $runtime_result = $runtime_result.Replace(" 0 h"," ")
                    } If ($runtime_result.Contains(" 0 min")) {
                        $runtime_result = $runtime_result.Replace(" 0 min"," ")
                        } If ($runtime_result.Contains(" 0 sec")) {
                        $runtime_result = $runtime_result.Replace(" 0 sec"," ")
                } # if ($runtime_result: first)

        # Display the runtime in console
        $timestamp_end = Get-Date -Format HH:mm:ss
        $end_text = "$timestamp_end - The unzipping session finished."
        Write-Output $end_text
        $empty_line | Out-String
        $runtime_text = "$($zip_files.Count) $item_text $verb unzipped in $runtime_result."
        Write-Output $runtime_text
        $empty_line | Out-String

    } Else {
        $empty_line | Out-String
    } # Else (If $zip_files.Count)
} # End




# [End of Line]


<#


   _____
  / ____|
 | (___   ___  _   _ _ __ ___ ___
  \___ \ / _ \| | | | '__/ __/ _ \
  ____) | (_) | |_| | | | (_|  __/
 |_____/ \___/ \__,_|_|  \___\___|


http://www.leeholmes.com/guide                                                                # Lee Holmes: "Windows PowerShell Cookbook (O'Reilly)" (Get-FileHash script)


  _    _      _
 | |  | |    | |
 | |__| | ___| |_ __
 |  __  |/ _ \ | '_ \
 | |  | |  __/ | |_) |
 |_|  |_|\___|_| .__/
               | |
               |_|
#>

<#

.SYNOPSIS
Unzips zip files to generically named new folders.

.DESCRIPTION
Unzip-Silently uses the Shell.Application to unzip files that are defined with
the -FilePath parameter. By default the -FilePath parameter accepts plain
filenames (then the current directory gets searched for the inputted filename)
or 'FullPath' filenames, which include the path to the file as well (such as
C:\Windows\archive.zip). If the -Include parameter is used, also paths to a folder
may be entered as -Filepath parameter values, and then all the zip files inside the
first directory level of the specified folder (as indicated by the common command
'dir' for example) are added to the list of files to be processed. Furthermore,
if the -Recurse parameter is used in adjunction with the -Include parameter in
the command launching Unzip-Silently, the search for zip files under the directory,
which is defined with the -Filepath parameter, is done recursively (i.e. the zip
files are searched from every subfolder level).

The naming principle of the new folders follows the original names of the zipped
files. The contents of the zip files are extracted to new folders, which are
created by default to the same folder, where each zip file is located. The default
output destination folder, under which the new folder(s) is/are created, may be
changed with the -Output parameter. When creating new folders Unzip-Silently
tries to preserve pre-existing content rather than overwrite any existing folders
(or files eventually), so if a folder seems to already exist, a similarly named
folder with a (possibly higher) number is created instead.

After the contents of the zip files has been extracted, the MD5 and SHA256 hash
values of the zip files (in machines that have PowerShell version 4 or later
installed with the inbuilt Get-FileHash cmdlet and in machines that are running
PowerShell version 2 or 3 by calling a Check-FileHash function, which is based
on Lee Holmes' Get-FileHash script in "Windows PowerShell Cookbook (O'Reilly)")
along with other performance related info are displayed.

To delete the original zip file(s), the parameter -Purge may be added to the
launching command. Please note that if any of the individual parameter values 
include space characters, the individual value should be enclosed in quotation 
marks (single or double) so that PowerShell can interpret the command correctly.

.PARAMETER FilePath
with aliases -FilenameWithPathName, -FullPath, -Source, -File, -ZipFile and -Zip.
The -FilePath parameter determines, which zip file(s) (or folders that might
contain zip file(s)) is/are selected for content extraction, and in essence define
the objects for Unzip-Silently.

By default the -FilePath parameter accepts plain filenames (then the current
directory gets searched for the inputted filename) or 'FullPath' filenames, which
include the path to the file as well (such as C:\Windows\archive.zip).
If the -Include parameter is used, also paths to a folder may be entered as
-Filepath parameter values, then all the zip files inside the first directory level
of the specified folder (as indicated by the common command 'dir' for example) are
added to the list of files to be processed. Furthermore, if the -Recurse parameter
is used in adjunction with the -Include parameter in the command launching
Unzip-Silently, the search for zip files under the directory, which is defined with
the -Filepath parameter, is done recursively (i.e. the zip files are searched from
every subfolder level).

To enter multiple zip files for content extraction, please separate each individual
entity with a comma. If the filename or the directory name includes space
characters, please enclose the whole string (the individual entity in question) in
quotation marks (single or double). It's not mandatory to write -FilePath in the
unzip command to invoke the -FilePath parameter, as is shown in the Examples below,
since Unzip-Silently is trying to decipher the inputted queries as good as it is
machinely possible within a 50 KB size limit. The -FilePath parameter also takes an
array of strings and objects could be piped to this parameter, too. If no value for
the -FilePath parameter is defined in the command launching Unzip-Silently, the user
will be prompted to enter a -FilePath value.

.PARAMETER Output
with an alias -OutputFolder. Specifies the folder, under which the new folder(s)
with the extracted zip file content is/are to be saved. For best results the value
should be a valid file system path, which points to a directory (for example
C:\Windows\). When creating new folders Unzip-Silently tries to preserve
pre-existing content rather than overwrite any existing folders (or files
eventually), so if a folder seems to already exist (under the defined -Output
folder), a similarly named folder with a (possibly higher) number is created
instead inside the directory indicated by the -Output parameter. If no value for
the -Output parameter is defined in the command launching Unzip-Silently, the
zip files are unzipped to new folders, which are created to the same folder,
where each zip file is located.

.PARAMETER Include
with aliases -IncludeZipFilesInTheFolderDefinedWithFilepathParameter,
-IncludesFolders, -Folders and -Folder. If the -Include parameter is added to the
command launching Unzip-Silently, also paths to a folder may be succesfully entered
as -Filepath parameter values: all the zip files inside the first directory level of
the specified folder (as indicated by the common command 'dir' for example) are
added to the list of files to be processed.

.PARAMETER Recurse
If the -Recurse parameter is used in adjunction with the -Include parameter in the
command launching Unzip-Silently, the search for zip files under the directory,
which is defined with the -Filepath parameter is done recursively (i.e. the zip
files are searched from every subfolder level).

.PARAMETER Purge
with aliases -DeleteZip, -DeleteOriginal and -Delete. If the -Purge parameter is
added to the command launching Unzip-Silently, the original zip file(s) is/are
deleted after the contents of the zip file(s) has been extracted.

.OUTPUTS
Unzips zip files.
If the -Purge parameter is added to the command launching Unzip-Silently,
the original zip file(s) will be deleted. Displays information about zip file
content extraction in console. For each zip file content extraction procedure
also a progress bar is shown in a separate window, which closes after the
extraction has been done. Another progress bar is also shown in console, if multiple
zip files are being processed.

.NOTES
Please note that all the parameters can be used in one unzip command and that each
of the parameters can be "tab completed" before typing them fully (by pressing the
[tab] key).

    Homepage:           https://github.com/auberginehill/unzip-silently
    Short URL:          http://tinyurl.com/zkg7s9l
    Version:            1.0

.EXAMPLE
./Unzip-Silently -FilePath archive.zip
Run the script. Please notice to insert ./ or .\ before the script name.
The current directory gets searched for the inputted filename ("archive.zip") and
the contents of the archive.zip would be extracted to the current directory
(where the "archive.zip" is located) under a newly created folder called "archive".

During the unzip procedure Unzip-Silently tries to preserve pre-existing content
rather than overwrite any existing folders (or files eventually), so if a folder
called "archive" seems to already exist, a similarly named folder with a number is
created instead (inside which the contents of "archive.zip" is extracted). Please
note, that the word -FilePath may be omitted in this example and that the -FilePath
value ("archive.zip") doesn't need to be enveloped in quotation marks, since it
doesn't contain any space characters.

.EXAMPLE
help ./Unzip-Silently -Full
Display the help file.

.EXAMPLE
./Unzip-Silently -FilePath "C:\Windows\explorer.zip" -Output "C:\Scripts"
Run the script and extract the contents of "C:\Windows\explorer.zip" to
"C:\Scripts\explorer". During the unzip procedure Unzip-Silently tries to preserve
pre-existing content rather than overwrite any existing folders (or files
eventually), so if a folder called "C:\Scripts\explorer" seems to already exist,
a similarly named folder with a number is created instead (inside which the
contents of "explorer.zip" is extracted). Please note, that the word -FilePath may
be omitted in this example and that the paths don't need to be enveloped in
quotation marks, because

    /Unzip-Silently C:\Windows\explorer.zip -Output C:\Scripts

will result in the same outcome.

.EXAMPLE
./Unzip-Silently C:\Users\Dropbox\, C:\dc01 -Include -Output C:\Scripts -Recurse
In this example "C:\Users\Dropbox\" and "C:\dc01" represent folders. Zip files under
every directory level of "C:\Users\Dropbox\" and "C:\dc01" are searched and the
contents of every found zip file is extracted under its own folder inside the
"C:\Scripts" folder.

.EXAMPLE
./Unzip-Silently -Source "C:\Windows\a certain archive.zip", "C:\Users\Dropbox\" -Folder -Purge
Will extract the contents of "C:\Windows\a certain archive.zip" under the folder
"C:\Windows\a certain archive". Will also look for zip files to process from the
first directory level of "C:\Users\Dropbox\" (as indicated by the common command
'dir C:\Users\Dropbox\' for example), and extracts the contents of every found zip
file under its own folder inside the "C:\Users\Dropbox\" directory. After the
contents of the zip file(s) has been extracted, the original zip file(s) is/are
deleted.

This command will work, because -Source is an alias of -FilePath and -Folder is
an alias of -Include. The -FilePath (a.k.a. -Source a.k.a. -FilenameWithPathName
a.k.a. -FullPath a.k.a. -File, a.k.a. -ZipFile, a.k.a. -Zip) variable value is
case-insensitive (as is most of the PowerShell), but since the zip filename contains
space characters, the whole string (entity) needs to be enveloped with quotation
marks. The -Source parameter may be left out from this command, since, for example,

    ./Unzip-Silently "c:\wINDOWs\A Certain Archive.zip", c:\users\dropbox -Folder -Purge

is the exact same command in nature.

.EXAMPLE
Set-ExecutionPolicy remotesigned
This command is altering the Windows PowerShell rights to enable script execution for
the default (LocalMachine) scope. Windows PowerShell has to be run with elevated rights
(run as an administrator) to actually be able to change the script execution properties.
The default value of the default (LocalMachine) scope is "Set-ExecutionPolicy restricted".


    Parameters:

    Restricted      Does not load configuration files or run scripts. Restricted is the default
                    execution policy.

    AllSigned       Requires that all scripts and configuration files be signed by a trusted
                    publisher, including scripts that you write on the local computer.

    RemoteSigned    Requires that all scripts and configuration files downloaded from the Internet
                    be signed by a trusted publisher.

    Unrestricted    Loads all configuration files and runs all scripts. If you run an unsigned
                    script that was downloaded from the Internet, you are prompted for permission
                    before it runs.

    Bypass          Nothing is blocked and there are no warnings or prompts.

    Undefined       Removes the currently assigned execution policy from the current scope.
                    This parameter will not remove an execution policy that is set in a Group
                    Policy scope.


For more information, please type "Get-ExecutionPolicy -List", "help Set-ExecutionPolicy -Full",
"help about_Execution_Policies" or visit https://technet.microsoft.com/en-us/library/hh849812.aspx
or http://go.microsoft.com/fwlink/?LinkID=135170.

.EXAMPLE
New-Item -ItemType File -Path C:\Temp\Unzip-Silently.ps1
Creates an empty ps1-file to the C:\Temp directory. The New-Item cmdlet has an inherent -NoClobber mode
built into it, so that the procedure will halt, if overwriting (replacing the contents) of an existing
file is about to happen. Overwriting a file with the New-Item cmdlet requires using the Force. If the
path name and/or the filename includes space characters, please enclose the whole -Path parameter value
in quotation marks (single or double):

    New-Item -ItemType File -Path "C:\Folder Name\Unzip-Silently.ps1"

For more information, please type "help New-Item -Full".

.LINK
http://www.leeholmes.com/guide
https://social.technet.microsoft.com/Forums/scriptcenter/en-US/6988d856-09ae-41c5-aa79-3d78a9e4d03a/powershell-use-shellapplication-to-zip-files?forum=ITCG
https://msdn.microsoft.com/en-us/powershell/reference/5.1/microsoft.powershell.utility/new-object
https://msdn.microsoft.com/en-us/library/system.io.directory_methods(v=vs.110).aspx
https://msdn.microsoft.com/en-us/library/system.io.path_methods(v=vs.110).aspx
https://technet.microsoft.com/en-us/library/ff730939.aspx
http://go.microsoft.com/fwlink/?LinkID=113418
http://www.winserverhelp.com/2010/04/powershell-tutorial-loops-for-foreach-while-do-while-do-until/2/
https://www.credera.com/blog/technology-insights/perfect-progress-bars-for-powershell/
http://serverfault.com/questions/18872/how-to-zip-unzip-files-in-powershell#201604
http://www.computerperformance.co.uk/powershell/powershell_com_shell.htm
https://gist.github.com/tcotav/6058400
http://poshcode.org/5905
http://poshcode.org/5668
http://poshcode.org/2154

#>

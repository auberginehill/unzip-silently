<!-- Visual Studio Code: For a more comfortable reading experience, use the key combination Ctrl + Shift + V
     Visual Studio Code: To crop the tailing end space characters out, please use the key combination Ctrl + A Ctrl + K Ctrl + X (Formerly Ctrl + Shift + X)
     Visual Studio Code: To improve the formatting of HTML code, press Shift + Alt + F and the selected area will be reformatted in a html file.
     Visual Studio Code shortcuts: http://code.visualstudio.com/docs/customization/keybindings (or https://aka.ms/vscodekeybindings)
     Visual Studio Code shortcut PDF (Windows): https://code.visualstudio.com/shortcuts/keyboard-shortcuts-windows.pdf

  _    _           _             _____ _ _            _   _       
 | |  | |         (_)           / ____(_) |          | | | |      
 | |  | |_ __  _____ _ __ _____| (___  _| | ___ _ __ | |_| |_   _ 
 | |  | | '_ \|_  / | '_ \______\___ \| | |/ _ \ '_ \| __| | | | |
 | |__| | | | |/ /| | |_) |     ____) | | |  __/ | | | |_| | |_| |
  \____/|_| |_/___|_| .__/     |_____/|_|_|\___|_| |_|\__|_|\__, |
                    | |                                      __/ |
                    |_|                                     |___/                                 -->


## Unzip-Silently.ps1

<table>
   <tr>
      <td style="padding:6px"><strong>OS:</strong></td>
      <td style="padding:6px">Windows</td>
   </tr>
   <tr>
      <td style="padding:6px"><strong>Type:</strong></td>
      <td style="padding:6px">A Windows PowerShell script</td>
   </tr>
   <tr>
      <td style="padding:6px"><strong>Language:</strong></td>
      <td style="padding:6px">Windows PowerShell</td>
   </tr>
   <tr>
      <td style="padding:6px"><strong>Description:</strong></td>
      <td style="padding:6px">Unzip-Silently uses the <code>Shell.Application</code> to unzip files that are defined with the <code>-FilePath</code> parameter. By default the <code>-FilePath</code> parameter accepts plain filenames (then the current directory gets searched for the inputted filename) or 'FullPath' filenames, which include the path to the file as well (such as <code>C:\Windows\archive.zip</code>). If the <code>-Include</code> parameter is used, also paths to a folder may be entered as <code>-Filepath</code> parameter values, and then all the zip files inside the first directory level of the specified folder (as indicated by the common command '<code>dir</code>' for example) are added to the list of files to be processed. Furthermore, if the <code>-Recurse</code> parameter is used in adjunction with the <code>-Include</code> parameter in the command launching Unzip-Silently, the search for zip files under the directory, which is defined with the <code>-FilePath</code> parameter, is done recursively (i.e. the zip files are searched from every subfolder level).
      <br />
      <br />The naming principle of the new folders follows the original names of the zipped files. The contents of the zip files are extracted to new folders, which are created by default to the same folder, where each zip file is located. The default output destination folder, under which the new folder(s) is/are created, may be changed with the <code>-Output</code> parameter. When creating new folders Unzip-Silently tries to preserve pre-existing content rather than overwrite any existing folders (or files eventually), so if a folder seems to already exist, a similarly named folder with a (possibly higher) number is created instead.
      <br />
      <br />After the contents of the zip files has been extracted, the MD5 and SHA256 hash values of the zip files (in machines that have PowerShell version 4 or later installed with the inbuilt <code>Get-FileHash</code> cmdlet and in machines that are running PowerShell version 2 or 3 by calling a Check-FileHash function, which is based on <strong>Lee Holmes</strong>' <dfn>Get-FileHash</dfn> <a href="http://poshcode.org/2154">script</a> in "<a href="http://www.leeholmes.com/guide">Windows PowerShell Cookbook (O'Reilly)</a>") along with other performance related info is displayed in console.
      <br />
      <br />To delete the original zip file(s), the parameter <code>-Purge</code> may be added to the launching command. Please note that if any of the individual parameter values include space characters, the individual value should be enclosed in quotation marks (single or double) so that PowerShell can interpret the command correctly.</td>
   <tr>
      <td style="padding:6px"><strong>Homepage:</strong></td>
      <td style="padding:6px"><a href="https://github.com/auberginehill/unzip-silently">https://github.com/auberginehill/unzip-silently</a>
      <br />Short URL: <a href="http://tinyurl.com/zkg7s9l">http://tinyurl.com/zkg7s9l</a></td>
   </tr>
   <tr>
      <td style="padding:6px"><strong>Version:</strong></td>
      <td style="padding:6px">1.0</td>
   </tr>
   <tr>
        <td style="padding:6px"><strong>Sources:</strong></td>
        <td style="padding:6px">
            <table>
                <tr>
                    <td style="padding:6px">Emojis:</td>
                    <td style="padding:6px"><a href="https://github.com/auberginehill/emoji-table">Emoji Table</a></td>
                </tr>
                <tr>
                    <td style="padding:6px">Lee Holmes:</td>
                    <td style="padding:6px"><a href="http://www.leeholmes.com/guide">Windows PowerShell Cookbook (O'Reilly)</a>: Get-FileHash <a href="http://poshcode.org/2154">script</a></td>
                </tr>
            </table>
        </td>
   </tr>
   <tr>
      <td style="padding:6px"><strong>Downloads:</strong></td>
      <td style="padding:6px">For instance <a href="https://raw.githubusercontent.com/auberginehill/unzip-silently/master/Unzip-Silently.ps1">Unzip-Silently.ps1</a>. Or <a href="https://github.com/auberginehill/unzip-silently/archive/master.zip">everything as a .zip-file</a>.</td>
   </tr>
</table>




### Screenshot

<ul><ul><ul>
<img class="screenshot" title="screenshot" alt="screenshot" height="80%" width="80%" src="https://raw.githubusercontent.com/auberginehill/unzip-silently/master/Unzip-Silently_2.png">
</ul></ul></ul>



### Parameters

<table>
    <tr>
        <th>:triangular_ruler:</th>
        <td style="padding:6px">
            <ul>
                <li>
                    <h5>Parameter <code>-FilePath</code></h5>
                    <p>with aliases <code>-FilenameWithPathName</code>, <code>-FullPath</code>, <code>-Source</code>, <code>-File</code>, <code>-ZipFile</code> and <code>-Zip</code>. The <code>-FilePath</code> parameter determines, which zip file(s) (or folders that might contain zip file(s)) is/are selected for content extraction, and in essence define the objects for Unzip-Silently.</p>
                    <p>By default the <code>-FilePath</code> parameter accepts plain filenames (then the current directory gets searched for the inputted filename) or 'FullPath' filenames, which include the path to the file as well (such as <code>C:\Windows\archive.zip</code>). If the <code>-Include</code> parameter is used, also paths to a folder may be entered as <code>-FilePath</code> parameter values – then all the zip files inside the first directory level of the specified folder (as indicated by the common command '<code>dir</code>' for example) are added to the list of files to be processed. Furthermore, if the <code>-Recurse</code> parameter is used in adjunction with the <code>-Include</code> parameter in the command launching Unzip-Silently, the search for zip files under the directory, which is defined with the <code>-FilePath</code> parameter, is done recursively (i.e. the zip files are searched from every subfolder level).</p>
                    <p>To enter multiple zip files (or folders that might contain zip file(s) for content extraction, please separate each individual entity with a comma. If the filename or the directory name includes space characters, please enclose the whole string (the individual entity in question) in quotation marks (single or double). It's not mandatory to write <code>-FilePath</code> in the unzip command to invoke the <code>-FilePath</code> parameter, as is shown in the Examples below, since Unzip-Silently is trying to decipher the inputted queries as good as it is machinely possible within a 50 KB size limit. The <code>-FilePath</code> parameter also takes an array of strings and objects could be piped to this parameter, too. If no value for the <code>-FilePath</code> parameter is defined in the command launching Unzip-Silently, the user will be prompted to enter a <code>-FilePath</code> value.</p>
                </li>
            </ul>
        </td>
    </tr>
    <tr>
        <th></th>
        <td style="padding:6px">
            <ul>
                <p>
                    <li>
                        <h5>Parameter <code>-Output</code></h5>
                        <p>with an alias <code>-OutputFolder</code>. Specifies the folder, under which the new folder(s) with the extracted zip file content is/are to be saved. For best results the <code>-Output</code> parameter value should be a valid file system path, which points to a existing directory (for example <code>C:\Windows\</code>). When creating new folders  (under the defined -Output folder) Unzip-Silently tries to preserve pre-existing content rather than overwrite any folders (or files eventually), so if a folder seems to already exist, a similarly named folder with a (possibly higher) number is created instead inside the directory indicated by the <code>-Output</code> parameter. If no value for the <code>-Output</code> parameter is defined in the command launching Unzip-Silently, the zip files are unzipped to new folders, which are created to the same folder, where each zip file is located.</p>
                    </li>
                </p>
                <p>
                    <li>
                        <h5>Parameter <code>-Include</code></h5>
                        <p>with aliases <code>-IncludeZipFilesInTheFolderDefinedWithFilepathParameter</code>, <code>-IncludesFolders</code>, <code>-Folders</code> and <code>-Folder</code>. If the <code>-Include</code> parameter is added to the command launching Unzip-Silently, also paths to a folder may be succesfully entered as <code>-FilePath</code> parameter values: all the zip files inside the first directory level of the specified folder (as indicated by the common command '<code>dir</code>' for example) are added to the list of files to be processed.</p>
                    </li>
                </p>
                <p>
                    <li>
                        <h5>Parameter <code>-Recurse</code></h5>
                        <p>If the <code>-Recurse</code> parameter is used in adjunction with the <code>-Include</code> parameter in the command launching Unzip-Silently, the search for zip files under the directory, which is defined with the <code>-FilePath</code> parameter is done recursively (i.e. the zip files are searched from every subfolder level).</p>
                    </li>
                </p>
                <p>
                    <li>
                        <h5>Parameter <code>-Purge</code></h5>
                        <p>with aliases <code>-DeleteZip</code>, <code>-DeleteOriginal</code> and <code>-Delete</code>. If the <code>-Purge</code> parameter is added to the command launching Unzip-Silently, the original zip file(s) is/are deleted after the contents of the zip file(s) has been extracted.</p>
                    </li>
                </p>                                                      
            </ul>
        </td>
    </tr>
</table>




### Outputs

<table>
    <tr>
        <th>:arrow_right:</th>
        <td style="padding:6px">
            <ul>
                <li>Unzips zip files.</li>
            </ul>
        </td>
    </tr>
    <tr>
        <th></th>
        <td style="padding:6px">
            <ul>
                <p>
                    <li>If the <code>-Purge</code> parameter is added to the command launching Unzip-Silently, the original zip file(s) will be deleted.</li>
                </p>
                <p>
                    <li>For each zip file content extraction procedure a progress bar is shown in a separate window, which closes after the extraction has been done. Another progress bar is also shown in console, if multiple zip files are being processed.</li>
                </p>                
            </ul>
        </td>
    </tr>
</table>




### Notes

<table>
    <tr>
        <th>:warning:</th>
        <td style="padding:6px">
            <ul>
                <li>Please note that all the parameters can be used in one unzip command and that each of the parameters can be "tab completed" before typing them fully (by pressing the <code>[tab]</code> key).</li>
            </ul>
        </td>
    </tr>
</table>





### Examples

<table>
    <tr>
        <th>:book:</th>
        <td style="padding:6px">To open this code in Windows PowerShell, for instance:</td>
   </tr>
   <tr>
        <th></th>
        <td style="padding:6px">
            <ol>
                <p>
                    <li><code>./Unzip-Silently -FilePath archive.zip</code><br />
                    Run the script. Please notice to insert <code>./</code> or <code>.\</code> before the script name. The current directory gets searched for the inputted filename ("<code>archive.zip</code>") and the contents of the <code>archive.zip</code> would be extracted to the current directory (where the "<code>archive.zip</code>" is located) under a newly created folder called "<code>archive</code>".
                    <br />
                    <br />During the unzip procedure Unzip-Silently tries to preserve pre-existing content rather than overwrite any existing folders (or files eventually), so if a folder called "<code>archive</code>" seems to already exist, a similarly named folder with a number is created instead (inside which the contents of "<code>archive.zip</code>" is extracted). Please note, that the word <code>-FilePath</code> may be omitted in this example and that the <code>-Filepath</code> value ("<code>archive.zip</code>") doesn't need to be enveloped in quotation marks, since it doesn't contain any space characters.</li>
                </p>
                <p>
                    <li><code>help ./Unzip-Silently -Full</code><br />
                    Display the help file.</li>
                </p>
                <p>
                    <li><code>./Unzip-Silently -FilePath "C:\Windows\explorer.zip" -Output "C:\Scripts"</code><br />
                    Run the script and extract the contents of "<code>C:\Windows\explorer.zip</code>" to <code>"C:\Scripts\explorer"</code>. During the unzip procedure Unzip-Silently tries to preserve pre-existing content rather than overwrite any existing folders (or files eventually), so if a folder called <code>"C:\Scripts\explorer"</code> seems to already exist, a similarly named folder with a number is created instead (inside which the contents of "<code>explorer.zip</code>" is extracted). Please note, that the word <code>-FilePath</code> may be omitted in this example and that the paths don't need to be enveloped in quotation marks, because
                    <br />
                    <br /><code>/Unzip-Silently C:\Windows\explorer.zip -Output C:\Scripts</code>
                    <br />
                    <br />will result in the same outcome.</li>
                </p>
                <p>
                    <li><code>./Unzip-Silently C:\Users\Dropbox\, C:\dc01 -Include -Output C:\Scripts -Recurse</code><br />
                    In this example "<code>C:\Users\Dropbox\</code>" and "<code>C:\dc01</code>" represent folders. Zip files under every directory level of "<code>C:\Users\Dropbox\</code>" and "<code>C:\dc01</code>" are searched and the contents of every found zip file is extracted under its own folder inside the "<code>C:\Scripts</code>" folder.</li>
                </p>
                <p>
                    <li><code>./Unzip-Silently -Source "C:\Windows\a certain archive.zip", "C:\Users\Dropbox\" -Folder -Purge</code><br />
                    Will extract the contents of "<code>C:\Windows\a certain archive.zip</code>" under the folder "<code>C:\Windows\a certain archive</code>". Will also look for zip files to process from the first directory level of "<code>C:\Users\Dropbox\</code>" (as indicated by the common command '<code>dir C:\Users\Dropbox\</code>' for example), and extracts the contents of every found zip file under its own folder inside the "<code>C:\Users\Dropbox\</code>" directory. After the contents of the zip file(s) has been extracted, the original zip file(s) is/are deleted.
                    <br />
                    <br />This command will work, because <code>-Source</code> is an alias of <code>-FilePath</code> and <code>-Folder</code> is an alias of <code>-Include</code>. The <code>-FilePath</code> (a.k.a. <code>-Source</code> a.k.a. <code>-FilenameWithPathName</code> a.k.a. <code>-FullPath</code> a.k.a. <code>-File</code>, a.k.a. <code>-ZipFile</code>, a.k.a. <code>-Zip</code>) variable value is case-insensitive (as is most of the PowerShell), but since the zip filename contains space characters, the whole string (entity) needs to be enveloped with quotation marks. The <code>-Source</code> parameter may be left out from this command, since, for example,
                    <br />
                    <br /><code>./Unzip-Silently "c:\wINDOWs\A Certain Archive.zip", c:\users\dropbox -Folder -Purge</code>
                    <br />
                    <br />is the exact same command in nature.</li>
                </p>
                <p>
                    <li><p><code>Set-ExecutionPolicy remotesigned</code><br />
                    This command is altering the Windows PowerShell rights to enable script execution for the default (LocalMachine) scope. Windows PowerShell has to be run with elevated rights (run as an administrator) to actually be able to change the script execution properties. The default value of the default (LocalMachine) scope is "<code>Set-ExecutionPolicy restricted</code>".</p>
                        <p>Parameters:
                                <ol>
                                    <table>
                                        <tr>
                                            <td style="padding:6px"><code>Restricted</code></td>
                                            <td style="padding:6px">Does not load configuration files or run scripts. Restricted is the default execution policy.</td>
                                        </tr>
                                        <tr>
                                            <td style="padding:6px"><code>AllSigned</code></td>
                                            <td style="padding:6px">Requires that all scripts and configuration files be signed by a trusted publisher, including scripts that you write on the local computer.</td>
                                        </tr>
                                        <tr>
                                            <td style="padding:6px"><code>RemoteSigned</code></td>
                                            <td style="padding:6px">Requires that all scripts and configuration files downloaded from the Internet be signed by a trusted publisher.</td>
                                        </tr>
                                        <tr>
                                            <td style="padding:6px"><code>Unrestricted</code></td>
                                            <td style="padding:6px">Loads all configuration files and runs all scripts. If you run an unsigned script that was downloaded from the Internet, you are prompted for permission before it runs.</td>
                                        </tr>
                                        <tr>
                                            <td style="padding:6px"><code>Bypass</code></td>
                                            <td style="padding:6px">Nothing is blocked and there are no warnings or prompts.</td>
                                        </tr>
                                        <tr>
                                            <td style="padding:6px"><code>Undefined</code></td>
                                            <td style="padding:6px">Removes the currently assigned execution policy from the current scope. This parameter will not remove an execution policy that is set in a Group Policy scope.</td>
                                        </tr>
                                    </table>
                                </ol>
                        </p>
                    <p>For more information, please type "<code>Get-ExecutionPolicy -List</code>", "<code>help Set-ExecutionPolicy -Full</code>", "<code>help about_Execution_Policies</code>" or visit <a href="https://technet.microsoft.com/en-us/library/hh849812.aspx">Set-ExecutionPolicy</a> or <a href="http://go.microsoft.com/fwlink/?LinkID=135170">about_Execution_Policies</a>.</p>
                    </li>
                </p>
                <p>
                    <li><code>New-Item -ItemType File -Path C:\Temp\Unzip-Silently.ps1</code><br />
                    Creates an empty ps1-file to the <code>C:\Temp</code> directory. The <code>New-Item</code> cmdlet has an inherent <code>-NoClobber</code> mode built into it, so that the procedure will halt, if overwriting (replacing the contents) of an existing file is about to happen. Overwriting a file with the <code>New-Item</code> cmdlet requires using the <code>Force</code>. If the path name and/or the filename includes space characters, please enclose the whole <code>-Path</code> parameter value in quotation marks (single or double):
                        <ol>
                            <br /><code>New-Item -ItemType File -Path "C:\Folder Name\Unzip-Silently.ps1"</code>
                        </ol>
                    <br />For more information, please type "<code>help New-Item -Full</code>".</li>
                </p>
            </ol>
        </td>
    </tr>
</table>




### Contributing

<p>Find a bug? Have a feature request? Here is how you can contribute to this project:</p>

 <table>
   <tr>
      <th><img class="emoji" title="contributing" alt="contributing" height="28" width="28" align="absmiddle" src="https://assets-cdn.github.com/images/icons/emoji/unicode/1f33f.png"></th>
      <td style="padding:6px"><strong>Bugs:</strong></td>
      <td style="padding:6px"><a href="https://github.com/auberginehill/unzip-silently/issues">Submit bugs</a> and help us verify fixes.</td>
   </tr>
   <tr>
      <th rowspan="2"></th>
      <td style="padding:6px"><strong>Feature Requests:</strong></td>
      <td style="padding:6px">Feature request can be submitted by <a href="https://github.com/auberginehill/unzip-silently/issues">creating an Issue</a>.</td>
   </tr>
   <tr>
      <td style="padding:6px"><strong>Edit Source Files:</strong></td>
      <td style="padding:6px"><a href="https://github.com/auberginehill/unzip-silently/pulls">Submit pull requests</a> for bug fixes and features and discuss existing proposals.</td>
   </tr>
 </table>




### www

<table>
    <tr>
        <th><img class="emoji" title="www" alt="www" height="28" width="28" align="absmiddle" src="https://assets-cdn.github.com/images/icons/emoji/unicode/1f310.png"></th>
        <td style="padding:6px"><a href="https://github.com/auberginehill/unzip-silently">Script Homepage</a></td>
    </tr>
    <tr>
        <th rowspan="16"></th>
        <td style="padding:6px">Lee Holmes: <a href="http://www.leeholmes.com/guide">Windows PowerShell Cookbook (O'Reilly)</a>: Get-FileHash <a href="http://poshcode.org/2154">script</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://social.technet.microsoft.com/Forums/scriptcenter/en-US/6988d856-09ae-41c5-aa79-3d78a9e4d03a/powershell-use-shellapplication-to-zip-files?forum=ITCG">PowerShell - use shell.application to zip files</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://msdn.microsoft.com/en-us/powershell/reference/5.1/microsoft.powershell.utility/new-object">New-Object</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://msdn.microsoft.com/en-us/library/system.io.directory_methods(v=vs.110).aspx">Directory Methods</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://msdn.microsoft.com/en-us/library/system.io.path_methods(v=vs.110).aspx">Path Methods</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://technet.microsoft.com/en-us/library/ff730939.aspx">Adding a Simple Menu to a Windows PowerShell Script</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="http://go.microsoft.com/fwlink/?LinkID=113418">Test-Path</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="http://www.winserverhelp.com/2010/04/powershell-tutorial-loops-for-foreach-while-do-while-do-until/2/">PowerShell Tutorial – Loops (For, ForEach, While, Do-While, Do-Until)</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://www.credera.com/blog/technology-insights/perfect-progress-bars-for-powershell/">Perfect Progress Bars for PowerShell</a></td>
    </tr>     
    <tr>
        <td style="padding:6px"><a href="http://serverfault.com/questions/18872/how-to-zip-unzip-files-in-powershell">How to zip/unzip files in Powershell?</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="http://www.computerperformance.co.uk/powershell/powershell_com_shell.htm">PowerShell Shell.Application To Launch Windows Explorer</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://gist.github.com/tcotav/6058400">unzip.ps1</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="http://poshcode.org/5905">Expand-ZipFile</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="http://poshcode.org/5668">Unzip Files</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="http://poshcode.org/2154">Get-FileHash.ps1</a></td>
    </tr> 
    <tr>
        <td style="padding:6px">ASCII Art: <a href="http://www.figlet.org/">http://www.figlet.org/</a> and <a href="http://www.network-science.de/ascii/">ASCII Art Text Generator</a></td>
    </tr>
</table>




### Related scripts

 <table>
    <tr>
        <th><img class="emoji" title="www" alt="www" height="28" width="28" align="absmiddle" src="https://assets-cdn.github.com/images/icons/emoji/unicode/0023-20e3.png"></th>
        <td style="padding:6px"><a href="https://gist.github.com/auberginehill/aa812bfa79fa19fbd880b97bdc22e2c1">Disable-Defrag</a></td>
    </tr>
    <tr>
        <th rowspan="25"></th>
        <td style="padding:6px"><a href="https://github.com/auberginehill/firefox-customization-files">Firefox Customization Files</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://github.com/auberginehill/get-ascii-table">Get-AsciiTable</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://github.com/auberginehill/get-battery-info">Get-BatteryInfo</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://github.com/auberginehill/get-computer-info">Get-ComputerInfo</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://github.com/auberginehill/get-culture-tables">Get-CultureTables</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://github.com/auberginehill/get-directory-size">Get-DirectorySize</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://github.com/auberginehill/get-hash-value">Get-HashValue</a></td>
    </tr>    
    <tr>
        <td style="padding:6px"><a href="https://github.com/auberginehill/get-installed-programs">Get-InstalledPrograms</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://github.com/auberginehill/get-installed-windows-updates">Get-InstalledWindowsUpdates</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://github.com/auberginehill/get-powershell-aliases-table">Get-PowerShellAliasesTable</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://gist.github.com/auberginehill/9c2f26146a0c9d3d1f30ef0395b6e6f5">Get-PowerShellSpecialFolders</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://github.com/auberginehill/get-ram-info">Get-RAMInfo</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://gist.github.com/auberginehill/eb07d0c781c09ea868123bf519374ee8">Get-TimeDifference</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://github.com/auberginehill/get-time-zone-table">Get-TimeZoneTable</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://github.com/auberginehill/get-unused-drive-letters">Get-UnusedDriveLetters</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://github.com/auberginehill/emoji-table">Emoji Table</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://github.com/auberginehill/java-update">Java-Update</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://github.com/auberginehill/remove-duplicate-files">Remove-DuplicateFiles</a></td>
    </tr>    
    <tr>
        <td style="padding:6px"><a href="https://github.com/auberginehill/remove-empty-folders">Remove-EmptyFolders</a></td>
    </tr>    
    <tr>
        <td style="padding:6px"><a href="https://gist.github.com/auberginehill/13bb9f56dc0882bf5e85a8f88ccd4610">Remove-EmptyFoldersLite</a></td>
    </tr> 
    <tr>
        <td style="padding:6px"><a href="https://gist.github.com/auberginehill/176774de38ebb3234b633c5fbc6f9e41">Rename-Files</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://github.com/auberginehill/rock-paper-scissors">Rock-Paper-Scissors</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://github.com/auberginehill/toss-a-coin">Toss-a-Coin</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://github.com/auberginehill/update-adobe-flash-player">Update-AdobeFlashPlayer</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://github.com/auberginehill/update-mozilla-firefox">Update-MozillaFirefox</a></td>
    </tr>
</table>

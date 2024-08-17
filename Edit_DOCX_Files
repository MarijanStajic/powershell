# Below is a PowerShell script for modifying a .docx file. To do this, you need to convert the file to a .zip archive and then extract it to access the document.xml file for content modification. 
# After making the changes, you re-compress it into a .zip file and rename it back to .docx. To avoid corrupting the file during the .zip conversion, it is necessary to use PowerShell v5 or earlier, with .NET-style methods. 
# In later versions, the conversion from .docx to .zip, extraction, re-compression, and renaming process can corrupt the file

# Defaults variables

$FirstName = "Marijan"
$Lastname = "Stajic"
$Initials = "MSA"
$Password = "MonM0tD4P@sse"

# .docx file path
$template = "C:\path\to\file.docx"

# Folder path
$templateFolder = "C:\path\to\temp_folder"

# Unzip/Zip with Pre PowerShell v5, .NET style
Add-Type -AssemblyName System.IO.Compression.FileSystem

# Unzip .docx file
[System.IO.Compression.ZipFile]::ExtractToDirectory($template, $templateFolder)

# Replace content
$bodyFile = $templateFolder + "\word\document.xml"
$body = Get-Content $bodyFile
$body = $body.Replace("[NAME1]", $FirstName)
$body = $body.Replace("[NAME2]", $Initials)
$body = $body.Replace("[NAME3]", $Password)
$body | Out-File $bodyFile -Force -Encoding Default

# Zip .docx file
[System.IO.Compression.ZipFile]::CreateFromDirectory($templateFolder, "C:\path\to\final_file.docx")

# Remove folder
Remove-Item $templateFolder -Recurse -Confirm:$false | Out-Null

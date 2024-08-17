# By default, in PowerShell, when trying to list the existing disks on the system, it is not possible to see the 'Recovery' partitions. 
# They are not displayed clearly. To delete this partition in a PowerShell script, I used DiskPart, which can see partitions named 'Recovery.' 
#I stored the information in a TXT file, and then PowerShell was able to read it, determine the partition number, and delete it.

# Get how many disks are on the system
$diskNumbers = (Get-Disk).Number

foreach ($diskNumber in $diskNumbers) {

    # Toggle for debug mode output
    $debugDiskpart = $false

    # Create a diskpart script to list partitions on the specified disk
    $diskpartScriptContent = "select disk $diskNumber `n list partition"

    # Write the diskpart script to a temporary file
    $tempFile = [System.IO.Path]::GetTempFileName()
    [System.IO.File]::WriteAllText($tempFile, $diskpartScriptContent)

    # Conditionally output the diskpart script file path
    if ($debugDiskpart) {

        Write-Output "Diskpart Script File Path :"
        Write-Output $tempFile
    }

    # Run the diskpart script and capture the output
    $diskpartOutput = diskpart /s $tempFile

    # Convert the output to an array of lines
    $lines = $diskpartOutput -split "`n"

    # Conditionally output the lines array
    if ($debugDiskpart) {

        Write-Output "Lines Array :"
        $lines | ForEach-Object { Write-Output $_ }
    }

    # Find the line containing the recovery partition
    $recoveryLine = $lines | Where-Object { $_ -match "Recovery" }

    # Check if the recovery line contains the word "Recovery"
    if ($recoveryLine -match "Recovery") {

        Write-Output "Recovery partition has been found on Disk $diskNumber. ERROR"
    else {

        #Write-Output "No Recovery partition was found on the Disk $diskNumber"
    }
}

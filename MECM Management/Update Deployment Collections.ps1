# This script is used to update bulk collections based on a list of collection and device names
# This can be useful when having to manually build deployment collections from fixed lists.
#
# Must load the Configuration Manager ISE Connect script prior to running this script
#


$CollectionList_CSV_Path = ($env:USERPROFILE + "\Desktop\SCCM\_CollectionList.csv")
        # Columns: CollectionName               
        # relationship list between devices and collections
        # This also serves as a check to ensure devices can only be added specific collections
        # If this is invalid you will be prompted for the file location

$CMRootCollection = "All Systems"                       
        # Used to limit the device query string.
        # This can speed up the script in large environments
        # This can be important if running as a user that is limited in what collections they have access to
        # If you don't have access to this or it does not exist, it will fail

$DeviceCollectionAssignmentList_CSV_Path = ($env:USERPROFILE + "\Desktop\SCCM\_DeviceCollectionAssignmentList.csv")
        #Columns: CollectionName,DeviceName               
        # This is the list that defines what devices will be added to what collection
        # If this is invalid you will be prompted for the file location

$OutputFilePath = ($env:USERPROFILE + "\Desktop\SCCM\")
        # Path the output log is stored

$ScriptDebug = $false
        # Put the script into debug mode (I added this way too late)

$ClearCollections = $false
        # Clear the listed collections before adding the devices?  true or false

$ValidationWaitTimer = 10
        # Set the wait timeout value
        # This as a mostly arbitrary number, make it bigger to increase wait times if your environment is large or slow 

#---------------------NO EDITABLE VARIABLES BELOW THIS LINE---------------------

$FileTimeNow = Get-Date

$HardStop = $false
        # Error condition that causes the entire script to stop in a controlled way

if ($ScriptDebug -eq $true)
{
    $OutputFileName = ("CollectionUpdateDebug.Log")
}
Else
{
    $OutputFileName = ("CollectionUpdate " + $FileTimeNow.ToUniversalTime().ToString("MMddyyHHmmss") + ".Log")
        # Build a unique file name
}

$OutputFile = ($OutputFilePath + $OutputFileName)
        # Create the output path

$HardStop = $false
        # Set and define the error code for "do not continue" but don't fail

Write-Output( "Output file location: " + $OutputFile)

New-Item -ItemType Directory -Force -Path $OutputFilePath

#---------------------Function List---------------------

#Simple Line break to console and log file
Function LineBreak
{
    if ($ScriptDebug -eq $true)
    {
        Write-Output " " | Tee-Object -FilePath $OutputFile -Append
    }
    else
    {
        Write-Output " " 
    }
}

# Check the status of collections being used to see if they are
# done updating so that they can be used or updated again
Function Wait-CollectionRefresh
{
    Write-Output "Waiting for Collection refresh" | Tee-Object -FilePath $OutputFile -Append
    get-date -Format "ddd yyyy/MM/dd HH:mm K" | Tee-Object -FilePath $OutputFile -Append

    For ($i=0; $i -le 100; $i++) 
    {
    
    Write-Progress -Activity "Waiting for Collections to update" -Status "Wait time" -PercentComplete ($i) -CurrentOperation ($i.ToString() + "%")
    Start-Sleep -Milliseconds ($ValidationWaitTimer*10)
    
    $CollectionRefresh = $false

        ForEach ($Collection in $CollectionList)
        {
            $CollectionStatus = Get-CMCollection -Name $Collection.CollectionName

            if ($CollectionStatus.CurrentStatus -ne "1")
            {
                $CollectionRefresh = $true
                #Write-Output ($Collection.CollectionName + " Refreshing " + $CollectionRefresh)
            }
            Else
            {
                if ($ScriptDebug -eq $true)
                {
                    Write-Output "$Collection Ready"
                }
            }
        }

        if ($CollectionRefresh -eq $false)
        {
            $i = 100
        }
        ElseIf ($CollectionRefresh -eq $false -and $i -eq 100)
        {
            Write-Output ("Warning: maximum wait time passed before all collections completed refresh")  | Tee-Object -FilePath $OutputFile -Append
            get-date -Format "ddd yyyy/MM/dd HH:mm K" | Tee-Object -FilePath $OutputFile -Append
        } 
        ElseIf ($CollectionRefresh -eq $false)
        {
            Write-Output "Completed updating device collections" | Tee-Object -FilePath $OutputFile -Append
            get-date -Format "ddd yyyy/MM/dd HH:mm K" | Tee-Object -FilePath $OutputFile -Append
        }
    }

    Write-Progress -Activity "Waiting for Collections to update" -Status "Wait time" -Completed
    
    Write-Output "Completed updating device collections" | Tee-Object -FilePath $OutputFile -Append
    get-date -Format "ddd yyyy/MM/dd HH:mm K" | Tee-Object -FilePath $OutputFile -Append
}
#---------------------Do Things!---------------------

LineBreak

# Validate the root collection prior to doing anything destructive
if (-not (Get-CMCollection -Name $CMRootCollection))
{
    Write-Output ("ERROR: You do not have access to ""$CMRootCollection"" or it does not exist.")
    $HardStop = $true
}

if ($ScriptDebug -eq $true)
{
    Write-Output ("Root collection set as: $CMRootCollection")
}

LineBreak

# Validate the data file paths

if (-not (Test-Path $CollectionList_CSV_Path))
{
    Write-Output "Collection list ($CollectionList_CSV_Path) is invalid." | Tee-Object -FilePath $OutputFile -Append
    $inputfile = Read-Host "Please enter the location of the valid collection list file"
    $CollectionList = Import-CSV $inputfile
    $CollectionList_CSV_Path = $inputfile

    if([string]::IsNullOrEmpty($CollectionList))
    {
        Write-Output ("INVALID FILE: " + $inputfile)
        $HardStop = $true
    }
}
Else
{
    Write-Output "Using script defined path: $CollectionList_CSV_Path" | Tee-Object -FilePath $OutputFile -Append
    $CollectionList = Import-CSV $CollectionList_CSV_Path 
        # Import collection validation list
}

LineBreak
$inputfile = @()

# Validate the Device assignment file

if (-not (Test-Path $DeviceCollectionAssignmentList_CSV_Path))
{
    Write-Output ("Device assignment list ($DeviceCollectionAssignmentList_CSV_Path) is invalid.") | Tee-Object -FilePath $OutputFile -Append
    $inputfile = Read-Host "Please enter the location of the valid collection list file"
    $DeviceCollectionAssignmentList = Import-CSV $inputfile
    $DeviceCollectionAssignmentList_CSV_Path = $inputfile

    if([string]::IsNullOrEmpty($CollectionList))
    {
        Write-Output ("INVALID FILE: " + $inputfile)
        $HardStop = $true
    }
}
Else
{
    Write-Output "Using script defined path: $DeviceCollectionAssignmentList_CSV_Path" | Tee-Object -FilePath $OutputFile -Append
    $DeviceCollectionAssignmentList = Import-CSV $DeviceCollectionAssignmentList_CSV_Path
        # Import the list of devices and what collections they should be added to
}

LineBreak
Write-Output "Using ""$CollectionList_CSV_Path"" to clear and validate collection data" | Tee-Object -FilePath $OutputFile -Append

Linebreak
Write-Output "Using ""$DeviceCollectionAssignmentList_CSV_Path"" to assign devices to collections" | Tee-Object -FilePath $OutputFile -Append

LineBreak

#---------------------Validate, and clear collections if required---------------------

Write-Output "Begin updating device collections" | Tee-Object -FilePath $OutputFile -Append
get-date -Format "ddd yyyy/MM/dd HH:mm K" | Tee-Object -FilePath $OutputFile -Append
LineBreak

Write-Output "Validating and clearing collections..." | Tee-Object -FilePath $OutputFile -Append

ForEach ($Collection in $CollectionList)
{
    $ValidCollection = Get-CMCollection -Name $collection.CollectionName
        
    if ($ValidCollection.count -ne 0)
    {
        if ($ScriptDebug -eq $true)
        {
            Write-Output ("Passed: " + $ValidCollection.Name)  | Tee-Object -FilePath $OutputFile -Append
        }
    }
    elseif ($ValidCollection.count -eq 0)
    {
        Write-OUtput ("FAILED: " + $Collection.CollectionName) | Tee-Object -FilePath $OutputFile -Append
        $HardStop = $true
    }
}

if ($ClearCollections -eq $true -and $HardStop -eq $false)
{
    Write-Output "   (Cleaning Collections)"
    LineBreak

    ForEach ($Collection in $CollectionList)
    {
        Write-Output $Collection.CollectionName  | Tee-Object -FilePath $OutputFile -Append
        $ValidCollection = Get-CMCollection -Name $collection.CollectionName
        $CollectionDirect = Get-CMCollectionDirectMembershipRule -CollectionName $collection.CollectionName
        Write-Output (" - Removing " + $CollectionDirect.count + " direct add devices...")  | Tee-Object -FilePath $OutputFile -Append
        
        ForEach ($DirectDevice in $CollectionDirect)
        {
            Write-Output ("  -  Removing " + $DirectDevice.RuleName)
            Remove-CMCollectionDirectMembershipRule -CollectionName $collection.CollectionName -ResourceName $DirectDevice.RuleName -Force
        }

        LineBreak

        # Output information that may be useful to why a device is or is not in a collection
        $CollectionInClude = Get-CMCollectionIncludeMembershipRule -CollectionName $collection.CollectionName
        write-output (" | Include memberships ignored: " + $CollectionInClude.Count) | Tee-Object -FilePath $OutputFile -Append

        $CollectionExclude = Get-CMCollectionExcludeMembershipRule -CollectionName $collection.CollectionName
        write-output (" | Exclude memberships ignored: " + $CollectionExclude.Count) | Tee-Object -FilePath $OutputFile -Append

        $CollectionQuery = Get-CMCollectionQueryMembershipRule -CollectionName $collection.CollectionName
        write-output (" | Query memberships rules ignored: " + $CollectionQuery.Count) | Tee-Object -FilePath $OutputFile -Append

        LineBreak

        if ($ValidCollection.count -ne 0)
        {
            if ($ScriptDebug -eq $true)
            {
                Write-Output ("Passed: " + $ValidCollection.Name)  | Tee-Object -FilePath $OutputFile -Append
            }

        }
        elseif ($ValidCollection.count -eq 0)
        {
            Write-Output ("FAILED: " + $Collection.CollectionName)  | Tee-Object -FilePath $OutputFile -Append
            $HardStop = $true
        }
    }
}
Elseif ($ClearCollections -eq $false)
{    
    Write-Output "   (Not cleaning collections)"
    LineBreak

    ForEach ($Collection in $CollectionList)
    {
        $ValidCollection = Get-CMCollection -Name $collection.CollectionName

        if ($ValidCollection.count -ne 0)
        {
            Write-Output ("Passed: " + $ValidCollection.Name)  | Tee-Object -FilePath $OutputFile -Append
        }
        elseif ($ValidCollection.count -eq 0)
        {
            Write-Output ("FAILED: " + $Collection.CollectionName)  | Tee-Object -FilePath $OutputFile -Append
            $HardStop = $true
        }
    }
}

If ($HardStop -eq $true)
{
    LineBreak
    write-output "Critical error occured, will not clear collections." | Tee-Object -FilePath $OutputFile -Append
}
Else
{
    Wait-CollectionRefresh
}

LineBreak

#---------------------Add devices to the identified collections---------------------

Write-Output "Adding Devices to Collections..." | Tee-Object -FilePath $OutputFile -Append

LineBreak 

if ($HardStop -eq $false)
{
    ForEach ($DeviceToAdd in $DeviceCollectionAssignmentList)
                                                                                                        {
    $DeviceToAdd_Collection = Get-cmcollection -Name $DeviceToAdd.CollectionName
    $DeviceToAdd_Device = Get-CMDevice -Name $DeviceToAdd.DeviceName -CollectionName $CMRootCollection
    $AddError = $false

        if ($DeviceToAdd_Device.Count -ne 1)
        {
            Write-Output ("ERROR: " + $DeviceToAdd.DeviceName + " | DEVICE did not return as valid") | Tee-Object -FilePath $OutputFile -Append
            $AddError = $true
        }

        if ($DeviceToAdd_Collection.Count -ne 1)
        {
            Write-Output ("ERROR: " + $DeviceToAdd.CollectionName + " | COLLECTION did not return as valid") | Tee-Object -FilePath $OutputFile -Append
            $AddError = $true
        }

        if ($DeviceToAdd.CollectionName -notin $CollectionList.CollectionName)
        {
            Write-Output ("ERROR: " + $DeviceToAdd.CollectionName + " | COLLECTION is not in approved list") | Tee-Object -FilePath $OutputFile -Append
            $AddError = $true
        }
   
        if ($AddError -eq $true) 
        {
            $FailedAddCount++
        }
        elseIf ($AddError -eq $false) 
        {
            if ($ScriptDebug -eq $true)
            {
                Write-Output ("  + " + $DeviceToAdd.DeviceName + " to " + $DeviceToAdd.CollectionName) | Tee-Object -FilePath $OutputFile -Append
            }
            Else
            {
                Write-Output ("  + " + $DeviceToAdd.DeviceName + " to " + $DeviceToAdd.CollectionName) 
            }
            Add-CMDeviceCollectionDirectMembershipRule -CollectionName $DeviceToAdd.CollectionName -ResourceID (Get-CMDevice -Name $DeviceToAdd.DeviceName).ResourceID 
            $SuccessAddCount++
        }
    }
}
ElseIf ($HardStop -eq $true)
{
    write-output "Critical error occured, will not add devices to collections." | Tee-Object -FilePath $OutputFile -Append
}

LineBreak
write-output ($SuccessAddCount.ToString() + " devices added to collections")  | Tee-Object -FilePath $OutputFile -Append
Write-Output ($FailedAddCount.ToString() + " devices failed to add to collections")  | Tee-Object -FilePath $OutputFile -Append

LineBreak

if  ($HardStop -eq $false)
{
    Wait-CollectionRefresh
}

LineBreak

# Check what devices eneded up in the collection.
# This exists to show what devices where excluded from the collection due to external factors
# This does display invalid ojects as well
If ($HardStop -eq $false)
{
    $i = 0
    ForEach ($DeviceToCheck in $DeviceCollectionAssignmentList)
    {
        $i++
        Write-Progress -Activity "Validating Devices..." -Status "Status:" -PercentComplete (($i/$DeviceCollectionAssignmentList.Count)*100) -CurrentOperation ($i.ToString() + "/" + $DeviceCollectionAssignmentList.Count + " processed")
        try
        {
            $DeviceCheck = Get-CMDevice -CollectionName $DeviceToCheck.CollectionName -Name $DeviceToCheck.DeviceName
            #if ($ScriptDebug -eq $true){Start-Sleep 1} #when you really need to slow things down...
        }
        catch
        {
            Write-Error "Something terrible happened!" | Out-Null
            Write-Error $_.exception.message
            $FailedVerificationCount++
        }
        if ($DeviceCheck.count -eq 0)
        {
            Write-Output ($DeviceToCheck.DeviceName + " failed verification check in collection " + $DeviceToCheck.CollectionName)
            $FailedVerificationCount++
        }
    }
}
ElseIf ($HardStop -eq $true)
{
    write-output "Critical error occured, will not validate devices." | Tee-Object -FilePath $OutputFile -Append
}

LineBreak
write-output ($FailedVerificationCount.ToString() + " devices failed verification.")  | Tee-Object -FilePath $OutputFile -Append

#---------------------Clean things up---------------------

$CollectionList = @()
$DeviceCollectionAssignmentList = @()
$FailedAddCount = 0
$SuccessAddCount = 0
$TempAddError = $false
$HardStop = $false
$FailedVerificationCount = 0

LineBreak
Write-Output "Completed updating device collections" | Tee-Object -FilePath $OutputFile -Append
get-date -Format "ddd yyyy/MM/dd HH:mm K" | Tee-Object -FilePath $OutputFile -Append

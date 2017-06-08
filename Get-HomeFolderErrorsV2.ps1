<#
.SYNOPSIS
	Find Home Folder errors on computers in Microsoft Active Directory.
.DESCRIPTION
	By default, builds a list of all computers where "Server" is in the 
	OperatingSystem property unless the ComputerName or InputFile 
	parameter is used.
	
	Process each server looking for Home Folder errors (Event ID 1060, Source TermService) 
	within, by default, the last 30 days.
	
	Builds a list of unique user names and servers unable to process.
	
	Creates the two text files, by default, in the folder where the script is run.
	
	Optionally, can specify the output folder.
	Unless the InputFile parameter is used, needs the ActiveDirectory module.
	
	The script has been tested with PowerShell versions 2, 3, 4, 5, and 5.1.
	The script has been tested with Microsoft Windows Server 2008 R2, 2012, 
	2012 R2, and 2016 and Windows 10 Creators Update.
.PARAMETER ComputerName
	Computer name used to restrict the computer search.
	Script surrounds ComputerName with "*".
	
	For example, if "RDS" is entered, the script uses "*RDS*".
	
	This allows the script to reduce the number of servers searched.
	
	If both ComputerName and InputFile are used, ComputerName is used to filter
	the list of computer names read from InputFile.
	
	Alias is CN
.PARAMETER InputFile
	Specifies the optional input file containing computer account names to search.
	
	Computer account names can be either the NetBIOS or Fully Qualified Domain Name.
	
	ServerName and ServerName.domain.tld are both valid.
	
	If both ComputerName and InputFile are used, ComputerName is used to filter
	the list of computer names read from InputFile.
	
	The computer names contained in the input file are not validated.
	
	Using this parameter causes the script to not check for or load the ActiveDirectory module.
	
	Alias is IF
.PARAMETER OrganizationalUnit
	Restricts the retrieval of computer accounts to a specific OU tree. 
	Must be entered in Distinguished Name format. i.e. OU=XenDesktop,DC=domain,DC=tld. 
	
	The script retrieves computer accounts from the top level OU and all sub-level OUs.
	
	Alias OU
.PARAMETER StartDate
	Start date, in MM/DD/YYYY format.
	Default is today's date minus 30 days.
.PARAMETER EndDate
	End date, in MM/DD/YYYY HH:MM format.
	Default is today's date.
.PARAMETER Folder
	Specifies the optional output folder to save the output reports. 
.EXAMPLE
	PS C:\PSScript > .\Get-HomeFolderErrorsV2.ps1
	
	Will return all Home Folders errors for the last 30 days for all
	computers in Active Directory that have "server" in the OperatingSystem
	property.
	
.EXAMPLE
	PS C:\PSScript > .\Get-HomeFolderErrorsV2.ps1 -StartDate "04/01/2017" -EndDate "04/02/2017" 
	
	Will return all Home Folder errors from "04/01/2017" through "04/02/2017".
.EXAMPLE
	PS C:\PSScript > .\Get-HomeFolderErrorsV2.ps1 -StartDate "04/01/2017" -EndDate "04/01/2017" 
	
	Will return all Home Folder errors from "04/01/2017" through "04/01/2017".
.EXAMPLE
	PS C:\PSScript > .\Get-HomeFolderErrorsV2.ps1 -Folder \\FileServer\ShareName

	Retrieves all Home Folder errors for the last 30 days.
	
	Output files will be saved in the path \\FileServer\ShareName
.EXAMPLE
	PS C:\PSScript > .\Get-HomeFolderErrorsV2.ps1 -ComputerName XEN
	
	Retrieves all Home Folder errors for the last 30 days.
	
	The script will only search computers with "XEN" in the DNSHostName.
	
.EXAMPLE
	PS C:\PSScript > .\Get-HomeFolderErrorsV2.ps1 -ComputerName RDS -Folder \\FileServer\ShareName -StartDate "05/01/2017" -EndDate "05/15/2017"
	
	The script will only search computers with "RDS" in the DNSHostName.
	
	Output file will be saved in the path \\FileServer\ShareName

	Will return all Home Folder errors from "05/01/2017" through "05/15/2017".
	
.EXAMPLE
	PS C:\PSScript > .\Get-HomeFolderErrorsV2.ps1 -ComputerName CTX -Folder \\FileServer\ShareName -StartDate "05/01/2017" -EndDate "05/07/2017" -InputFile c:\Scripts\computers.txt
	
	The script will only search computers with "CTX" in the entries contained in the computers.txt file.

	Output file will be saved in the path \\FileServer\ShareName

	Will return all Home Folder errors from "05/01/2017" through "05/07/2017".
	
	InputFile causes the script to not check for or use the ActiveDirectory module.

.EXAMPLE
	PS C:\PSScript > .\Get-HomeFolderErrorsV2.ps1 -OU "ou=RDS Servers,dc=domain,dc=tld"
	
	Finds Home Folder errors in all computers found in the "ou=RDS Servers,dc=domain,dc=tld" OU tree.
	
	Retrieves all Home Folder errors for the last 30 days.
	
.INPUTS
	None.  You cannot pipe objects to this script.
.OUTPUTS
	No objects are output from this script.  This script creates two text files.
.NOTES
	NAME: Get-HomeFolderErrorsV2.ps1
	VERSION: 2.00
	AUTHOR: Carl Webster, Sr. Solutions Architect for Choice Solutions, LLC
	LASTEDIT: May 18, 2017
#>


#Created by Carl Webster, CTP 24-Aug-2016
#Sr. Solutions Architect for Choice Solutions, LLC
#webster@carlwebster.com
#@carlwebster on Twitter
#http://www.CarlWebster.com
#
#Version 1.00 released to the community on 26-Aug-2016

[CmdletBinding(SupportsShouldProcess = $False, ConfirmImpact = "None", DefaultParameterSetName = "Default") ]

Param(
	[parameter(ParameterSetName="Default",Mandatory=$False)] 
	[Datetime]$StartDate = ((Get-Date -displayhint date).AddDays(-30)),

	[parameter(ParameterSetName="Default",Mandatory=$False)] 
	[Datetime]$EndDate = (Get-Date -displayhint date),
	
	[parameter(ParameterSetName="Default",Mandatory=$False)] 
	[Alias("CN")]
	[string]$ComputerName,
	
	[parameter(ParameterSetName="Default",Mandatory=$False)] 
	[Alias("IF")]
	[string]$InputFile="",
	
	[parameter(ParameterSetName="Default",Mandatory=$False)] 
	[Alias("OU")]
	[string]$OrganizationalUnit="",
	
	[parameter(ParameterSetName="Default",Mandatory=$False)] 
	[string]$Folder=""
	
	)

Set-StrictMode -Version 2
	
Write-Host "$(Get-Date): Setting up script"

If(![String]::IsNullOrEmpty($InputFile))
{
	Write-Host "$(Get-Date): Validating input file"
	If(!(Test-Path $InputFile))
	{
		Write-Error "Input file specified but $InputFile does not exist. Script cannot continue."
		Exit
	}
}

If($Folder -ne "")
{
	Write-Host "$(Get-Date): Testing folder path"
	#does it exist
	If(Test-Path $Folder -EA 0)
	{
		#it exists, now check to see if it is a folder and not a file
		If(Test-Path $Folder -pathType Container -EA 0)
		{
			#it exists and it is a folder
			Write-Host "$(Get-Date): Folder path $Folder exists and is a folder"
		}
		Else
		{
			#it exists but it is a file not a folder
			Write-Error "Folder $Folder is a file, not a folder. Script cannot continue"
			Exit
		}
	}
	Else
	{
		#does not exist
		Write-Error "Folder $Folder does not exist.  Script cannot continue"
		Exit
	}
}

#test to see if OrganizationalUnit is valid
If(![String]::IsNullOrEmpty($OrganizationalUnit))
{
	Write-Host "$(Get-Date): Validating Organnization Unit"
	try 
	{
		$results = Get-ADOrganizationalUnit -Identity $OrganizationalUnit
	} 
	
	catch
	{
		#does not exist
		Write-Error "Organization Unit $OrganizationalUnit does not exist.`n`nScript cannot continue`n`n"
		Exit
	}	
}

If($Folder -eq "")
{
	$pwdpath = $pwd.Path
}
Else
{
	$pwdpath = $Folder
}

If($pwdpath.EndsWith("\"))
{
	#remove the trailing \
	$pwdpath = $pwdpath.SubString(0, ($pwdpath.Length - 1))
}

Function Check-LoadedModule
#Function created by Jeff Wouters
#@JeffWouters on Twitter
#modified by Michael B. Smith to handle when the module doesn't exist on server
#modified by @andyjmorgan
#bug fixed by @schose
#bug fixed by Peter Bosen
#This Function handles all three scenarios:
#
# 1. Module is already imported into current session
# 2. Module is not already imported into current session, it does exists on the server and is imported
# 3. Module does not exist on the server

{
	Param([parameter(Mandatory = $True)][alias("Module")][string]$ModuleName)
	#following line changed at the recommendation of @andyjmorgan
	$LoadedModules = Get-Module |% { $_.Name.ToString() }
	#bug reported on 21-JAN-2013 by @schose 
	
	[string]$ModuleFound = ($LoadedModules -like "*$ModuleName*")
	If($ModuleFound -ne $ModuleName) 
	{
		$module = Import-Module -Name $ModuleName -PassThru -EA 0
		If($module -and $?)
		{
			# module imported properly
			Return $True
		}
		Else
		{
			# module import failed
			Return $False
		}
	}
	Else
	{
		#module already imported into current session
		Return $True
	}
}

#only check for the ActiveDirectory module if an InputFile was not entered
If([String]::IsNullOrEmpty($InputFile) -and !(Check-LoadedModule "ActiveDirectory"))
{
	Write-Host "Unable to run script, no ActiveDirectory module"
	Exit
}

Function ProcessComputer 
{
	Param([string]$TmpComputerName)
	
	If(Test-Connection -ComputerName $TmpComputerName -quiet -EA 0)
	{
		try
		{
			Write-Host "$(Get-Date): `tRetrieving Home Folder event log entries"
			$Errors = Get-EventLog -logname system `
			-computername $TmpComputerName `
			-source "TermService" `
			-entrytype "Error" `
			-after $StartDate `
			-before $EndDate `
			-EA 0
		}
		
		catch
		{
			Write-Host "$(Get-Date): `tServer $($TmpComputerName) had error being accessed"
			Out-File -FilePath $Filename2 `
			-Append -InputObject "Server $($TmpComputerName) had error being accessed $(Get-Date)"
			Continue
		}
		
		If($? -and $Null -ne $Errors)
		{
			$data = @()
			
			#errors found, now search for home folder entries
			ForEach($Item in $Errors)
			{
				If($Item.Message -like "*home directory*")
				{
					$obj = New-Object -TypeName PSObject
					$obj | Add-Member -MemberType NoteProperty -Name UserName -Value $Item.ReplacementStrings[0]
					$obj | Add-Member -MemberType NoteProperty -Name DomainName -Value $Item.ReplacementStrings[1]
					
					$data += $obj
				}
			}	

			If(($data | measure-object).count -eq 0)
			{
				#no home folder errors found in the list of errors found
				$Errors = @()
				$ErrorCount = 0
			}
			Else
			{
				$Errors = $data | Sort DomainName, UserName -Unique
				
				$ErrorCount = ($Errors | measure-object).count
			}
			
			Write-Host "$(Get-Date): `t$($ErrorCount) Home Folder errors found"
			
			If($ErrorCount -gt 0)
			{
				#only add to the master array if needed
				$Script:AllMatches += $Errors
				$Script:AllMatches = $Script:AllMatches | Sort UserName -Unique
			}
			
			$ErrorArrayCount = ($Script:AllMatches | measure-object).count
			Write-Host "$(Get-Date): `t`t$($ErrorArrayCount) total Home Folder errors found"
		}
		Else
		{
			Write-Host "$(Get-Date): `tNo Home Folder errors found"
		}
	}
	Else
	{
		Write-Host "`tComputer $($TmpComputerName) is not online"
		Out-File -FilePath $Filename2 -Append -InputObject "Computer $($TmpComputerName) was not online $(Get-Date)"
	}
}

$startTime = Get-Date

[string]$Script:FileName1 = "$($pwdpath)\HomeFolderErrors.txt"
[string]$Script:FileName2 = "$($pwdpath)\HFOfflineServers.txt"
#make sure files contain the current date only
Out-File -FilePath $Script:FileName1 -InputObject (Get-Date)
Out-File -FilePath $Script:FileName2 -InputObject (Get-Date)

If(![String]::IsNullOrEmpty($ComputerName) -and [String]::IsNullOrEmpty($InputFile))
{
	#computer name but no input file
	Write-Host "$(Get-Date): Retrieving list of computers from Active Directory"
	$testname = "*$($ComputerName)*"
	If(![String]::IsNullOrEmpty($OrganizationalUnit))
	{
		$Computers = Get-AdComputer -filter {DNSHostName -like $testname} -SearchBase $OrganizationalUnit -SearchScope Subtree -properties DNSHostName, Name -EA 0 | Sort Name
	}
	Else
	{
		$Computers = Get-AdComputer -filter {DNSHostName -like $testname} -properties DNSHostName, Name -EA 0 | Sort Name
	}
}
ElseIf([String]::IsNullOrEmpty($ComputerName) -and ![String]::IsNullOrEmpty($InputFile))
{
	#input file but no computer name
	Write-Host "$(Get-Date): Retrieving list of computers from Input File"
	$Computers = Get-Content $InputFile
}
ElseIf(![String]::IsNullOrEmpty($ComputerName) -and ![String]::IsNullOrEmpty($InputFile))
{
	#both computer name and input file
	Write-Host "$(Get-Date): Retrieving list of computers from Input File"
	$testname = "*$($ComputerName)*"
	$Computers = Get-Content $InputFile | ? {$_ -like $testname}
}
Else
{
	Write-Host "$(Get-Date): Retrieving list of computers from Active Directory"
	If(![String]::IsNullOrEmpty($OrganizationalUnit))
	{
		$Computers = Get-AdComputer -filter {OperatingSystem -like "*server*"} -SearchBase $OrganizationalUnit -SearchScope Subtree -properties DNSHostName, Name -EA 0 | Sort Name
	}
	Else
	{
		$Computers = Get-AdComputer -filter {OperatingSystem -like "*server*"} -properties DNSHostName, Name -EA 0 | Sort Name
	}
}

If($? -and $Null -ne $Computers)
{
	If($Computers -is [array])
	{
		Write-Host "Found $($Computers.Count) servers to process"
	}
	Else
	{
		Write-Host "Found 1 server to process"
	}

	$Script:AllMatches = @()
	[int]$Script:TotalErrorCount = 0
	[int]$Script:ErrorCount = 0

	If(![String]::IsNullOrEmpty($InputFile))
	{
		ForEach($Computer in $Computers)
		{
			$TmpComputerName = $Computer
			Write-Host "Testing computer $($TmpComputerName)"
			ProcessComputer $TmpComputerName
		}
	}
	Else
	{
		ForEach($Computer in $Computers)
		{
			$TmpComputerName = $Computer.DNSHostName
			Write-Host "Testing computer $($TmpComputerName)"
			ProcessComputer $TmpComputerName
		}
	}

	$Script:AllMatches = $Script:AllMatches | Sort UserName -Unique
	
	Write-Host "$(Get-Date): Output Home Folder errors to file"
	If($Script:TotalErrorCount -gt 0)
	{
		Out-File -FilePath $Script:FileName1 -Append -InputObject $Script:AllMatches
	}
	Else
	{
		Out-File -FilePath $Script:FileName1 -Append -InputObject "No Home Folder errors were found"
	}

	If(Test-Path "$($Script:FileName2)")
	{
		Write-Host "$(Get-Date): $($Script:FileName2) is ready for use"
	}
	If(Test-Path "$($Script:FileName1)")
	{
		Write-Host "$(Get-Date): $($Script:FileName1) is ready for use"
	}

	Write-Host "$(Get-Date): Script started: $($StartTime)"
	Write-Host "$(Get-Date): Script ended: $(Get-Date)"
	$runtime = $(Get-Date) - $StartTime
	$Str = [string]::format("{0} days, {1} hours, {2} minutes, {3}.{4} seconds", `
		$runtime.Days, `
		$runtime.Hours, `
		$runtime.Minutes, `
		$runtime.Seconds,
		$runtime.Milliseconds)
	Write-Host "$(Get-Date): Elapsed time: $($Str)"
	$runtime = $Null
}
Else
{
	If(![String]::IsNullOrEmpty($ComputerName) -and [String]::IsNullOrEmpty($InputFile))
	{
		#computer name but no input file
		Write-Host "Unable to retrieve a list of computers from Active Directory"
	}
	ElseIf([String]::IsNullOrEmpty($ComputerName) -and ![String]::IsNullOrEmpty($InputFile))
	{
		#input file but no computer name
		Write-Host "Unable to retrieve a list of computers from the Input File $InputFile"
	}
	ElseIf(![String]::IsNullOrEmpty($ComputerName) -and ![String]::IsNullOrEmpty($InputFile))
	{
		#computer name and input file
		Write-Host "Unable to retrieve a list of matching computers from the Input File $InputFile"
	}
	Else
	{
		Write-Host "Unable to retrieve a list of computers from Active Directory"
	}
}

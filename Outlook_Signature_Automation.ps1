#Requires -Version 5.1
<#
.SYNOPSIS
    Updates Email Signatures for On-Premesis and Online Exchange Users.

.DESCRIPTION
    Collects user$user data from various sources then updates the local outlook email signature and the Outlook Web App/Outlook Online email signature.

.PARAMETER FilterUsers
    Adds filter to what users will updated

.PARAMETER TestSignature
    Saves a test html file using the template

.PARAMETER TestUsers
    Outputs user information to a CSV file

.EXAMPLE
    .\\Outlook_Signature_Automation.ps1 -FilterUsers 'user1,user2@company.com,user3@company.com,user4' -TestUsers

    Runs the script but only outputs a CSV of the listed users' info
#>

param(
    [Parameter(HelpMessage="Enter Users by UPN or SAM separated by commas")]
    [string]$FilterUsers,

    [Parameter(HelpMessage="Will save an html file using the template to the folder the script is run from")]
    [switch]$TestSignature,

    [Parameter(HelpMessage="Will save a CSV of user information generated to the folder the script is run from")]
    [switch]$TestUsers
)

#region Classes

class Log {
    
    #This class is a basic logger that also outputs the logged info to the powershell console.

    [string]$LogFilePath

    Log([string]$LogFolderPath) {
        $this.LogFilePath = $LogFolderPath + (Get-Date -Format 'yyyy-MM-dd_HH_mm_ss') + '.log'
        New-Item -ItemType File -Path $this.LogFilePath -force | Out-Null
    }

    WriteLog([string]$message, [string]$level = 'OTHER') {
        $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        $logEntry = "$timestamp | [$level] $message"
        Add-Content -Path $this.LogFilePath -Value $logEntry
    }

    WriteInfo([string]$message) {
        $this.WriteLog($message,'INFO')
        Write-Host $message
    }

    WriteWarning([string]$message) {
        $this.WriteLog($message,'WARNING')
        Write-Host $message -ForegroundColor Yellow
    }

    WriteError([string]$message) {
        $this.WriteLog($message,'ERROR')
        Write-Host $message -ForegroundColor Red
    }
}

#endregion

#region Initalize Classes

#The code MyInvocation.MyCommand is just a reference to the running script.
$global:Log = New-Object Log -ArgumentList ([string]($MyInvocation.MyCommand.Path) -replace [string]($MyInvocation.MyCommand.Name),'Logs\')

#endregion

#region Variables

#This takes the (optional) list of users and creates an object of each depending on what way the users are entered.
$FilterList = @()
foreach ($FUser in $FilterUsers.Split(',')) {
    if ($FUser.Contains('@')) {
        $UPN = $FUser
        $SAM = ''
    } else {
        $UPN = ''
        $SAM = $FUser
    }
    $UserObject = [PSCustomObject]@{ UPN = $UPN; SAM = $SAM }
    $FilterList += $UserObject
}


$ConfigPath = ([string]($MyInvocation.MyCommand.Path) -replace [string]($MyInvocation.MyCommand.Name),'signatureconfig.json')

if (!(Test-Path -Path $ConfigPath)) {
    $global:Log.WriteError("Configuration file not found at $ConfigPath")
    Read-Host -Prompt "Press Enter To Exit"
    exit 1
}

$Config = Get-Content -Path $ConfigPath | ConvertFrom-Json

$Template = $Config.Template
$UseCSVData = $Config.UserCSVData.UseCSV
$UserCSVPath = $Config.UserCsvData.CSVPath
$ExchangeOnlineInfo = $Config.ExchangeOnline
$CompanyName = $Config.CompanyName
$ComputerPrefix = $Config.ComputerPrefix
$SearchBase = $Config.SearchBase
$FindComputers = $Config.FindComputers

#Boolean values to evaluate which Outlook types are used
$UseLocalOutlook = $Config.UseLocalOutlook
$UseWebOutlook = $Config.UseWebOutlook

#FilterXML for Get-WinEvent filter
#Function will replace {COMPUTERNAME}
$FilterXML = @"
<QueryList>
<Query Id="0" Path="Security">
<Select Path="Security">
*[System[(EventID=4624)]]
and
*[EventData[
Data[@Name='TargetUserName'] != 'SYSTEM' and
Data[@Name='TargetUserName'] != 'localadmin' and
Data[@Name='TargetUserName'] != '{COMPUTERNAME}' and
Data[@Name='TargetDomainName'] != 'Window Manager' and
Data[@Name='TargetDomainName'] != 'Font Driver Host' and
(Data[@Name='LogonType'] = 2 or Data[@Name='LogonType'] = 11) and
Data[@Name='SubjectDomainName'] != "Window Manager"
]
]
</Select>
</Query>
</QueryList>
"@

#endregion

#region Functions

function Coalesce {
    param ($a,$b,$c,$d,$e)

    if ($null -ne $a) { return $a }
    elseif ($null -ne $b) { return $b }
    elseif ($null -ne $c) { return $c }
    elseif ($null -ne $d) { return $d }
    elseif ($null -ne $e) { return $e }
    else { return $null }
}

function Set-TLSConfiguration {
    #The ExchangeOnlineManagement module requires that TLS12 is used, the line below forces the script to use this.
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
}

function Verify-Modules {
    if ((Get-Module -Name ExchangeOnlineManagement -ListAvailable).Count -eq 0) {
        Install-Module ExchangeOnlineManagement -AllowClobber -Force
    }
    Import-Module -Name ActiveDirectory
    if ((Get-Module -Name ActiveDirectory -ListAvailable).Count -eq 0) {
        Install-Module ActiveDirectory -AllowClobber -Force
    }
    Import-Module -Name ActiveDirectory
}

function Connect-TenentExchange {
    param ($ExchangeOnlineInfo)

    if ($null -eq $ExchangeOnlineInfo.ApplicationID -or
    $null -eq $ExchangeOnlineInfo.CertificateThumbprint -or
    $null -eq $ExchangeOnlineInfo.CertificatePath -or
    $null -eq $ExchangeOnlineInfo.TenetName -or
    $ExchangeOnlineInfo.ApplicationID -eq '' -or
    $ExchangeOnlineInfo.CertificateThumbprint -eq '' -or
    $ExchangeOnlineInfo.CertificatePath -eq '' -or
    $ExchangeOnlineInfo.TenetName -eq '') {
        $global:Log.WriteError('Missing Information in ExchangeOnline config')
        Read-Host -Prompt "Press Enter To Exit"
        exit 1
    }

    #Gets the certificate object from the certificate store defined in the config file.
    $Certificate = Get-ChildItem "$($ExchangeOnlineInfo.CertificatePath)$($ExchangeOnlineInfo.CertificateThumbprint)"

    try {
        Connect-ExchangeOnline -AppId $ExchangeOnlineInfo.ApplicationID -Certificate $Certificate -Organization $ExchangeOnlineInfo.TenetName -ShowBanner:$false -ErrorAction Stop
    } catch {
        $e = $_.Exception
        $msg = $e.Message
        $global:Log.WriteError("Exception $e")
        #$global:Log.WriteError("Error Message: $msg")
    }
    if ((Get-ConnectionInformation | Select-Object Name | Where-Object {$_.Name -like 'ExchangeOnline*'} | Measure).Count -gt 0) {
        $global:Log.WriteInfo('Connection to Exchange Online Successful')
        return $true
    } else {
        $global:Log.WriteError('Connection to Exchange Online Failed')
        return $false
    }
}

function Disconnect-TenetExchange {
    if ((Get-ConnectionInformation | Select-Object Name | Where-Object {$_.Name -like 'ExchangeOnline*'} | Measure).Count -gt 0) {
        Disconnect-ExchangeOnline -Confirm:$false
        $global:Log.WriteInfo('Disconnected from Exchange Online')
    }
}

function Verify-OutlookInstall {
    param($ComputerName)

    $OutlookInstallPath = 'Software\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE'

    try {
        $base = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, $ComputerName)
        $key = $base.OpenSubKey($OutlookInstallPath)
        if (!$key) {
            $base.Close()
            $global:Log.WriteWarning("Outlook not installed on $ComputerName")
            $result = $false
        } else {
            $key.Close()
            $base.Close()
            $global:Log.WriteInfo("Outlook installed on $ComputerName")
            $result = $true
        }
    } catch [UnauthorizedAccessException],[System.IO.IOException] {
        $global:Log.WriteError("Cannot Access Computer $ComputerName")
        $result = $false
    }

    return $result
}

function Get-ComputerLastUser {
    param($ComputerName, $FilterXML)

    $FilterXML = $FilterXML -replace '{COMPUTERNAME}',$ComputerName

    try {
        $LastUser = Invoke-Command -ComputerName $ComputerName -ErrorAction Stop -ScriptBlock {
            $Events = Get-WinEvent -FilterXml $using:FilterXML -MaxEvents 1
            foreach ($Event in $Events) {
                [XML]$XMLEvent = $Event.ToXML()
                $XMLEvent.Event.SelectSingleNode("//*[@Name='TargetUserName']").'#text'
            }
        }
    } catch {
        $global:Log.WriteError("Event retrieval failed for $ComputerName")
        $LastUser = 'FAILED'
    }

    return $LastUser
}

function Get-ComputerCurrentUser {
    param($ComputerName)

    try {
        $CurrentUser = (Get-CimInstance -ClassName Win32_ComputerSystem -ComputerName $ComputerName -ErrorAction Stop).UserName
        if ($null -eq $CurrentUser -or $CurrentUser -eq '') {
            $CurrentUser = 'NONE'
        } else {
            $CurrentUser = $CurrentUser.Split('\\')[1]
        }
    } catch {
        $global:Log.WriteError("Current user failed for $ComputerName")
        $CurrentUser = 'FAILED'
    }

    return $CurrentUser
}

function Get-Computers {
    param ($FilterXML,$ComputerPrefix)

    $global:Log.WriteInfo('Getting Computers')

    $recentLogonDate = (Get-Date).AddDays(-14)
    $Computers = Get-ADComputer -Filter { Name -like "$ComputerPrefix*" -and lastLogonTimeStamp -ge $recentLogonDate } | Select-Object Name | Sort-Object Name

    if (($Computers | Measure).Count -eq 0) {
        $global:Log.WriteWarning('No computers found in domain')
    } else {
        $global:Log.WriteInfo("$(($Computers | Measure).Count) computers found in domain")
    }

    $OutlookComputers = @()

    foreach ($Computer in $Computers) {
        if (Test-Connection -Count 1 -ComputerName $Computer.Name -Quiet) {
            if (Verify-OutlookInstall -ComputerName $Computer.Name) {
                $LastUser = Get-ComputerLastUser -ComputerName $Computer.Name -FilterXML $FilterXML
                $CurrentUser = Get-ComputerCurrentUser -ComputerName $Computer.Name
                $ComputerObj = [PSCustomObject]@{Name = $Computer.Name; CurrentUser = $CurrentUser; LastUser = $LastUser}
                $OutlookComputers += $ComputerObj
            }
        } else {
            $global:Log.WriteError("Cannot connect to Computer $($Computer.Name)")
        }  
    }

    return $OutlookComputers
}

function Get-LocalUsers {
    param ($SearchBase)

    $global:Log.WriteInfo('Getting Local Users')

    $LocalUsers = @()

    $UserList = Get-ADUser -Filter 'Enabled -eq $true' -SearchBase $SearchBase -Properties DisplayName,SID,SamAccountName,UserPrincipalName,Title | Select-Object -Property DisplayName,SID,SamAccountName,UserPrincipalName,Title | Sort-Object DisplayName

    if (($UserList | Measure).Count -eq 0) {
        $global:Log.WriteWarning("No local users found in SearchBase $SearchBase")
    } else {
        foreach ($User in $UserList) {
            $UserObject = [PSCustomObject]@{UPN = $user.UserPrincipalName; Name = $user.DisplayName; Title = $user.Title; SID = $user.SID.Value; SAM = $user.SamAccountName}
            $LocalUsers += $UserObject
        }
        $global:Log.WriteInfo("$(($LocalUsers | Measure).Count) users found")
    }

    return $LocalUsers
}

function Get-WebUsers {

    $global:Log.WriteInfo('Getting Web Users')

    $WebUsers = @()
    
    $UserList = $usersinfo = Get-User -Filter "RecipientType -eq 'UserMailBox' -and RecipientTypeDetails -eq 'UserMailBox'" | Select-Object -Property UserPrincipalName,DisplayName,Title | Sort-Object DisplayName
    
    if (($UserList | Measure).Count -eq 0) {
        $global:Log.WriteWarning("No web users found")
    } else {
        foreach ($User in $UserList) {
            $UserObject = [PSCustomObject]@{UPN = $user.UserPrincipalName; Name = $user.DisplayName; Title = $user.Title}
            $WebUsers += $UserObject
        }
        $global:Log.WriteInfo("$(($WebUsers | Measure).Count) users found")
    }
    
    return $WebUsers
}

function Get-CSVUsers {
    param ($CSVPath)

    $global:Log.WriteInfo('Getting CSV Users')

    $CSVUsers = @()

    if (Test-Path -Path $CSVPath) {
        $UserList = Import-CSV -Path $CSVPath | Sort-Object Name

        foreach ($User in $UserList) {
            $UserObject = [PSCustomObject]@{UPN = $user.UPN; Name = $user.Name; Title = $user.Title; SAM = $user.SAM; Computer = $user.Computer}
            $CSVUsers += $UserObject
        }
    } else {
        $global:Log.WriteError("CSV not found at path $CSVPath")
    }

    return $CSVUsers
}

function Merge-UserData {

    #This function attempts to consolidate user information from various sources in order to make the most complete list possible.
    #The order of priority is the User Data CSV, then the Local User data, then the Web User Data. If any fields are missing from one, it will use the next.

    param($LocalUsers,$WebUsers,$CSVUsers,$Computers)

    $global:Log.WriteInfo('Merging User Data')

    $MergedUsers = @()

    if (!($null -eq $LocalUsers -and $null -eq $WebUsers -and $null -eq $CSVUsers)) {
        if ($null -ne $LocalUsers) { $LocalCount = ($LocalUsers | Measure).Count } else { $LocalCount = 0 }
        if ($null -ne $WebUsers) { $WebCount = ($WebUsers | Measure).Count } else { $WebCount = 0 }
        if ($null -ne $CSVUsers) { $CSVCount = ($CSVUsers | Measure).Count } else { $CSVCount = 0 }
        
        $MostUsers = (Get-Variable -Name LocalCount,WebCount,CSVCount | Sort Value | Select -Last 1).Name
        $BaseUsers = (Get-Variable -Name $(switch ($MostUsers) {"LocalCount" {"LocalUsers";Break} "WebCount" {"WebUsers";Break} "CSVCount" {"CSVUsers";Break} }) -ValueOnly).Name
        foreach ($User in $BaseUsers) {
            $LUser = $LocalUsers | Where-Object { $_.Name -eq $User }
            $WUser = $WebUsers | Where-Object { $_.Name -eq $User }
            $CSVUser = $CSVUsers | Where-Object { $_.Name -eq $User }

            $UPN = Coalesce $CSVUser.UPN $LUser.UPN $WUser.UPN
            $Name = $User
            $Title = Coalesce $CSVUser.Title $LUser.Title $WUser.Title ''
            $SID = Coalesce $CSVUser.SID $LUser.SID ''
            $SAM = Coalesce $CSVUser.SAM $LUser.SAM ''

            $ComputerUser = $Computer | Where-Object { $_.CurrentUser -eq $SAM -or $_.LastUser -eq $SAM }

            $Computer = Coalesce $CSVUser.Computer $ComputerUser.Name

            $UserObject = [PSCustomObject]@{UPN = $UPN; Name = $Name; Title = $Title; SID = $SID; SAM = $SAM; Computer = $Computer; UserSignature = ''}
            $MergedUsers += $UserObject
        }
    } else {
        $global:Log.WriteError("All user lists empty")
    }

    return $MergedUsers
}

function Filter-Users {

    #This function uses the list of users from the script parameters to filter down the merged user list.

    param($AllUsers,$FilterUsers)

    if ($null -ne $FilterUsers -and $FilterUsers -ne '') {
        $UPNUsers = Compare-Object -ReferenceObject $AllUsers -DifferenceObject $FilterUsers -IncludeEqual -Property UPN | Where-Object {$_.SideIndicator -eq '==' -and $_.UPN -ne '' -and $null -ne $_.UPN}
        $SAMUsers = Compare-Object -ReferenceObject $AllUsers -DifferenceObject $FilterUsers -IncludeEqual -Property SAM | Where-Object {$_.SideIndicator -eq '==' -and $_.SAM -ne '' -and $null -ne $_.SAM}
        $FilteredUsers = @()
        foreach ($UPNUser in $UPNUsers) {
            $UserObject = $AllUsers | Where-Object { $_.UPN -eq $UPNUser.UPN }
            $FilteredUsers += $UserObject
        }
        foreach ($SAMUser in $SAMUsers) {
            $UserObject = $AllUsers | Where-Object { $_.SAM -eq $SAMUser.SAM }
            $FilteredUsers += $UserObject
        }
    } else {
        $FilteredUsers = $AllUsers
    }

    return $FilteredUsers
}

function Create-UserSignature {
    param($TemplatePath, $User, $TemplateName, $ImageName, $CompanyName)

    #The variables below are the placeholders used in the html file.
    $fName = '{NAME}'
    $fTitle = '{TITLE}'
    $fEmail = '{EMAIL}'
    $fImage = '{IMAGE}'
    $fB64 = '{B64}'

    $LocalImageHTML = "<img width=154 height=137 src=`"data:image/jpeg;base64,{B64}`" alt=`"$CompanyName`">"
    $WebImageHTML = "<img width=154 height=137 src=`"`" alt=`"$CompanyName`"><span id=`"dataURI`" style=`"display:none`">data:image/jpeg;base64,{B64}</span>"

    $Template = Get-Content -Path "$TemplatePath$TemplateName" -RAW
    
    $UserSignature = $Template -replace $fName,$User.Name -replace $fTitle,$User.Title -replace $fEmail,$User.UPN

    if ($null -ne $ImageName) {
        #The below code converts the image to Base64 so that the signature doesn't need to reference files. This is the only way the web signature images can function.
        if ($PSVersionTable.PSVersion.Major -ge 7) {
            $Image64 = [Convert]::ToBase64String((Get-Content -Path $TemplatePath$ImageName -AsByteStream))
        } else {
            $Image64 = [Convert]::ToBase64String((Get-Content -Path $TemplatePath$ImageName -Encoding Byte))
        }

        $LocalSignature = $UserSignature -replace $fImage,$LocalImageHTML -replace $fB64,$Image64
        $WebSignature = $UserSignature -replace $fImage,$WebImageHTML -replace $fB64,$Image64
    } else {
        $LocalSignature = $UserSignature
        $WebSignature = $UserSignature
    }

    return [PSCustomObject]@{LocalSignature = $LocalSignature; WebSignature = $WebSignature}
}

function Update-LocalSignature {
    param ($User,$CompanyName)

    $UserID = $User.SAM
    $UPN = $User.UPN
    $ComputerName = $User.Computer
    $LocalSig = $User.UserSignature.LocalSignature

    $global:Log.WriteInfo("Updating Local Signature for $UserID on $ComputerName")

    if ($null -eq $UserID -or $UserID -eq '' -or
    $null -eq $UPN -or $UPN -eq '' -or
    $null -eq $ComputerName -or $ComputerName -eq '' -or
    $null -eq $LocalSig -or $LocalSig -eq '') {
        $global:Log.WriteError('Missing Information, skipping User')
    } else {
        $LocalSignatureFolder = "C:\Users\$UserID\AppData\Roaming\Microsoft\Signatures\"
        $SignatureName = "$CompanyName ($UPN)"

        try {
            Invoke-Command -ComputerName $ComputerName -ErrorAction Stop -ScriptBlock {
                if (Test-Path "$using:LocalSignatureFolder$using:SignatureName.htm") {
                    Remove-Item -Path "$using:LocalSignatureFolder$using:SignatureName.htm" | Out-Null
                }
                New-Item -Path "$using:LocalSignatureFolder$using:SignatureName.htm" -force | Out-Null
                Set-Content -Path "$using:LocalSignatureFolder$using:SignatureName.htm" -Value $using:LocalSig | Out-Null
            }
            $global:Log.WriteInfo("Signature Updated for $UserID on $ComputerName")
        } catch {
            $global:Log.WriteError("Signature Update Failed for $UserID on $ComputerName")
        }
    }
}

function Disable-RoamingSignatures {
    param ($User)

    $SID = $User.SID
    $ComputerName = $User.Computer

    $RoamingTogglePath = "$SID\Software\Microsoft\Office\16.0\Outlook\Setup\"
    $RoamingToggleKey = 'DisableRoamingSignaturesTemporaryToggle'
    $RoamingToggleValue = 1

    try {
        $base = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey([Microsoft.Win32.RegistryHive]::Users, $ComputerName)
        if ($base) {
            $RoamingKey = $base.OpenSubKey($RoamingTogglePath, $true)
            if ($RoamingKey) {
                if ($null -eq $RoamingKey.GetValue($RoamingToggleKey)) {
                    $RoamingKey.SetValue($RoamingToggleKey,$RoamingToggleValue,[Microsoft.Win32.RegistryValueKind]::DWord)
                } elseif ($RoamingKey.GetValue($RoamingToggleKey) -ne $RoamingToggleValue) {
                    $RoamingKey.SetValue($RoamingToggleKey,$RoamingToggleValue)
                }
                $RoamingKey.Close()
            }
            $base.Close()
            $global:Log.WriteInfo("Roaming Toggle registry value updated on $ComputerName")
        } else {
            $global:Log.WriteWarning("Unable to update Roaming Toggle value on $ComputerName")
        }        
    } catch [UnauthorizedAccessException],[System.IO.IOException] {
        $global:Log.WriteError("Cannot Access Computer $ComputerName")
    }
}

function Update-SignatureRegistry {
    param ($User, $CompanyName)

    $SID = $User.SID
    $ComputerName = $User.Computer
    $UPN = $User.UPN
    $UserID = $User.SAM

    $SignatureName = "$CompanyName ($UPN)"

    $ProfilesPath = "$SID\Software\Microsoft\Office\16.0\Outlook\Profiles\"
    $ProfileTypeKey = '9375CFF0413111d3B88A00104B2A6676'

    try {
        $base = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey([Microsoft.Win32.RegistryHive]::Users, $ComputerName)
        if ($base) {
            $ProfileKey = $base.OpenSubKey($ProfilesPath)
            if ($ProfileKey) {
                foreach ($Profile in $ProfileKey.GetSubKeyNames()) {
                    $AccountKey = $base.OpenSubKey("$ProfilesPath$Profile\$ProfileTypeKey")
                    if ($AccountKey) {
                        foreach ($Account in $AccountKey.GetSubKeyNames()) {
                            $SignatureKey = $base.OpenSubKey("$ProfilesPath$Profile\$ProfileTypeKey\$Account",$true)
                            if ($SignatureKey) {
                                $NewSignature = $SignatureKey.GetValue('New Signature')
                                $ReplyForwardSignature = $SignatureKey.GetValue('Reply-Forward Signature')
                                if ($null -eq $NewSignature) {
                                    $SignatureKey.SetValue('New Signature',$SignatureName,[Microsoft.Win32.RegistryValueKind]::String)
                                } elseif ($NewSignature -ne $SignatureName) {
                                    $SignatureKey.SetValue('New Signature',$SignatureName)
                                }
                                <#
                                if ($null -eq $ReplyForwardSignature) {
                                    $SignatureKey.SetValue('Reply-Forward Signature',$SignatureName,[Microsoft.Win32.RegistryValueKind]::String)
                                } elseif ($ReplyForwardSignature -ne $SignatureName) {
                                    $SignatureKey.SetValue('Reply-Forward Signature',$SignatureName)
                                }
                                #>
                                $SignatureKey.Close()
                            }
                        }
                        $AccountKey.Close()
                    }
                }
                $ProfileKey.Close()
            }
            $base.Close()
            $global:Log.WriteInfo("Signature Registry updated for $UserID on $ComputerName")
        } else {
            $global:Log.WriteWarning("Unable to update Signature Registry for $UserID on $ComputerName")
        }        
    } catch [UnauthorizedAccessException],[System.IO.IOException] {
        $global:Log.WriteError("Cannot Access Computer $ComputerName")
    }
}

function Update-WebSignature {
    param ($User, $CompanyName)

    $UPN = $User.UPN
    $WebSig = $User.UserSignature.WebSignature

    $global:Log.WriteInfo("Updating Web Signature for $UPN")

    try {
        Set-MailboxMessageConfiguration -Identity $UPN -SignatureName $CompanyName -SignatureHTML $WebSig -DefaultSignature $true -UseDefaultSignatureOnMobile $true -AutoAddSignature $true #-AutoAddSignatureOnReply $true
        $global:Log.WriteInfo("Signature Updated for $UPN")
    } catch {
        $global:Log.WriteError("Signature Update Failed for $UPN")
    }
}



#endregion

#region Main

$global:Log.WriteInfo('Starting Script')
 
if ($TestSignature) {

    $global:Log.WriteInfo('Generating Test Signatures')

    $TestUser = [PSCustomObject]@{ UPN = "testuser@$(($CompanyName -replace ' ','').ToLower()).com"; Title = 'Signature Tester'; Name = 'Test User'}
    $TestUserSignature = Create-UserSignature -User $TestUser -TemplatePath $Template.FolderPath -TemplateName $Template.FileName -ImageName $Template.ImageName -CompanyName $CompanyName
    $LocalHtmPath = ([string]($MyInvocation.MyCommand.Path) -replace [string]($MyInvocation.MyCommand.Name),'Local_Signature_Test.htm')
    $WebHtmlPath = ([string]($MyInvocation.MyCommand.Path) -replace [string]($MyInvocation.MyCommand.Name),'Web_Signature_Test.html')
    New-Item -ItemType File -Path $LocalHtmPath -force | Out-Null
    Set-Content -Path $LocalHtmPath -Value $TestUserSignature.LocalSignature | Out-Null
    New-Item -ItemType File -Path $WebHtmlPath -force | Out-Null
    Set-Content -Path $WebHtmlPath -Value $TestUserSignature.WebSignature | Out-Null

    $global:Log.WriteInfo('Test Signatures Complete')

} else {

    Set-TLSConfiguration

    if ($UseLocalOutlook -eq $false -and $UseWebOutlook -eq $false) {
        $global:Log.WriteError("At least one of the Outlook types needs to be set to 'true' in the Config file")
        exit 1
    }

    if ($UseLocalOutlook) {
        $LocalUsers = Get-LocalUsers -SearchBase $SearchBase
        if ($FindComputers) {
            $Computers = Get-Computers -FilterXML $FilterXML -ComputerPrefix $ComputerPrefix
        }
    }

    if ($UseWebOutlook) {
        if (Connect-TenentExchange -ExchangeOnlineInfo $ExchangeOnlineInfo) {
            $WebUsers = Get-WebUsers
            Disconnect-TenetExchange
        }
    }

    if ($UseCSVData) {
        $CSVUsers = Get-CSVUsers -CSVPath $UserCSVPath
    }

    $MergedUsers = Merge-UserData -LocalUsers $LocalUsers -WebUsers $WebUsers -CSVUsers $CSVUsers -Computers $Computers

    $global:Log.WriteInfo('Generating Signatures')

    foreach ($MUser in $MergedUsers) {
        $MUser.UserSignature = Create-UserSignature -User $MUser -TemplatePath $Template.FolderPath -TemplateName $Template.FileName -ImageName $Template.ImageName -CompanyName $CompanyName
    }

    $Users = Filter-Users -AllUsers $MergedUsers -FilterUsers $FilterList

    if ($TestUsers) {
        $OutCSVPath = ([string]($MyInvocation.MyCommand.Path) -replace [string]($MyInvocation.MyCommand.Name),'TestUserData.csv')
        $Users | Export-Csv -Path $OutCSVPath

        $global:Log.WriteInfo("Output $(($Users | Measure).Count) users to CSV")
    } else {
        foreach ($User in $Users) {
            if ($UseLocalOutlook) {
                if ($null -ne $User.Computer -and $User.Computer -ne '') {
                    Update-LocalSignature -User $User -CompanyName $CompanyName
                    Disable-RoamingSignatures -User $User
                    Update-SignatureRegistry -User $User -CompanyName $CompanyName
                } else {
                    $global:Log.WriteError("Computer not found for $User.Name")
                }
            }

            if ($UseWebOutlook) {
                if (Connect-TenentExchange -ExchangeOnlineInfo $ExchangeOnlineInfo) {
                    Update-WebSignature -User $User -CompanyName $CompanyName
                    Disconnect-TenetExchange
                }
            }
        }

        $global:Log.WriteInfo("$(($Users | Measure).Count) users updated")
    }
}

$global:Log.WriteInfo('End of Script')

#endregion

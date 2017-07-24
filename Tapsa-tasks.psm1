Function Get-Dice {
<#
.SYNOPSIS
 Throw dice
.DESCRIPTION
 Returns random value 1-6
.EXAMPLE
 Get-Dice
 Returns random value 1-6
#>
    [CmdletBinding()]
    Param()

    $result = Get-Random -Minimum 1 -Maximum 7

    return $result
}


Function Get-Employee {
<#
.SYNOPSIS
 Search Active Directory accounts by title or name
.DESCRIPTION
 Retrieves names and job titles based on search terms. Name or job title must be used as parameter.
.PARAMETER search
 One or multiple search terms separated by ','. Search by job title or name.
.EXAMPLE
 Get-Employee -search janitor,teacher
 Retrieves all janitors and teachers
.EXAMPLE
 Get-Employee -search John
 Retrieves everyone named John
#>
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory=$True)]
        [string]$search
    )

    $searchlist = $search.Split(",") # Separate search terms

    Try 
    {
        foreach ($search in $searchlist) 
        {
            # Search AD user by Title
            $wildsearch = "*$search*"
            $thisresult = Get-ADUser -Filter {title -like $wildsearch -and enabled -eq $true} -Properties title | fl name,title

            # Search AD user by givenname/surname/fullname
            if ("$thisresult" -eq '') 
            {
                $surname    = $search.Split(" ")|select -Last 1
                $givenname  = $search.Split(" ")|select -First 1
                $fullname   = "$surname $givenname"
                $thisresult = Get-ADUser -Filter {givenname -like $search -or surname -like $search -or name -like $search -or name -like $fullname} -Properties title | fl name,title | Out-String
            }
            
            if ("$thisresult" -eq '') 
            {
                $thisresult = "$search not found"
            }

            $result += "`n`n"+"$thisresult".trim()
       }
    }
    Catch 
    {
        write-host "fail $error"
        $result = "Something went wrong"
    }

    return $result
}


function Start-AADsync {
<#
.SYNOPSIS
 Start Azure AD synchronization.
.DESCRIPTION
 This function executes AAD Sync delta synchronization command on a remote server that has Azure AD syncronization tool installed.
.PARAMETER ComputerName
 A single computer name.
.EXAMPLE
 Start-AADsync -ComputerName server01
 Run Azure AD sync on server01
#>
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory=$True)]
        [string]$ComputerName
    )
    
    Try 
    {
        # Credentials
        $user = "admin@domain.com"
        $pass = ConvertTo-SecureString -string "Password1245" -Force -AsPlainText
        $cred = new-object -typename System.Management.Automation.PSCredential -argumentlist $user,$pass
        
        # Run remote commands
        Invoke-Command -Credential $cred -ComputerName $ComputerName {
            Import-Module ADSync
            Start-ADSyncSyncCycle -PolicyType Delta
        }

        # Send information about the progress. Note this works only from Skype-bot!! You need to remove this part if you want to use the function elsewhere.
        #############################################################
        $info = "Synchronization started, please wait 1 minute"
        $msg = New-Object "System.Collections.Generic.Dictionary[Microsoft.Lync.Model.Conversation.InstantMessageContentType,String]"
        $msg.Add(0, $info)
        Send -msg $msg
        sleep 60
        ############################################################# END OF BOT RELATED CODE

        $result = "Synchronization finished"
    }
    Catch 
    {
        write-host "fail $error"
        $result = "Something went wrong"
    }
    return $result
}

Function Get-Lunch {
<#
.SYNOPSIS
 Retreives lunch details from Sharepoint Online intranet-site.
.DESCRIPTION
 This function connects to Sharepoint Online (Office 365) and retrieves this week's lunch menu from a sharepoint-list. It is done by parsing the RSS-feed data of the list. Lunch for current day is displayed.
.EXAMPLE
 Get-Lunch
 Retrieves lunch info
#>
    [CmdletBinding()]
    Param()

    $date = get-date -Format %d.%M.20%y # make sure this is in the same format as your date in the sharepoint list

    # Credentials
    $user        = "user@domain.com"
    $password    = ConvertTo-SecureString -string "Password1" -Force -AsPlainText
    $credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($user, $password)

    Try
    {
        Try 
        {
            if (!(Test-Path .\$date.xml))  # Get a new list if the xml list doesn't exist
            {
                # Sharepoint client assemblies
                [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client")         | Out-Null
                [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client.Runtime") | Out-Null

                # Sharepoint list details
                $WebUrl     = "https://domain.sharepoint.com/somesite"     # sharepoint online site
                $ListId     = "{B26413C4-75A2-482A-87D3-4284094D69D0}"         # ID of the sharepoint list
                $listRssUrl = "$WebUrl/_layouts/15/listfeed.aspx?List=$ListId" # RSS URL
                
                # Download RSS
                $client = New-Object System.Net.WebClient
                $client.Credentials = $credentials
                $client.Headers.Add("X-FORMS_BASED_AUTH_ACCEPTED", "f")
                $content = $client.DownloadString($listRssUrl)
                $client.Dispose()
                
                # Save RSS as XML
                [xml]$rss = $content
                $rss.save(".\$date.xml")                
            }
            # Read XML and parse content
            $lunchxml = [xml](Get-Content ".\$date.xml")
            $menu     = $lunchxml.rss.channel.item.description."#cdata-section" -replace "<.*?>" -split "´n"
        }
        Catch 
        {
            write-host $error
            $result = "Something went wrong"
        }

        # pick today's lunch details from menu
        $lunch = $menu  | where { ($_.contains("$date") -eq $true)}

        if ($lunch -eq $null) 
        {
            $lunch = "I don't know the lunch for the date: $date" 
        }
        $result = $lunch
    }
    Catch
    {
        write-host "Food fail: $error"
        $result = "Everything went wrong"
    }
    return $result
}

function Get-License {
<#
.SYNOPSIS
 Check user's Office 365 license or assign a new license
.DESCRIPTION
 Connects to Office 365 and retrieves license information of a single user or assigns a new license
.PARAMETER upn
 User email-address / userprincipalname
.PARAMETER assign
 Assign O365 license, $true or $false (default)
.EXAMPLE
 Get-License -user john.smith@domain.com
.EXAMPLE
 Get-License -user john.smith@domain.com -assign $true
 Assign O365 license
#>    
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True)]
        [string]$upn,
        [string]$assign=$false
    )

    Try 
    {
        # Connect to Office 365
        $user        = "admin@domain.com"
        $pass        = ConvertTo-SecureString -string "Password1234" -Force -AsPlainText
        $O365Cred    = new-object -typename System.Management.Automation.PSCredential -argumentlist $user,$pass
        $O365Session = New-PSSession –ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell -Credential $O365Cred -Authentication Basic -AllowRedirection 
        
        Import-PSSession $O365Session -AllowClobber | Out-Null
        Connect-MsolService –Credential $O365Cred

        # Office 365 tenant
        $tenant = "domain"

        # Remove skype messaging related stuff
        $upn = $upn.Replace("o365 license","").Replace("set","").Trim()
        $upn = $upn.split(" ") | select -First 1 # Skype automatically formats email-addresses as "name@domain.com <mailto:name@domain.com>" so we select just the first one.

        # Assign Office 365 license if Assign parameter is $true
        if ($assign) 
        {
            Set-MsolUser -UserPrincipalName $upn -UsageLocation "FI" # Usage location for Finland
            Set-MsolUserLicense -UserPrincipalName $upn -AddLicenses ($tenant+":STANDARDWOFFPACK")
        }

        # retrieve license info
        $result = Get-MsolUser -UserPrincipalName $upn | fl displayname,licenses
    }
    Catch
    {
        write-host "fail: $error"
        $result = "Something went wrong"
    }
    Finally 
    {
        # Close Office 365 session
        Remove-PSSession $O365Session
    }
    return $result
}
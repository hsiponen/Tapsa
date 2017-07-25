# Task Automation PowerShell Assistant (Tapsa) bot for Lync/SfB
#
#   ---
#
#
#    Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
#
#    The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.
#
#    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
#
#	---
#
#
# Tapsa is a Skype for Business / Lync -chatbot capable of running powershell functions
#
# Requirements:
# PowerShell 3.0 or later
# Lync SDK (Lync .NET assemblies), follow this guide (the LyncSDK.exe installation part): https://blogs.technet.microsoft.com/heyscriptingguy/2016/08/19/use-powershell-to-integrate-with-the-lync-2013-sdk-for-skype-for-business-part-1/
# Microsoft Skype for Business or Lync 2010/2013
# ActiveDirectory PowerShell module (only for permission check (check-permissions) and example functions (Tapsa-module))
#
# Recommended use: run from PowerShell ISE
# Bot will run as long as PowerShell session is running
#
# Microsoft Lync Model API documentation
# https://msdn.microsoft.com/library/office/hh243705(v=office.14).aspx


$ModelPath    = "C:\Program Files (x86)\Microsoft Office 2013\LyncSDK\Assemblies\Desktop\Microsoft.Lync.Model.dll"    # Location of Lync Model .NET assembly
$ControlsPath = "C:\Program Files (x86)\Microsoft Office 2013\LyncSDK\Assemblies\Desktop\Microsoft.Lync.Controls.dll" # Location of Lync Controls .NET assembly
$TapsaModule  = "C:\Tapsa-tasks.psm1"            # module for bot tasks, add your own.
$logfile      = "C:\botlog.txt"                  # Location of log file

$botUser = "username@domain.com" # Lync username
$botPwd  = "password"            # Lync password

$permittedGroup = "tapsa_users" # AD group permitted to use the bot

Get-EventSubscriber | Unregister-Event

Try
{
    Import-Module $ModelPath
    Import-Module $ControlsPath
    Remove-Module Tapsa-tasks -ErrorAction Ignore # This removes the task module from previous run just in case it was modified
    Import-Module $TapsaModule
}
Catch
{
    Write-Host "module import error. $error"
    break
}


function check-permissions ([string]$upn,[string]$group) {
    Try
    {
        $permitted = $false
        $users     = Get-ADGroupMember $group -Recursive | select -ExpandProperty samaccountname | get-aduser
        foreach ($user in $users) 
        {
            if ($user.userprincipalname -eq $upn) 
            {
                $permitted = $true
            }
        }
        return $permitted
    }
    Catch
    {
        Write-Host $error
    }

}

function run-command($msgStr)
{
    try 
    {
        switch ($msgStr) # do something based on the received message
        {
            {($null -ne ("food","lunch"    | where { $msgStr -match $_}))} { $result = get-lunch }                                                # example: lunch tomorrow
            {($null -ne ("who is"          | where { $msgStr -match $_}))} { $result = get-employee -search $msgStr.Replace("who is","").Trim() } # example: who is HR-assistant
            {($null -ne ("hello","hi","yo" | where { $msgStr -match $_}))} { $result = "hi","hey :)","hello :)" | get-random }                    # example: hi
            {$msgStr -eq "sync ad"} { $result = Start-AADsync -ComputerName server01 }                                                            # example: sync ad
            {$msgStr -eq "dice"}    { $result = get-dice }                                                                                        # example: dice
            {$msgStr.Contains("o365 license")} {                                                                                                  # example: o365 license user@domain.com 
                if ($msgStr.Contains("set")) { $result = Get-License -upn $msgStr -assing $true }
                else                         { $result = Get-License -upn $msgStr }}
            default { $result = ""}
        }

        # Lync only accepts string-type messages
        $result = $result | Out-String

	    return $result.Trim()
    }
    Catch 
    {
        write-host $error
        write-host $msgStr
    }
}

function send($msg)
{
	$null = $Modality.BeginSendMessage($msg, $null, $msg)
}

function log-message($usr,$msg)
{
    Try
    {
        $time = get-date -UFormat "%d.%m.%Y %H:%M:%S"
        "$time $usr : $msg" | Out-File -Encoding utf8 -Append -FilePath $logfile
    }
    catch 
    {
        Write-Host $error
        Write-Host $msg
    }
}


# Obtain the entry point to the Lync.Model API
$client = [Microsoft.Lync.Model.LyncClient]::GetClient()

# Sign in if not already signed
if ($client.State -ne [Microsoft.Lync.Model.ClientState]::SignedIn)
{
    $client.EndSignIn(
        $client.BeginSignIn("$botUser", "$botUser", "$botPwd", $null, $null))
}

# Add event to pickup new conversations and register events for new participants
$conversationMgr = $client.ConversationManager
Register-ObjectEvent -InputObject $conversationMgr -EventName "ConversationAdded" -SourceIdentifier "NewIncomingConversation" -Action {
	$client = [Microsoft.Lync.Model.LyncClient]::GetClient()
	foreach ($con in $client.ConversationManager.Conversations)
	{
		# For each participant in the conversation
		$peers = $con.Participants | Where { !$_.IsSelf }
		foreach ($c in $peers)
		{
			if (!(Get-EventSubscriber $c.Contact.uri))
			{
				Register-ObjectEvent -InputObject $c.Modalities[1] -EventName "InstantMessageReceived" -SourceIdentifier $c.Contact.uri -action $global:action
			}
		}
	} 	
}

# Register events for current open conversation participants
foreach ($con in $client.ConversationManager.Conversations)
{
	# For each participant in the conversation
	$peers = $con.Participants | Where { !$_.IsSelf }

	foreach ($c in $peers)
	{
        Try 
        {
		    if (!(Get-EventSubscriber $c.Contact.uri -ErrorAction Ignore))
		    {
			    Register-ObjectEvent -InputObject $c.Modalities[1] -EventName "InstantMessageReceived" -SourceIdentifier $c.Contact.uri -action $global:action
		    }
        }
        Catch [system.ArgumentExeption] 
        {}
	}
}

$global:action = {
	# New conversation / receive message
	$Conversation = $Event.Sender.Conversation
	$msg = New-Object "System.Collections.Generic.Dictionary[Microsoft.Lync.Model.Conversation.InstantMessageContentType,String]"
	$Global:Modality = $Conversation.Modalities[1]	
	[string]$msgStr = $Event.SourceArgs.Text
	
    # trim and convert user message to lowercase
    $msgStr = $msgStr.ToString().ToLower().Trim()

    Try 
    {
        # Log user's message
        $sender = $Event.SourceIdentifier
        $sender = "$sender".Replace('sip:','')
        log-message -usr $sender -msg $msgStr

        # Run a command based on the message. Allowed only for a certain AD-group
        if (check-permissions -upn $sender -group $permittedGroup) 
        {
            $response = run-command -msgStr $msgStr
        }
        else 
        {
            $response = "Not permitted"
        }

        # If there's a response, send it to user & write to log
        if ($response -ne "") 
        {
            # Send response to user
            $msg.Add(0, $response)
            send -msg $msg 

            # Log response
            log-message -usr "Tapsa" -msg $response
        }
    }
    Catch 
    {
        Write-Host $error
        write-host $msgStr
        $msg.Add(0, "Something went very wrong")
        send -msg $msg
    }   
}

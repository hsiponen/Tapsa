# Tapsa - bot for Lync/SfB (PowerShell script)
Henri Siponen 2017

Tapsa is a Skype for Business / Lync -chatbot capable of running powershell functions. The bot is just a powershell script. Not an executable application.

![gif](/gif/lunch.gif)

Microsoft Lync Model API documentation  
https://msdn.microsoft.com/library/office/hh243705(v=office.14).aspx

Tapsa can be used as a software robot for automating Windows Server administration tasks. At least that's how I've been using it for more than two years. 
Of course it can do much more. If you can do it with PowerShell, you can do it from Lync with Tapsa. For example, our employees use it to 
check todays lunch menu and the menu of nearby restaurants.
It's surprisingly fun and convenient to run powershell tasks from a Lync-client (which I'm using all the time anyway).

This bot comes with an example module that has a few functions. The bot we use in production has more than 60 functions and more than half of them are actual server/cloud administration tasks that our IT-support uses daily. They rarely need admin privileges anymore 
since Tapsa can do all that work. Saves us at least two full work days on average per month. Propably more since it doesn't make any mistakes. I will propably keep adding some of those functions to Tapsa-tasks.psm1.

## Requirements
* Windows 7 or later
* PowerShell 3.0 or later
* Lync SDK (Lync .NET assemblies), follow this guide (the LyncSDK.exe installation part): https://blogs.technet.microsoft.com/heyscriptingguy/2016/08/19/use-powershell-to-integrate-with-the-lync-2013-sdk-for-skype-for-business-part-1/
* Microsoft Skype for Business or Lync 2010/2013
* ActiveDirectory PowerShell module (only for permission check (check-permissions) and example functions (Tapsa-module))

Recommended use: run from PowerShell ISE  
The bot will run as long as the PowerShell session is running.

## How to use
1. Make sure you meet the requirements above
2. Clone the repository or download the Tapsa.ps1 and Tapsa-tasks.psm1 (or use your own powershell modules)
3. Open the Tapsa.ps1 script in Windows PowerShell ISE
4. Modify filepaths and credentials and save the script
7. Start Skype for Business or Lync
8. Run the script
9. Try sendin a message "hi" to the bot from another Lync account
10. You should get a reply
11. Next:
  * Make sure you understand what the script does and where you need to add your own commands.
  * You need to create your own functions. Tapsa-tasks.psm1 is just provided as an example.



![gif](/gif/teppo.gif)
# Task Automation PowerShell Assistant (Tapsa) bot for Lync/SfB
Henri Siponen 2017

Tapsa is a Skype for Business / Lync -chatbot capable of running powershell functions

![gif](/gif/lunch.gif)

Requirements:
PowerShell 3.0 or later
Lync SDK (Lync .NET assemblies), follow this guide (the LyncSDK.exe installation part): https://blogs.technet.microsoft.com/heyscriptingguy/2016/08/19/use-powershell-to-integrate-with-the-lync-2013-sdk-for-skype-for-business-part-1/
Microsoft Skype for Business or Lync 2010/2013
ActiveDirectory PowerShell module (only for permission check (check-permissions) and example functions (Tapsa-module))

Recommended use: run from PowerShell ISE
The bot will run as long as the PowerShell session is running

Microsoft Lync Model API documentation
https://msdn.microsoft.com/library/office/hh243705(v=office.14).aspx

This is a software robot for automating Windows Server administration tasks. That's how I've been using it for two years. 
Of course it can do much more. If you can do it with PowerShell, you can do it from Lync with Tapsa. For example, our employees use it to 
check todays lunch menu and the menu of nearby restaurants.
It's surprisingly fun and convenient to run powershell tasks from a Lync-client (which I'm using all the time anyway).

This bot comes with an example module that has a few functions. The bot we use in production has more than 50 functions and about 
half of them are actual server/cloud administration tasks that our IT-support uses daily. They rarely need admin privileges anymore 
since Tapsa can do all that work. Saves us at least two full work days on average per month. Propaly more since it doesn't make any mistakes.


![gif](/gif/teppo.gif)
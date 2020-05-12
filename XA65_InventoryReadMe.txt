Before you can start using PowerShell to document anything in the XenApp 6.5 farm you first need to install the SDK (for the Help file) and Citrix Group Policy commands. From your XenApp 6.5 server, go to http://tinyurl.com/XenApp65PSSDK 

From your XenApp 6.5 server, go to http://tinyurl.com/XenApp65PSSDK 
Scroll down and click on Download XenApp 6.5 Powershell SDK — Version 6.5
Extract the file to C:\XA65SDK. Click Start, Run, type in C:\XA65SDK\XASDK6.5.exe and press Enter 
Click Run 
Select I accept the terms of this license agreement and click Next 
Select Update the execution policy (to AllSigned) and Click Next.
Note: If you do not update the execution policy to AllSigned, the Citrix supplied XenApp PowerShell scripts will not load.
Click Install 
After a few seconds, the installation completes. Click Finish 

In your Internet browser; go to http://tinyurl.com/XenApp6PSPolicies

Scroll down and click on Citrix.GroupPolicy.Commands.psm1

Save the file in two different places:

C:\Windows\System32\WindowsPowerShell\v1.0\Modules, in a new folder named Citrix.GroupPolicy.Commands

C:\Windows\SysWOW64\WindowsPowerShell\v1.0\Modules, in a new folder named Citrix.GroupPolicy.Commands

Close your Internet browser.

Click Start, Administrative Tools, Windows PowerShell Modules.

To prepare for processing the Citrix farm policies, type in import-module Citrix.GroupPolicy.Commands

If you use Configuration Logging, you will need to use a UDL file in order for the History section of the script to work.  For an explanation, see http://tinyurl.com/CreateUDLFile.

The UDL file will need to be placed in the same folder as the XA65_Inventory.ps1 script.  The UDL file will need to be named XA65ConfigLog.udl.  You will need to edit the UDL file and add  ;Password=ConfigLogDatabasePassword to the end of the last line in the file.  For example, here is mine (line is one line):

Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;User ID=administrator;Initial Catalog=XA65ConfigLog;Data Source=SQL;Password=abcd1234

How to use this script?

I saved the script as XA65_Inventory.ps1 in the C:\PSScripts folder. From the PowerShell prompt, change to the C:\PSScripts folder, or the folder where you saved the script. From the PowerShell prompt, type in:

.\XA65_Inventory.ps1 |out-file .\XA65Farm.doc and press Enter.

Open XA65Farm.txt in either WordPad or Microsoft Word 

To use the script with Remoting, read the following article:

http://carlwebster.com/using-my-citrix-xenapp-6-5-powershell-doumentation-script-with-remoting/

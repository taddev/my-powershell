# PowerTab 0.97

you can doubleclick setup.cmd to unblock if needed and then start PowerTabsetup.ps1 directly
next - next -next upgrade 

for rest see changes.txt

# PowerTab 0.95

PowerShell TabExpansion extension

### Instalation ### 

On first installation you can just start PowerTabSetup.PS1 to Install / Configure PowerTab

follow the instructions or just say yes to all questions (next next next setup)

you will need to Dot Source the Setup to directly start using the new tabcompletion functions
if you start it normaly the setup will work fine but you need to start a new PowerShell session to use it.
this looks like this :

. ./PowerTabSetup.PS1

### Upgrading from former versions ###

# full setup

the most simple way to upgrade is to copy all files to the installation directory and run PowerTabSetup.ps1 again
then after setup remove the addition to the profile from the former version.

# keep database

you do not have to overwrite the database, but a new table needs to be added, so if you do not create one setup will ask to add the config table

# Manual 

you can run 

New-PowerTabConfig.ps1 to update the database

also the profile needs to be updated like this :

################ Start of PowerTab 0.95 TabCompletion Code ########################
#
#  added by PowerTab 0.95 Setup For Loading of Custom TabExpansion,
#
# /\/\o\/\/
# http://ThePowerShellGuy.com
#

# Load the PowerTab Utility Functions

. 'F:\PowerShell\PowerTab\TabExpansionLib.ps1'     

# Import the TabExpansion DataBase

Import-TabExpansionDataBase 

# Initialize PowerTab Configuration

F:\PowerShell\PowerTab\Init-TabExpansion.ps1 

# load other functions 

. 'F:\PowerShell\PowerTab\TabExpansion.ps1'          # Load Main Tabcompletion function
. 'F:\PowerShell\PowerTab\Out-DataGridView.ps1'      # Used for GUI TabExpansion
. 'F:\PowerShell\PowerTab\Out-ConsoleList.ps1'       # Used for RawUi ConsoleList
. 'F:\PowerShell\PowerTab\ConsoleLib.ps1'       # Used for RawUi ConsoleList
. 'F:\PowerShell\PowerTab\Get-ScriptParameters.ps1'  # Get Parameters of Scripts


# load External Library for Share Enumeration

[void][System.Reflection.Assembly]::LoadFile('F:\PowerShell\PowerTab\shares.dll')

################ End of PowerTab 0.95 TabCompletion Code ##########################

thats all

Enjoy


## history 


# PowerTab 0.8

PowerShell TabExpansion extension

Start PowerTabSetup.PS1 to Install / COnfigure PowerTab 0.72


For more Information and examples about PowerTab Tabexpansions see :

The PowerTab PowerShell Tab Extension Overview Page

http://thepowershellguy.com/blogs/posh/pages/powertab.aspx

This script makes use of the Shares.DLL from the following C# library
 
http://www.codeproject.com/cs/internet/networkshares.asp

you can find the complete source in the Lib directory

Enjoy, Greetings /\/\o\/\/

http://ThePowerShellGuy.com

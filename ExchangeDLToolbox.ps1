######################################################################################
# Exchange DL Toolbox V1.3 - Released 9/9/2022                                       #
#                                                                                    #
# Script Created by Ryan Strayhan:                                                   #
# Released Under MIT License                                                         # 
######################################################################################


#####################################################################
# ------- FUNCTIONS FOR DISTRIBUTION LIST USER ADMIN -------------- #
#####################################################################

# This function creates a DL and sets all of it's parameters.
function Add-DL-Single
{Clear-Host
  [string]$DLNameSet = Read-Host "Enter the name of the new DL."
  [string]$DLDNSet = Read-Host "Enter the Display Name for the DL."
  [string]$DLAliasSet = Read-Host "Enter the Alias for the DL."
  [string]$DLEmailSet = Read-Host "Ener the desired Email Address for the DL."
  [string]$DLOwnerSet = Read-Host "Enter the Email Address of the DL's Owner/Maintainer."
  New-DistributionGroup -Name $DLNameSet -DisplayName $DLDNSet -Alias $DLAliasSet -PrimarySmtpAddress $DLEmailSet -ManagedBy $DLOwnerSet -Type SelfService
  Pause
}
# This function adds a single user to a DL and gives Member access as the Default permission.
function Add-DLUser-Single
{Clear-Host
  [string]$UserNameDLAdd = Read-Host "Enter the display name, email address or AD Username of the USER that needs access to a DL. By Default, they will be given 'Member' access only'"
  [string]$DistListAdd = Read-Host "Enter the display name, email address or AD Username of the DL they need access to"
  #[string]$DLPermissionAdd = Read-Host "What kind of permissions does the user need e.g. FullAccess, SendAs"
  Add-DistributionGroupMember -Identity $DistListAdd -Member $UserNameDLAdd
  Pause
}
# This function removes a single user from the DL.
function Remove-DLUser-Single
{Clear-Host
  [string]$UserNameDLRem = Read-Host "Enter the display name, email address or AD Username of the USER that you want to remove from the DL"
  [string]$DistListRem = Read-Host "Enter the display name, email address or AD Username of the DL"
  Remove-DistributionGroupMember -Identity $DistListRem -User $UserNameDLRem -Confirm:$false
  Pause
}
# This function adds a Maintainer/Owner to a DL.
function Add-DLOwn-Single
{Clear-Host
  [string]$UserNameOwnerDL = Read-Host "Enter the display name, email address or AD Username of the new Maintainer"
  [string]$DistListOwnAdd = Read-Host "Enter the display name, email address or AD Username of the DL"
  Set-DLManager -Identity $UserNameOwnerDL -AddManager $DistListOwnAdd
  Pause
}
# This function sets the DL Owner whaen none was previously specified.
function Add-DLOwn-New
{Clear-Host
  [string]$SetDLOwnName = ReadHost "Enter the display name, email address or AD Username of the new Maintainer."
  [string]$SetDLListOwnAdd = ReadHost "Enter the name, email address, or the AD Username of the DL."
  Set-DistributionGroup -Name $SetDLListOwnAdd -ManagedBy $SetDLOwnName
  Pause
}
# This function will remove a single maintainer
function Rem-DLOwn-Single
{Clear-Host
  [string]$DLRemOwn = ReadHost "Enter the name, email address, or AD UserName of the Maintainer to be removed."
  [string]$DistListOwnRem = ReadHost "Enter the name, email address, or the AD UserName of the DL."
  Set-DLManager -$DistListOwnRem -RemoveManager$DLRemOwn
  Pause
}
# This function lists all users who are members of a DL
function ListDLMembers
{Clear-Host
  [string]$DLUserList = Read-Host "Enter the display name, email address or AD Username of the DL to list the users who have access and their permission."
  Get-DistributionGroupMember -Identity $DLUserList
  Pause
}
# This function lists only the Maintainers for a DL.
function ListDLOwn
{Clear-Host
  [string]$DLOwnList = Read-Host "Enter the display name, email address, or AD UserName of the DL to show its Maintainers."
  Get-DLManager -DLName
  Pause
}
# This function shows the list of Maintainers with the full AD path.
function ListDLOwnFull
{Clear-Host
  [string]$DLOwnListFull = "Enter the display name, email address, or AD UserName of the DL to show its Maintainers with full AD path."
  Get-DistributionGroup $DLOwnListFull | select -ExpandProperty ManagedBy
  Pause
}
# This function allows the Meta Group Naming Policy to be bypassed for Display Name.
function MetaPolicyBypass
{Clear-Host
  Write-Host "Please Note: THIS FUNCTION WILL OVERRIDE META GROUP NAMING CONVENTIONS. Make sure you have proper permission to make these changes."
  [string]$DLName = "Enter the display name of the Existing DL needing Meta Group Naming Policy Bypass."
  [string]$DLBypDispName = "Enter the new Display Name requested"
  Set-DistributionGroup -Identity $DLName -Name $DLBypDispName -IgnoreNamingPolicy
  Pause
}

#####################################################################
# ------------ FUNCTIONS FOR DL ACCOUNT ADMINISTRATION -------------#
#####################################################################

# This function shows the full details of a DL
function DLFullDetails
  {Clear-Host
  [string]$DLCheckFull = Read-Host "Enter the display name, email address or AD Username of the DL to its full information."
  Get-DistributionGroup -Identity $DLCheckFull | Format-List
  Pause
}
# This function lets you set the Display Name of a DL.
function DLDispName
  {Clear-Host
  [string]$DLNameDN = Read-Host "Enter the display name, email address or AD Username of the DL."
  [string]$NewDLDispName = Read-Host "Enter the display name, email address or AD Username of the new Display Name."
  Set-DistributionGroup -Identity $DLNameDN -DisplayName $NewDLDispName
  Pause
}
# This function lets you set the Display Name of a DL and overide Meta Group Naming Policy.
function DLDispNameOVR
  {Clear-Host
  Write-Host "Please Note: THIS FUNCTION WILL OVERRIDE META GROUP NAMING CONVENTIONS. Make sure you have proper permission to make these changes."  
  [string]$DLNameDNOR = Read-Host "Enter the display name, email address or AD Username of the DL."
  [string]$NewDLDispNameOR = Read-Host "Enter the display name, email address or AD Username of the new Display Name."
  Set-DistributionGroup  -Name $DLNameDNOR -DisplayName $NewDLDispNameOR -ForceUpgrade
  Pause
}
# This function lets you change a DL's Alias.
function DLSetAlias
  {Clear-Host
  [string]$DLAliasName = Read-Host "Enter the display name, email address or AD Username of the DL."
  [string]$AliasNameSet = Read-Host "Enter the display name, email address or AD Username of new Alias for this DL."
  Set-DistributionGroup -Identity $DLAliasName -Alias $AliasNameSet
  Pause
} 
# This function lets you change a DL's Email Address and Alias with Policy Override.
function DLSetAliasOVR
  {Clear-Host
  [string]$DLEMAlName = Read-Host "Enter the display name, email address or AD Username of the DL."
  [string]$DLEMAlSMTPAliasSet = Read-Host "Enter the display name, email address or AD Username of the USER changing access to this calendar"
  [string]$DLEMAlAliasSet = Read-Host "Enter the display name, email address or AD Username of the USER changing access to this calendar"
  Set-DistributionGroup -Identity $DLEMAlName -PrimarySMTPAddress $DLEMAlSMTPAliasSet -Alias $DLEMAlAliasSet -EmailAddressPolicyEnabled $false
  Pause
}
# This function lets you change a DL's SAM Account Name.
function DLSetSAMName
  {Clear-Host
  [string]$DLSAMName = Read-Host "Enter the display name, email address or AD Username of the DL."
  [string]$DLNewSAMName = Read-Host "Enter the display name, email address or AD Username of new SAM Account Name for this DL."
  Set-DistributionGroup -Identity $DLSAMName -SamAccountName $DLNewSAMName
  Pause
}
# This function lets you change a DL's Email, Alias, NAme (internal), and Display Name .
function DLSetAllOVR
  {Clear-Host
  [string]$DLSetAllEmail = Read-Host "Enter the display name, email address or AD Username of the DL."
  [string]$DLSetAllName = Read-Host "Enter the display name, email address or AD Username of new SAM Account Name for this DL."
  [string]$DLSetAllDispName = Read-Host "Enter the display name, email address or AD Username of the DL."
  [string]$DLSetAllAlias = Read-Host "Enter the display name, email address or AD Username of the DL."
  Set-DistributionGroup $DLSetAllEmail -Name $DLSetAllName -Alias $DLSetAllAlias -DisplayName "$DLSetAllDispName" -EmailAddressPolicyEnabled $false
  Pause
}
# This function lets you change a DL's Email with Naming Policy Override.
function DLSetEmailOVR                                                                  
  {Clear-Host
  [string]$DLNameOVR = Read-Host "Enter the display name, email address or AD Username of the DL."
  [string]$DLNewEmailOVR = Read-Host "Enter the FULL email address for this DL."
  t-DistributionGroup -Identity $DLNameOVR -PrimarySmtpAddress $DLNewEmailOVR -EmailAddressPolicyEnabled $false
  Pause
}
#####################################################################
# ------------ FUNCTIONS FOR DL UTILITIES --------------            #
#####################################################################

# This function sets the DLs 'Send As' permission for a user or group. No notifications are sent
function SetDLSendAs
  {Clear-Host
  [string]$DLSendAsName = Read-Host "Enter the display name, email address or AD Username of the CALENDAR"
  [string]$SetSendAsUser = Read-Host "Enter the display name, email address or AD Username of the USER changing access to this calendar"
  Add-RecipientPermission -Identity $DLSendAsName -Trustee $SetSendAsUser -AccessRights SendAs -Confirm:$false
  Pause
}

# This function sets the DL to where only members and maintainers can communicate.
function RestrictDL
  {Clear-Host
  [string]$DLRestrict = Read-Host "Enter the display name, email address or AD Username of the DL to Restrict."
  Get-DistributionGroup “cp-leads” | select -ExpandProperty AcceptMessagesOnlyFromSendersOrMembers | ft Name
  Pause
}

# This function sets the DL access level for adding new members.
function SetDLAccess
  {Clear-Host
  [string]$DLAccessName = Read-Host "Enter the display name, email address or AD Username of the DL to Modify."
  [string]$DLAccessIn = Read-Host "For a member to join, is access Open, Closed, or ApprovalRequired?"
  [string]$DLAccessOut = Read-Host "For a member to depart, is access Open, Closed, or ApprovalRequired?"
  Set-DistributionGroup $DLAccessName –MemberJoinRestriction $DLAccessIn –MemberDepartRestriction $DLAccessOut
  Pause
}

# This function outputs the members of a DL in CSV format, up to 1,000 members
function DLCSVMemList
{Clear-Host
  Write-Host "NOTE: The CSV file will be saved as C:\Distribution-List-Members.csv"
  [string]$DLMemListName = Read-Host "Enter the display name, email address or AD Username of the DL to access."
  Get-DistributionGroupMember -Identity $DLMemListName | Select-Object Name, PrimarySMTPAddress | Export-CSV "C:\MET\$DLMemListName-Distribution-List-Members-$(get-date -f MM-dd-yyyy_hh.mm.ss).csv" -NoTypeInformation -Encoding UTF8
  Pause
}

<# # This function outputs the members of a DL in CSV format, All members, over 1,000 members
function DLCSV1KList
{Clear-Host
  Write-Host "NOTE: The CSV file will be saved as C:\Distribution-List-Members.csv"
  [string]$DL1KListName = Read-Host "Export DL Members Names and Email Address to CSV File, saved in C:\, All Users (over 1000)."
  Get-DistributionGroupMember -Identity $DL1KListName -ResultSize Unlimited | Select-Object Name, PrimarySMTPAddress | Export-CSV "C:\Distribution-List-Members.csv" -NoTypeInformation -Encoding UTF8
  Pause
} #>

# This function lists for a DL, the SAM Account Name, Organizational Unit, Organizational Unit Root, and OrganizationID
function DLNameDetails
  {Clear-Host
  [string]$DLDetails = Read-Host "Enter the display name, email address or AD Username of the DL to list details for."
  Get-DistributionGroup $DLDetails | select sam*, org*
  Pause
}

#####################################################################
# ---------- Functions for Session Management --------------
#####################################################################

# For Future Usage

# This function starts a session in Exchange Online to carry out the function for carrying out Powershell actions
# function Show-Session-Global
# { Clear-Host
#   # Writes a message on screen to confirm that the login script to create a new session has started
#   Write-Host "Running the requested script..." -ForegroundColor DarkBlue -BackgroundColor Gray
#   # This sets the variable Session in the global scope and is the command that you would run before commands usually when you do something with Exchange in Powershell
#   $global:Session =
#     New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $global:UserCredential -Authentication "Basic" -AllowRedirection;
#     # The below line then imports that session using the Variable we have just created to allow us to send the output of this function into the option we need
#     # This is why the Menu Options contain these functions so this login part can be changed just here if this changes in future  
#     Import-PSSession -Session $global:Session
# }

# This should enable a connection to Exchange Management Console (On-Prem) upon running the script.
# Connect to Exchange Management Console (On-Prem)
Enable-EMS

#####################################################################
# ---------- Functions for Menus and Submenus ----------
#####################################################################
function Show-MainMenu-Toolbox
{
   # This is the create function to display the one level flat Menu and the Opt variables are called later for the the Write-Hosts.
   # Just change the text in the Double Speech marks and it will change in the menu.
   # You can add as many options as you want here make sure you wrap the new string in 2 round brackets above and below.
   # Leave Title as it is in the String Variable here and also below as you will call this later in the Do for User Input later on as -Title.
   # Each of the code blocks have variables that should relate to the function of that section for easier readibility and to make any 
   # changes to this script easier in the future.
    param (
        [string]$MainTitle = "My Menu"
          )
          (
        [string]$MainOpt1 = "Exchange Distribution List Creation and User Administration"
          )
          (
        [string]$MainOpt2 = "Exchange Distribution List Administration"
          )
          (
        [string]$MainOpt3 = "Exchange Distribution List Utilities"
          )
    Clear-Host
    Write-Host "================== $MainTitle =================="
    Write-Host ""
    Write-Host " 1: $MainOpt1"
    Write-Host " 2: $MainOpt2"
    Write-Host " 3: $MainOpt3"
    #Write-Host " 4: $MainOpt4"
    #Write-Host " 5: $MainOpt5"
    #Write-Host " S: Turn On/Off Sounds"
    Write-Host ""
    Write-Host "=======Type 'Q' to Quit The Application========"
}
function Show-SubMenu-DLAdmin
{
   # This is the create function to display the DL Sub Menu and the DLAcctOpt variables are called later for the the Write-Hosts.
   # The Title has also been changed in the Write-Host and Parameter to CalTitle so that this menu has it's own Title
    param (
        [string]$CalTitle = "My Menu"
          )
          (
        [string]$DLAcctOpt1 = "Show full details of a DL"
          )
          (
        [string]$DLAcctOPT2 = "Set the DL's Display Name"
          )
          (
        [string]$DLAcctOpt3 = "Set the DL's Display Name (Group Naming Policy Override)"
          )
          (
        [string]$DLAcctOpt4 = "Set the DL's Alias"
          )
          (
        [string]$DLAcctOpt5 = "Set the DL's Email Address and Name (Group Naming Policy Override)"
          )
          (
        [string]$DLAcctOpt6 = "Set the DL's SAM Account Name"
          )
          (
        [string]$DLAcctOpt7 = "Set the DL's new Name (Internal), Display Name, and Alias. (Policy Override)"
          )
          (
        [string]$DLAcctOpt8 = "Set a new email for a DL. (Group Naming Policy Override)"
          )
          <# (
        [string]$DLAcctOpt9 = ""                           
          ) #>
    Clear-Host
    Write-Host "================== $CalTitle =================="
    Write-Host ""
    Write-Host " 1: $DLAcctOpt1"
    Write-Host " 2: $DLAcctOpt2"
    Write-Host " 3: $DLAcctOpt3"
    Write-Host " 4: $DLAcctOpt4"
    Write-Host " 5: $DLAcctOpt5"
    Write-Host " 6: $DLAcctOpt6"
    Write-Host " 7: $DLAcctOpt7"
    Write-Host ""
}
function Show-SubMenu-DLUsers
{
   # This is the create function to display the calendar Sub Menu and the DLUserOpt variables are called later for the the Write-Hosts.
   # The Title has also been changed in the Write-Host and Parameter to DLUserTitle so that this menu has it's own Title
    param (
        [string]$DLUserTitle = "My Menu"
          )
          (
        [string]$DLUserOpt1 = "Set up a new Distribution List."
          )
          (
        [string]$DLUserOpt2 = "Add a User to a Distribution List."
          )
          (
        [string]$DLUserOpt3 = "Remove a User from a Distribution List.."
          )
          (
        [string]$DLUserOpt4 = "Set a User as Maintainer for a Distribution List."
          )
          (
        [string]$DLUserOpt5 = "Set a User as Owner for a Distribution List."
          )
          (
        [string]$DLUserOpt6 = "Remove an Owner or Maintainer of a Distribution List. Will still be a member after this action is taken."
          )
          (
        [string]$DLUserOpt7 = "List all DL Members."
          )
          (
        [string]$DLUserOpt8 = "List All DL Maintainers."
          )
          (
        [string]$DLUserOpt9 = "List Maintainers, but with Full AD Path for names."
          )
          (
        [string]$DLUserOpt10 = "Turn off adherence to Meta Group Naming Policy."
          )
    Clear-Host
    Write-Host "================== $DLUserTitle =================="
    Write-Host ""
    Write-Host "  1: $DLUserOpt1"
    Write-Host "  2: $DLUserOpt2"
    Write-Host "  3: $DLUserOpt3"
    Write-Host "  4: $DLUserOpt4"
    Write-Host "  5: $DLUserOpt5"
    Write-Host "  6: $DLUserOpt6"
    Write-Host "  7: $DLUserOpt7"
    Write-Host "  8: $DLUserOpt8"
    Write-Host "  9: $DLUserOpt9"
    Write-Host " 10: $DLUserOpt10"
    Write-Host ""
}
function Show-SubMenu-DLUtils
{
   # This is the create function to display all the scripts in a Sub Menu and the DLUtlOpt variables are called later for the the Write-Hosts.
   # The Title has also been changed in the Write-Host and Parameter to FullTitle so that this menu has it's own Title
    param (
        [string]$FullTitle = "My Menu"
          )
          (
        [string]$DLUtlOpt1 = "Sets 'Send As' permission for user or group."
          )
          (
        [string]$DLUtlOpt2 = "Restrict DL so ony Users and accept message from Senders or Members."
          )
          (
        [string]$DLUtlOpt3 = "Set how a user can join a DL. Three ways: Open, Closed, ApprovalRequired (no space)"
          )
          (
        [string]$DLUtlOpt4 = "Export DL Members Names and Email Address to CSV File, saved in C:\, up to 1000"
          )
          <# (
        [string]$DLUtlOpt5 = "Export DL Members Names and Email Address to CSV File, saved in C:\, All Users (over 1000)"
          ) #>
          (
        [string]$DLUtlOpt6 = "Lists for a DL, the SAM Account Name, Organizational Unit, Organizational Unit Root, and OrganizationIDs."
          )
          <# (
        [string]$FullOpt7 = "List all Users that have Delegate Permissions on a Calendar"
          )
          (
        [string]$FullOpt8 = "Add a User to a Calendar"
          )
          (
        [string]$FullOpt9 = "Remove a User from a Calendar"
          )
          (
        [string]$FullOpt10 = "Change a User's Access Permissions to a Calendar"
          )
          (
        [string]$FullOpt11 = "Remove an Individual User From Multiple Calendars"
          )
          (
        [string]$FullOpt12 = "Update Your Saved Credentials"
          ) #>
    Clear-Host
    Write-Host "================== $FullTitle ==================" 
    Write-Host ""
    Write-Host " 1: $DLUtlOpt1"
    Write-Host " 2: $DLUtlOpt2"
    Write-Host " 3: $DLUtlOpt3"
    Write-Host " 4: $DLUtlOpt4"
    #Write-Host " 5: $DLUtlOpt5"
    Write-Host " 5: $DLUtlOpt6"
    Write-Host ""
}
# --------------------------------- END OF FUNCTION LIST -----------------------------------------------------
#
# Do Loop for Menu containing user switches and a nested looping system
# This also calls the Menus with a changed positional paremeter switch so you can change these next to e.g. "-MainTitle" and "-MailTitle"
# Use a new variable for each menu or it'll loop indefinitely or just have no menu e.g. $MainTile and $MailTitle
# You can add as many options as you want to each layer and then branch out from there
# Functions have been used so they can be called instead of code blocks in the User Switches for easier readibility and editing of the options later
do
{
    Clear-Host
    Show-MainMenu-Toolbox -MainTitle "Exchange Distribution List Toolbox v1.0"
    $UserMain = Read-Host "Please enter your selection from the list"
    switch ($UserMain)
     {
           "1" {
                
                do
                    {
                        Clear-Host
                        Show-SubMenu-DLUsers -DLUserTitle "Exchange Distribution List Creation and User Administration"
                        $UserSubMail = Read-Host -Prompt 'Enter 1 - 10 or B for Back'
                        Switch ($UserSubMail)
                         {
                            '1' {Add-DL-Single}
                            '2' {Add-DLUser-Single}
                            '3' {Remove-DLUser-Single}
                            '4' {Add-DLOwn-Single}
                            '5' {Add-DLOwn-New}
                            '6' {Rem-DLOwn-Single}  
                            '7' {ListDLMembers}
                            '8' {ListDLOwn}
                            '9' {ListDLOwnFull}
                            '10' {MetaPolicyBypass}
                         }
                         #pause
                    }
                    until ($UserSubMail -eq "b")
           } "2" {
               Clear-Host
               do
                {
                    Clear-Host
                    Show-SubMenu-DLAdmin -CalTitle "Exchange Distribution List Account Administration"
                    $UserSubCal = Read-Host -Prompt 'Enter 1 - 6 or B for Back'
                    Switch ($UserSubCal)
                     {
                        '1' {DLFullDetails}          
                        '2' {DLDisplayName}
                        '3' {DLDispNameOVR} 
                        '4' {DLSetAlias}     
                        '5' {DLSetAliasOVR}
                        '6' {DLSetSAMName}
                        '7' {DLSetAllOVR}
                        '8' {DLSetEmailOVR}
                     }
                     #pause
                }
                until ($UserSubCal -eq "b")
           } "3" {
            Clear-Host
               do
               {
               Clear-Host
               Show-SubMenu-DLUtils -FullTitle "Exchange Distribution List Utilities"
               $UserSubFull = Read-Host -Prompt 'Enter 1 - 6 or B for Back'
               Switch ($UserSubFull)
                {
                   '1' {SetDLSendAs}        
                   '2' {RestrictDL}
                   '3' {SetDLAccess}
                   '4' {DLCSVMemList}     
                   #'5' {DLCSV1KList}
                   '5' {DLNameDetails}
                   #'7' {UserCalList}
                   #'8' {Add-Cal-Single}
                   #'9' {UserBoxList}
                  #'10' {Set-Cal-Single}
                  #'11' {Remove-Cal-Multi}
                  #'12' {Update-Creds}
                  }
                  #pause
                }
                until ($UserSubFull -eq "b")
          } "4" {Clear-Host
                 Update-Creds
          }"q" {
                return
           }
     }
     #pause
}
until ($UserMain -eq "q")
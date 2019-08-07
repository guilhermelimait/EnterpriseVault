<#
  .DEVELOPED BY: 
    Guilherme Lima
  .PLATFORM:
    O365
  .WEBSITE:
    http://solucoesms.com.br
  .LINKEDIN:
    https://www.linkedin.com/in/guilhermelimait/
  .DESCRIPTION:
    This script will check in which AD Group they are (active or inactive users) and based on their custom attribute 5 add to the correct
    Enterprise Vault policy
#>


import-module ActiveDirectory

#insert the name of the servers to check the last mailboxes created
$servers = @("server-mbx-01","server-mbx-02","server-mbx-03")
#insert the days of check (it will vary from your local schedule
$days = 15

#EV groups to add/remove the user to/from
$evGroupInactive  = "EnterpriseVault_Inactive_Mailboxes_SA"
$evGroupActive = "Enterprisevault_All_Users_SA"
$evGroupShortcut = "Enterprisevault_Shortcut_Group_SA"
$adInactiveGroup = (Get-ADGroup $evGroupInactive).distinguishedName
$adStandardGroup = (Get-ADGroup $evGroupActive).distinguishedName
$adShortcutGroup = (Get-ADGroup $evGroupShortcut).distinguishedName

#Creating the input and output log records
$Input = new-item -type file -name "EV-Input-$(get-date -f MM-dd-yyyy_HH_mm_ss).csv"
$output = new-item -type file -name "EV-Output$(get-date -f MM-dd-yyyy_HH_mm_ss).csv"
add-content $Input -value "SamAccountName,LinkedMasterAccount"
#Identify the mailboxes informed on servers
foreach($server in $servers){
	Get-Mailbox -server $server -Resultsize unlimited | Where {$_.WhenCreated -gt (Get-Date).AddDays(-$days)} | select-object SamAccountName, LinkedMasterAccount | add-content $Input
}

#$input = ".\teste.csv"
#$output = ".\saida.csv"

#Remove the unnecessary characters on file
(get-content $Input) | foreach-object { 
	$_ -replace '"','' `
	-replace ';',',' `
	-replace '@{SamAccountName=','' `
	-replace ' LinkedMasterAccount=','' `
	-replace '}','' `
	} | set-content $Input
	$inputcontent = select-string -pattern "\w" -path $Input | foreach-object{$_.line}
	$inputcontent | set-content $Input

#Give permissions from legacy AD account to mailbox	
add-content $output "`n`n+ Adding permissions to mailbox +`n`n"
import-csv $Input | foreach {
	Add-MailboxPermission  $_.SamAccountName  -User  $_.LinkedMasterAccount  -AccessRights  FullAccess
	$SAMUser = $_.SamAccountName
	$LinkedUser = $_.LinkedMasterAccount
	add-content $output "Permission from legacy account $LinkedUser to the mailbox $SAMUser created successfully`n"
}

#Check if user is member of the active group, if not, the user is added 
add-content $output "`n`n+ Adding to correct EV Group +`n`n"

import-csv $Input | ForEach {
	$SAMUser = $_.SamAccountName
	$CA5 = get-mailbox $_.SamAccountName | select -expandproperty CustomAttribute5

	Function EVStandard { #Function that give permissions to standard ev policy
	#	If ((Get-ADUser $SAMUser -Properties MemberOf | Select -ExpandProperty MemberOf) -contains $adStandardGroup){
	#		if((Get-ADUser $SAMUser -Properties MemberOf | Select -ExpandProperty MemberOf) -contains $adShortcutGroup){
				remove-ADGroupMember -Identity $adShortcutGroup -Members $SAMUser -Confirm:$False	
	#		}else{
			add-ADGroupMember -Identity $adStandardGroup -Members $SAMUser
			add-content $output "Adding $SAMUser with RU Code $CA5 to the group $evGroupActive`n"
			write-host "Adding $SAMUser with RU Code $CA5 to the group $evGroupActive`n"
	#	   }
	#	}
	}
	Function EVShortcut{ #Function that give permissions to shortcut ev policy
	#	If ((Get-ADUser $SAMUser -Properties MemberOf | Select -ExpandProperty MemberOf) -contains $adShortcutGroup){
	#		If ((Get-ADUser $SAMUser -Properties MemberOf | Select -ExpandProperty MemberOf) -contains $adStandardGroup){
				remove-ADGroupMember -Identity $adStandardGroup -Members $SAMUser -Confirm:$False
	#	   }else{
				add-ADGroupMember -Identity $adShortcutGroup -Members $SAMUser
				add-content $output "Adding $SAMUser with RU Code $CA5 to the group $evGroupShortcut`n"
				Write-host "Adding $SAMUser with RU Code $CA5 to the group $evGroupShortcut`n"
	#	   }
	#	}
	}
	
	switch -wildcard ($CA5){ #Function that will check the RU code from new users
		"*VALOR1*"{EVShortcut}
		"*VALOR2*"{EVShortcut}
		"*VALOR3*"{EVShortcut}
		"*VALOR4*"{EVShortcut}
		"*VALOR5*"{EVShortcut}
		"*VALOR6*"{EVStandard}
		"*VALOR7*"{EVShortcut}
		"*VALOR8*"{EVShortcut}
		"*VALOR9*"{EVStandard}
		"*VALOR10*"{EVShortcut}
		"*VALOR11*"{EVStandard}
		"*VALOR12*"{EVShortcut}
	}
}

add-content $output "`n`n+ Removing from inactive EV Group +`n`n"
#Check if user is member of the inactive group, if yes, the user is removed 
import-csv $Input | ForEach {
	$SAMUser = $_.SamAccountName
	$CA5 = get-mailbox $_.SamAccountName | select -expandproperty CustomAttribute5
	If ((Get-ADUser $SAMUser -Properties MemberOf | Select -ExpandProperty MemberOf) -contains $adInactiveGroup){
		remove-ADGroupMember -Identity $adInactiveGroup -Members $SAMUser -Confirm:$False
		add-content $output "Removing $SAMUser with RU Code $CA5 of the group $evGroupInactive`n"
	}
}

#sleep 100

$body = get-content($Output) -delimiter "\n"
$param = @{
    SmtpServer = 'MAILSERVER01.DOMAIN.COM'
    From = 'EV.Monitoring@domain.com'
    To = 'guilherme.lima@domain.com.br'
    Subject = 'EV-Automated'
    Body = $body
    #Attachments = 'D:\articles.csv'
}
 
Send-MailMessage @param


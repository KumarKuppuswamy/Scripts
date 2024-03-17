[array]$val = $null
$users = Get-Content "2009scjsurveyhelpdesk@wtwco.com"
if(!$users)
{
	Write-Host "No data"
}Else
	{
		ForEach($user in $users)
		{
			$mbxs = Get-Recipient "$user"
			if(!$mbxs)
			{
				Write-Host "No Object"
			}Else
				{
					ForEach($mbx in $mbxs)
					{
						$mbs = Get-Mailbox "$user" | Where-Object {$_.RecipientTypeDetails -eq "SharedMailbox"} | select DisplayName,PrimarySMTPAddress,IsDirSynced,@{name="EmailAddresses";expression={$_.EmailAddresses -join ";"}},@{name="Manager";expression={(Get-User $_.Name).Manager}},@{Name="GrantSendOnBehalfTo";expression={($_.GrantSendOnBehalfTo | Get-User | Select-Object -ExpandProperty WindowsEmailAddress) -join "|"}},UserPrincipalName
						if(!$mbs)
						{
							Write-Host "$user - check manually"
						}Else
							{
								ForEach($mb in $mbs)
								{
									$perm = Get-MailboxPermission $mb.PrimarySMTPAddress | where-object {($_.user -notlike "S-1-5-21") -and ($_.user -notlike "NT AUTHORITY\SELF") -and ($_.isinherited -eq $false)} 
									$upn = $mb.UserPrincipalName
									$lic = Get-MsolUser -UserPrincipalName $upn 									

									$upn
										if($upn -like "*@willistowerswatson.com*")
										{
										    $accs = Get-ADUser -Filter {UserPrincipalName -eq $upn} -Properties * -Server INT.DIR.WILLIS.COM | select UserAccountControl,@{name="Accounthistory";expression={$_.Accountnamehistory -join ";"}},msExchRecipientTypeDetails,msExchRemoteRecipientType
												
												$obj1 = New-Object PSObject
												$obj1 | Add-Member NoteProperty -Name "DisplayName" -Value $mb.DisplayName
												$obj1 | Add-Member NoteProperty -Name "PrimarySMTPaddress" -Value $mb.PrimarySMTPaddress
												$obj1 | Add-Member NoteProperty -Name "Sync" -Value $mb.IsDirSynced
												$obj1 | Add-Member NoteProperty -Name "Manager" -Value $mb.Manager
												$obj1 | Add-Member NoteProperty -Name "EmailAddress" -Value $mb.EmailAddresses
												$obj1 | Add-Member NoteProperty -Name "GrantSendOnBehalfTo" -Value $mb.GrantSendOnBehalfTo
												$obj1 | Add-Member NoteProperty -Name "Permissions" -Value ($perm.user -join ";")
												$obj1 | Add-Member NoteProperty -Name "UserAccountControl" -Value $accs.UserAccountControl
												$obj1 | Add-Member NoteProperty -Name "UserPrincipalName" -Value $upn
												$obj1 | Add-Member NoteProperty -Name "msExchRecipientTypeDetails" -Value $accs.msExchRecipientTypeDetails
												$obj1 | Add-Member NoteProperty -Name "msExchRemoteRecipientType" -Value $accs.msExchRemoteRecipientType
												$obj1 | Add-Member NoteProperty -Name "License" -Value $lic.isLicensed
												$val += $obj1  
										
										}Elseif($upn -like "*@towerswatson.com*")
											{
												$accs = Get-ADUser -Filter {UserPrincipalName -eq $upn} -Properties * -Server INTERNAL.TOWERSWATSON.COM | select UserAccountControl,@{name="Accounthistory";expression={$_.Accountnamehistory -join ";"}},msExchRecipientTypeDetails,msExchRemoteRecipientType
												
												$obj1 = New-Object PSObject
												$obj1 | Add-Member NoteProperty -Name "DisplayName" -Value $mb.DisplayName
												$obj1 | Add-Member NoteProperty -Name "PrimarySMTPaddress" -Value $mb.PrimarySMTPaddress
												$obj1 | Add-Member NoteProperty -Name "Sync" -Value $mb.IsDirSynced
												$obj1 | Add-Member NoteProperty -Name "Manager" -Value $mb.Manager
												$obj1 | Add-Member NoteProperty -Name "EmailAddress" -Value $mb.EmailAddresses
												$obj1 | Add-Member NoteProperty -Name "GrantSendOnBehalfTo" -Value $mb.GrantSendOnBehalfTo
												$obj1 | Add-Member NoteProperty -Name "Permissions" -Value ($perm.user -join ";")
												$obj1 | Add-Member NoteProperty -Name "UserAccountControl" -Value $accs.UserAccountControl
												$obj1 | Add-Member NoteProperty -Name "UserPrincipalName" -Value $upn
												$obj1 | Add-Member NoteProperty -Name "msExchRecipientTypeDetails" -Value $accs.msExchRecipientTypeDetails
												$obj1 | Add-Member NoteProperty -Name "msExchRemoteRecipientType" -Value $accs.msExchRemoteRecipientType
												$obj1 | Add-Member NoteProperty -Name "License" -Value $lic.isLicensed

												$val += $obj1  
											}Else
												{
												  Write-Host "No UPN"
												 }
								}	
											
							}				
					
					}

				}
		}
	}

$val
$val | Export-csv Testing.csv -notypeinformation
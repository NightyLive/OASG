#Initialisation des variables
$SignatureName = "DransEnergie" #Nom de la signature
$FileSource = "\\domain.local\NETLOGON\signature\DransEnergie.html"
$FileTarget = "$env:APPDATA\Microsoft\Signatures\$SignatureName.html"
$RegBasePath = "HKCU:\Software\Microsoft\Office" #Dossier de base
$ErrorAction = "SilentlyContinue"

#Initialisation des fonctions
$shell = new-object -comobject "WScript.Shell"
$UserName = $env:username
$Filter = “(&(objectCategory=User)(samAccountName=$UserName))” 
$Searcher = New-Object System.DirectoryServices.DirectorySearcher 
$Searcher.Filter = $Filter
$ADUserPath = $Searcher.FindOne() 
$ADUser = $ADUserPath.GetDirectoryEntry() 

#Récupération des informations de l'utilisateur
$ADDisplayName = $ADUser.DisplayName
$ADEmailAddress = $ADUser.mail
$ADTitle = $ADUser.title
$ADTelePhoneNumber = $ADUser.TelephoneNumber
$ADTelePhoneNumberInternal = $ADUser.TelephoneNumber
$ADTelePhoneNumber = $ADTelePhoneNumber.ToString()
if (($ADTelePhoneNumber) -And ($ADTelePhoneNumber.Length -lt 4)) { $ADTelePhoneNumber = "+41 27 111 22 22" } #Si numéro est interne -> remplacer par numéro de l'entreprise
$ADMobileNumber = $ADUser.mobile
$ADDepartment = $ADUser.department
$ADCompany = $ADUser.company

#Affichage des informations
clear
echo ""
echo ""
echo "  Utilisateur: $UserName"
echo "  Fonction/Titre: $ADTitle"
echo "  Nom/Prénom: $ADDisplayName"
echo "  Email:  $ADEmailAddress"
echo "  Telephone interne: $ADTelePhoneNumberInternal"
echo "  Telephone externe: $ADTelePhoneNumber"
echo "  Telephone mobile: $ADMobileNumber"
echo "  Departement: $ADDepartment"
echo "  Société: $ADCompany"
echo ""

#Confirmation de la création de la signature
$choice = $shell.popup("Cette procédure va remplacer votre signature `"$SignatureName`", veuillez vérifier que ces informations sont correctes. Voulez-vous continuer ?",0,"Attention !",4+32)
if ($choice -eq '7') { exit }

#Demande pour le numéro de natel
$ShowMobile = $shell.popup("Voulez-vous afficher votre numéro de natel dans la signature ?",0,"Vie privée",4+32)
if ($ShowMobile -eq '7') { $ADMobileNumber = "&nbsp;" }
if ($ShowMobile -eq '7') { $TextReplace = "&nbsp;" } else { $TextReplace = "<strong>P</strong>" }

#Generation de la signature
Remove-Item $FileTarget #Suppression du fichier sinon il n'est pas mis à jour
Get-Content $FileSource -Encoding UTF8 | ForEach-Object {
    $_ -replace 'Prénom Nom', "$ADDisplayName" `
       -replace 'Fonction', "$ADTitle" `
       -replace '\+41 11 222 33 xx', "$ADTelePhoneNumber" `
       -replace '\+41 22 xxx xx xx', "$ADMobileNumber" `
       -replace '<strong>P</strong>', "$TextReplace"
    } | Set-Content $FileTarget -Encoding UTF8 #-Force #Si -Force plus besoin de Remove-Item?
#Start-Process $FileTarget #Ouvre le fichier dans le navigateur

#Defini la signature pour les différentes versions d'office
$RegVersion = "14.0"
Remove-ItemProperty -Path "$RegBasePath\$RegVersion\Outlook\Setup" -Name First-Run -Force -EA $ErrorAction #-Confirm
Set-ItemProperty -Path "$RegBasePath\$RegVersion\Common\MailSettings" -Name NewSignature -Value $SignatureName -Force -EA $ErrorAction
Set-ItemProperty -Path "$RegBasePath\$RegVersion\Common\MailSettings" -Name ReplySignature -Value $SignatureName -Force -EA $ErrorAction
$RegVersion = "15.0"
Remove-ItemProperty -Path "$RegBasePath\$RegVersion\Outlook\Setup" -Name First-Run -Force -EA $ErrorAction #-Confirm
Set-ItemProperty -Path "$RegBasePath\$RegVersion\Common\MailSettings" -Name NewSignature -Value $SignatureName -Force -EA $ErrorAction
Set-ItemProperty -Path "$RegBasePath\$RegVersion\Common\MailSettings" -Name ReplySignature -Value $SignatureName -Force -EA $ErrorAction
$RegVersion = "16.0"
Remove-ItemProperty -Path "$RegBasePath\$RegVersion\Outlook\Setup" -Name First-Run -Force -EA $ErrorAction #-Confirm
Set-ItemProperty -Path "$RegBasePath\$RegVersion\Common\MailSettings" -Name NewSignature -Value $SignatureName -Force -EA $ErrorAction
Set-ItemProperty -Path "$RegBasePath\$RegVersion\Common\MailSettings" -Name ReplySignature -Value $SignatureName -Force -EA $ErrorAction
#clear

#Message avertissement
$shell.popup("Votre signature nouvelle est définie, il est nécessaire de redémarrer Outlook afin que les changements soient pris en compte !",0,"Merci",0x0)

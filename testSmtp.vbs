'**************************
'Declarations de variables
'**************************

Dim destinataire	'Destinataire du mail
Dim smtp		'Serveur SMTP
Dim username		'Login du serveur SMTP
Dim password		'Mot de passe du serveur SMTP
Dim port		'Port du serveur SMTP
Dim serveur 		'Nom du serveur qui test le SMTP


'************************
'Declarations des objets
'************************

Set objEmail = CreateObject("CDO.Message") 'Objet pour l'envoi de mail


'******************************
'Initialisations des variables
'******************************

destinataire	= "destinataire@domain.com"	'L'adresse destinataire du mail
smtp 			= "smtp.domain.com"			'Nom du serveur SMTP
username 		= "username"				'Login du serveur SMTP
password 		= "password"				'Mot de pass du serveur SMTP
from			= "from@domain.com"
sender			= "TEST SMTP"
organization	= "ORGANISATION"
port 			= 25						'Port du serveur SMTP
serveur 		= "MON SERVEUR"				'Nom du serveur qui test l'envoi via le SMTP


'**********
'Fonctions
'**********


'Fonction qui permet l'envoi de mail
Function sendMail(contenu)
	Dim iMsg, iConf, Flds

	Set iMsg = CreateObject("CDO.Message")
	Set iConf = CreateObject("CDO.Configuration")
	Set Flds = iConf.Fields

	schema = "http://schemas.microsoft.com/cdo/configuration/"
	Flds.Item(schema & "sendusing") = 2
	Flds.Item(schema & "smtpserver") = smtp
	Flds.Item(schema & "smtpserverport") = port
	Flds.Item(schema & "smtpauthenticate") = 1
	Flds.Item(schema & "sendusername") = username
	Flds.Item(schema & "sendpassword") = password
	Flds.Update

	With iMsg
	.To = destinataire
	.From = from
	.Subject = "TEST SMTP NCI"
	.HTMLBody = contenu
	.Sender = sender
	.Organization = organization
	Set .Configuration = iConf
	SendEmailGmail = .Send
	End With

	set iMsg = nothing
	set iConf = nothing
	set Flds = nothing
End Function

'**************
'Procédure Sub
'**************

Sub testSMTP
	
	sendMail ("Le serveur "&serveur&" arrive a lancer un mail avec le serveur smtp : " &smtp&" sur le port : "&port&" avec le login : "&username&" et le mot de passe : "&password)

End Sub

'***************************************************************************************************************************************************

testSMTP


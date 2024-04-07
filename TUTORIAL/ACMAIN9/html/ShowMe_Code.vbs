Dim L_ShowMe_ErrorMessage(10)
Dim msouierrSuccess
Dim msouierrFail
Dim msouierrNotValidId
Dim msouierrNoDialog
Dim msouierrWrongDialog
Dim msouierrAdminDisabled
Dim msouierrDisabled
Dim msouierrOn
Dim msouierrOff
Dim msouierrUnknown
Dim msouierrAppModal
Dim sSecurityMSG

'LOCALIZABLE: There needs to be a "Dim" statement for each "L_Security??_ErrorMessage" constant:
	Dim L_SecurityT1_ErrorMessage
	Dim L_SecurityT2_ErrorMessage
	Dim L_SecurityT3_ErrorMessage
	Dim L_SecurityE1_ErrorMessage
	Dim L_SecurityE2_ErrorMessage
	Dim L_SecurityE3_ErrorMessage
	Dim L_SecurityE4_ErrorMessage
	Dim L_App_DialogTitle


'------------------------
' Detect if IE is >=4.0 -
'------------------------
Function DetectBrowserVersion()
   Dim iVersion
   iVersion=navigator.appversion
   If Left(iVersion,1)>=4 Then
       DetectBrowserVersion="4.0>"
   Else
       DetectBrowserVersion="3.0x"
   End if
End Function

'----------------------------------------
' Display the appropriate error message -
'----------------------------------------

Sub DisplayError(retVal)
    Call InitConstants
    Msgbox L_ShowMe_ErrorMessage(retVal), 4144, L_APP_DialogTitle
End Sub

Sub InitConstants()
	'NON-LOCALIZABLE: Return values from OUACtrl.ocx. Used by the "Show Me" jumps.
   msouierrSuccess=0
   msouierrFail=1
   msouierrNotValidId=2
   msouierrNoDialog=3
   msouierrWrongDialog=4
   msouierrAdminDisabled=5
   msouierrDisabled=6
   msouierrOn=7
   msouierrOff=8
   msouierrUnknown=9
   msouierrAppModal=10

   'LOCALIZABLE: Possible error messages displayed to the user, in order of frequency
   L_ShowMe_ErrorMessage(msouierrFail)="Impossibile completare automaticamente la procedura. Eseguire i passaggi manualmente."		'Message to display when there is a general Show Me failure
   L_ShowMe_ErrorMessage(msouierrAppModal)="È già visualizzata una finestra di dialogo."																'Message to display when the application is already displaying a dialog
   L_ShowMe_ErrorMessage(msouierrDisabled)="Impossibile completare automaticamente la procedura. Eseguire i passaggi manualmente."	'Message to display when the application is in a state that makes the feature disabled
   L_ShowMe_ErrorMessage(msouierrNoDialog)="Finestra di dialogo non visualizzata."																'Message to display when the application doesn't display the requested dialog
   L_ShowMe_ErrorMessage(msouierrAdminDisabled)="Il comando che si sta tentando di utilizzare è stato disattivato dall'amministratore."	'Message to display when the feature we're trying to use is disabled by an administrator
   L_ShowMe_ErrorMessage(msouierrWrongDialog)="La finestra di dialogo specificata non contiene l'opzione indicata."								'Message to display when we attempt to "click" a non-existent control on a dialog ("Do It" jumps -- not really used)
   L_ShowMe_ErrorMessage(msouierrNotValidId)="Errore interno. Eseguire i passaggi manualmente."										'Message to display when our Show Me code calls the wrong TCID (This should never display!)
   

   '***********************************************************************************
   'NOTE TO VENDORS: These string resources need to be the same as the ones in
   '              "IE 3.0x Fixes.xls"!!  Please do the following:
   '                 - Click the "Copy MsgBox" on the worksheet.
   '                 - Remove the existing lines below between "BEGIN" and "END"
   '                 - Paste the contents of the clipboard between "BEGIN" and "END"
   '                 - Insert a "Dim" statement at the top of this file for each
	'                   constant
	'
   '*** BEGIN *** BEGIN *** BEGIN *** BEGIN *** BEGIN *** BEGIN *** BEGIN *** BEGIN ***
L_SecurityT1_ErrorMessage="Impossibile visualizzare la procedura per la presenza di impostazioni di protezione "
L_SecurityT2_ErrorMessage="del browser troppo restrittive o per l'errata installazione del controllo ActiveX "
L_SecurityT3_ErrorMessage="Ouactrl.ocx."
L_SecurityE1_ErrorMessage="- Impostare un livello inferiore di protezione del browser"
L_SecurityE2_ErrorMessage="- Se questo messaggio viene visualizzato dopo l'impostazione di un livello inferiore,"
L_SecurityE3_ErrorMessage="  rivolgersi all'amministratore di sistema per la verifica dell'installazione del controllo"
L_SecurityE4_ErrorMessage="  ActiveX Ouactrl.ocx, posto nella cartella in cui è installato Microsoft Office"
sSecurityMSG=L_SecurityT1_ErrorMessage & chr(13) & L_SecurityT2_ErrorMessage & chr(13) & L_SecurityT3_ErrorMessage & chr(13) & chr(13) & L_SecurityE1_ErrorMessage & chr(13) & L_SecurityE2_ErrorMessage & chr(13) & L_SecurityE3_ErrorMessage & chr(13) & L_SecurityE4_ErrorMessage
L_App_DialogTitle="Guida di Microsoft Office"
   '*** END *** END *** END *** END *** END *** END *** END *** END *** END *** END ***
	
End Sub

Sub InitErrorMsgs()
	'Leave this here just in case we forgot to remove a call.
End Sub

'---------------------------------------------
' Mouse over, mouse out, rollover procedures -
'---------------------------------------------
Sub ColorSteps(sColor)
	If sColor="LightBlue" Then
		steps.className = "Highlight"
	Else
		steps.className = "Normal"
	End If
End Sub


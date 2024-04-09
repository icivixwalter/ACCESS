Attribute VB_Name = "FILE_MDL20_OggettoFileSistem"
Option Compare Database

'Descrizione
'Insieme in sola lettura di tutte le unit� disponibili.
'Osservazioni
'Per visualizzare nell'insieme Drives le unit� con
'supporto rimovibile non � necessario che il supporto sia inserito.
'Nel seguente codice viene spiegato come visualizzare
'un insieme Drives ed eseguire un ciclo nell'insieme utilizzando
'l'istruzione For Each...Next:



'//ROUTINE INVIDUDA I DISCHI ATTIVI INTERNI ED ESTERNI
Sub ShowDriveList()
    Dim fs, d, dc, s, n As Variant
    '//fs=Oggetto File sistem
    Set fs = CreateObject("Scripting.FileSystemObject")
    '//dc = oggetto Driver
    Set dc = fs.Drives
    '//Itera
    For Each d In dc
        
        s = s & d.DriveLetter & " : "
        If d.DriveType = Remote Then
            n = d.ShareName
        
        Else
            n = d.VolumeName
        End If
        '//vbCrLf= ritorno a capo
        s = s & n & vbCrLf
    Next
    MsgBox s
End Sub



'//PROPRIETA FILE SISTM
'//=================================================================================================//

'Propriet� IsReady
'Descrizione
'Restituisce True se l'unit� specificata � pronta;
'False in caso contrario.
'Sintassi
'oggetto.IsReady
'L 'argomento oggetto corrisponde sempre a un oggetto Drive.

'Osservazioni
'Nel caso di unit� con supporto rimovibile e unit� CD-ROM,
'IsReady restituisce True solo quando il supporto appropriato
'� inserito e pronto per l'accesso.

'Nel seguente codice viene spiegato come utilizzare la propriet�
'IsReady:

Private Sub ESEGUI()
ShowDriveInfo "C:\"
End Sub


Private Sub ShowDriveInfo(drvpath)
    Dim fs, d, s, t As Variant
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set d = fs.GetDrive(drvpath)
    Select Case d.DriveType
        Case 0: t = "Sconosciuta"
        Case 1: t = "Rimovibile"
        Case 2: t = "Fissa"
        Case 3: t = "Rete"
        Case 4: t = "CD-ROM"
        Case 5: t = "Disco RAM"
    End Select
    s = "Drive " & d.DriveLetter & ": - " & t
    If d.IsReady Then
        s = s & vbCrLf & "Unit� pronta."
    Else
        s = s & vbCrLf & "Unit� non pronta."
    End If
    MsgBox s
End Sub

'//PROPRIETA FILE SISTM   *** FINE ***
'//=================================================================================================//




'//***************************************************//
'//
'//             METODI
'//
'//***************************************************//


'//METODO DRIVEEXIST
'//<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<//
'Metodo DriveExists

'Descrizione

'Restituisce True se l'unit� specificata esiste; False in caso contrario.

'sintassi

'oggetto.DriveExists (unit�spec)

'La sintassi del metodo DriveExists � composta dalle seguenti parti:

'Parte Descrizione
'oggetto Obbligatoria. � sempre il nome di un oggetto FileSystemObject.
'unit�spec Obbligatoria. Lettera di unit� o percorso completo.


'Osservazioni

'Per le unit� con supporto rimovibile, il metodo DriveExists restituisce True anche se non � presente alcun supporto. Utilizzare la propriet� IsReady dell'oggetto Drive per determinare se un'unit� � pronta.


Private Sub ESEGUI_ShowDriveExist()
    ShowDriveExist "C:\"
End Sub


Private Sub ShowDriveExist(drvpath)
    Dim fs, d, s, t As Variant
    Dim dc As Object
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    '//dc = oggetto Driver
    Set dc = fs.Drives
    
    For Each d In dc
        
        s = s & d.DriveLetter & " : "
                
        If d.IsReady Then
            s = s & vbCrLf & "Unit� pronta."
        Else
            s = s & vbCrLf & "Unit� non pronta."
        End If
        
        If d.DriveType = Remote Then
            n = d.ShareName
        
        Else
            n = d.VolumeName
        End If
        '//vbCrLf= ritorno a capo
        s = s & n & vbCrLf
    Next
    
    
    MsgBox s
    Set d = fs.GetDrive(drvpath)


End Sub

'//METODO DRIVEEXIST    *** FINE ***
'//<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<//



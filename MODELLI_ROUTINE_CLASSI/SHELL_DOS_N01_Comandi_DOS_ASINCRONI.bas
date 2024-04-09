Attribute VB_Name = "SHELL_DOS_N01_Comandi_DOS_ASINCRONI"
Option Compare Database


'// FUNZIONE Shell _
Eseguire un programma eseguibile e restituisce un valore Variant (Double) _
che rappresenta l'ID del programma attivit� in caso contrario, in caso contrario restituisce zero.

'//vbHide = 0 _
Finestra � nascosta e lo stato attivo alla finestra nascosta. La costante vbHide non � applicabile su piattaforme Macintosh. _
vbNormalFocus = 1 _
Finestra con lo stato attivo e ripristinato nella posizione e le dimensioni originali. _
vbMinimizedFocus = 2 _
Finestra verr� visualizzata come icona con lo stato attivo. _
vbMaximizedFocus = 3 _
Finestra ingrandita con lo stato attivo. _
vbNormalNoFocus = 4 _
Finestra viene ripristinata alle dimensioni e posizione pi� recente. La finestra attiva rimane attiva. _
vbMinimizedNoFocus = 6 _
Finestra verr� visualizzata come icona. La finestra attiva rimane attiva. _
Osservazioni _
Se la funzione Shell esegue correttamente il file specificato, verr� restituito l'ID attivit� del programma avviato. _
ID attivit� � un numero univoco che identifica il programma in esecuzione. Se la funzione Shell non � possibile avviare _
il programma specificato, verr� restituito un errore. _
Su Macintosh, vbNormalFocus, vbMinimizedFocuse vbMaximizedFocus tutti posizionare l'applicazione in primo piano. vbHide, _
vbNoFocusvbMinimizeFocus tutti posizionare l'applicazione in background.


'//Shell ( percorso  [, windowstyle ] )
'//[VBA]Esecuzione sincrona file Eseguibili _
Questo codice funziona ovviamete solo con File ESEGUIBILI, _
nel caso non lo siano la stringa deve essere composta specificando quale EXE � associato _
al file... recuperandolo con altri metodi come API(FindExecutable) _
nel caso non lo siano la stringa deve essere composta specificando quale EXE � associato _
Consente di eseguire un File, di norma Batch o Exe in modalit� SINCRONA, quindi il ritorno dalla chiamata avverr� solo al termine. _
Il sistema non ha TIMEOUT.

Private Sub PROVA_COMANDO_SHELL_DOS()
'// Parametro vbHide = la finestra dos � nascosta
'X = ShellEX("C:\CASA\GE_CASA\GE_MARINO\GESTIONE_SPESE\GESTIONE_PROCEDURE\GE_CASA_SALVATAGGIO_ARCHIVI_XLS\CANCELLA_FILE_XLS.bat", vbHide, True)

'// Parametro VbNormaFocus = la finestra dos � visibile ed in esecuzione asincrona
X = ShellEX("C:\CASA\GE_CASA\GE_MARINO\GESTIONE_SPESE\GESTIONE_PROCEDURE\GE_CASA_SALVATAGGIO_ARCHIVI_XLS\CANCELLA_FILE_XLS.bat", vbNormalFocus, True)


'//MsgBox "VALORE RESTITUITO DALLA SHELL X: " & X

End Sub


'//ROUTINE PER LA CANCELLAZIONE _
Da utilizzare solo alla chiusura della form perch� altrimenti i file collegati non vengono cancellati
Public Sub DOS_CANCELLA_FILE_XLS()

    '// Parametro VbNormaFocus = la finestra dos � visibile ed in esecuzione asincrona ed il comando dos attivato � il CANCELLA.BAT
    X = ShellEX("C:\CASA\GE_CASA\GE_MARINO\GESTIONE_SPESE\GESTIONE_PROCEDURE\GE_CASA_SALVATAGGIO_ARCHIVI_XLS\CANCELLA_FILE_XLS.bat", vbNormalFocus, True)

End Sub

'//SPOSTA I FILE TMP ALLA CHIUSURA DELLA FORM _
alla chiusura della form vengono prima spostati i file xls dalla cartella tmp alla cartella archivio
Public Sub DOS_SPOSTA_FILE_TMP()

'// Parametro VbNormaFocus = la finestra dos � visibile ed in esecuzione asincrona ed il comando dos attivato � il SPOSTA_FILE_XLS.BAT
    X = ShellEX("C:\CASA\GE_CASA\GE_MARINO\GESTIONE_SPESE\GESTIONE_PROCEDURE\GE_CASA_SALVATAGGIO_ARCHIVI_XLS\SPOSTA_FILE_XLS.bat", vbNormalFocus, True)


End Sub

Function ShellEX(ByVal Percorso As String, _
            ByVal windowstyle As Integer, _
            ByVal Wait As Boolean) As Boolean
     
     
     
     On Error GoTo Err_Shell
     Dim wshell As Object

     ShellEX = False
     Set wshell = CreateObject("WScript.shell")
     wshell.Run Percorso, windowstyle, Wait

    '//Libero la memoria e restituisco True alla funzione
     Set wshell = Nothing
     ShellEX = True
     
Exit_Here:
     Exit Function
Err_Shell:
     Resume Exit_Here
End Function




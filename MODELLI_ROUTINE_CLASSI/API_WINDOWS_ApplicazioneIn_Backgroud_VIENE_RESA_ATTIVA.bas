Attribute VB_Name = "API_WINDOWS_ApplicazioneIn_Backgroud_VIENE_RESA_ATTIVA"
Option Compare Database



'(D) Come faccio a scoprire � Excel o Word � in esecuzione in background?
'(A), � possibile utilizzare la funzione fIsAppRunning per verificare
'se un'applicazione � in esecuzione. Passare il nome Appplication a questa funzione.
'Un argomento opzionale � passato come True o False se si desidera attivare l'applicazione.
'Per esempio, se si desidera attivare Word se � trovato in esecuzione, provate questo nella finestra di debug:
'Print fIsAppRunning("Parola", True)
'Se si vuole solo sapere se Word � in esecuzione o meno, si pu� semplicemente
'chiamare la funzione come questa:
'Print fIsAppRunning("parola")
'Nota: i nomi nuova classe possono essere aggiunti alla struttura Select Case
'per estendere le funzionalit�.


'***************** Code Start ***************
'This code was originally written by Dev Ashish.
'It is not to be altered or distributed,
'except as part of an application.
'You are free to use it in any application,
'provided the copyright notice is left unchanged.
'
'Code Courtesy of
'Dev Ashish
'
Private Const SW_HIDE = 0
Private Const SW_SHOWNORMAL = 1
Private Const SW_NORMAL = 1
Private Const SW_SHOWMINIMIZED = 2
Private Const SW_SHOWMAXIMIZED = 3
Private Const SW_MAXIMIZE = 3
Private Const SW_SHOWNOACTIVATE = 4
Private Const SW_SHOW = 5
Private Const SW_MINIMIZE = 6
Private Const SW_SHOWMINNOACTIVE = 7
Private Const SW_SHOWNA = 8
Private Const SW_RESTORE = 9
Private Const SW_SHOWDEFAULT = 10
Private Const SW_MAX = 10

Private Declare Function apiFindWindow Lib "user32" Alias _
    "FindWindowA" (ByVal strClass As String, _
    ByVal lpWindow As String) As Long

Private Declare Function apiSendMessage Lib "user32" Alias _
    "SendMessageA" (ByVal Hwnd As Long, ByVal Msg As Long, ByVal _
    wParam As Long, lParam As Long) As Long
    
Private Declare Function apiSetForegroundWindow Lib "user32" Alias _
    "SetForegroundWindow" (ByVal Hwnd As Long) As Long
    
Private Declare Function apiShowWindow Lib "user32" Alias _
    "ShowWindow" (ByVal Hwnd As Long, ByVal nCmdShow As Long) As Long
    
Private Declare Function apiIsIconic Lib "user32" Alias _
    "IsIconic" (ByVal Hwnd As Long) As Long
    
'//VISUALIZZA GLI APPLICATIVI IN BACKGROUND
Private Sub CHIAMA_fIsAppRunning()
fIsAppRunning "excel", True
End Sub


Function fIsAppRunning(ByVal strAppName As String, _
                       Optional fActivate As Boolean) As Boolean
    Dim lngH As Long, strClassName As String
    Dim lngX As Long, lngTmp As Long
    Const WM_USER = 1024
    On Local Error GoTo fIsAppRunning_Err
    fIsAppRunning = False
    Select Case LCase$(strAppName)
        Case "excel":       strClassName = "XLMain"
        Case "word":        strClassName = "OpusApp"
        Case "access":      strClassName = "OMain"
        Case "powerpoint95": strClassName = "PP7FrameClass"
        Case "powerpoint97": strClassName = "PP97FrameClass"
        Case "notepad":     strClassName = "NOTEPAD"
        Case "paintbrush":  strClassName = "pbParent"
        Case "wordpad":     strClassName = "WordPadClass"
        Case Else:          strClassName = vbNullString
    End Select
    
    If strClassName = "" Then
        lngH = apiFindWindow(vbNullString, strAppName)
    Else
        lngH = apiFindWindow(strClassName, vbNullString)
    End If
    If lngH <> 0 Then
        apiSendMessage lngH, WM_USER + 18, 0, 0
        lngX = apiIsIconic(lngH)
        If lngX <> 0 Then
            '//l'applicazione in backgrud viene resa attiva
            lngTmp = apiShowWindow(lngH, SW_SHOWNORMAL)
        End If
        If fActivate Then
            lngTmp = apiSetForegroundWindow(lngH)
            MsgBox "Applicazione gia attiva"
        End If
        fIsAppRunning = True
    End If
fIsAppRunning_Exit:
    Exit Function
fIsAppRunning_Err:
    fIsAppRunning = False
    Resume fIsAppRunning_Exit
End Function
'******************** Code End ****************




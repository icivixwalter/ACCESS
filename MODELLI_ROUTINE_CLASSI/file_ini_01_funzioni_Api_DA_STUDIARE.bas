Attribute VB_Name = "file_ini_01_funzioni_Api_DA_STUDIARE"
Private Declare Function GetPrivateProfileString Lib "kernel32.dll" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Long

Private Declare Function WritePrivateProfileString Lib "kernel32.dll" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Long

Public Function IniRead(ByVal Filename As String, ByVal Section As String, ByVal Key As String, Optional ByVal lpDefault As String = "", Optional ByVal bRaiseError As Boolean = False) As String
'***********************************************************
' Func: IniRead SB 20080520
' Desc: Legge il valore di una Chiave di una Sezione precisa del file .Ini
' Par : FileName Nome e percorso completo del file .INI
' Section Sezione del file .Ini contenente la chiave
' Key Chiave del file .Ini da leggere
' [ldDefault] Valore di default in caso di lettura non riuscita
' [bRaiseError] Boolean. Se true viene generata un'eccezione in caso d'errore
' Ret : String Valore stringa letto.
' Note:
'***********************************************************
Dim RetVal As String
Dim LenResult As Integer
Dim ErrString As String
    RetVal = Space(255)
    LenResult = GetPrivateProfileString(Section, Key, lpDefault, RetVal, RetValLength, Filename)
    If LenResult = 0 And bRaiseError Then
        ErrString = "Impossibile eseguire l'operazione: la sezione o la chiave sono errate oppure l'accesso al file non è consentito"
        'Err.Raise Err.Raise(9998, Nothing, ErrString)
    End If
IniRead = Mid(RetVal, 1, LenResult)
End Function

Public Function IniWrite(ByVal Filename As String, ByVal Section As String, ByVal Key As String, ByVal Value As String, Optional ByVal bRaiseError As Boolean = False) As Boolean
'***********************************************************
' Func: IniWrite SB 20080520
' Desc: Legge il valore di una Chiave di una Sezione precisa del file .Ini
' Par : FileName Nome e percorso completo del file .INI
' Section Sezione del file .Ini contenente la chiave
' Key Chiave del file .Ini da scrivere
' Value Stringa da assegnare alla chiave.
' [bRaiseError] Boolean. Se true viene generata un'eccezione in caso d'errore
' Ret : String Valore stringa letto.
' Note:
'***********************************************************
Dim LenResult As Integer
Dim ErrString As String
    LenResult = WritePrivateProfileString(Section, Key, Value, Filename)
    If LenResult = 0 And bRaiseError Then
        ErrString = "Impossibile eseguire l'operazione: la sezione o la chiave sono errate oppure l'accesso al file non è consentito"
        'Err.Raise Err.Raise(9999, Nothing, ErrString)
    End If
IniWrite = IIf(LenResult = 0, False, True)

End Function

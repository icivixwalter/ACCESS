
Option Explicit
On Error Resume Next 'Se fa errore, non segnala

Dim a,b,title 'Variabili

title = "xOrCrYpT 0.2" 'titolo
a = InputBox("Lettera da criptare:",title) 'richiesta della lettera
b = InputBox("Chiave:",title) 'richiesta della chiave

MsgBox "Ecco la lettera criptata: " & encrypt(a,b),vbInformation,title
'visualizza la lettera criptata

Function encrypt(char,key)
'funzione di crittografia
    encrypt = Asc(CStr(char)) XOR CInt(key)
    'restituisce lo XOR fra il carattere e la chiave
End Function


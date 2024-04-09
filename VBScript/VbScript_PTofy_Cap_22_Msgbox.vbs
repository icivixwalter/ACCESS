Option Explicit
On Error Resume Next

Dim title,risposta

title = "QuIz! 2.0" 'Titolo dell'applicazione

risposta=MsgBox("Il coccige è la parte terminale della colonna vertebrale ?",vbYesNo+vbQuestion,title)
'Visualizza una MsgBox

if risposta=vbYes then
'Se la risposta è si

MsgBox "Esatto!",vbInformation,title
'Visualizza esatto

else
'altrimenti

MsgBox "Hai sbagliato!" & vbCrLf & "Che ci vuoi fare...",vbExclamation,title
'Visualizza sbagliato
end if
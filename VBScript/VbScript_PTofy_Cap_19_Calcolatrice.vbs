Option Explicit
Dim WshShell 'Variabile oggetto
set WshShell = CreateObject("wscript.Shell") 'Oggetto shell
WshShell.Exec("calc") 'Chiama la calcolatrice di Windows 
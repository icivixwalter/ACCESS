Attribute VB_Name = "Modulo1_file_ini_II_da_studaire"
Option Explicit
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" _
(ByVal lpApplicationname As String, _
ByVal lpKeyName As Any, _
ByVal lpDefault As String, _
ByVal lpReturnedString As String, _
ByVal nSize As Long, _
ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" _
(ByVal lpApplicationname As String, _
ByVal lpKeyName As Any, _
ByVal lpString As Any, _
ByVal lpFileName As String) As Long

'Grandezza, in caratteri, del buffer puntato dal
'parametro lpReturnedString.
Public Const MaxBuf As Integer = 255


Dim FileIni As Variant
Dim path As String

'Per leggere e scrivere su un file INI dobbiamo utilizzare 2 funzioni delle API. GetPrivateProfileString per leggere, _
 e WritePrivateProfileString per scrivere. _
 http://www.franknet.altervista.org/Tips_VBasic/sistema/Leggere%20e%20Scrivere%20un%20File%20INI.htm






Private Sub cmdScrivi()
path = "c:\CASA\LINGUAGGI\ACCESS\ACCESS_FILE_INI\"
FileIni = "FILE_INI"
'scrive la prima Sezione (Utente), chiave (Default), valore (Gianni Marzio)
ScriviIni FileIni, "Utente", "Default", "Gianni Marzio"
'scrive la seconda Sezione (NomeUtenti), chiave (NumeroUtenti), valore (3)
ScriviIni FileIni, "NomeUtenti", "NumeroUtenti", "3"
'scrive la seconda Sezione (NomeUtenti), chiave (Nome1), valore (Tizio)
ScriviIni FileIni, "NomeUtenti", "Nome1", "Tizio"
'scrive la seconda Sezione (NomeUtenti), chiave (Nome2), valore (Caio)
ScriviIni FileIni, "NomeUtenti", "Nome2", "Caio"
'scrive la seconda Sezione (NomeUtenti), chiave (Nome3), valore (Sempronio)
ScriviIni FileIni, "NomeUtenti", "Nome3", "Sempronio"
End Sub




Private Sub cmdLeggi()
Dim NomeDefault As String, NomeUtenti(3) As String
Dim NumUtenti As Integer, i As Integer

path = "c:\CASA\LINGUAGGI\ACCESS\ACCESS_FILE_INI\"

'legge la prima Sezione (Utente), chiave (Default)
NomeDefault = LeggiIni(FileIni, "Utente", "Default")
'legge la seconda Sezione (NomeUtenti), chiave (NumeroUtenti)
NumUtenti = Val(LeggiIni(FileIni, "NomeUtenti", "NumeroUtenti"))
'carica dal file ini il nome degli utenti
 Debug.Print NomeDefault
For i = 1 To NumUtenti
    'legge la seconda Sezione (NomeUtenti), chiave (Nome) & incremento di (i)
    NomeUtenti(i) = LeggiIni(FileIni, "NomeUtenti", "Nome" & CStr(i))
    Debug.Print NomeUtenti(i)
Next i
End Sub



'Leggere una riga di un tipico file *.ini:
Public Function LeggiIni(ByVal nomeFileIni As String, nomeSezione As String, nomeChiave As String) As String
'nomeSezione = (lpApplicationName) nome della sezione in cui cercare il valore
'nomeChiave = (lpKeyName) nome della chiave in cui cercare il valore
'default = (lpDefault) se la chiave non può essere nel file ini, la funzione GetPrivateProfileString copia la stringa default nel buffer di lpReturnedString buffer. Questo parametro non può essere NULL.
'nomeValore = (lpReturnedString) punta al buffer che riceve la stringa trovata
'MaxBuf = (nsize) grandezza, in caratteri, del buffer puntato dal parametro lpReturnedString.
'nomeFileIni = (lpFileName) nome del file *.ini da leggere
Dim default As String, nomeValore As String
Dim ret As Long

ContrNomeFile nomeFileIni
default = Chr$(0)
nomeValore = String$(MaxBuf, 0)
ret = GetPrivateProfileString(nomeSezione, nomeChiave, default, nomeValore, MaxBuf, nomeFileIni)
If ret <> 0 Then
    LeggiIni = Left(nomeValore, ret)
Else
    LeggiIni = ""
End If
End Function

'Scrivere una riga di un tipico file *.ini:
Public Sub ScriviIni(ByVal nomeFileIni As String, nomeSezione As String, nomeChiave As String, tempStringa As String)
'nomeSezione = (lpApplicationName) nome della sezione in cui cercare il valore
'nomeChiave = (lpKeyName) nome della chiave in cui cercare il valore
'nomeStringa = (lpString) stringa che deve essere scritta nel file. Se questo parametro è NULL, la chiave associata sarà cancellata.
'nomeFileIni = (lpFileName) nome del file *.ini da leggere
Dim ret As Long, nomeStringa As String
If tempStringa <> "" Then nomeStringa = tempStringa
ContrNomeFile nomeFileIni
ret = WritePrivateProfileString(nomeSezione, nomeChiave, nomeStringa, nomeFileIni)
End Sub

Public Sub ContrNomeFile(nomeIni As String)
'Controlla se il file sia completo di estensione e se è senza un percorso specificato gli assegna il percorso dell'applicazione
If InStr(nomeIni, ".") = 0 Then nomeIni = nomeIni & ".ini"
If InStr(nomeIni, "\") = 0 Then nomeIni = path & "\" & nomeIni
End Sub





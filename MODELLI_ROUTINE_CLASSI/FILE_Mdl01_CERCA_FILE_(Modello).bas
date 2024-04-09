Attribute VB_Name = "FUNZIONI_Mdl01_CERCA_FILE"
Option Compare Database

 '//OPZIONI
    '//........................................................
   ' Option Compare Text                     'Le Opzioni di comparazione testo
    Option Explicit                         'Le Opzioni esplicite per le variabili

    '//*** Fine ***
    '//OPZIONI
    '//........................................................
            

Private Declare Function apiSearchTreeForFile Lib "ImageHlp.dll" Alias _
        "SearchTreeForFile" (ByVal lpRoot As String, ByVal lpInPath _
        As String, ByVal lpOutPath As String) As Long

Private Sub CHIAMA_fReturnFilePath()

fSearchFile "Folium_7582_2015", "c:\GESTIONI\GESTIONE_LLPP\02_SCANNER\ScannerTmp\"
End Sub


Function fSearchFile(ByVal strFilename As String, _
            ByVal strSearchPath As String) As String
'Returns the first match found
    Dim lpBuffer As String
    Dim lngResult As Long
    fSearchFile = ""
    lpBuffer = String$(1024, 0)
    lngResult = apiSearchTreeForFile(strSearchPath, strFilename, lpBuffer)
    If lngResult <> 0 Then
        If InStr(lpBuffer, vbNullChar) > 0 Then
            fSearchFile = Left$(lpBuffer, InStr(lpBuffer, vbNullChar) - 1)
            
        End If
    End If
End Function



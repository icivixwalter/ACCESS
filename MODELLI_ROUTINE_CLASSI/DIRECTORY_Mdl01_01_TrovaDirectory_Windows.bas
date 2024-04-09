Attribute VB_Name = "DIRECTORY_Mdl01_01_TrovaDirectory_Windows"
Option Compare Database

'//TROVA LA DIRECTORY WINDOWS
Public Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Public Function WinDir() As String
    WinDir = String$(256, 0)
    WinDir = Left$(WinDir, GetWindowsDirectory(WinDir, 255))
    MsgBox "TROVATA LA DIRECTORY : --> " & WinDir

End Function

Attribute VB_Name = "DOS_SHELL_Mdl21_FUNZIONE_SHELL"
Option Compare Database
Option Explicit

Sub ATTIVA_SHELLDOS()
 Call ShellDOS_Exit
End Sub
Function ShellDOS_Exit() As Integer
         On Local Error GoTo ShellDOS_Exit_Err
         Dim MyCommand As String
         Dim TaskId As Integer
         ' Create command string to show the contents of current
         ' directory. Upon completion the window closes.
         MyCommand = "COMMAND.COM /C DIR /P"
         '??MyCommand = "COMMAND.COM DIR C:\CASA\LINGUAGGI\ACCESS\ACCESS_MODELLO_ANALISI\MODULI_CLASSI\*.*"
         
         
         TaskId = Shell(MyCommand, 1)
         ShellDOS_Exit = True
ShellDOS_Exit_End:
         Exit Function
ShellDOS_Exit_Err:
         MsgBox Error$
         Resume ShellDOS_Exit_End
      End Function

      Function ShellDOS_Stay() As Integer
         On Local Error GoTo ShellDOS_Stay_Err
         Dim MyCommand As String
         Dim TaskId As Integer
         ' Create command string to show the contents of current
         ' directory. Upon completion the window remain opens
         ' at the MS-DOS prompt.
         MyCommand = "COMMAND.COM /K DIR /P"
         TaskId = Shell(MyCommand, 1)
         ShellDOS_Stay = True
ShellDOS_Stay_End:
         Exit Function
ShellDOS_Stay_Err:
         MsgBox Error$
         Resume ShellDOS_Stay_End
      End Function
                    

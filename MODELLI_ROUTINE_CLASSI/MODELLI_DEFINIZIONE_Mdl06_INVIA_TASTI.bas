Attribute VB_Name = "MODELLI_DEFINIZIONE_Mdl06_INVIA_TASTI"
Option Compare Text
Option Explicit




Function INVIA_TASTI(par_sScelta_db As String, par_sName_tab As String)

'.................................
' Dim DB, RS
Dim dbs As Database
Dim rRst As Recordset
Dim sSql As String
Dim Vv1 As Variant


'.................................
' Dim Variabili oggetto
Dim appAccess As New Access.Application


'.................................
' Dim Variabili generali
'Dim I As Integer
Dim X As Integer
Dim sPath As String
Dim sControlloPath As String
Dim sRiga As String
Dim sFile As String

            Dim ReturnValue, i

    
    On Error GoTo Err_INVIA_TASTI

    '------------------------------------------------------------------------------------
    '   RESET VARIABILI, CANCELLAZIONE DATI IN TABELLA, OPEN RST
            
            
             '   Apro il file di aggiornamento
            ' Avvia MsAccess.exe
           ' Shell ("d:\Programmi\Microsoft Office\Office\MsAccess.exe")

            
            'ReturnValue = Shell("C:\CASA\ANALISI\A-CASA\SEZ1_CND.mdb", 1)
            ReturnValue = Shell("d:\Programmi\Microsoft Office\Office\MsAccess.exe C:\CASA\ANALISI\A-CASA\SEZ1_CND.mdb", 1)
            AppActivate ReturnValue      'attiva il db da aggiornare
            
            'Set AppAccess = GetObject("C:\CASA\ANALISI\A-CASA\SEZ1_CND.mdb", "Access.Application")
            Set appAccess = GetObject("C:\CASA\ANALISI\A-CASA\SEZ1_CND.mdb")
            appAccess.DoCmd.OpenTable (par_sName_tab), acViewDesign
            appAccess.DoCmd.SelectObject acTable, (par_sName_tab), False
            appAccess.DoCmd.OpenTable (par_sName_tab), acViewDesign
                                
                                            SendKeys "{ESC}", True
                                            SendKeys "{DOWN}", True
                                            SendKeys "{DOWN}", True
                                            SendKeys "%{F4}", True          ' tasto Alt+F4
                                            SendKeys "{ESC}", True
            

    
    
'------------------------------------------------------------------------------
'   LA CHIUSURA E LA GESTIONE ERRORI
   
Exit_INVIA_TASTI:
Exit Function

Err_INVIA_TASTI:
    MsgBox Err.Number & " - " & Err.Description, vbCritical, "Routine INVIA_TASTI"

    Resume Exit_INVIA_TASTI

    
End Function

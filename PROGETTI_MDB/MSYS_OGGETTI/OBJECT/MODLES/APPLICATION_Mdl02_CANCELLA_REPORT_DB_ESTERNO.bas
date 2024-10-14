Attribute VB_Name = "APPLICATION_Mdl02_CANCELLA_REPORT_DB_ESTERNO"
Option Compare Database
Option Explicit

Dim appAccess As Object
    Dim strDB As String
    Dim strReportName As String

Sub RichiamaDatiAccess()
' Dichiarare variabile oggetto nella sezione Dichiarazioni di un modulo
    strDB = "c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MDB\MENU_TB03_OGGETTI_DA_CANCELLARE\MENU_TB03_OGGETTI_DA_CANCELLARE.mdb"
    strReportName = "Report____________________Report________________________________"
    VerificaReportAccess strDB, strReportName
End Sub

Sub VerificaReportAccess(strDB As String, _
     strReportName As String)
     
     Dim formName As String
      Dim rpt As Object
      
    ' Restituisce riferimento a oggetto Application
    ' Di Microsoft Access.
    Set appAccess = New Access.Application
    ' Apre database in Microsoft Access.
    appAccess.OpenCurrentDatabase strDB
    ' Verifica esistenza del report.
    On Error GoTo ErrorHandler
    'appAccess.CurrentProject.AllReports (strReportName)
    'MsgBox "Il report " & strReportName & _
        " è stato trovato nel database Northwind."
    
    
    'FormExists = False
    For Each rpt In appAccess.CurrentProject.AllReports
        Debug.Print rpt.Name
        If rpt.Name = strReportName Then
            'FormExists = True
            appAccess.DoCmd.DeleteObject acReport, rpt.Name
             
            Exit For
        End If
    Next rpt
    
    appAccess.CloseCurrentDatabase
    Set appAccess = Nothing
    
    
    
Exit Sub
ErrorHandler:
    MsgBox "Il report " & strReportName & _
        " non esiste nel database Northwind."
    appAccess.CloseCurrentDatabase
    Set appAccess = Nothing
End Sub


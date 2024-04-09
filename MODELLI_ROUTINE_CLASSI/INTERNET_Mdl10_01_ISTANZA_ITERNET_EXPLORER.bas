Attribute VB_Name = "INTERNET_Mdl10_01_ISTANZA_ITERNET_EXPLORER"
Option Compare Database

Public Sub m()

    Dim objIE As Object
     Dim s As String
     
     Set objIE = CreateObject("InternetExplorer.Application")
     objIE.Visible = True
     
     objIE.Navigate ("http://www.maurogsc.eu/prove/paginatest.htm")
     
     
     'Application.Wait (Now + TimeValue("0:00:05"))
     
     s = objIE.Document.All("testo").innerText
     MsgBox s
     
     objIE.Quit
     Set objIE = Nothing



End Sub


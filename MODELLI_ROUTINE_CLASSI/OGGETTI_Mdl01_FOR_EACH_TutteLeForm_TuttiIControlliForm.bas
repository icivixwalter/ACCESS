Attribute VB_Name = "UTILITA_Mdl37_FOR_EACH_TutteLeForm_TuttiControlli"
 Option Compare Database
Option Explicit
 
 '//LOOP CONTROLLO FORM _
 attenzione occorrono le seguenti librerie altrimenti da errore:
 '//Visual Basic For Application _
    Microsoft Access  9.0 Obiect Library _
    Microsoft VbScript Regular Expressiins 5.5. _
    Microsoft DAO 3.6 Obiect Library _
    Microsoft Visual Basic for Application Extensibility 5.3 _
    Microsoft Activex Data Obiects 2.1 Library _
    Microsoft Scripting Runtime
 
 Public Sub LoopFormProps()
    On Error GoTo err_Handler
    Dim frmLoop As Object
    Dim ctrl As Control
    Dim propLoop As Property
    Dim strFormName As String
   
    '//ITERO IN TUTTE LE FORM E IN TUTTI I CONTROLLI DELLE FORM
    For Each frmLoop In CurrentProject.AllForms
        strFormName = frmLoop.Name
        DoCmd.OpenForm strFormName, acDesign, , , , acHidden
    
        With Forms(strFormName)
            For Each propLoop In .Properties
                Debug.Print "Form [" & .Name & "](" & propLoop.Name & "):" & propLoop.Value
            Next
            For Each ctrl In .Controls
                For Each propLoop In ctrl.Properties
                    Debug.Print "Form [" & .Name & "] - Control[" & ctrl.Name & "](" & propLoop.Name & "):" & propLoop.Value
                Next
            Next
        End With
        DoCmd.Close acForm, strFormName
    Next
 
    Exit Sub
err_Handler:
    Select Case Err.Number
        Case 2186
            'Property not available in design view
            Resume Next
        Case Else
            MsgBox "Err: " & Err.Number & " - " & Err.Description
    End Select
End Sub




Attribute VB_Name = "Modulo1___PROVA_FOR_EACH_AL_DI_FUORI_DELLA FORM"

Option Compare Database


'//>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>><<<//
'//CICLO_FOR_EACH_DEI_CONTROLLI_ESEMPIO_RICERCA_PROPRIETA_NEXT
'//>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>><<<//
'//Sub  :
'//Note :

'Esempio di istruzione For Each...Next'
'In questo esempio l'istruzione For Each...Next
'viene utilizzata per eseguire una ricerca nella proprietà Text di tutti gli elementi
'di un insieme e individuare la presenza della stringa "Salve".
'Nell'esempio, MyObject è un oggetto di testo ed è un elemento dell'insieme di dati
'MyCollection. Sono entrambi nomi generici utilizzati unicamente a titolo esemplificativo.


Sub CicloForEach_TUTTI_CONTROLLI()


Dim Found, MyObject, MyCollection
Found = False        ' Inizializza la variabile.
For Each MyObject In MyCollection    ' Esegue un'iterazione in ogni elemento.
    
    Debug.Print MyObject.Name
    Debug.Print "Proprieta text "
    Debug.Print MyObject.Text
    
    If MyObject.Text = "Salve" Then    ' Se Text è uguale a "Salve".
        Found = True    ' Imposta Found su True.
        Exit For            ' Esce dal ciclo.
    End If
Next

End Sub




'//>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>><<<//
'//CICLO_FOR_EACH_DEI_CONTROLLI_ESEMPIO_RICERCA_PROPRIETA_NEXT *** FINE ***
'//>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>><<<//



'//non funziona ?? come impostare l'oggetto form al di fuori
'//di una maschera.
Sub IteraDbCorrente()

Dim obj         As AccessObject
Dim dbs         As Object
Dim frm         As Form
Dim ctl         As Control
Dim intCount    As Integer
Dim intFile     As Integer
Dim appAccess As Object
    
    '//il database corrente
    Set dbs = Application.CurrentProject
    intFile = FreeFile
    Set Form = Application.Forms
    
    For Each ctl In Form.Controls
    
      Next ctl


     For Each obj In dbs.AllForms

            If obj.IsLoaded = True Then DoCmd.Close acForm, obj.Name, acSaveNo
            DoCmd.OpenForm obj.Name, acDesign, , , , acHidden
            Set frm = Access.Forms(obj.Name)
                    
            For Each ctl In frm.Controls
                Debug.Print ctl.Name
                For intCount = 0 To ctl.Properties.Count
                     Debug.Print "Maschera(" & obj.Name & _
                                 ")--> Controllo(NRO " & intCount; " " & ctl.Name & _
                     ") --> Proprietà(" & ctl.Properties(intCount).Name; ") = " & ctl.Properties(intCount).Value
                    ' Quì invece del DEBUG vai a scrivere in un FILE con 1 riga di codice...
                    
                    Open "C:\FormsControlsProperties.txt" For Append Shared As #intFile
                        Print #intFile, "Maschera(" & obj.Name & _
                                        ")--> Controllo(" & ctl.Name & _
                            ") --> Proprietà(" & ctl.Properties(intCount).Name; ") = " & ctl.Properties(intCount).Value
                    Close #intFile
            DoEvents
               Next intCount

                ' Salta una RIGA ogni Maschera
               Open "C:\FormsControlsProperties.txt" For Append Shared As #intFile
               Print #intFile, "-----"
               Close #intFile

            Next ctl

            DoCmd.Close acForm, obj.Name, acSaveNo

        Next obj

        Set dbs = Nothing
        Set frm = Nothing
    

End Sub

Sub prova()
Dim obj         As AccessObject
Dim dbs         As Object
Dim frm         As Access.Form
Dim ctl         As Access.Control
Dim intCount    As Integer
Dim intFile     As Integer
    Dim appAccess As Object
    
    '//il database corrente
    Set dbs = Application.CurrentProject
    intFile = FreeFile
    
    
    
    '//Oppure la forma crea oggetto
    '//Set appAccess = New Access.Application

    Set appAccess = CreateObject("Access.Application")
    
    
    
    Dim MyControl As Control
    Dim MyForm As Form
    Dim NroMyCotrol_i As Integer
    

    Dim Found, MyObject, MyCollection
    Found = False        ' Inizializza la variabile.
    Set MyObject = appAccess
    'Set MyCollection = appAccess.AllForms
    For Each MyObject In MyForm    ' Esegue un'iterazione in ogni elemento.
        If MyObject.Text = "Salve" Then    ' Se Text è uguale a "Salve".
            Found = True    ' Imposta Found su True.
            Exit For            ' Esce dal ciclo.
        End If
    Next
    
    
    For Each MyControl In Controls
       iCount = iCount + 1
       '//Numero dei controlli della form
       NroMyCotrol_i = iCount
    
    Next
    '//Distruzione oggetto
    Set appAccess = Nothing
End Sub


'//>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>><<<//
'//CICLO_FOR_EACH_DEI_CONTROLLI_ESEMPIO_RICERCA_PROPRIETA_NEXT *** FINE ***
'//>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>><<<//






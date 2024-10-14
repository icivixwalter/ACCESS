Attribute VB_Name = "Modulo1_APRI_DB_ESTERNO_CON_APPLICATION"
Option Compare Database
Option Explicit



'L 'esempio successivo mostra come utilizzare Microsoft Access come componente COM. Da Microsoft Excel, Visual Basic o da un'altra applicazione che agisce come componente COM, creare un riferimento a Microsoft Access scegliendo Riferimenti dal menu Strumenti nella finestra Moduli. Selezionare la casella di controllo posta accanto a Microsoft Access 9.0 Object Library. Immettere quindi il codice riportato di seguito in un modulo di Visual Basic all'interno di tale applicazione e chiamare la routine RicercaDati.

'Nell 'esempio, a una routine che crea una nuova istanza della classe Application vengono passati un nome di database e un nome di report, viene aperto il database e stampato il report indicato.

' Dichiara variabile di oggetto nella sezione Dichiarazioni di un modulo
Dim strDB As String
Dim strNomeReport As String
Dim appAccess As Object
Dim dbPath As String
Dim formName As String
Dim icount As Integer








'//@CANCELLA@FORM@DB@ESTERNO_(Cancella le @form nel db esterno con l'oggetto @application)
Sub ApriDatabaseEsternoECancellaForm()
    
    
    ' Path al database esterno
    Const strConPathToSamples = "c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MDB\MENU_TB03_OGGETTI_DA_CANCELLARE\MENU_TB03_OGGETTI_DA_CANCELLARE.mdb"
    
    ' Nome della form da cancellare
    formName = "MODELLO_FrS01_01_MODELLO_ELENCO"
    icount = 0
    ' Creare un'istanza dell'applicazione Access
    Set appAccess = CreateObject("Access.Application")
    
    ' Aprire il database esterno e creo un oggetto access che punta alla _
        form del database esterno. Tale oggetto application viene passato alla _
        routine
    appAccess.OpenCurrentDatabase strConPathToSamples
    
    ' Controlla se la form esiste nel database esterno
    If FormExists_b(appAccess, formName) Then
        ' Cancellare la form
        appAccess.DoCmd.DeleteObject acForm, formName
        Debug.Print "Form cancellata: " & formName
        icount = icount = 1
        MsgBox "FORM CANCELLATE NEL DB ESTERNO NRO :" & icount
    Else
        Debug.Print "La form non esiste: " & formName
    End If
    
    ' Chiusura del database esterno
    appAccess.CloseCurrentDatabase
    Set appAccess = Nothing

End Sub

'//l'iterazione degli oggetti AllForms non puo avvenire con l'oggetto application del db esterno _
   ma prima si deve creare un oggetto access
Function FormExists_b(appAccess As Object, formName As String) As Boolean
    Dim frm As Object
    FormExists_b = False
    For Each frm In appAccess.CurrentProject.AllForms
        If frm.Name = formName Then
            FormExists_b = True
            Exit Function
        End If
    Next frm
End Function


Attribute VB_Name = "APPLICATION_Mdl01_CANCELLA_FORM_DB_ESTERNO"
'//modulo ---> APPLICATION_Mdl01_CANCELLA_FORM_DB_ESTERNO _
    Utilizzo del metodo application per aprire un db esterno, iterare negli oggetti form e _
    cancellare la form individuata se esiste.

    'L 'esempio successivo mostra come utilizzare Microsoft Access come componente COM. _
        Da Microsoft Excel, Visual Basic o da un'altra applicazione che agisce come componente COM, _
        creare un riferimento a Microsoft Access scegliendo Riferimenti dal menu Strumenti nella finestra Moduli. _
        Selezionare la casella di controllo posta accanto a Microsoft Access 9.0 Object Library. _
        Immettere quindi il codice riportato di seguito in un modulo di Visual Basic all'interno di tale applicazione e chiamare la routine RicercaDati.

        'Nell 'esempio, a una routine che crea una nuova istanza della classe Application vengono passati un nome di database _
         e un nome di form, se esiste viene cancellato utilizzando i metodo cmd.deleteObject del db esterno


Option Compare Database
Option Explicit

Dim strDB As String                     '//path e database esterno
Dim appAccess As Object                 '//istanza oggetto application creata del db esterno
Dim dbPath As String
Dim formName As String                  '//il nome della form da cancellare
Dim icount As Integer                   '//contatore form cancellate
Dim frm As Object                       '//oggetto form db esterno
    
Dim FormNameCancellate_s As String




'//@CANCELLA@FORM NEL DATABASE ESTERNO
'//************************************************************************************************************//
'//NOTE: utilizzo due routine la prima apre il db esterno creando un oggetto Application che punta al db _
        la seconda itera con l'oggetto application passato come parametro nel db esterno e controlla se _
        esiste la form da cancellare.


'//@CANCELLA@FORM@DB@ESTERNO_(Cancella le @form nel db esterno con l'oggetto @application)
Sub ApriDatabaseEsternoECancellaForm()
    
    
    ' Path al database esterno
    Const strConPathToSamples = "c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MDB\MENU_TB03_OGGETTI_DA_CANCELLARE\MENU_TB03_OGGETTI_DA_CANCELLARE.mdb"
    
    'reset e  Nome della form da cancellare
    formName = ""
    icount = 0
    formName = "MsysDbEstTb05Frm01_Mt00_}--------------------------------------@"
    
    
    
    ' Creare un'istanza dell'applicazione Access
    Set appAccess = CreateObject("Access.Application")
    
    ' Aprire il database esterno e creo un oggetto access che punta alla _
        form del database esterno. Tale oggetto application viene passato alla _
        routine
    appAccess.OpenCurrentDatabase strConPathToSamples
    
    ' Controlla se la form esiste nel database esterno chiamando la routine a cui passo oggetto _
      application e il nome della form. Restituisce true se trovata.
    If FormExists_b(appAccess, formName) Then
        
        ' Cancellare la form
        appAccess.DoCmd.DeleteObject acForm, formName
        Debug.Print "Form cancellata: " & formName & Chr$(13)
        
        '//salvo le form cancellate
        FormNameCancellate_s = FormNameCancellate_s & "Form cancellata: " & formName & Chr$(13) & Chr$(10)
        '//contatore
        icount = icount + 1
        
        
    Else
        Debug.Print "La form non esiste: " & formName
    End If
                '//messaggio finale
                MsgBox "FORM CANCELLATE NEL DB ESTERNO NRO : " & icount & " elenco : " & Chr$(13) & FormNameCancellate_s
                
    ' Chiusura del database esterno
    appAccess.CloseCurrentDatabase
    Set appAccess = Nothing

End Sub

'//ITERA NEGLI OGGETTI DEL DB ESTERNO NELLE ALLFORMS PER CONTROLLARE SE ESISTE LA FORM
'//-----------------------------------------------------------------------------------------//
'//l'iterazione degli oggetti AllForms non puo avvenire con l'oggetto application del db esterno _
   ma prima si deve creare un oggetto access application passato come parametro _
   insieme al nome della form.
Function FormExists_b(appAccess As Object, formName As String) As Boolean
    
    '//imposto a false il flag
    FormExists_b = False
    
    '//ITERO NEGLI OGGETTI FORM DEL DB ESTERNO
    For Each frm In appAccess.CurrentProject.AllForms
        If frm.Name = formName Then
            '//Se trova la form da cancellare imposta a True
            FormExists_b = True
            Exit Function
        End If
    Next frm
    
End Function
'//-----------------------------------------------------------------------------------------//

'//@CANCELLA@FORM NEL DATABASE ESTERNO **** FINE ****
'//************************************************************************************************************//

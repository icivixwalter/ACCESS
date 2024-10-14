Attribute VB_Name = "APPLICATION_Mdl03_CANCELLA_TABELLA_IN_DB_ESTERNO"
Option Compare Database
Option Explicit

'//ATTENZIONE UTILIZZO WORKSPACE ed implicitamente l'oggetto application, dovrebbe essere piu esplicito _
    l'utilizzo dell'oggetto application vedi modulo di cancellazione del Report esterno.

Sub DeleteSpecificTableInExternalDB()
    Dim dbPath As String
    Dim externalDB As DAO.Database
    Dim tdf As DAO.TableDef
    Dim tableName As String
    Dim tableExists As Boolean

    ' Nome della tabella da cancellare
    tableName = "GEST_MENU_Tb03_}-----------------------------------------------@"

    ' Percorso del database esterno
    dbPath = "C:\Casa\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MDB\PROVA_CANCELLAZIONE_DB_ESTERNO\MENU_TB03_OGGETTI_DA_CANCELLARE.mdb"

    ' Apri il database esterno
    Set externalDB = DBEngine.Workspaces(0).OpenDatabase(dbPath)

    ' Controlla se la tabella esiste
    tableExists = False
    For Each tdf In externalDB.TableDefs
        If tdf.Name = tableName Then
            tableExists = True
            Exit For
        End If
    Next tdf

    ' Se la tabella esiste, la cancella
    If tableExists Then
        On Error Resume Next
        externalDB.TableDefs.Delete tableName
        If Err.Number = 0 Then
            Debug.Print "Tabella '" & tableName & "' cancellata con successo."
        Else
            Debug.Print "Errore nella cancellazione della tabella '" & tableName & "'."
            Err.Clear
        End If
        On Error GoTo 0
    Else
        Debug.Print "La tabella '" & tableName & "' non esiste nel database esterno."
    End If

    ' Chiudi il database esterno
    externalDB.Close
    Set externalDB = Nothing
End Sub


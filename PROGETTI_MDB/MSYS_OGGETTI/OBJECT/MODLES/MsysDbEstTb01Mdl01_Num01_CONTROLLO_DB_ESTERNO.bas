Attribute VB_Name = "MsysDbEstTb01Mdl01_Num01_CONTROLLO_DB_ESTERNO"
'//MODULO = MsysDbEstTb01Mdl01_Num01_CONTROLLO_DB_ESTERNO _
            MODULO PER IL CONTROLLO DELLE TABELLE DI SISTEMA.




Option Compare Database


'//DIM variabili GENERICHE
'//.....................................................................//

Dim dbs As Database
Dim daoRS As DAO.Recordset

Dim qdfProva As QueryDef
Dim qdfCiclo As QueryDef
Dim prpCiclo As Property



'Parametri Table
Dim sxNomeTable As String
Dim sxCodiceTable As String
Dim sxParametroTable As String
Dim ixLungTable As Integer


'Parametri Query
Dim sxNomeQuery As String
Dim sxCodiceQuery As String
Dim sxParametroQuery As String
Dim ixLungQuery As Integer


Dim iCount  As Integer
Dim i As Integer
Dim iTotOggetti As Integer
Dim Bolean1 As Boolean


Dim sxTipoDatabase As String
Dim sxNomeDatabase As String

Dim sxQueryOrigine As String
Dim sxQueryDestinazione As String

Dim sxMessaggioBox  As String

'//DIM variabili GENERICHE
'//.....................................................................//



'//DIM variabili per la gestione del DATABASE ESTERNO
'//.....................................................................//
'//database disco e path e campi
Dim myDISCO_s As String
Dim MyPath_s As String
Dim myDATABASE_s As String
Dim myDirectory_s As String

'//campi
Dim myScel_b As Boolean
Dim myNOTA_OGGETTO_s As String
Dim myNOTEex_s As String

    '//PER IL DB ESTERNO

'//DIM path, db e tabella
Dim dbPath As String
Dim externalDB As DAO.Database
Dim tdf As DAO.TableDef
Dim Qrydf As DAO.QueryDef

    Dim tableName As String
    Dim tableExists_b As Boolean
    Dim tableExists As Boolean              '//ESISTE LA TABLLEA tableExists=TRUE OPPURE FALSE

'//dim le collection per separare il tipo di tabelle
Dim systemTables As Collection
Dim connectedTables As Collection
Dim physicalTables As Collection


'//dim le collection per separare il tipo di tabelle
Dim systemQUERYes As Collection
Dim connectedQUERYes As Collection
Dim physicalQUERYes As Collection


    

  
'//.....................................................................//






'//ACCESSO AL DB ESTERNO - controllo tabelle
'//============================================================================================//
'//NOTE : la routine e la funzione aprono una istanza presso il db esterno access per il _
        controllo delle tabelle. La routine utilizza 3 collezioni di oggetti che per quanto _
        riguarda la systemTables = precarico quali tabelle sono di sistema poi crea altre _
        due collection che vengono popolate se il controllo if IsSystemTable restituisce false _
        perche chiama la funzione per il controllo se la tabella appartiene al sistema, per esclusione _
        appartiene alla fisiche o alle collegate, e quindi popola le due collection che poi vengono _
        stampate _
        'Percorso del database esterno COME ESEMPIO.
        'dbPath = "c:\Casa\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MDB\MENU_TB03_OGGETTI_DA_CANCELLARE\MENU_TB03_OGGETTI_DA_CANCELLARE.mdb"

    
'//codice : @ROUTINE@ATTIVA_(@controllo@esterno@db)
Public Sub ListTablesInExternalDB()

    On Error GoTo Err_ListTablesInExternalDB
    
    
    '//CREO LA COLLEZIONE TABELLE DI SISTEMA E LA POPOLO
    '//popola la collection delle tabelle di sistema
    ' Aggiungi i nomi delle tabelle di sistema da escludere
    Set systemTables = New Collection
    '//Aggiugno gli elementi della collezione
    systemTables.Add "MSysACEs"                     'MSysACEs
    systemTables.Add "MSysAccessObjects"            'MSysAccessObjects'
    systemTables.Add "MSysAccessStorage"            'MSysAccessStorage
                        
    systemTables.Add "MSysNameMap"                  'MSysNameMap
    systemTables.Add "MSysObjects"                  'MSysObjects'
    systemTables.Add "MSysQueries"                  'MSysQueries'
    systemTables.Add "MSysAccessXML"                'MSysAccessXML
    systemTables.Add "MSysRelationships"            'MSysRelationships'
                        
    systemTables.Add "MSysNavPaneGroupCategories"    'MSysNavPaneGroupCategories'
    systemTables.Add "MSysNavPaneGroupToObjects"     'MSysNavPaneGroupToObjects'
    systemTables.Add "MSysNavPaneObjectIDs"           'MSysNavPaneObjectIDs'
    systemTables.Add "MSysNavPaneGroups"              'MSysNavPaneGroups


    ' Inizializza le collezioni per tabelle fisiche e collegate
    Set connectedTables = New Collection
      
    
    
    '//APRO TABELLA PER INDIVIDUARE IL DATABASE ESTERNO
    '//.....................................................................................................//
            
            '//reset
    
            myDISCO_s = ""
            MyPath_s = ""
            myDATABASE_s = ""
            myScel_b = False
    
     '//Apro il Database
     Set daoDB = DBEngine.Workspaces(0).Databases(0)
     '//Apro un Recordset dal parametro ssql
     
            sSql = ""
            sSql = sSql & "SELECT MSysTb05_DB_EST.DISCO_s, "
            sSql = sSql & "MSysTb05_DB_EST.PATH_s, "
            sSql = sSql & "MSysTb05_DB_EST.DATABASE_s, "
            sSql = sSql & "MSysTb05_DB_EST.Scel_b "
            sSql = sSql & "FROM MSysTb05_DB_EST "
            sSql = sSql & "WHERE (((MSysTb05_DB_EST.Scel_b)=True)) "
            sSql = sSql & "WITH OWNERACCESS OPTION;"
            
            
            '//DEBUG CONTROLLO ED APERTURA
            Debug.Print sSql
     
     Set daoRS = daoDB.OpenRecordset(sSql)
        
    If daoRS.EOF = False And daoRS.BOF = False Then
        '//Posizione Primo record
        daoRS.MoveFirst
            While Not daoRS.EOF
              '//Blocco iterazione
                 DoEvents
                    
                    '//CONTROLLO SE SCELTO IL DB DA CONTROLLARE
                    If daoRS.Fields("Scel_b") = True Then
                        
                        If IsNull(daoRS.Fields("DISCO_s")) Or IsNull(daoRS.Fields("PATH_s")) Or _
                           IsNull(daoRS.Fields("DATABASE_s")) Then
                           
                           MsgBox "ATTENZIONE DISCO/PATH/DB SONO NULLI - > " & " DISCO: " & daoRS.Fields("DISCO_s") & Chr$(13) _
                                 & " PATH: " & daoRS.Fields("PATH_s") & Chr$(13) _
                                 & " DATABASE: " & daoRS.Fields("DATABASE_s") & Chr$(13) _
                                 & " USCITA DALLA ROUTINE!!!", vbCritical
                                '//Uscita Rs e chiusura oggetti
                                daoRS.Close
                                Set daoRS = Nothing
                                
                                GoTo Exit_ListTablesInExternalDB

                           
                        End If
                        
                        
                        
                        '//imposto a scelta si
                        myScel_b = True
                        '//imposto la path trovata + directory e mdb
                        dbPath = daoRS.Fields("DISCO_s") & daoRS.Fields("PATH_s") & daoRS.Fields("DATABASE_s")
                        
                        '//DISCO, PATH , DB
                        myDISCO_s = daoRS.Fields("DISCO_s")
                        MyPath_s = daoRS.Fields("PATH_s")
                        myDATABASE_s = daoRS.Fields("DATABASE_s")
                        
                        '//directory completa DISO + PATH
                        myDirectory_s = daoRS.Fields("DISCO_s") & daoRS.Fields("PATH_s")
                          
                            'TODO: fare un controllo di esistenza path e db!!
                            Debug.Print dbPath
                        
                        '//trovato la path vado a fine rs per uscire dal db
                        daoRS.MoveLast
                        
                    End If
                    
                                                    
                '//Record Successivo
                daoRS.MoveNext
    
        Wend
    
            
        '//Uscita Rs e chiusura oggetti
        daoRS.Close
        Set daoRS = Nothing
        
        End If  '//If DAORs.EOF = False And DAORs.BOF = False Then
                
                
                    '
            '//IMPOSTAZIONE PATH E CONTROLLO ESISTENZA DIRECTORY
            '//------------------------------------------------------------------------------//
            '//NOTE     : controllo l'esistenza della path definita dai salvataggi se non esiste _
                        esco dalla routine.
                
                '//VALORIZZO I PARAMETRI
                par_Directory_s = dbPath
                ParametroFile_i = myDATABASE_s
                
                        
                'Str1 = Dir(Path_s, 16)
                'Vv1 = Dir("*.TXT", 2)
                'MyPath = "c:\CASA\GE_CASA\GE_MARINO\GESTIONE_SPESE\ARCHIVI_XLS\"    ' Imposta il percorso.
                'MYNAME = Dir(MyPath, vbDirectory)    ' Recupera la prima voce.
                'MYNAME = Dir(par_Directory_s, vbDirectory)    ' Recupera la prima voce.
                Vv1 = Dir(par_Directory_s, vbDirectory)   ' Recupera la prima voce.
                
               ' Vv1 = Dir("c:\", vbDirectory)
                
                If Vv1 = "" Then
                        MsgBox "NON ESISTE LA DIRECTORY /DATABASE ---> " & myDirectory_s & myDATABASE_s & " - USCITA DALLA ROUTINE"
                        GoTo Exit_ListTablesInExternalDB
                End If
            '//-------------------------------------------------------------------------------//


                
    '//** FINE **
    '//ITERAZIONE_RECORSET
    '//.....................................................................................................//
    
            '//SOLO SE ESISTE IL DB FACCIO IL CONTROLLO
            If myScel_b = True Then
                
                '//svuoto la tabella tmp + RESET
                    CurrentDb.Execute "MsysDbEstTb01Qry01_Dlt01_OBJECT_TMP"
                    
                    iCount = 0
                    
                
                '//APRO LE COLLECTION E IL DB ESTERNO
                 Set physicalTables = New Collection
                        'todo: FARE UN CONTROLLO PRELIMINARE DELLA PATH E DEL FILE!!
                ' Apri il database esterno
                Set externalDB = DBEngine.Workspaces(0).OpenDatabase(dbPath)
                                
                                    
                                ' Scansiona tutte le tabelle nel database esterno
                                '//...............................................................................//
                                    For Each tdf In externalDB.TableDefs
                                        ' Verifica se la tabella è di sistema con il confronte nella collection precaricata _
                                          e chiama funzione di controllo
                                          
                                        
                                
                                           '//resetto le variabili ad oni ciclo
                                           myNOTA_OGGETTO_s = ""
                                    
                                        
                                        
                                        
                                        If IsSystemTable(tdf.Name, systemTables) Then
                                            Debug.Print tdf.Name & " (Tavola di sistema)"
                                                
                                                iCount = iCount + 1
                                                
                                            myNOTA_OGGETTO_s = "TABELLA MSYS (Tavola di sistema)"
                                            myNOTEex_s = "OGGETTO DATABASE ESTERNO " & myDATABASE_s
                                            
                                                '//QUI AGGIUNGERE SQL INSERIMENTO IN TABELLA TMP
                                                '//..........................................................................//
                                                        sSql = ""
                                                        sSql = sSql & "INSERT INTO "
                                                        sSql = sSql & "MsysDbEstTb01_OBJECT_TMP "
                                                        sSql = sSql & "( NRO_OGGETTO_i, TIPOGGETTO_s, COD_PROGETTO_s, NOTA_OGGETTO_s, NOTEex_s,DISCO_s,PATH_s, DATABASE_s, Name1_s ) "
                                                        sSql = sSql & "SELECT "
                                                        sSql = sSql & iCount & " AS [NRO], "
                                                        sSql = sSql & "'TABLE' AS TIPOGGETTO_s, "
                                                        sSql = sSql & "'MsysDbEst' AS COD_PROGETTO_s,"
                                                        sSql = sSql & "'" & myNOTA_OGGETTO_s & "' AS [NOTE], "
                                                        sSql = sSql & "'" & myNOTEex_s & "' AS [NOTE_EX], "
                                                        
                                                        '//AGGIUNTO DISCO + PATH + DB
                                                        sSql = sSql & "'" & myDISCO_s & "' AS [DISCO_s], "
                                                        sSql = sSql & "'" & MyPath_s & "' AS [PATH_s], "
                                                        sSql = sSql & "'" & myDATABASE_s & "' AS [DATABASE_s], "
                                                        '//ULTIMO campo senza virgola
                                                        sSql = sSql & "'" & tdf.Name & "' AS [Name1_s] "
                                                        
                                                                                                                
                                                        
                                                        sSql = sSql & "WITH OWNERACCESS OPTION;"
                                                        
                                                        '//CONTROLLO ED ESECUZIONE
                                                        Debug.Print
                                                        
                                                        CurrentDb.Execute (sSql)
                                                '//..........................................................................//
                                                
                                            
                                          'se la tabella è collegata aggiunge alla collection TABELLE FISICHE
                                        ElseIf Len(tdf.Connect) > 0 Then
                                        
                                            connectedTables.Add tdf.Name
                                            myNOTA_OGGETTO_s = "TABELLA Collegata"
                                            myNOTEex_s = "OGGETTO DATABASE ESTERNO " & myDATABASE_s
                                            
                                                iCount = iCount + 1
                                            
                                                '//QUI AGGIUNGERE SQL INSERIMENTO IN TABELLA TMP
                                                '//..........................................................................//
                                                        sSql = ""
                                                        sSql = sSql & "INSERT INTO "
                                                        sSql = sSql & "MsysDbEstTb01_OBJECT_TMP "
                                                        'sSql = sSql & "( NRO_OGGETTO_i, TIPOGGETTO_s, COD_PROGETTO_s, NOTA_OGGETTO_s, NOTEex_s, Name1_s ) "
                                                        sSql = sSql & "( NRO_OGGETTO_i, TIPOGGETTO_s, COD_PROGETTO_s, NOTA_OGGETTO_s, NOTEex_s,DISCO_s,PATH_s, DATABASE_s, Name1_s ) "
                                                        sSql = sSql & "SELECT "
                                                        sSql = sSql & iCount & " AS [NRO], "
                                                        sSql = sSql & "'TABLE' AS TIPOGGETTO_s, "
                                                        sSql = sSql & "'MsysDbEst' AS COD_PROGETTO_s,"
                                                        sSql = sSql & "'" & myNOTA_OGGETTO_s & "' AS [NOTE], "
                                                        sSql = sSql & "'" & myNOTEex_s & "' AS [NOTE_EX], "
                                                        
                                                        '//AGGIUNTO DISCO + PATH + DB
                                                        sSql = sSql & "'" & myDISCO_s & "' AS [DISCO_s], "
                                                        sSql = sSql & "'" & MyPath_s & "' AS [PATH_s], "
                                                        sSql = sSql & "'" & myDATABASE_s & "' AS [DATABASE_s], "
                                                        '//ULTIMO campo senza virgola
                                                        sSql = sSql & "'" & tdf.Name & "' AS [Name1_s] "
                                                        
                                                        sSql = sSql & "WITH OWNERACCESS OPTION;"
                                                        
                                                        '//CONTROLLO ED ESECUZIONE
                                                        Debug.Print sSql
                                                        
                                                        CurrentDb.Execute (sSql)
                                                '//..........................................................................//
                                                
                                            
                                            
                                        Else
                                            
                                            'la tabella è fisica e la aggiunge alla collection TABELLE COLLEGATE
                                            physicalTables.Add tdf.Name
                                            
                                            myNOTA_OGGETTO_s = "TABELLA (Fisica)"
                                            myNOTEex_s = "OGGETTO DATABASE ESTERNO " & myDATABASE_s
                                                    
                                                    iCount = iCount + 1
                                            
                                                '//QUI AGGIUNGERE SQL INSERIMENTO IN TABELLA TMP
                                                '//..........................................................................//
                                                        sSql = ""
                                                        sSql = sSql & "INSERT INTO "
                                                        sSql = sSql & "MsysDbEstTb01_OBJECT_TMP "
                                                        'sSql = sSql & "( NRO_OGGETTO_i, TIPOGGETTO_s, COD_PROGETTO_s, NOTA_OGGETTO_s, NOTEex_s, Name1_s ) "
                                                        sSql = sSql & "( NRO_OGGETTO_i, TIPOGGETTO_s, COD_PROGETTO_s, NOTA_OGGETTO_s, NOTEex_s,DISCO_s,PATH_s, DATABASE_s, Name1_s ) "
                                                        sSql = sSql & "SELECT "
                                                        sSql = sSql & iCount & " AS [NRO], "
                                                        sSql = sSql & "'TABLE' AS TIPOGGETTO_s, "
                                                        sSql = sSql & "'MsysDbEst' AS COD_PROGETTO_s,"
                                                        sSql = sSql & "'" & myNOTA_OGGETTO_s & "' AS [NOTE], "
                                                        sSql = sSql & "'" & myNOTEex_s & "' AS [NOTE_EX], "
                                                       
                                                           '//AGGIUNTO DISCO + PATH + DB
                                                        sSql = sSql & "'" & myDISCO_s & "' AS [DISCO_s], "
                                                        sSql = sSql & "'" & MyPath_s & "' AS [PATH_s], "
                                                        sSql = sSql & "'" & myDATABASE_s & "' AS [DATABASE_s], "
                                                        '//ULTIMO campo senza virgola
                                                        sSql = sSql & "'" & tdf.Name & "' AS [Name1_s] "
                                                                                                             
                                                        sSql = sSql & "WITH OWNERACCESS OPTION;"
                                                        
                                                        '//CONTROLLO ED ESECUZIONE
                                                        Debug.Print
                                                        
                                                        CurrentDb.Execute (sSql)
                                                '//..........................................................................//
                                                
                                            
                                            
                                        End If
                                    
                                    Next tdf
                                            
                                                ' Stampa le tabelle collegate con le collezioni
                                                Debug.Print
                                                Debug.Print
                                                Debug.Print "Tabelle collegate:"
                                                'iterazione nella collection delle tabelle collegate
                                                For Each t In connectedTables
                                                    Debug.Print t & " (Collegata)"
                                                Next t
                                            
                                                ' Stampa le tabelle fisiche
                                                Debug.Print
                                                Debug.Print
                                                Debug.Print "Tabelle fisiche:"
                                                '//iterazione nella collezione fisica
                                                For Each t In physicalTables
                                                    Debug.Print t & " (Fisica)"
                                                Next t
                                            
                                
                                '//...............................................................................//
                                    
            
                        
                        
            
                 
                
                            ' Chiudi il database esterno
                            externalDB.Close
                            Set externalDB = Nothing
                
                Else '//If myScel_b
                    
                            '\\controllo finale sulla scelta
                            
                            If myScel_b = False Then
                                MsgBox "ATTENZIONE DB ESTERNO NON TROVATO USCITA DALLA ROUTINE!!", vbCritical
                                                          
                            End If
            
                        
            End If '//If myScel_b
            
            
'USCITA ED ERRORI
'..............................................................
Exit_ListTablesInExternalDB:
    Exit Sub

Err_ListTablesInExternalDB:
    MsgBox Err.Description
    Resume Exit_ListTablesInExternalDB

                                                      
End Sub




'//FUNZIONE PER IL CONTROLLO DELLE TABELLE DI SISTEMA
'//controlla la tabella corrente con l'elenco della collection systemTables


Private Function IsSystemTable(tableName As String, systemTables As Collection) As Boolean


    Dim tbl As Variant
    IsSystemTable = False
    For Each tbl In systemTables
        If tableName = tbl Then
            IsSystemTable = True
            Exit For
        End If
    Next tbl
'USCITA ED ERRORI
'..............................................................
Exit_IsSystemTable:
    Exit Function

Err_IsSystemTable:
    MsgBox Err.Description
    Resume Exit_IsSystemTable

                                                      
End Function



'//============================================================================================//


'//@CANCELLA@TABELLA_(in @db@esterno)
Public Sub DeleteSpecificTableInExternalDB_Psub()

    On Error GoTo Err_DeleteSpecificTableInExternalDB_Psub


    Dim dbPath As String
    Dim externalDB As DAO.Database
    Dim tdf As DAO.TableDef
    Dim tableName As String
    Dim tableExists_b As Boolean
    
    
    '//reset ed assegnazione
    iCount = 0
    
    '//ESEMPIO DISATTIVATO
    '//.................................................................//
    ' Nome della tabella da cancellare
    'tableName = "GEST_MENU_Tb03_}-----------------------------------------------@"

    ' Percorso del database esterno - esempio
    '//dbPath = "C:\Casa\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MDB\PROVA_CANCELLAZIONE_DB_ESTERNO\MENU_TB03_OGGETTI_DA_CANCELLARE.mdb"
    '//.................................................................//
    
    '//preparo la stringa ssql delle tabelle da cancellare
        sSql = ""
        sSql = sSql & ""
        sSql = sSql & "SELECT "
        sSql = sSql & "MsysDbEstTb01_OBJECT.TIPOGGETTO_s, "
        sSql = sSql & "MsysDbEstTb01_OBJECT.NOTA_OGGETTO_s, "
        sSql = sSql & "MsysDbEstTb01_OBJECT.Name1_s, "
        
        '//aggiunti per costruire la path
        sSql = sSql & "MsysDbEstTb01_OBJECT.DISCO_s, "
        sSql = sSql & "MsysDbEstTb01_OBJECT.PATH_s, "
        sSql = sSql & "MsysDbEstTb01_OBJECT.DATABASE_s, "
        
        '//ULTIMO campo senza virgola
        sSql = sSql & "MsysDbEstTb01_OBJECT.Scel_b "
        sSql = sSql & "FROM  "
        sSql = sSql & "MsysDbEstTb01_OBJECT "
        sSql = sSql & "WHERE (((MsysDbEstTb01_OBJECT.Scel_b) = True)) "
        sSql = sSql & "WITH OWNERACCESS OPTION;"

   
 
 
    '//APRI RS
    '//=======================================================================//
        
        '//apro il rs se popolato
        
        '//DEBUG CONTROLLO ED APERTURA
        Debug.Print sSql
                
        '//I AREA DI LAVORO = apro l'area di lavoro per il rs corrente
        Set daoDB = DBEngine.Workspaces(0).Databases(0)
        '//Apro il Database
        Set daoRS = daoDB.OpenRecordset(sSql)
        
        '//true = record non popolato non si apre il rs
        If daoRS.EOF = False And daoRS.BOF = False Then
        
           dbPath = daoRS.Fields("DISCO_s") & daoRS.Fields("PATH_s") & daoRS.Fields("DATABASE_s")
        '//ATTENZIONE
        ' II AREA DI LAVORO =  Apri il database esterno con area di lavoro 2 -- TODO: fornire la path
        Set externalDB = DBEngine.Workspaces(0).OpenDatabase(dbPath)

        
        
        '//Posizione Primo record
        daoRS.MoveFirst
        
            '//finche non è ultimo record
            While Not daoRS.EOF
            
                 '//Blocco iterazione
                 DoEvents
                
                '//salvo nella variabile il nome della tabella da ricercare
                tableName = daoRS.Fields("Name1_s")
                    
                    ' Controlla se la tabella esiste
                    tableExists = False
                    
                    For Each tdf In externalDB.TableDefs
                        
                        If tdf.Name = tableName Then
                            '//se esiste imposto a true la variabile
                            tableExists_b = True
                            Exit For
                        End If
                    Next tdf
                
                ' Se la tabella esiste = true, la cancella
                If tableExists_b Then
                    On Error Resume Next
                    externalDB.TableDefs.Delete tableName
                    
                  
                    
                    '//conteggio eliminazioni
                    iCount = iCount + 1
                    If Err.Number = 0 Then
                        Debug.Print "Tabella '" & tableName & "' cancellata con successo."
                                                   '//MESSAGGIO
                    MsgBox "TABELLA CANCELLATA CON SUCCESSO  :  " & iCount, vbExclamation

                    Else
                        Debug.Print "Errore nella cancellazione della tabella '" & tableName & "'."
                          '//MESSAGGIO
                    MsgBox "ERRORE NELLA CANCELLAZIIONE DELLA TABELLA NON ESISTE LA TABELLA NEL DB ESTERNO: " & iCount, vbExclamation
                  
                        Err.Clear
                    End If
                    On Error GoTo 0
                Else
                    Debug.Print "La tabella '" & tableName & "' non esiste nel database esterno."
                    
                    '//MESSAGGIO
                    MsgBox "NON ESISTE LA TABELLA NEL DB ESTERNO: " & iCount, vbExclamation
                    
                    
                End If
                
                    '//Record Successivo
                    daoRS.MoveNext
                
            
            Wend 'While Not DaoRs.EOF
                    
                    
                      MsgBox "TABELLA CANCELLATA CON SUCCESSO  :  " & iCount, vbExclamation

                    

                    '//Uscita Rs e chiusura oggetti
                    daoRS.Close
                    Set daoRS = Nothing
                    daoDB.Close
                    Set daoDB = Nothing
            
                ' Chiudi il database esterno
                externalDB.Close
                Set externalDB = Nothing
            
        
        End If '//If DaoRs.EOF = False And DaoRs.BOF = False Then
    

'//=======================================================================//



'USCITA ED ERRORI
'..............................................................
Exit_DeleteSpecificTableInExternalDB_Psub:
    Exit Sub

Err_DeleteSpecificTableInExternalDB_Psub:
    MsgBox Err.Description
    Resume Exit_DeleteSpecificTableInExternalDB_Psub

                                                      
End Sub




'//CANCELLA LA TABELLA NEL DB ESTERNO
'//====================================================================================//


'// FUNZIONE DI CANCELLAZIONE DELLE TABELLE
'//============================================================================================//

    
'//codice : @ROUTINE@ATTIVA_(@controllo@esterno@db @delete@cancella@tabelle@fisiche e @tabelle@Scollegate)
Public Sub DELETE_TablesInExternalDB()

    On Error GoTo Err_DELETE_TablesInExternalDB
    
    
    '//CREO LA COLLEZIONE TABELLE DI SISTEMA E LA POPOLO
    '//popola la collection delle tabelle di sistema
    ' Aggiungi i nomi delle tabelle di sistema da escludere
    Set systemTables = New Collection
    
    '//Aggiugno gli elementi della collezione
    systemTables.Add "MSysACEs"                     'MSysACEs
    systemTables.Add "MSysAccessObjects"            'MSysAccessObjects'
    systemTables.Add "MSysAccessStorage"            'MSysAccessStorage
                        
    systemTables.Add "MSysNameMap"                  'MSysNameMap
    systemTables.Add "MSysObjects"                  'MSysObjects'
    systemTables.Add "MSysQueries"                  'MSysQueries'
    systemTables.Add "MSysAccessXML"                'MSysAccessXML
    systemTables.Add "MSysRelationships"            'MSysRelationships'
                        
    systemTables.Add "MSysNavPaneGroupCategories"    'MSysNavPaneGroupCategories'
    systemTables.Add "MSysNavPaneGroupToObjects"     'MSysNavPaneGroupToObjects'
    systemTables.Add "MSysNavPaneObjectIDs"           'MSysNavPaneObjectIDs'
    systemTables.Add "MSysNavPaneGroups"              'MSysNavPaneGroups


    ' Inizializza le collezioni per tabelle fisiche e collegate
    Set connectedTables = New Collection
      
    
    
    '//APRO TABELLA PER INDIVIDUARE IL DATABASE ESTERNO
    '//.....................................................................................................//
            
            '//reset
    
            myDISCO_s = ""
            MyPath_s = ""
            myDATABASE_s = ""
            myScel_b = False
    
     '//Apro il Database
     Set daoDB = DBEngine.Workspaces(0).Databases(0)
     '//Apro un Recordset dal parametro ssql
     
            sSql = ""
            sSql = sSql & "SELECT MSysTb05_DB_EST.DISCO_s, "
            sSql = sSql & "MSysTb05_DB_EST.PATH_s, "
            sSql = sSql & "MSysTb05_DB_EST.DATABASE_s, "
            sSql = sSql & "MSysTb05_DB_EST.Scel_b "
            sSql = sSql & "FROM MSysTb05_DB_EST "
            sSql = sSql & "WHERE (((MSysTb05_DB_EST.Scel_b)=True)) "
            sSql = sSql & "WITH OWNERACCESS OPTION;"
            
            
            '//DEBUG CONTROLLO ED APERTURA
            Debug.Print sSql
     
     Set daoRS = daoDB.OpenRecordset(sSql)
        
    If daoRS.EOF = False And daoRS.BOF = False Then
        '//Posizione Primo record
        daoRS.MoveFirst
            While Not daoRS.EOF
              '//Blocco iterazione
                 DoEvents
                    
                    '//CONTROLLO SE SCELTO IL DB DA CONTROLLARE
                    If daoRS.Fields("Scel_b") = True Then
                        
                        If IsNull(daoRS.Fields("DISCO_s")) Or IsNull(daoRS.Fields("PATH_s")) Or _
                           IsNull(daoRS.Fields("DATABASE_s")) Then
                           
                           MsgBox "ATTENZIONE DISCO/PATH/DB SONO NULLI - > " & " DISCO: " & daoRS.Fields("DISCO_s") & Chr$(13) _
                                 & " PATH: " & daoRS.Fields("PATH_s") & Chr$(13) _
                                 & " DATABASE: " & daoRS.Fields("DATABASE_s") & Chr$(13) _
                                 & " USCITA DALLA ROUTINE!!!", vbCritical
                                '//Uscita Rs e chiusura oggetti
                                daoRS.Close
                                Set daoRS = Nothing
                                
                                GoTo Exit_DELETE_TablesInExternalDB

                           
                        End If
                        
                        
                        
                        '//imposto a scelta si
                        myScel_b = True
                        '//imposto la path trovata + directory e mdb
                        dbPath = daoRS.Fields("DISCO_s") & daoRS.Fields("PATH_s") & daoRS.Fields("DATABASE_s")
                        
                        '//DISCO, PATH , DB
                        myDISCO_s = daoRS.Fields("DISCO_s")
                        MyPath_s = daoRS.Fields("PATH_s")
                        myDATABASE_s = daoRS.Fields("DATABASE_s")
                        
                        '//directory completa DISO + PATH
                        myDirectory_s = daoRS.Fields("DISCO_s") & daoRS.Fields("PATH_s")
                          
                            'TODO: fare un controllo di esistenza path e db!!
                            Debug.Print dbPath
                        
                        '//trovato la path vado a fine rs per uscire dal db
                        daoRS.MoveLast
                        
                    End If
                    
                                                    
                '//Record Successivo
                daoRS.MoveNext
    
        Wend
    
            
        '//Uscita Rs e chiusura oggetti
        daoRS.Close
        Set daoRS = Nothing
        
        End If  '//If DAORs.EOF = False And DAORs.BOF = False Then
                
                
                    '
            '//IMPOSTAZIONE PATH E CONTROLLO ESISTENZA DIRECTORY
            '//------------------------------------------------------------------------------//
            '//NOTE     : controllo l'esistenza della path definita dai salvataggi se non esiste _
                        esco dalla routine.
                
                '//VALORIZZO I PARAMETRI
                par_Directory_s = dbPath
                ParametroFile_i = myDATABASE_s
                
                        
                'Str1 = Dir(Path_s, 16)
                'Vv1 = Dir("*.TXT", 2)
                'MyPath = "c:\CASA\GE_CASA\GE_MARINO\GESTIONE_SPESE\ARCHIVI_XLS\"    ' Imposta il percorso.
                'MYNAME = Dir(MyPath, vbDirectory)    ' Recupera la prima voce.
                'MYNAME = Dir(par_Directory_s, vbDirectory)    ' Recupera la prima voce.
                Vv1 = Dir(par_Directory_s, vbDirectory)   ' Recupera la prima voce.
                
               ' Vv1 = Dir("c:\", vbDirectory)
                
                If Vv1 = "" Then
                        MsgBox "NON ESISTE LA DIRECTORY /DATABASE ---> " & myDirectory_s & myDATABASE_s & " - USCITA DALLA ROUTINE"
                        GoTo Exit_DELETE_TablesInExternalDB
                End If
            '//-------------------------------------------------------------------------------//


                
    '//** FINE **
    '//ITERAZIONE_RECORSET
    '//.....................................................................................................//
    
            '//SOLO SE ESISTE IL DB FACCIO IL CONTROLLO
            If myScel_b = True Then
                
                '//svuoto la tabella tmp + RESET
                    CurrentDb.Execute "MsysDbEstTb01Qry01_Dlt01_OBJECT_TMP"
                    
                    iCount = 0
                    
                
                '//APRO LE COLLECTION E IL DB ESTERNO
                 Set physicalTables = New Collection
                        'todo: FARE UN CONTROLLO PRELIMINARE DELLA PATH E DEL FILE!!
                ' Apri il database esterno
                Set externalDB = DBEngine.Workspaces(0).OpenDatabase(dbPath)
                                
                                '//prova query
                                
                                
                           
                                ' Scansiona tutte le tabelle nel database esterno
                                '//...............................................................................//
                                    For Each tdf In externalDB.TableDefs
                                        ' Verifica se la tabella è di sistema con il confronte nella collection precaricata _
                                          e chiama funzione di controllo
                                          
                                        
                                
                                           '//resetto le variabili ad oni ciclo
                                           myNOTA_OGGETTO_s = ""
                                    
                                            
                                            
                                            '//tabella corrente
                                            Debug.Print "controllo tabella corrente da esaminare"
                                            Debug.Print tdf.Name
                                        
                                        
                                        If IsSystemTable(tdf.Name, systemTables) Then
                                            Debug.Print tdf.Name & " (Tavola di sistema)"
                                                
                                      
                                                
                                      
                                            
                                                '//TABELLA DI SISTEMA NON SI CANCELLA
                                                '//..........................................................................//
                                                            '// NON CANCELLO TABLE DI SISTEMA
                                                '//..........................................................................//
                                                
                                            
                                          'se la tabella è collegata aggiunge alla collection TABELLE FISICHE
                                        ElseIf Len(tdf.Connect) > 0 Then
                                        
                                            
                                                
                                            
                                                '//@CANCELLAZIONE@TABELLE_(CANCELLO LA @TABELLA@COLLEGATA SE ESISTE)
                                                '//..........................................................................//
                                                '// Note : Aggiungo la tabella fisica all'insieme pe la futura cancellazione
                                                        
                                                        '//TABELLA COLLEGATA aggiunta all'insieme tabelle collegate
                                                        tableExists = True
                                                        connectedTables.Add tdf.Name
                                                        CANCELLA tdf.Name
                                                        iCount = iCount + 1
                                                         
                                                '//..........................................................................//
                                                
                                            
                                            
                                        Else
                                            
                                            
                                      
                                                    
                                                  
                                            
                                                '// '//@CANCELLAZIONE@TABELLE_(CANCELLO LA @TABELLA@FISICA SE ESISTE)
                                                '//..........................................................................//
                                                '// Note : cancellazione della tabella fisica
                                                        'la tabella è fisica e la aggiunge alla collection TABELLE FISICHE
                                                            
                                                            physicalTables.Add tdf.Name
                                                            tableExists = True
                                                            CANCELLA tdf.Name
                                                              iCount = iCount + 1
                                      
                                                '//..........................................................................//
                                                
                                            
                                            
                                        End If
                                        
                                                            tableExists = False
                                    
                                    Next tdf
                                    
                                                
                                            
                                                ' Stampa le tabelle collegate con le collezioni
                                                Debug.Print
                                                Debug.Print
                                                Debug.Print "Tabelle collegate:"
                                                'iterazione nella collection delle tabelle collegate
                                                For Each t In connectedTables
                                                    Debug.Print t & " (Collegata)"
                                           
                                                Next t
                                            
                                                ' Stampa le tabelle fisiche
                                                Debug.Print
                                                Debug.Print
                                                Debug.Print "Tabelle fisiche:"
                                                '//iterazione nella collezione fisica
                                                For Each t In physicalTables
                                                    Debug.Print t & " (Fisica)"
                                                    
                                                    
                                                    
                                                Next t
                                            
                                                     MsgBox "TABELLE CANCELLATE CON SUCCESSO  :  " & iCount, vbExclamation

                                '//...............................................................................//
                                    
                 
                
                            ' Chiudi il database esterno
                            externalDB.Close
                            Set externalDB = Nothing
                
                Else '//If myScel_b
                    
                            '\\controllo finale sulla scelta
                            
                            If myScel_b = False Then
                                MsgBox "ATTENZIONE DB ESTERNO NON TROVATO USCITA DALLA ROUTINE!!", vbCritical
                                                          
                            End If
            
                        
            End If '//If myScel_b
            
            
'USCITA ED ERRORI
'..............................................................
Exit_DELETE_TablesInExternalDB:
    Exit Sub

Err_DELETE_TablesInExternalDB:
    MsgBox Err.Description
    Resume Exit_DELETE_TablesInExternalDB

                                                      
End Sub


'//CANCELLA LA TABELLA NEL DB ESTERNO
'//============================================================================================//
Private Sub CANCELLA(par_TableName As String)



 Set externalDB2 = DBEngine.Workspaces(0).OpenDatabase(dbPath)
 ' Se la tabella esiste, la cancella
    If tableExists Then
        On Error Resume Next
        externalDB2.TableDefs.Delete par_TableName
        If Err.Number = 0 Then
            Debug.Print "Tabella '" & par_TableName & "' cancellata con successo."
        Else
            Debug.Print "Errore nella cancellazione della tabella '" & par_TableName & "'."
            Err.Clear
        End If
        On Error GoTo 0
    Else
        Debug.Print "La tabella '" & tableName & "' non esiste nel database esterno."
    End If

    ' Chiudi il database esterno
    externalDB2.Close
    Set externalDB2 = Nothing
End Sub


'//CANCELLA LA TABELLA NEL DB ESTERNO *** FINE ***
'//====================================================================================//




'//************************************************************************************//
'//             GESTIONE DEGLI OGGETTI QUERY
'//************************************************************************************//



'//ACCESSO AL DB ESTERNO - controllo query
'//============================================================================================//
'//NOTE : la routine e la funzione aprono una istanza presso il db esterno access per il _
        controllo delle QUERY. La routine utilizza 3 collezioni di oggetti che per quanto _
        riguarda la systemTables = precarico quali QUERY sono di sistema poi crea altre _
        due collection che vengono popolate se il controllo if IsSystemTable restituisce false _
        perche chiama la funzione per il controllo se la tabella appartiene al sistema, per esclusione _
        appartiene alla fisiche o alle collegate, e quindi popola le due collection che poi vengono _
        stampate _
        'Percorso del database esterno COME ESEMPIO.
        'dbPath = "c:\Casa\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MDB\MENU_TB03_OGGETTI_DA_CANCELLARE\MENU_TB03_OGGETTI_DA_CANCELLARE.mdb"

                
'//codice : @ROUTINE@ATTIVA@QUERY_(@controllo@esterno@db delle @query@esterne)
Public Sub ListQUERYESInExternalDB()

    On Error GoTo Err_ListQUERYESInExternalDB
                
                
                '//CREO LA COLLEZIONE QUERY DI SISTEMA E LA POPOLO
                '//popola la collection delle query di sistema
                ' Aggiungi i nomi delle query di sistema da escludere
                Set systemQUERYes = New Collection
                '//Aggiugno gli elementi della collezione
                systemQUERYes.Add "MSysACEs"                     'MSysACEs
                systemQUERYes.Add "MSysAccessObjects"            'MSysAccessObjects'
                systemQUERYes.Add "MSysAccessStorage"            'MSysAccessStorage
                      
                systemQUERYes.Add "MSysNameMap"                  'MSysNameMap
                systemQUERYes.Add "MSysObjects"                  'MSysObjects'
                systemQUERYes.Add "MSysQueries"                  'MSysQueries'
                systemQUERYes.Add "MSysAccessXML"                'MSysAccessXML
                systemQUERYes.Add "MSysRelationships"            'MSysRelationships'
                              
                systemQUERYes.Add "MSysNavPaneGroupCategories"    'MSysNavPaneGroupCategories'
                systemQUERYes.Add "MSysNavPaneGroupToObjects"     'MSysNavPaneGroupToObjects'
                systemQUERYes.Add "MSysNavPaneObjectIDs"           'MSysNavPaneObjectIDs'
                systemQUERYes.Add "MSysNavPaneGroups"              'MSysNavPaneGroups


                ' Inizializza le collezioni per query fisiche e collegate
                Set connectedQUERYes = New Collection
                  
                
                
                '//APRO TABELLA PER INDIVIDUARE IL DATABASE ESTERNO
                '//.....................................................................................................//
                        
                        '//reset
                
                        myDISCO_s = ""
                        MyPath_s = ""
                        myDATABASE_s = ""
                        myScel_b = False
                
                 '//Apro il Database
                 Set daoDB = DBEngine.Workspaces(0).Databases(0)
                 '//Apro un Recordset dal parametro ssql
                 
                        sSql = ""
                        sSql = sSql & "SELECT MSysTb05_DB_EST.DISCO_s, "
                        sSql = sSql & "MSysTb05_DB_EST.PATH_s, "
                        sSql = sSql & "MSysTb05_DB_EST.DATABASE_s, "
                        sSql = sSql & "MSysTb05_DB_EST.Scel_b "
                        sSql = sSql & "FROM MSysTb05_DB_EST "
                        sSql = sSql & "WHERE (((MSysTb05_DB_EST.Scel_b)=True)) "
                        sSql = sSql & "WITH OWNERACCESS OPTION;"
                        
                        
                        '//DEBUG CONTROLLO ED APERTURA
                        Debug.Print sSql
                 
                 Set daoRS = daoDB.OpenRecordset(sSql)
                    
                If daoRS.EOF = False And daoRS.BOF = False Then
                    '//Posizione Primo record
                    daoRS.MoveFirst
                        While Not daoRS.EOF
                          '//Blocco iterazione
                             DoEvents
                                
                                '//CONTROLLO SE SCELTO IL DB DA CONTROLLARE
                                If daoRS.Fields("Scel_b") = True Then
                                    
                                    If IsNull(daoRS.Fields("DISCO_s")) Or IsNull(daoRS.Fields("PATH_s")) Or _
                                       IsNull(daoRS.Fields("DATABASE_s")) Then
                                       
                                       MsgBox "ATTENZIONE DISCO/PATH/DB SONO NULLI - > " & " DISCO: " & daoRS.Fields("DISCO_s") & Chr$(13) _
                                             & " PATH: " & daoRS.Fields("PATH_s") & Chr$(13) _
                                             & " DATABASE: " & daoRS.Fields("DATABASE_s") & Chr$(13) _
                                             & " USCITA DALLA ROUTINE!!!", vbCritical
                                            '//Uscita Rs e chiusura oggetti
                                            daoRS.Close
                                            Set daoRS = Nothing
                                            
                                            GoTo Exit_ListQUERYESInExternalDB

                                       
                                    End If
                                    
                                    
                                    
                                    '//imposto a scelta si
                                    myScel_b = True
                                    '//imposto la path trovata + directory e mdb
                                    dbPath = daoRS.Fields("DISCO_s") & daoRS.Fields("PATH_s") & daoRS.Fields("DATABASE_s")
                                    
                                    '//DISCO, PATH , DB
                                    myDISCO_s = daoRS.Fields("DISCO_s")
                                    MyPath_s = daoRS.Fields("PATH_s")
                                    myDATABASE_s = daoRS.Fields("DATABASE_s")
                                    
                                    '//directory completa DISO + PATH
                                    myDirectory_s = daoRS.Fields("DISCO_s") & daoRS.Fields("PATH_s")
                                      
                                        'TODO: fare un controllo di esistenza path e db!!
                                        Debug.Print dbPath
                                    
                                    '//trovato la path vado a fine rs per uscire dal db
                                    daoRS.MoveLast
                                    
                                End If
                                
                                                                
                            '//Record Successivo
                            daoRS.MoveNext
                
                    Wend
                
                        
                    '//Uscita Rs e chiusura oggetti
                    daoRS.Close
                    Set daoRS = Nothing
                    
                    End If  '//If DAORs.EOF = False And DAORs.BOF = False Then
                            
                            
                                '
                        '//IMPOSTAZIONE PATH E CONTROLLO ESISTENZA DIRECTORY
                        '//------------------------------------------------------------------------------//
                        '//NOTE     : controllo l'esistenza della path definita dai salvataggi se non esiste _
                                    esco dalla routine.
                            
                            '//VALORIZZO I PARAMETRI
                            par_Directory_s = dbPath
                            ParametroFile_i = myDATABASE_s
                            
                                    
                            'Str1 = Dir(Path_s, 16)
                            'Vv1 = Dir("*.TXT", 2)
                            'MyPath = "c:\CASA\GE_CASA\GE_MARINO\GESTIONE_SPESE\ARCHIVI_XLS\"    ' Imposta il percorso.
                            'MYNAME = Dir(MyPath, vbDirectory)    ' Recupera la prima voce.
                            'MYNAME = Dir(par_Directory_s, vbDirectory)    ' Recupera la prima voce.
                            Vv1 = Dir(par_Directory_s, vbDirectory)   ' Recupera la prima voce.
                            
                           ' Vv1 = Dir("c:\", vbDirectory)
                            
                            If Vv1 = "" Then
                                    MsgBox "NON ESISTE LA DIRECTORY /DATABASE ---> " & myDirectory_s & myDATABASE_s & " - USCITA DALLA ROUTINE"
                                    GoTo Exit_ListQUERYESInExternalDB
                            End If
                        '//-------------------------------------------------------------------------------//


                            
                '//** FINE **
                '//ITERAZIONE_RECORSET
                '//.....................................................................................................//
                
                        '//SOLO SE ESISTE IL DB FACCIO IL CONTROLLO
                        If myScel_b = True Then
                            
                            '//svuoto la tabella tmp + RESET
                                CurrentDb.Execute "MsysDbEstTb01Qry01_Dlt01_OBJECT_TMP"
                                
                                iCount = 0
                                
                            
                            '//APRO LE COLLECTION E IL DB ESTERNO
                             Set physicalQUERYes = New Collection
                                    'todo: FARE UN CONTROLLO PRELIMINARE DELLA PATH E DEL FILE!!
                            ' Apri il database esterno
                            Set externalDB = DBEngine.Workspaces(0).OpenDatabase(dbPath)
                                            
                                                
                                            ' Scansiona tutte le tabelle nel database esterno
                                            '//...............................................................................//
                                                
                                    
  
                                                For Each Qrydf In externalDB.QueryDefs
                                                    ' Verifica se la tabella è di sistema con il confronte nella collection precaricata _
                                                      e chiama funzione di controllo
                                                      
                                                    
                                            
                                                       '//resetto le variabili ad oni ciclo
                                                       myNOTA_OGGETTO_s = ""
                                                
                                                    
                                                          Debug.Print Qrydf.Name & " (LE QUERY ESAMINATE)"
                                                    
                                                    If IsSystemQUERY(Qrydf.Name, systemQUERYes) Then
                                                        Debug.Print Qrydf.Name & " (QUERY di sistema)"
                                                            
                                                            iCount = iCount + 1
                                                            
                                                        myNOTA_OGGETTO_s = "QUERY MSYS (QUERY di sistema)"
                                                        myNOTEex_s = "OGGETTO DATABASE ESTERNO " & myDATABASE_s
                                                        
                                                            '//QUI AGGIUNGERE SQL INSERIMENTO IN TABELLA TMP
                                                            '//..........................................................................//
                                                                    sSql = ""
                                                                    sSql = sSql & "INSERT INTO "
                                                                    sSql = sSql & "MsysDbEstTb01_OBJECT_TMP "
                                                                    sSql = sSql & "( NRO_OGGETTO_i, TIPOGGETTO_s, COD_PROGETTO_s, NOTA_OGGETTO_s, NOTEex_s,DISCO_s,PATH_s, DATABASE_s, Name1_s ) "
                                                                    sSql = sSql & "SELECT "
                                                                    sSql = sSql & iCount & " AS [NRO], "
                                                                    sSql = sSql & "'TABLE' AS TIPOGGETTO_s, "
                                                                    sSql = sSql & "'MsysDbEst' AS COD_PROGETTO_s,"
                                                                    sSql = sSql & "'" & myNOTA_OGGETTO_s & "' AS [NOTE], "
                                                                    sSql = sSql & "'" & myNOTEex_s & "' AS [NOTE_EX], "
                                                                    
                                                                    '//AGGIUNTO DISCO + PATH + DB
                                                                    sSql = sSql & "'" & myDISCO_s & "' AS [DISCO_s], "
                                                                    sSql = sSql & "'" & MyPath_s & "' AS [PATH_s], "
                                                                    sSql = sSql & "'" & myDATABASE_s & "' AS [DATABASE_s], "
                                                                    '//ULTIMO campo senza virgola
                                                                    sSql = sSql & "'" & Qrydf.Name & "' AS [Name1_s] "
                                                                    
                                                                                                                            
                                                                    
                                                                    sSql = sSql & "WITH OWNERACCESS OPTION;"
                                                                    
                                                                    '//CONTROLLO ED ESECUZIONE
                                                                    Debug.Print
                                                                    
                                                                    CurrentDb.Execute (sSql)
                                                            '//..........................................................................//
                                                            
                                                        
                                                      'se la tabella è collegata aggiunge alla collection TABELLE FISICHE
                                                    ElseIf Len(Qrydf.Connect) > 0 Then
                                                    
                                                        connectedQUERYes.Add Qrydf.Name
                                                        myNOTA_OGGETTO_s = "QUERY Collegata??"
                                                        myNOTEex_s = "OGGETTO DATABASE ESTERNO " & myDATABASE_s
                                                        
                                                            iCount = iCount + 1
                                                        
                                                            '//QUI AGGIUNGERE SQL INSERIMENTO IN TABELLA TMP
                                                            '//..........................................................................//
                                                                    sSql = ""
                                                                    sSql = sSql & "INSERT INTO "
                                                                    sSql = sSql & "MsysDbEstTb01_OBJECT_TMP "
                                                                    'sSql = sSql & "( NRO_OGGETTO_i, TIPOGGETTO_s, COD_PROGETTO_s, NOTA_OGGETTO_s, NOTEex_s, Name1_s ) "
                                                                    sSql = sSql & "( NRO_OGGETTO_i, TIPOGGETTO_s, COD_PROGETTO_s, NOTA_OGGETTO_s, NOTEex_s,DISCO_s,PATH_s, DATABASE_s, Name1_s ) "
                                                                    sSql = sSql & "SELECT "
                                                                    sSql = sSql & iCount & " AS [NRO], "
                                                                    sSql = sSql & "'TABLE' AS TIPOGGETTO_s, "
                                                                    sSql = sSql & "'MsysDbEst' AS COD_PROGETTO_s,"
                                                                    sSql = sSql & "'" & myNOTA_OGGETTO_s & "' AS [NOTE], "
                                                                    sSql = sSql & "'" & myNOTEex_s & "' AS [NOTE_EX], "
                                                                    
                                                                    '//AGGIUNTO DISCO + PATH + DB
                                                                    sSql = sSql & "'" & myDISCO_s & "' AS [DISCO_s], "
                                                                    sSql = sSql & "'" & MyPath_s & "' AS [PATH_s], "
                                                                    sSql = sSql & "'" & myDATABASE_s & "' AS [DATABASE_s], "
                                                                    '//ULTIMO campo senza virgola
                                                                    sSql = sSql & "'" & tdf.Name & "' AS [Name1_s] "
                                                                    
                                                                    sSql = sSql & "WITH OWNERACCESS OPTION;"
                                                                    
                                                                    '//CONTROLLO ED ESECUZIONE
                                                                    Debug.Print sSql
                                                                    
                                                                    CurrentDb.Execute (sSql)
                                                            '//..........................................................................//
                                                            
                                                        
                                                        
                                                    Else
                                                        
                                                        'la QUERY è fisica e la aggiunge alla collection QUERY FISICA
                                                        physicalQUERYes.Add Qrydf.Name
                                                        
                                                        myNOTA_OGGETTO_s = "QUERY (Fisica)"
                                                        myNOTEex_s = "OGGETTO DATABASE ESTERNO " & myDATABASE_s
                                                                
                                                                iCount = iCount + 1
                                                        
                                                            '//QUI AGGIUNGERE SQL INSERIMENTO IN TABELLA TMP
                                                            '//..........................................................................//
                                                                    sSql = ""
                                                                    sSql = sSql & "INSERT INTO "
                                                                    sSql = sSql & "MsysDbEstTb01_OBJECT_TMP "
                                                                    'sSql = sSql & "( NRO_OGGETTO_i, TIPOGGETTO_s, COD_PROGETTO_s, NOTA_OGGETTO_s, NOTEex_s, Name1_s ) "
                                                                    sSql = sSql & "( NRO_OGGETTO_i, TIPOGGETTO_s, COD_PROGETTO_s, NOTA_OGGETTO_s, NOTEex_s,DISCO_s,PATH_s, DATABASE_s, Name1_s ) "
                                                                    sSql = sSql & "SELECT "
                                                                    sSql = sSql & iCount & " AS [NRO], "
                                                                    sSql = sSql & "'QUERY' AS TIPOGGETTO_s, "
                                                                    sSql = sSql & "'MsysDbEst' AS COD_PROGETTO_s,"
                                                                    sSql = sSql & "'" & myNOTA_OGGETTO_s & "' AS [NOTE], "
                                                                    sSql = sSql & "'" & myNOTEex_s & "' AS [NOTE_EX], "
                                                                   
                                                                       '//AGGIUNTO DISCO + PATH + DB
                                                                    sSql = sSql & "'" & myDISCO_s & "' AS [DISCO_s], "
                                                                    sSql = sSql & "'" & MyPath_s & "' AS [PATH_s], "
                                                                    sSql = sSql & "'" & myDATABASE_s & "' AS [DATABASE_s], "
                                                                    '//ULTIMO campo senza virgola
                                                                    sSql = sSql & "'" & Qrydf.Name & "' AS [Name1_s] "
                                                                                                                         
                                                                    sSql = sSql & "WITH OWNERACCESS OPTION;"
                                                                    
                                                                    '//CONTROLLO ED ESECUZIONE
                                                                    Debug.Print
                                                                    
                                                                    CurrentDb.Execute (sSql)
                                                            '//..........................................................................//
                                                            
                                                        
                                                        
                                                    End If
                                                
                                                Next Qrydf
                                                        
                                                            ' Stampa le QUERY collegate ?? con le collezioni
                                                            Debug.Print
                                                            Debug.Print
                                                            Debug.Print "QUERY collegate ??:"
                                                            'iterazione nella collection delle tabelle collegate
                                                            For Each t In connectedQUERYes
                                                                Debug.Print t & " (QUERY Collegata)"
                                                            Next t
                                                        
                                                            ' Stampa le QUERY fisiche
                                                            Debug.Print
                                                            Debug.Print
                                                            Debug.Print "QUERY fisiche:"
                                                            '//iterazione nella collezione fisica
                                                            For Each t In physicalQUERYes
                                                                Debug.Print t & " (Fisica)"
                                                            Next t
                                                        
                                            
                                            '//...............................................................................//
                                                
                        
                                    
                                    
                        
                             
                            
                                        ' Chiudi il database esterno
                                        externalDB.Close
                                        Set externalDB = Nothing
                            
                            Else '//If myScel_b
                                
                                        '\\controllo finale sulla scelta
                                        
                                        If myScel_b = False Then
                                            MsgBox "ATTENZIONE DB ESTERNO NON TROVATO USCITA DALLA ROUTINE!!", vbCritical
                                                                      
                                        End If
                        
                                    
                        End If '//If myScel_b
                        
                        
            'USCITA ED ERRORI
            '..............................................................
Exit_ListQUERYESInExternalDB:
                Exit Sub

Err_ListQUERYESInExternalDB:
                MsgBox Err.Description
                Resume Exit_ListQUERYESInExternalDB

                                                                  
End Sub


'//FUNZIONE PER IL CONTROLLO DELLE QUERY DI SISTEMA
'//controlla la tabella corrente con l'elenco della collection systemTables


Private Function IsSystemQUERY(QUERYName As String, systemQUERYs As Collection) As Boolean


    Dim qry As Variant              'oggetto query variant
    IsSystemQUERY = False
    For Each qry In systemQUERYes
        Vv1 = Left(QUERYName, 3)
      
   
        Debug.Print "--------------CONTROLLO 3 CARATTERI QUERY -----------------------"
        Debug.Print Vv1
        
        If QUERYName = qry Or Left(QUERYName, 3) = "~sq" Then
            IsSystemQUERY = True
            Exit For
        End If
    Next qry
'USCITA ED ERRORI
'..............................................................
Exit_IsSystemQUERY:
    Exit Function

Err_IsSystemQUERY:
    MsgBox Err.Description
    Resume Exit_IsSystemQUERY

                                                      
End Function




'//************************************************************************************//
'//             GESTIONE DEGLI OGGETTI QUERY  *** FINE ***
'//************************************************************************************//





'//PROVA DI GESTIONE FORMS
'//**********************************************************************************************//
'//codice : @ROUTINE@ATTIVA_(@controllo@esterno@db @delete@cancella@Form@fisiche e @Form@Scollegate)


'// prova la funzione di cancellazione delle form *** fine ***
'//..................................................................................//

Public Sub DELETE_FormsInExternalDB()
    On Error GoTo Err_DELETE_FormsInExternalDB

    ' Creare la collezione per le Form di sistema
    Set systemForms = New Collection
    ' Aggiungere i nomi delle Form di sistema da escludere
    systemForms.Add "MSysACEs"
    systemForms.Add "MSysAccessObjects"
    systemForms.Add "MSysAccessStorage"
    systemForms.Add "MSysNameMap"
    systemForms.Add "MSysObjects"
    systemForms.Add "MSysQueries"
    systemForms.Add "MSysAccessXML"
    systemForms.Add "MSysRelationships"
    systemForms.Add "MSysNavPaneGroupCategories"
    systemForms.Add "MSysNavPaneGroupToObjects"
    systemForms.Add "MSysNavPaneObjectIDs"
    systemForms.Add "MSysNavPaneGroups"

    ' Inizializza le collezioni per Form fisiche e collegate
    Set connectedFORMs = New Collection
    Set physicalForms = New Collection

    ' Apri il Database
    Set daoDB = DBEngine.Workspaces(0).Databases(0)
    
    sSql = "SELECT MSysTb05_DB_EST.DISCO_s, MSysTb05_DB_EST.PATH_s, MSysTb05_DB_EST.DATABASE_s, MSysTb05_DB_EST.Scel_b " & _
           "FROM MSysTb05_DB_EST " & _
           "WHERE MSysTb05_DB_EST.Scel_b = True " & _
           "WITH OWNERACCESS OPTION;"
    
    Debug.Print sSql
    Set daoRS = daoDB.OpenRecordset(sSql)

    If Not daoRS.EOF And Not daoRS.BOF Then
        daoRS.MoveFirst
        While Not daoRS.EOF
            DoEvents
            If daoRS.Fields("Scel_b") = True Then
                If IsNull(daoRS.Fields("DISCO_s")) Or IsNull(daoRS.Fields("PATH_s")) Or IsNull(daoRS.Fields("DATABASE_s")) Then
                    MsgBox "ATTENZIONE DISCO/PATH/DB SONO NULLI - > " & " DISCO: " & daoRS.Fields("DISCO_s") & Chr$(13) & _
                           " PATH: " & daoRS.Fields("PATH_s") & Chr$(13) & _
                           " DATABASE: " & daoRS.Fields("DATABASE_s") & Chr$(13) & _
                           " USCITA DALLA ROUTINE!!!", vbCritical
                    daoRS.Close
                    Set daoRS = Nothing
                    GoTo Exit_DELETE_FormsInExternalDB
                End If

                myScel_b = True
                dbPath = daoRS.Fields("DISCO_s") & daoRS.Fields("PATH_s") & daoRS.Fields("DATABASE_s")
                myDISCO_s = daoRS.Fields("DISCO_s")
                MyPath_s = daoRS.Fields("PATH_s")
                myDATABASE_s = daoRS.Fields("DATABASE_s")
                myDirectory_s = daoRS.Fields("DISCO_s") & daoRS.Fields("PATH_s")
                Debug.Print dbPath
                daoRS.MoveLast
            End If
            daoRS.MoveNext
        Wend

        daoRS.Close
        Set daoRS = Nothing
    End If

    ' Verifica se esiste la directory del database esterno
    par_Directory_s = dbPath
    ParametroFile_i = myDATABASE_s
    Vv1 = Dir(par_Directory_s, vbDirectory)

    If Vv1 = "" Then
        MsgBox "NON ESISTE LA DIRECTORY /DATABASE ---> " & myDirectory_s & myDATABASE_s & " - USCITA DALLA ROUTINE"
        GoTo Exit_DELETE_FormsInExternalDB
    End If

    If myScel_b = True Then
        CurrentDb.Execute "MsysDbEstTb01Qry01_Dlt01_OBJECT_TMP"
        iCount = 0

        ' Apri il database esterno
        Set externalDB = DBEngine.Workspaces(0).OpenDatabase(dbPath)

        ' Scansiona tutte le form nel database esterno
        Dim cont As Container
        Dim doc As Document
        Set cont = externalDB.Containers("Forms")
        
        Dim oggetto As Variant
        Dim appAccess As Object
        Set appAccess = CreateObject("Access.Application.9")
        
        
        
        For Each doc In cont.Documents
            myNOTA_OGGETTO_s = ""
            Debug.Print "controllo FORM corrente da esaminare"
            Debug.Print doc.Name

            If IsSystemFORM(CStr(doc.Name), systemForms) Then
                Debug.Print doc.Name & " (Form di sistema)"
            Else
                physicalForms.Add doc.Name
                CANCELLA_FORM_ESTERNA doc.Name
                iCount = iCount + 1
            End If
        Next doc

        Debug.Print "Form di sistema:"
        For Each t In connectedFORMs
            Debug.Print t & " (Collegata)"
        Next t

        Debug.Print "Form fisiche:"
        For Each t In physicalForms
            Debug.Print t & " (Fisica)"
        Next t

        MsgBox "FORM CANCELLATE CON SUCCESSO  :  " & iCount, vbExclamation

        externalDB.Close
        Set externalDB = Nothing
    Else
        If myScel_b = False Then
            MsgBox "ATTENZIONE DB ESTERNO NON TROVATO USCITA DALLA ROUTINE!!", vbCritical
        End If
    End If

Exit_DELETE_FormsInExternalDB:
    Exit Sub

Err_DELETE_FormsInExternalDB:
    MsgBox Err.Description
    Resume Exit_DELETE_FormsInExternalDB

End Sub



' Cancella la form nel database esterno
Private Sub CANCELLA_FORM_ESTERNA(par_FormName As String)
    Set externalDB2 = DBEngine.Workspaces(0).OpenDatabase(dbPath)
    On Error Resume Next
    
    '//bloccata perche .Documents.Delete  non funziona nel db esterno
    '//externalDB2.Containers("Forms").Documents.Delete par_FormName
    
        ' Elimina la form utilizzando DoCmd.DeleteObject = metodo alternativo con la macro!
    externalDB2.DoCmd.DeleteObject acForm, par_FormName

        
    
    If Err.Number = 0 Then
        Debug.Print "La Form '" & par_FormName & "' cancellata con successo."
    Else
        Debug.Print "Errore nella cancellazione della Form '" & par_FormName & "'."
        Err.Clear
    End If
    On Error GoTo 0

    externalDB2.Close
    Set externalDB2 = Nothing
End Sub



' Verifica se una form è di sistema
Private Function IsSystemFORM(ByVal formName As String, ByVal systemForms As Collection) As Boolean
    Dim f As Variant
    IsSystemFORM = False
    For Each f In systemForms
        If f = formName Then
            IsSystemFORM = True
            Exit For
        End If
    Next f
End Function


'//PROVA DI GESTIONE FORMS  *** FINE ***
'//**********************************************************************************************//

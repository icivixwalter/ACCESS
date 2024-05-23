Attribute VB_Name = "MsysDbEstTb01Mdl01_Num01_CONTROLLO_DB_ESTERNO"
'//MODULO = MsysDbEstTb01Mdl01_Num01_CONTROLLO_DB_ESTERNO _
            MODULO PER IL CONTROLLO DELLE TABELLE DI SISTEMA.




Option Compare Database


Dim dbs As Database
Dim DaoRs As DAO.Recordset

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
        
    '//database disco e path e campi
    Dim myDISCO_s As String
    Dim myPATH_s As String
    Dim myDATABASE_s As String
    Dim myDirectory_s As String
    
    '//campi
    Dim myScel_b As Boolean
    Dim myNOTA_OGGETTO_s As String
    Dim myNOTEex_s As String
    
    '//DIM path, db e tabella
    Dim dbPath As String
    Dim externalDB As DAO.Database
    Dim tdf As DAO.TableDef
    '//dim le collection per separare il tipo di tabelle
    Dim systemTables As Collection
    Dim connectedTables As Collection
    Dim physicalTables As Collection
    
    '//recupero il db dalla tabella
    
    '//popola la collection delle tabelle di sistema
    ' Aggiungi i nomi delle tabelle di sistema da escludere
    Set systemTables = New Collection
    systemTables.Add "MSysAccessObjects"
    systemTables.Add "MSysACEs"
    systemTables.Add "MSysObjects"
    systemTables.Add "MSysQueries"
    systemTables.Add "MSysAccessXML"
    systemTables.Add "MSysRelationships"
    systemTables.Add "MSysNavPaneGroupCategories"
    systemTables.Add "MSysNavPaneGroupToObjects"
    systemTables.Add "MSysNavPaneObjectIDs"
    systemTables.Add "MSysNavPaneGroups"

    ' Inizializza le collezioni per tabelle fisiche e collegate
    Set connectedTables = New Collection
      
    
    
    '//APRO TABELLA PER INDIVIDUARE IL DATABASE ESTERNO
    '//.....................................................................................................//
            
            '//reset
    
            myDISCO_s = ""
            myPATH_s = ""
            myDATABASE_s = ""
            myScel_b = False
    
     '//Apro il Database
     Set DaoDB = DBEngine.Workspaces(0).Databases(0)
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
     
     Set DaoRs = DaoDB.OpenRecordset(sSql)
        
    If DaoRs.EOF = False And DaoRs.BOF = False Then
        '//Posizione Primo record
        DaoRs.MoveFirst
            While Not DaoRs.EOF
              '//Blocco iterazione
                 DoEvents
                    
                    '//CONTROLLO SE SCELTO IL DB DA CONTROLLARE
                    If DaoRs.Fields("Scel_b") = True Then
                        
                        If IsNull(DaoRs.Fields("DISCO_s")) Or IsNull(DaoRs.Fields("PATH_s")) Or _
                           IsNull(DaoRs.Fields("DATABASE_s")) Then
                           
                           MsgBox "ATTENZIONE DISCO/PATH/DB SONO NULLI - > " & " DISCO: " & DaoRs.Fields("DISCO_s") & Chr$(13) _
                                 & " PATH: " & DaoRs.Fields("PATH_s") & Chr$(13) _
                                 & " DATABASE: " & DaoRs.Fields("DATABASE_s") & Chr$(13) _
                                 & " USCITA DALLA ROUTINE!!!", vbCritical
                                '//Uscita Rs e chiusura oggetti
                                DaoRs.Close
                                Set DaoRs = Nothing
                                
                                GoTo Exit_ListTablesInExternalDB

                           
                        End If
                        
                        
                        
                        '//imposto a scelta si
                        myScel_b = True
                        '//imposto la path trovata + directory e mdb
                        dbPath = DaoRs.Fields("DISCO_s") & DaoRs.Fields("PATH_s") & DaoRs.Fields("DATABASE_s")
                        
                        myDATABASE_s = DaoRs.Fields("DATABASE_s")
                        myDirectory_s = DaoRs.Fields("DISCO_s") & DaoRs.Fields("PATH_s")
                          
                            'TODO: fare un controllo di esistenza path e db!!
                            Debug.Print dbPath
                        
                        '//trovato la path vado a fine rs per uscire dal db
                        DaoRs.MoveLast
                        
                    End If
                    
                                                    
                '//Record Successivo
                DaoRs.MoveNext
    
        Wend
    
            
        '//Uscita Rs e chiusura oggetti
        DaoRs.Close
        Set DaoRs = Nothing
        
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
                                                        sSql = sSql & "( NRO_OGGETTO_i, TIPOGGETTO_s, COD_PROGETTO_s, NOTA_OGGETTO_s, NOTEex_s, Name1_s ) "
                                                        sSql = sSql & "SELECT "
                                                        sSql = sSql & iCount & " AS [NRO], "
                                                        sSql = sSql & "'TABLE' AS TIPOGGETTO_s, "
                                                        sSql = sSql & "'MsysDbEst' AS COD_PROGETTO_s,"
                                                        sSql = sSql & "'" & myNOTA_OGGETTO_s & "' AS [NOTE], "
                                                        sSql = sSql & "'" & myNOTEex_s & "' AS [NOTE_EX], "
                                                        
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
                                                        sSql = sSql & "( NRO_OGGETTO_i, TIPOGGETTO_s, COD_PROGETTO_s, NOTA_OGGETTO_s, NOTEex_s, Name1_s ) "
                                                        sSql = sSql & "SELECT "
                                                        sSql = sSql & iCount & " AS [NRO], "
                                                        sSql = sSql & "'TABLE' AS TIPOGGETTO_s, "
                                                        sSql = sSql & "'MsysDbEst' AS COD_PROGETTO_s,"
                                                        sSql = sSql & "'" & myNOTA_OGGETTO_s & "' AS [NOTE], "
                                                        sSql = sSql & "'" & myNOTEex_s & "' AS [NOTE_EX], "
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
                                                        sSql = sSql & "( NRO_OGGETTO_i, TIPOGGETTO_s, COD_PROGETTO_s, NOTA_OGGETTO_s, NOTEex_s, Name1_s ) "
                                                        sSql = sSql & "SELECT "
                                                        sSql = sSql & iCount & " AS [NRO], "
                                                        sSql = sSql & "'TABLE' AS TIPOGGETTO_s, "
                                                        sSql = sSql & "'MsysDbEst' AS COD_PROGETTO_s,"
                                                        sSql = sSql & "'" & myNOTA_OGGETTO_s & "' AS [NOTE], "
                                                        sSql = sSql & "'" & myNOTEex_s & "' AS [NOTE_EX], "
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





Private Sub MODELLO()

    On Error GoTo Err_MODELLO

    
            'Reset Variabili Oggetti Form
            m_sxTIPOGGETTO = ""
            m_sxPROPRIETA = ""
            m_sxMETODO = ""
            m_sxEVENTO = ""
            
            
'USCITA ED ERRORI
'..............................................................
Exit_MODELLO:
    Exit Sub

Err_MODELLO:
    MsgBox Err.Description
    Resume Exit_MODELLO

                                                      
End Sub


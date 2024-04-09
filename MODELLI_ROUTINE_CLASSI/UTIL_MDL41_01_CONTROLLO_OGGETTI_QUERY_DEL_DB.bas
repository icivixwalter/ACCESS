Attribute VB_Name = "UTIL_MDL41_01_CONTROLLO_OGGETTI_QUERY_DEL_DB"
Option Compare Text
Option Explicit

'//................................
'// DIM OGGETTI E PARAMETRI
Dim obTipoDB As Object

'//................................
'// DIM SCELTA DATABASE DA APRIRE
Dim DbNomeDatabase_s As String
Dim sOption_SEZ As Integer          '//Numero Sezione di Progetto Scelta
Dim sName_tab As String             '//Nome Tabella da creare

Dim dbs As Database
Dim Apridbs As Database
Dim ApriRs As Recordset

Dim intCiclo As Integer             '//Intero per ciclo lettura oggetti

'//................................
'// DIM TABELLA DA ESPORTARE
Dim ExportTable1 As String
Dim ExportTable2 As String

'//................................
'// DIM CREAZIONE NUOVE TABELLE
Dim TableDefNuovo    As TableDef
Dim ApriTableDef    As TableDef
Dim FieldNuovo  As Field
Dim idxNuovo_Contatore_Univoco  As Index        '// Oggetto Index
Dim idxNuovo_Testo_Univoco  As Index
Dim iIndice1    As Integer                            '// per il passaggio dei parametri
Dim prpPROPERTY1   As Property                  '// proprieta Description

Dim sNomeField  As String

'//................................
'// DIM OGGETTO QUERY

Dim sName_Query As String

'//................................
'// DIM LE OPZIONI
Dim Opzione_1 As Integer
Dim Opzione_2 As Integer
Dim sTipo_Oggetto As String

Dim iOperazione As Integer
Dim Flag As Boolean
        
        
'//................................
'// DIM LE VARIABILI LOCALI GENERALI
Dim vV1 As Variant
Dim sStr1 As String
Dim bBoolean1 As Boolean
Dim iInt1    As Integer
Dim dblDBL1   As Double
Dim objObject1 As Object


'//****************************************************************************************************************************
'//           PROCEDURA DI ITERAZIONE DEGLI OGGETTI DI DATABASE
'// i parametri ricevuti sono : par_Scelta db = path db
'//
'//
'//****************************************************************************************************************************

'//Funzione ITERAZIONE OGGETTI
'//Iterazione_oggetti_QUERY()
'//PARAMETRI IN IMPUT:
'//par_DbNomeDatabase_s_s =database ; par_Name_tab_s =nome tabella da ricercare; par_Name_Query_s =nome query da ricercare
'//par_sTipo_Oggetto = tipo oggetto ; tipo di operazione =?  ; par_Indice1 =? ; par_Flag = True = individuato
'//
'//OUTPUT : RESTITUISCE una stringa

Public Function Iterazione_oggetti_QUERY(par_DbNomeDatabase_s_s As String, _
                                         par_Name_Query_s As String, _
                                         par_sTipo_Oggetto As String, _
                                         par_iOperazione_i As Integer, _
                                         par_Indice1 As Integer, _
                                         par_Flag As Boolean) As String


        On Error GoTo Err_Iterazione_oggetti_QUERY


        '//------------------------------------------------------------------------------------------
        '//1.1    SET VARIABILI, OPEN DB, CONTROLLO OGGETTO NEL DB
         
                    '//OPEN DB: Apre il database scelto per la creazione o cancellazione
                    '//         con il parametro Option = false per indicare che l'//apertura è condivisa /
                    '//                          Option = True apertura esclusiva
                    
                If par_DbNomeDatabase_s_s = "" Then
                        '//apro il db corrente se non viene indicato un database esterno
                        Set dbs = CurrentDb
                    Else
                        '//apro un database esterno con la path >0 passata con i parametri
                        Set dbs = OpenDatabase(par_DbNomeDatabase_s_s, False)
                End If
            Select Case par_sTipo_Oggetto
                
                           
                Case Is = "Cmd_Elenco_Query"
                
                        '//ITERAZIONE QUERYDEFS <<LE QUERY>>
                        '//______________________________________________________________________________
                        '//OGGETTI QUERY
                        '//CONTROLLO OGGETTO NEL DB: Controllo l'//esistenza della QUERY
                            
                            
                            '//pulisco la tabella delle query prima del salvataggio dei nuovi nomi
                            CurrentDb.Execute ("DELETE MSys_ELENCO_OGGETTI_QUERY_salvate.* FROM MSys_ELENCO_OGGETTI_QUERY_salvate;")
                            
                            '//apro la tabella per il salvataggio delle query
                            Set ApriRs = dbs.OpenRecordset("MSys_ELENCO_OGGETTI_QUERY_salvate")
                            
                        '//itero nel db
                        With dbs
                            
                                Debug.Print .QueryDefs.Count & _
                                    " QueryDefs in database  " & DbNomeDatabase_s
                                    
                                    
                                For intCiclo = 0 To .QueryDefs.Count - 1
                                    
                                    '//controlli
                                    '//................................................................
                                        Debug.Print "  " & .QueryDefs(intCiclo).Name
                                        
                                        '//Stampo il contenuto sql della query
                                        Debug.Print
                                        Debug.Print "  " & .QueryDefs(intCiclo).sQL
                                        
                                        '//stampo il contenuto della query scelta direttamente
                                        Debug.Print "  " & .QueryDefs("MSsys_VS01_MSys_ELENCO_OGGETTI_QUERY").sQL
                                    '//................................................................
                                        
                                        '//SALVATAGGIO NOME QUERY
                                        '//................................................................
                                            ApriRs.AddNew
                                            ApriRs.Fields("Name1") = .QueryDefs(intCiclo).Name
                                            
                                            
                                            ApriRs.Update
                                        '//................................................................
                                        
                                    '//controllo nome query passata con parametro
                                    If .QueryDefs(intCiclo).Name = par_Name_Query_s Then
                                        
                                        bBoolean1 = True    '// Flag True = trovato QUERY
                                        
                                        '//recupero il contenuto della query
                                        Iterazione_oggetti_QUERY = .QueryDefs(intCiclo).sQL
                                        
                                        
                                        GoTo Exit_Iterazione_oggetti_QUERY  '//esci
                                    Else
                                        bBoolean1 = False   '// Flag False = non trovato QUERY
                                        Iterazione_oggetti_QUERY = bBoolean1 '//riassegno alla funzione il valore del flag
                                    End If
                                    
                                Next intCiclo
                                
                            End With
                    
                    Case Else
                            MsgBox "Scelta comando non effettuata", vbExclamation
                End Select
                            '//chiudo i rs aperti + Rilascio la memoria
                            ApriRs.Close
                            Set ApriRs = Nothing
                                    
'//------------------------------------------------------------------------------
'//                       FINE FUNCTIONE E GESTIONE ERRORI

Exit_Iterazione_oggetti_QUERY:
    dbs.Close
    Set dbs = Nothing

Err_Iterazione_oggetti_QUERY:
    MsgBox "ERRORE FUNCTION PUBLIC    " & Err.Number & " - " & Err.Description, vbCritical, "Iterazione_oggetti_QUERY"
    Resume Exit_Iterazione_oggetti_QUERY
    
End Function








'//****************************************************************************************************************************
'//           PROCEDURA CONTROLLO OGGETTI DATABASE
'// i parametri ricevuti sono : par_Scelta db = path db
'//                           : par_TableDefNuovo = Tabella da Accodare
'//                           : par_Name_tab_s = Nome della Tabella da Cancellare
'//                           : par_iOperazione_i = Tipo di operazione da eseguire (1=inserire tabella; 2=cancellare tabella)
'//                           : par_Flag = Segnala se l'//operazione è stata effettuata
'//
'//****************************************************************************************************************************

Public Function pfCONTROLLO_OGGETTI_DB(par_DbNomeDatabase_s_s As String, _
                                   par_Name_tab_s As String, par_Name_Query_s As String, _
                                   par_iOperazione_i As Integer, par_Indice1 As Integer, par_Flag As Boolean) As Boolean


        
        
        On Error GoTo Err_pfCONTROLLO_OGGETTI_DB
           
                
                
        
        '//------------------------------------------------------------------------------------------
        '//1.1    SET VARIABILI, OPEN DB, CONTROLLO OGGETTO NEL DB
         
                
                '//SET VARIABILI:
                DbNomeDatabase_s = par_DbNomeDatabase_s_s
                sName_tab = par_Name_tab_s
                iOperazione = par_iOperazione_i
                    '//iIndice1 = par_Index
                Flag = par_Flag
                sName_Query = ""
                
                '//OPEN DB: Apre il database scelto per la creazione o cancellazione
                Set dbs = OpenDatabase(DbNomeDatabase_s)


                
                '//CONTROLLO OGGETTO NEL DB: Controllo l'//esistenza della tabella Nell'//insieme Tabledefs
                '//......................................................................................
                    With dbs
                        
                        Debug.Print .TableDefs.Count & _
                            " TableDefs in database  " & DbNomeDatabase_s
                        For intCiclo = 0 To .TableDefs.Count - 1
                            Debug.Print "  " & .TableDefs(intCiclo).Name
                        
                            If .TableDefs(intCiclo).Name = par_Name_tab_s Then
                                bBoolean1 = True    '// Flag True = trovato tabelle si può cancellare
                                pfCONTROLLO_OGGETTI_DB = bBoolean1 '//riassegno alla funzione il valore del flag
                                GoTo Exit_pfCONTROLLO_OGGETTI_DB  '//esci
                            Else
                                bBoolean1 = False   '// Flag False = non trovato TABELLA
                                pfCONTROLLO_OGGETTI_DB = bBoolean1 '//riassegno alla funzione il valore del flag
                            End If
                            
                        Next intCiclo

                    End With
                 '//......................................................................................


'//EXIT E GESTIONE ERRORI
'//-----------------------------------------------------------------------------------------------

Exit_pfCONTROLLO_OGGETTI_DB:
    dbs.Close
    Set dbs = Nothing

Exit Function

Err_pfCONTROLLO_OGGETTI_DB:
    MsgBox "ERRORE FUNCTION PUBLIC    " & Err.Number & " - " & Err.Description, vbCritical, "pfCONTROLLO_OGGETTI_DB"
    Resume Exit_pfCONTROLLO_OGGETTI_DB
    


End Function


'//PROVA_LA_PROCEDURA
'//============================================================================================================================//
Sub CHIAMA_pfCONTROLLO_OGGETTO_QUERY_DB()

    sStr1 = pfCONTROLLO_OGGETTO_QUERY_DB("", "", "GiurVS05_N01_PARAMETRI_GESTIONE", 0, 0, False)

End Sub
'//============================================================================================================================//



'// PROCEDURA CONTROLLO OGGETTO QUERY DEL DB
'//============================================================================================================================//
'//
'//PARAMETRI------------------->: i parametri ricevuti sono
'//                             01) par_Scelta db = path db, la scelda del database, se null è interno.
'//                             02) par_TableDefNuovo = nome della Tabella da Accodare
'//                             03) par_Name_tab_s = il nome della Tabella da Cancellare
'//                             04) par_iOperazione_i = Tipo di operazione da eseguire (1=inserire tabella; 2=cancellare tabella)
'//                             05) par_Flag = Segnala se l'operazione è stata effettuata, True = Eseguita, False = non eseguita
'//
'//RESTITUISCE : Un valore Stringa
'//
'//ATTIVITA ------------------->:Itera nell'//oggetto database corrente individuando la QUERY inviata come parametro stringa, _
                            e recupero il valore sql della sua proprieta .SQL
'//
'//****************************************************************************************************************************

Public Function pfCONTROLLO_OGGETTO_QUERY_DB(par_DbNomeDatabase_s_s As String, _
                                             par_Name_tab_s As String, _
                                             par_Name_Query_s As String, _
                                             par_iOperazione_i As Integer, _
                                             par_Indice1 As Integer, _
                                             par_Flag As Boolean) As String
        
On Error GoTo Err_pfCONTROLLO_OGGETTO_QUERY_DB

    '//RESET
    Set dbs = CurrentDb

        '//QUALIFICAZIONE DB ED itero nel db
        '//-----------------------------------------------------------------------------------------------------//
            With dbs
                
                    Debug.Print .QueryDefs.Count & _
                        " QueryDefs in database  " & DbNomeDatabase_s
                        
                    '//CICLO_OGGETTI
                    '//-----------------------------------------------------------------------------------------------------//
                    For intCiclo = 0 To .QueryDefs.Count - 1
                        '//controlli
                        '//...........................................................................................//
                            
                            Debug.Print "--------------------NOME_QUERY--------------------------"
                            Debug.Print "  " & .QueryDefs(intCiclo).Name
                            
                            '//Stampo il contenuto sql della query
                            Debug.Print "la Stringa Sql della Quuery"
                            Debug.Print "  " & .QueryDefs(intCiclo).sQL
                            
                            '//stampo il tipo di query scelta direttamente
                            Debug.Print "IL TIPO DELLA Query"
                            Debug.Print "  " & .QueryDefs(intCiclo).Type
                            
                            Debug.Print "-----------------------------------------------------------"
                            Debug.Print
                        '//...........................................................................................//
                            
                            
                        '//controllo nome query passata con parametro
                        '//...........................................................................................//
                        If .QueryDefs(intCiclo).Name = par_Name_Query_s Then
                            
                            bBoolean1 = True    '// Flag True = trovato QUERY
                            
                            '//recupero il contenuto della query
                            pfCONTROLLO_OGGETTO_QUERY_DB = .QueryDefs(intCiclo).sQL
                            MsgBox "OGGETTO INDIVIDUATO ---> " & pfCONTROLLO_OGGETTO_QUERY_DB
                            
                            
                        Else
                            
                        End If
                        '//...........................................................................................//
                        
                    Next intCiclo
                '//CICLO_OGGETTI        *** FINE ***
                '//-----------------------------------------------------------------------------------------------------//
                
                    
                End With
        '//QUALIFICAZIONE DB ED itero nel db *** FINE ***
        '//-----------------------------------------------------------------------------------------------------//
        
        
           
                '//chiudo i rs aperti + Rilascio la memoria
                dbs.Close
                Set dbs = Nothing
      

'//EXIT E GESTIONE ERRORI
'//-----------------------------------------------------------------------------------------------

Exit_pfCONTROLLO_OGGETTO_QUERY_DB:
    
Exit Function

Err_pfCONTROLLO_OGGETTO_QUERY_DB:
    MsgBox "ERRORE FUNCTION PUBLIC    " & Err.Number & " - " & Err.Description, vbCritical, "pfCONTROLLO_OGGETTO_QUERY_DB"
    Resume Exit_pfCONTROLLO_OGGETTO_QUERY_DB

End Function

'// PROCEDURA CONTROLLO OGGETTO QUERY DEL DB        *** fine ***
'//============================================================================================================================//











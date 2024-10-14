Attribute VB_Name = "UTIL_Nro42_N01_CONTROLLO_OGGETTI_TABLE_DB_ESTERNO"
' UTIL_Nro42_N01_CONTROLLO_OGGETTI_TABLE_DB_ESTERNO
Option Compare Database

Dim Count_i As Integer



Private Sub ATTIVA()

Dim par_Scelta As String
Dim par_TableDefNuovo_s As String
Dim par_Name_tab_s As String

Dim par_iOperazione_i  As Integer
Dim par_Flag As Boolean


End Sub

'// CONTROLLO_OGGETTI_TABELLA
'//============================================================================================================================//
'//
'//NOTE ----------------------->: Ciclo oggetti
'//PARAMETRI------------------->: i parametri ricevuti sono
'//                             01) par_Scelta db = path db, la scelda del database, se null è interno.
'//                             02) par_TableDefNuovo = nome della Tabella da Accodare
'//                             03) par_Name_tab_s = il nome della Tabella da Cancellare
'//                             04) par_iOperazione_i = Tipo di operazione da eseguire (1=inserire tabella; 2=cancellare tabella)
'//                             05) par_Flag = Segnala se l'operazione è stata effettuata, True = Eseguita, False = non eseguita
'//
'//RESTITUISCE ---------------->: Un valore Stringa IL NOME DELLA TABELLA
'//
'//ATTIVITA ------------------->:Itera nell'//oggetto database corrente individuando la TABELLA inviata come parametro stringa, _
                                e recupero il valore sql della sua proprieta .SQL
'//NOTE------------------------>:controllo TUTTE LE TABELLE del db restituendo il _
                                nome delle stesse.
'//Codice---------------------->:DIRECTORY_CREAZIONE_N01_CreoNuovaDirectory_pFunct.01
'//
'//
'// CONTROLLO_OGGETTI_TABELLA       *** FINE ***
'//============================================================================================================================//
'

'//DA SISTEMARE ????*****
'Public Function CONTROLLO_OGGETTO_TABELLE_DB_pFunct(par_DbNomeDatabase_s_s As String, _
                                             par_Name_tab_s As String, _
                                             par_Name_TABELLA_s As String, _
                                             par_iOperazione_i As Integer, _
                                             par_Indice1 As Integer, _
                                             par_Flag As Boolean) As String
                                             
Public Function CONTROLLO_OGGETTO_TABELLE_DB_pFunct() As String
        
On Error GoTo Err_CONTROLLO_OGGETTO_TABELLE_DB_pFunct



'//IMPOSTAZIONE PATH E CONTROLLO ESISTENZA DIRECTORY
'//------------------------------------------------------------------------------//
'//Codice---------->:CONTROLLO_OGGETTO_TABELLE_DB_pFunct.01
'//NOTE------------>:Ciclo oggetti db che rescontrollo l'esistenza della path definita dai salvataggi se non esiste _
'//NOTE------------>:controllo TUTTE LE TABELLE del db restituendo il _
                     nome delle stesse.
        '//Str1 = CONTROLLO_OGGETTO_TABELLE_DB_pFunct

'//
'//------------------------------------------------------------------------------//


    '//RESET
    Dim dbs  As Database
    Dim fldLoop As Field
    Dim relLoop As Relation
    Dim tdfloop As TableDef
 
 
 '//DATABASE ESTERNO
 Set dbsNorthwind = OpenDatabase("c:\Casa\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MDB\PROVA_CANCELLAZIONE_DB_ESTERNO\MENU_TB03_OGGETTI_DA_CANCELLARE.mdb")
 'Set dbsNorthwind = CurrentDb

    '//ATTIVO SOLO IL DB INTERNO
    If par_DbNomeDatabase_s_s = "" Then
        Set dbs = CurrentDb
    Else
        '//IL DB ESTERNO DA COSTRUIRE??
        MsgBox "PROCEDURA DI GESTIONE TABELLA database Access Esterno non ATTIVATA - !!"
        Stop
    End If
         
        '//RESET
        Count_i = 0
         
         
        '//QUALIFICAZIONE DB ED itero nel db
        '//-----------------------------------------------------------------------------------------------------//
            With dbs
                    
                    Debug.Print .TableDefs.Count & _
                        " TableDefs in database  " & DbNomeDatabase_s
                        
                    '//CICLO_OGGETTI
                    '//-----------------------------------------------------------------------------------------------------//
                    For intCiclo = 0 To .TableDefs.Count - 1
                        '//controlli
                        '//...........................................................................................//
                        
                            
                            DoEvents
                            
                            '//CONTATORE TABELLE
                            Count_i = Count_i + 1
                            
                            Debug.Print "                       CONTROLLO TABELLE "
                            Debug.Print "//===========================================================================//"
                            Debug.Print "NRO : " & Count_i & " )"
                            Debug.Print "  " & .TableDefs(intCiclo).Name
                            
                            '//Stampo il NOME della TABELLA
                            Debug.Print "IL NOME DELLA TABELLA"
                            Debug.Print "  " & .TableDefs(intCiclo).Name
                            
                            '//stampo GLI ATTRIBUTI tipo DELLA TABELLA
                            Debug.Print
                            Debug.Print "GLI ATTRIBUTI DELLA TABELLA"
                            Debug.Print "  " & .TableDefs(intCiclo).Attributes
                            
                            '//visualizzo se la tabella è connessa
                            Debug.Print
                            Debug.Print "CONTROLLO SE LA TABELLA E CONNESSA"
                            Debug.Print "  " & .TableDefs(intCiclo).SourceTableName
                            
                                
                                '// puo dare errore.
                                '.TableDefs(intCiclo).RefreshLink
                            
                            
                           Debug.Print "//===========================================================================//"
                           
                            Debug.Print
                        '//...........................................................................................//
                            
                            
                        '//controllo nome TABELLA passata con parametro
                        '//...........................................................................................//
                        If .TableDefs(intCiclo).Name = par_Name_TABELLA_s Then
                            
                            bBoolean1 = True    '// Flag True = trovato TABELLA
                            
                            '//recupero il contenuto della TABELLA
                            CONTROLLO_OGGETTO_TABELLE_DB_pFunct = .TableDefs(intCiclo).Name
                            MsgBox "OGGETTO INDIVIDUATO ---> " & CONTROLLO_OGGETTO_TABELLE_DB_pFunct
                            
                            
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

Exit_CONTROLLO_OGGETTO_TABELLE_DB_pFunct:
    
Exit Function

Err_CONTROLLO_OGGETTO_TABELLE_DB_pFunct:
    MsgBox "ERRORE FUNCTION PUBLIC    " & Err.Number & " - " & Err.Description, vbCritical, "CONTROLLO_OGGETTO_TABELLE_DB_pFunct"
    Resume Exit_CONTROLLO_OGGETTO_TABELLE_DB_pFunct

End Function

'// CONTROLLO_OGGETTI_TABELLA        *** fine ***
'//============================================================================================================================//





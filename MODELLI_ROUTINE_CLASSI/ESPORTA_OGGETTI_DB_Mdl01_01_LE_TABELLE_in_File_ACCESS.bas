Attribute VB_Name = "ESPORTA_OGGETTI_DB_Mdl01_01_LE_TABELLE_in_File_ACCESS"
Option Compare Database
Option Explicit

'//++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++//
'//       LE VARIABILI DI MODULO
'//
'//++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++//

'//===================================================================================================//
'//

'//LE VARIABILI DEL MODULO DEFINITE
'//::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::://

'//01
'//LE VARIABILI DATABASE
'//....................................................................//
    
    '//Database e Recordset di tipo DAO
    Dim DaoDB                                       As DAO.Database
    Dim DaoWks                                      As DAO.Workspace
    Dim DaoRs                                       As DAO.Recordset

    '//Database e recordset di tipo ADO
    Dim ADODB                                       As Database
    Dim AdodaoRs                                    As Recordset
    
    
    '//Stringa ssql e Path e per i Database ADO + DAO
    Dim sSql                                        As String                       '//STRINGA SQL
    Dim Path_s                                      As String                       '//la path del database


    '//Contatori
    Dim icount                                      As Integer
    Dim dbl_count                                   As Double
    Dim NRO_CMD_i                                   As Integer                      '//Numero del comando
    
    
    
    'Le Variabili del Metodo TransferDatabase
    '//-------------------------------------------------------------------------------
    '// Tipo di database traferibili (paradox, excell ecc..)
    Dim Tipo_Database_Dbase4_s                      As String                       '//...
    Dim Tipo_Database_Paradox3_s                    As String
    
    '//Macro Trasferisci database = parametri
    Dim Paths_s                                     As String                       '//Path database di destinazione
    Dim Tipo_Oggetto_s                              As String                       '//Tipo di oggetto (es. acTable = Tabella, acQuery = Query ecc..)
    Dim DbNomeDatabase_s                            As String                       '//Nome database (es. Database.mdb)
    '//Gli oggetti del database corrente trasferibili
    Dim OggettoOrigine_s                            As String                       '//nome oggetto di origine (es. della tabella, query ecc..)
    Dim OggettoDestinazione_s                       As String                       '//nome oggetto di destinazione ((es. della tabella, query ecc..)
    Dim SoloStruttura_b                             As Boolean                      '//solo struttura della tabella
  
    Dim SalvaIdConnessione_b                        As Boolean                      '//??
    '//-------------------------------------------------------------------------------
    

'//....................................................................//

'//02
'//LE VARIABILI DEI MESSAGGI
'//....................................................................//

    '//Variabili box messaggi                                                       '//MESSAGGI WINDOWS ACCESS
    Dim MsgBox_Title_s                              As String                       '//TITOLO DEL MESSAGGIO da emettere
    Dim MsgBox_s                                    As String                       '//messaggio del modulo da costruire
    
'//....................................................................//


'//03
'//LE VARIABILI DEI MESSAGGI DI ERRORE DELLE ROUTINE O FUNZIONI -
'//....................................................................//

    '//Variabili box messaggi                                                       '//MESSAGGI DI ERRORE DELLE ROUTINE O LE FUNZIONI
    Dim ROUT_NRO_i                                  As String                       '//Numero di routine/funzione
    Dim ROUT_ERR_MSG_s                              As String                       '//messaggio di Errore della routine o funzione
    Dim ROUT_NAME_MSG_s                             As String                       '//messaggio che identifica il nome della routine o funzione
    Dim ROUT_TIPO_MSG_s                             As String                       '//Tipo di messaggio
    Dim ROUT_TIPO_OPERAZ_MSG_s                      As String                       '//Messaggio sul Tipo di operazione eseguita
    

'//....................................................................//



'//04
'//LE VARIABILI GENERALI DELLA CLASSE FORM -
'//....................................................................//

 '//Variabili delle form e sotto form                                               '//LE VARIABILI PER LA GESTIONE DELLE FORM
    Dim FormName_s                                  As String                       '//Nome della form
    Dim SottFormName_s                              As String                       '//Nome della Sotto form collegata

'....................................................................



'//05
'//LE VARIABILI GENERICHE DI FUNZIONAMENTO
'//....................................................................//

    'Le variabili generiche
    Dim Byte1 As Byte
    Dim Dbl1 As Double
    Dim int1 As Integer
    Dim Long1 As Long
    Dim Bool1 As Boolean
    Dim Str1 As String
    Dim vV1 As Variant
    
'//....................................................................//




'//LE VARIABILI DEL MODULO DEFINITE *** FINE ***
'//::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::://


'//++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++//
'//       LE VARIABILI DI MODULO  *** fine ***
'//
'//++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++//



Private Sub p()

'//Imposto le variabili
'Paths_s = "c:\CASA\GE_CASA\GE_MARINO\GESTIONE_SPESE\GESTIONE_PROCEDURE\GE_CASA_MDB\GE_CASA_TB90_SALVATAGGI_ARCHIVI.mdb"
'ESPORTA_Oggetti_Dababase_PFunct "c:\CASA\GE_CASA\GE_MARINO\GESTIONE_SPESE\GESTIONE_PROCEDURE\GE_CASA_MDB\GE_CASA_TB90_SALVATAGGI_ARCHIVI.mdb", _
'                "Oggetto -> acTable = Tabella", _
'                "GE_CASA_TB01_MASTRO_TMP", _
 '               "GE_CASA_TB01_MASTRO_TMP", _
 '               False
End Sub



Private Sub ESPORTA_OGGETTI_COLLETTIVI()


ESPORTA_Oggetti_Dababase_PFunct _
                "c:\CASA\GE_CASA\GE_MARINO\GESTIONE_SPESE\GESTIONE_PROCEDURE\GE_CASA_MDB\", _
                "GE_CASA_TB90_SALVATAGGI_ARCHIVI.mdb", _
                "Oggetto -> acTable = Tabella", _
                "GE_CASA_TB01_MASTRO_TMP", _
                "GE_CASA_TB01_MASTRO_TMP", _
                False
End Sub




'------------------------------------------------------------
' ESPORTA_Oggetti_Dababase_PFunct
'
'------------------------------------------------------------
Function ESPORTA_Oggetti_Dababase_PFunct(par_Paths_s As String, _
                                         par_DbNomeDatabase_s As String, _
                                         par_TipoOggetto_s As String, _
                                         par_OggettoOrigine_s As String, _
                                         par_OggettoDestinazione_s As String, _
                                         par_SoloStruttura_b As Boolean)
                         
On Error GoTo ESPORTA_Oggetti_Dababase_PFunct_Err

    '//Esempio completo : _
    'DoCmd.TransferDatabase acExport, "Microsoft Access", "c:\CASA\GE_CASA\GE_MARINO\GESTIONE_SPESE\GESTIONE_PROCEDURE\GE_CASA_MDB\GE_CASA_TB90_SALVATAGGI_ARCHIVI.mdb", acTable, "GE_CASA_TB01_MASTRO_TMP", "GE_CASA_TB01_MASTRO_TMP", False
    
    '//ESPORTO OGGETTI TABELLA CON IL PARAMETRO par_TipoOggetto_s = "Tabella"
    '//----------------------------------------------------------------------------------------//
    If par_TipoOggetto_s = "Oggetto -> acTable = Tabella" Then
        
        '// METODO TRASFERIMENTO DATABASE
        '// ................................................................. _
        NOTE : La funzione trasfer database richiede, 5 parametri: _
        01) il TIPO DI TRAFERIMENTO _
        rappresentato dalla variabile acExport, NON IDENTIFICATA DA NESSUN PARAMETRO; _
        02) il TIPO DI DATABASE (es. Microsoft Access, dBase 5.0 ecc) NON IDENTIFICATA DA NESSUN PARAMETRO; _
        03) il NOME DEL DATABASE con il percorso completo ottenuto con i seguenti parametri : par_Paths_s par_DbNomeDatabase_s; _
        04) il tipo di OGGETTO (Tabella, Query ecc.) rappresentata dalla variabile par_TipoOggetto_s; _
        05) l'OGGETTO DI ORIGINE ossia il nome della tabella, query ecc. PRELEVATA da salvare e rappresentata da par_OggettoOrigine_s; _
        06) l'OGGETTO DI DESTINAZIONE ossia il nome della tabella, query ecc. che sarà salvata _
            nel database di destinazione rappresentata da par_OggettoDestinazione_s; _
        07) SOLOSTRUTTURA l'indicazione boolean = True Trasferisco solo la struttura della tabella, _
            False = trasferisco i dati e la struttura della tabella.

            DoCmd.TransferDatabase acExport, "Microsoft Access", par_Paths_s & par_DbNomeDatabase_s, acTable, _
            par_OggettoOrigine_s, par_OggettoDestinazione_s, par_SoloStruttura_b
            
            
           '//SOSPESA = dava troppi messaggi uno per tabella e rallentava, sosituito _
              con la query di accodamento di seguito riportata che salva i messaggi _
              nella tabella tmp che saranno visualizzatti tutti insieme in un report.
           ' MsgBox "ESPORTAZIONE OGGETTO DATABASE di Tipo-> " & Chr$(13) _
           ' & " 01) Tipo di oggetto " & par_TipoOggetto_s & Chr$(13) _
           ' & " 02) Oggetto Origine ->" & par_OggettoOrigine_s & Chr$(13) _
           ' & " 03) Oggetto Destinazione -> " & par_OggettoDestinazione_s & Chr$(13) _
           ' & " 04) Path di destinazione -> " & par_Paths_s & Chr$(13) _
           ' & " 05) Solo Struttura  ->" & par_SoloStruttura_b & Chr$(13)
           
           '//SALVATAGGIO DEL MESSAGGIO _
              nella tabella tmp da visualizzare _
              alla fine del ciclo con un Report. Il modello della _
              query è il -> GE_CASA_QryTb85_71_INSERT_accodoMessaggi_in_TMP
           '//++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++//
                    '//reset
                    Str1 = ""
                       
                    '//struttura messaggio
                    Str1 = Str1 & "ESPORTAZIONE OGGETTO DATABASE di Tipo-> " & Chr$(13) _
                     & " 01) Tipo di oggetto " & par_TipoOggetto_s & Chr$(13) _
                     & " 02) Oggetto Origine ->" & par_OggettoOrigine_s & Chr$(13) _
                     & " 03) Oggetto Destinazione -> " & par_OggettoDestinazione_s & Chr$(13) _
                     & " 04) Path di destinazione -> " & par_Paths_s & Chr$(13) _
                     & " 05) Solo Struttura  ->" & par_SoloStruttura_b & Chr$(13)
                    
                    
                    sSql = ""
                    sSql = sSql & "INSERT INTO GE_CASA_Tb85_MESSAGGI_TMP ( MESSAGGIO_m, DATAINS, TIMEOPER ) "
                    sSql = sSql & "SELECT '" & Str1 & "' AS MSG," & Date & " AS DATA, '" & Time() & "' AS [TIME];"
                    
                    '//controllo ed esecuzione
                    Debug.Print
                    Debug.Print " controlla il messaggio costruito str1"
                    Debug.Print "-------------------------------------------------------"
                    Debug.Print sSql
                    Debug.Print
                    
                    CurrentDb.Execute sSql
                    
              '//++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++//
        '// .................................................................
    End If
    '//----------------------------------------------------------------------------------------//
    'DoCmd.TransferDatabase _
    Esporta = acExport, _
    TipoDb = "Microsoft Access", _
    '//NOME DB CON IL PERCORSO _
    paths_s ="c:\CASA\GE_CASA\GE_MARINO\GESTIONE_SPESE\GESTIONE_PROCEDURE\GE_CASA_MDB\GE_CASA_TB90_SALVATAGGI_ARCHIVI.mdb", _
    TipoOggetto_s = acTable, _
    OggettoOrigine_s "GE_CASA_TB01_MASTRO_TMP", _
    OggettoDestinazione_s "GE_CASA_TB01_MASTRO_TMP", _
    SoloStruttura_b = False


ESPORTA_Oggetti_Dababase_PFunct_Exit:
    Exit Function

ESPORTA_Oggetti_Dababase_PFunct_Err:
    MsgBox Error$
    Resume ESPORTA_Oggetti_Dababase_PFunct_Exit

End Function



Attribute VB_Name = "UTIL_Nro98_N01_LE_VARIABILI_DI_MODULO"
Option Compare Database
Option Explicit

'//++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++//
'//       LE VARIABILI DI MODULO
'//
'//++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++//


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
    Dim iCount                                      As Integer
    Dim dbl_count                                   As Double
    Dim NRO_CMD_i                                   As Integer                      '//Numero del comando
    
    'Le Variabili del Metodo TransferDatabase
    
    Dim Tipo_Trasferimento_s                        As String                       '//definire funzioni??...
    Dim Tipo_Database_Dbase4_s                      As String                       '//...
    Dim Tipo_Database_Paradox3_s                    As String
    Dim SoloStruttura_b                             As Boolean
    Dim NomeDabataseDb_s                            As String
    Dim DbOrigine_s                                 As String
    Dim DbDestinazione_s                            As String
    Dim SalvaIdConnessione_b                        As Boolean

'//....................................................................//


'//LE VARIABILI DEI MESSAGGI
'//....................................................................//

    '//Variabili box messaggi                                                       '//MESSAGGI WINDOWS ACCESS
    Dim MsgBox_Title_s                              As String                       '//TITOLO DEL MESSAGGIO da emettere
    Dim MsgBox_s                                    As String                       '//messaggio del modulo da costruire
    
'//....................................................................//



'//LE VARIABILI DEI MESSAGGI DI ERRORE DELLE ROUTINE O FUNZIONI -
'//....................................................................//

    '//Variabili box messaggi                                                       '//MESSAGGI DI ERRORE DELLE ROUTINE O LE FUNZIONI
    Dim ROUT_NRO_i                                  As String                       '//Numero di routine/funzione
    Dim ROUT_ERR_MSG_s                              As String                       '//messaggio di Errore della routine o funzione
    Dim ROUT_NAME_MSG_s                             As String                       '//messaggio che identifica il nome della routine o funzione
    
'//....................................................................//




'//LE VARIABILI GENERALI DELLA CLASSE FORM -
'//....................................................................//

 '//Variabili delle form e sotto form                                               '//LE VARIABILI PER LA GESTIONE DELLE FORM
    Dim FormName_s                                  As String                       '//Nome della form
    Dim SottFormName_s                              As String                       '//Nome della Sotto form collegata

'....................................................................




'//LE VARIABILI GENERICHE DI FUNZIONAMENTO
'//....................................................................//

    'Le variabili generiche
    Dim Byte1       As Byte
    Dim Vv1         As Variant
    Dim Dbl1        As Double
    Dim Int1        As Integer
    Dim Long1       As Long
    Dim Bool1       As Boolean
    Dim Date1       As Date
    Dim String1     As String

    

'//....................................................................//

'//++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++//
'//       LE VARIABILI DI MODULO  *** fine ***
'//
'//++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++//



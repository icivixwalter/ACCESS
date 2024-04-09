


Option Compare Database
Option Explicit

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'       LE VARIABILI DI MODULO
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++


'....................................................................
'Dim di FORM_X.01.01 : Variabili di File
Dim DaoDB As DAO.Database
Dim DaoWks As DAO.Workspace
Dim DaoRs As DAO.Recordset

Dim ADODB As Database
Dim AdodaoRs As Recordset

Dim sSql As String


'Variabili generiche
'................................................
Dim Byte1 As Byte
Dim Str1 As String
Dim Int1 As Integer
Dim Lng1 As Long
Dim Dbl1 As Double
Dim Flag1 As Boolean
Dim Vv1 as Variant

'................................................
'DIM Variabili  Routine sub
Dim iNRO_CMD As Integer
Dim sMESS_CMD As String
Dim sMESS_TIPO As String
Dim sFile As String

'....................................................................

'Le Variabili del Metodo TransferDatabase

Dim sxTipo_Trasferimento                    As String
Dim sxTipo_Database_Dbase4                  As String
Dim sxTipo_Database_Dbase5                  As String
Dim sxTipo_Database_Paradox3                As String
Dim blnxSoloStruttura                       As Boolean
Dim sxNomeDabase                            As String
Dim sxOrigine                               As String
Dim sxDestinazione                          As String
Dim sxSalvaIdConnessione                    As Boolean


'....................................................................
'Dim di FORM_X.01.03  : Variabili di Generali CLASSE FORM
Dim sFormName   As String
Dim sSottFormName   As String

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'       LE VARIABILI DI MODULO          *** FINE ***
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++




'//******************************************************************************************************************************************************//
'//                                     EVENTI_FORM
'//******************************************************************************************************************************************************//

'//ROUTINE-------->:Form_Load()
'//========================================================================================================================================//
Private Sub Form_Load()

  '//I comandi della form
    Dim CMD_NRO_i As Integer                    '//Numero del comando
    Dim CMD_ERR_MSG_s As String                 '//Messaggio del Comando
    Dim CMD_TIPO_s As String                    '//TIPO Comando (ES. Evento Load, Update ecc.)


On Error GoTo Err_Form_Load


'//RESET VARIABILI
CMD_NRO_i = 1
CMD_ERR_MSG_s = "ERRORE NELL'EVENTO FORM LOAD -> FORM MASTRO_COLL_XLS"
CMD_TIPO_s = "EVENTO FORM LOAD -> FORM MASTRO_COLL_XLS"



    '//CARICAMENTO OGGETTI
    '//----------------------------------------------------------------------------------------------
    '//GE_CASA_Frm01_S01_S01_MASTRO_Coll_Xls.01.02
    '//NOTA:  Al caricamento della form reimposto i valori degli oggetti
    '//       della form corrente.
    
           
        '//REIMPOSTO_LA_CASELLA_COMBINATA
        '//________________________________________________________________________________________
        
             '//Imposto il valore della casella combinata
            Me.Cmd_01_TXT.RowSourceType = "Elenco valori"
            
           ' Str1 = "01_CONTROLLA;02_CANCELLA DATI TMP;03_ACCODA DATI IN ARCHIVIO;04_UTILE 01"
            Me.Cmd_01_TXT.RowSource = "01_CONTROLLA;" & _
                                      "02_CANCELLA DATI TMP;" & _
                                      "03_ACCODA DATI IN ARCHIVIO;" & _
                                      "04_UTILE 01;" & _
                                      "05_UTILE 02"
            
            '//Ipostazione di base appena visibile
            Me.Cmd_01_TXT = "COMANDI INTERNI FORM"
            
            
        '//REIMPOSTO_LA_CASELLA_COMBINATA *** FINE ***
        '//________________________________________________________________________________________
  
    
    '//CARICAMENTO OGGETTI *** FINE ***
    '//----------------------------------------------------------------------------------------------
  



'//::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'//GE_CASA_Frm01_S01_S01_MASTRO_Coll_Xls.03.01)    GESTIONE ERRORI E USCITA ROUTINE
'//NOTA:

Exit_Form_Load:
    Exit Sub

Err_Form_Load:

        Vv1 = MsgBox("ERRORE NRO : " & Err.Number & " - TIPO DI ERRORE ==>: " & Err.Description & Chr$(13) _
        & " - ROUTINE SUB: " & CMD_NRO_i & " - " & CMD_TIPO_s & " - " & CMD_ERR_MSG_s, vbCritical)
        
        Resume Exit_Form_Load

End Sub
'//ROUTINE-------->:Form_Load() *** FINE ***
'//========================================================================================================================================//


'//******************************************************************************************************************************************************//
'//                                     EVENTI_FORM *** FINE ***
'//******************************************************************************************************************************************************//



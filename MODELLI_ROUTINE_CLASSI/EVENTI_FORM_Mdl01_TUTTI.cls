VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_GE_CASA_Frm01_S01_S01_MASTRO_Coll_Xls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


'//++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++//
'//       LE VARIABILI DI MODULO

'//LE VARIABILI DATABASE
'//....................................................................//
    Dim DaoDB As DAO.Database
    Dim DaoWks As DAO.Workspace
    Dim DaoRs As DAO.Recordset

    Dim ADODB As Database
    Dim AdodaoRs As Recordset
    Dim sSql As String                          '//STRINGA SQL
    Dim Path_s As String                        '//la path


    '//Contatori
    Dim iCount As Integer
    Dim NRO_CMD_i As Integer                    '//Numero del comando
    Dim dbl_count As Double

  
    'Le variabili generiche
    Dim Vv1 As Variant
    Dim Dbl1 As Double
    Dim Int1 As Integer
    Dim Long1 As Long
    Dim Str1 As Long

'....................................................................

'//++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++//




Private Sub Cmd_01_TXT_AfterUpdate()

'Stop

End Sub

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
            Me.Cmd_01_TXT = "COMANDI INTERNI FORM"
    
        '//REIMPOSTO_LA_CASELLA_COMBINATA *** FINE ***
        '//________________________________________________________________________________________
  
    
    '//CARICAMENTO OGGETTI *** FINE ***
    '//----------------------------------------------------------------------------------------------
  




'--------------------------------------------------------------------------------------------------
'GE_CASA_Frm01_S01_S01_MASTRO_Coll_Xls.03.01)    GESTIONE ERRORI E USCITA ROUTINE
'
'NOTA:

Exit_Form_Load:
    Exit Sub

Err_Form_Load:

        Vv1 = MsgBox("ERRORE NRO : " & Err.Number & " - TIPO DI ERRORE : " & Err.Description & Chr$(13) _
        & " - ROUTINE SUB: " & CMD_NRO_i & " - " & CMD_TIPO_s & " - " & CMD_ERR_MSG_s, vbCritical)
        
        Resume Exit_Form_Load

End Sub

'//ROUTINE-------->:Form_Load() *** FINE ***
'//========================================================================================================================================//


'//******************************************************************************************************************************************************//
'//                                     EVENTI_FORM *** FINE ***
'//******************************************************************************************************************************************************//


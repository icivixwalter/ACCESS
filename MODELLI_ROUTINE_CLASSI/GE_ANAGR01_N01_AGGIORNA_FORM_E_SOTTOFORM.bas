Attribute VB_Name = "GE_ANAGR01_N01_AGGIORNA_FORM_E_SOTTOFORM"
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>:>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'Option

Option Compare Text
Option Explicit

'Variabili di database
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>:>>>>>>>>>>>>>>>>>>>>>>>>>>>>

'DAO
Dim DaoDB As DAO.Database
Dim DaoWks As DAO.Workspace
Dim DaoRs As DAO.Recordset
Dim DaoRs_2 As DAO.Recordset

'ADO
Dim ADODB As Database
Dim AdodaoRs As Recordset

'Contatori
Dim iCount As Integer
Dim dbl_count As Double

'Le variabili generiche
Dim sSql As String                                          ' Stringa di estrazione

'Variabili generali
Dim Str1 As String
Dim Int1 As Integer
Dim Lng1 As Long
Dim Dbl1 As Double
Dim Bln1 As Boolean
Dim Vv1 As Variant


'Gestione parametri condominio
Dim sxCODCOND As String
Dim ixANNOESERC As Integer
Dim sxDATAINIZIO As String
Dim sxDATAFINE As String
Dim sxGESTIONE As String


'Larghezza e numero di colonna
Dim sLarg_Col As String
Dim iNum_Col As Integer


'Variabili DELLA FORM
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>:>>>>>>>>>>>>>>>>>>>>>>>>>>>>

Dim sxFrm_CHIAMANTE_GENERALE As String                  '... Form CHIAMANTE generale
Dim sxFrm_CHIAMANTE_GENERALE_CORRENTE As String         '... Form CHIAMANTE generale OLD
Dim sxFrm_CHIAMANTE_GENERALE_PRECEDENTE As String       '... Form CHIAMANTE precedente







'AGGIORNA RECORD TMP
'Aggiorno il record temporaneo che descrive il soggetto in esame dalla procedura.
Public Sub Aggiorna_Form_e_SottoForm()

On Error GoTo Err_Aggiorna_Form_e_SottoForm


'Aggiorno i campi testo
'-----------------------------------------------------------------------------
  
        
        
        'Aggiorno FORM QUADRO FABBRICATI
        'Application.Forms("GEANG_Frm01_M01_ANAGRAFICA")!SottForm_QUADRO_FABB.Requery
        
        'Aggiorno FORM QUADRO TERRENI
        'Application.Forms("GEANG_Frm01_M01_ANAGRAFICA")!SottForm_QUADRO_TERR.Requery
        
  
  
  
          
  
        

'fine Aggiorno i campi testo
'-----------------------------------------------------------------------------



        
                
'USCITA E GESTIONE ERRORI
'..................................................................................................
                
                
Exit_Aggiorna_Form_e_SottoForm:
    Exit Sub

Err_Aggiorna_Form_e_SottoForm:
    MsgBox Err.Description
    Resume Exit_Aggiorna_Form_e_SottoForm

End Sub




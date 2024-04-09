Attribute VB_Name = "GE_ANAGR01_N02_VISUALIZZA_DATI_NELLA_FORM"

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







'VISUALIZZO CAMPI FORM
'Associo i dati dei parametri ai campi della form

Public Sub Visualizza_Campi_Form_Comunicazioni(par_sxCodice As String, _
                                                par_sxRagione_Sociale As String, _
                                                par_sxCodice_Fiscale As String, _
                                                par_sxComune_residenza As String, _
                                                par_sxIndirizzo_residenza As String, _
                                                par_sxCivico As String, _
                                                par_sxCODPRAT As String, _
                                                par_sxNRO_PRAT As String, _
                                                par_sxTRIM As String, _
                                                par_sxFASC As String, _
                                                par_sxFALD As String)
                                                



On Error GoTo Err_Visualizza_Campi_Form_Comunicazioni



            'Aggiorno i campi testo
            '-----------------------------------------------------------------------------
              
                    
                    Application.Forms("GEANG_Frm01_M01_ANAGRAFICA")!TXT_01 = par_sxCodice
                    Application.Forms("GEANG_Frm01_M01_ANAGRAFICA")!TXT_02 = par_sxCodice_Fiscale
                    Application.Forms("GEANG_Frm01_M01_ANAGRAFICA")!TXT_03 = par_sxRagione_Sociale
                    Application.Forms("GEANG_Frm01_M01_ANAGRAFICA")!TXT_04 = par_sxComune_residenza
                    Application.Forms("GEANG_Frm01_M01_ANAGRAFICA")!TXT_05 = par_sxIndirizzo_residenza
                   
                   'Application.Forms("GEANG_Frm01_M01_ANAGRAFICA")!TXT_06 = Me.[Civico]
                    
                    Application.Forms("GEANG_Frm01_M01_ANAGRAFICA")!TXT_NRO_PRAT = par_sxNRO_PRAT
                    Application.Forms("GEANG_Frm01_M01_ANAGRAFICA")!Txt_FASC = par_sxFASC
                    Application.Forms("GEANG_Frm01_M01_ANAGRAFICA")!Txt_Fald = par_sxFALD
                    'Str1 = "- Nro " & par_vxNRO_COMU_Coll & "  - " & par_vxFASC & "/" & par_vxFALD
                    Application.Forms("GEANG_Frm01_M01_ANAGRAFICA")!Txt_CODPRAT = par_sxCODPRAT
            
            
            'fine Aggiorno i campi testo
            '-----------------------------------------------------------------------------
                
'USCITA E GESTIONE ERRORI
'..................................................................................................
                
                
Exit_Visualizza_Campi_Form_Comunicazioni:
    Exit Sub

Err_Visualizza_Campi_Form_Comunicazioni:
    MsgBox Err.Description
    Resume Exit_Visualizza_Campi_Form_Comunicazioni

End Sub





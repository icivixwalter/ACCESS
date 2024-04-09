Attribute VB_Name = "FUNZIONI_Mdl90_MODELLO_FUNZIONE"
Option Compare Database
'//Variabili di database e GENERALI
'//>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>:>>>>>>>>>>>>>>>>>>>>>>>>>>>>

'DAO
Dim DaoDB As DAO.Database
Dim DaoWks As DAO.Workspace
Dim DaoRs As DAO.Recordset
Dim DaoRs_Lett As DAO.Recordset

'ADO
Dim ADODB As Database
Dim AdodaoRs As Recordset

'Contatori
Dim iCount As Integer
Dim iTOTcount As Integer
Dim dbl_count As Double

'Le variabili generiche
Dim sSql As String                                          ' Stringa di estrazione
Dim sSql_Lett As String                                     ' Stringa di estrazione

'Variabili generali
Dim Str1 As String
Dim Int1 As Integer
Dim Int2 As Integer
Dim Int3 As Integer
Dim Lng1 As Long
Dim Dbl1 As Double
Dim Bln1 As Boolean
Dim Vv1 As Variant
Dim obj1 As Object

'Gestione parametri comandi
Dim sxParamCmd_3 As String


'Larghezza e numero di colonna
Dim sLarg_Col As String
Dim iNum_Col As Integer


'//ERRORI PROCEDURA_FUNCTION O ROUTINE
Dim ProceduraMessaggioErrore_s As String
Dim ProceduraAttivaEseguita_s As String

'//Variabili di database e GENERALI     *** FINE ***
'//>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>:>>>>>>>>>>>>>>>>>>>>>>>>>>>>


'***********************************************************************************************************************
' Modello_Funzione_Pfunct()
'
'***********************************************************************************************************************

'***********************************************************************************************************************
' Modello_FUNZIONE_N01_PFucnt()
'
'***********************************************************************************************************************

'//DENOMINAZIONE---------> Modello_FUNZIONE_N01_PFucnt
'//=================================================================================================================//
'//ATTIVITA--------------> GESTIONE DEI FALDONI
'//NOTE------------------> ....
'//PARAMETRI-------------> Nessuno
'//VALORE_DI_RITORNO-----> Nulla
'//CODICE----------------> Modello_FUNZIONE_N01_PFucnt.01.00
'//=================================================================================================================//
Function Modello_FUNZIONE_N01_PFucnt()

On Error GoTo Modello_FUNZIONE_N01_PFucnt_Err

    '//APRO_FORM_FALDONI = Chiamo procedura
    '//---------------------------------------------------------------------------------------//
    '//CODICE               -> Modello_FUNZIONE_N01_PFucnt.01.01

            '//Call Modello_FUNZIONE_N01_PFucnt
    '//---------------------------------------------------------------------------------------//
    
   
    
    '//APRO_FORM_FALDONI = Apertura form + aggiornamento nome faldoni
    '//---------------------------------------------------------------------------------------//
    '//CODICE               -> Modello_FUNZIONE_N01_PFucnt.01.02
    
                    '//RESET
                ProceduraMessaggioErrore_s = ""
                ProceduraAttivaEseguita_s = ""
   

                '//IMPOSTO LE VARIABILI
                ProceduraMessaggioErrore_s = "GESTIONE DEI FALDONI"
                ProceduraAttivaEseguita_s = "Modello_FUNZIONE_N01_PFucnt"

        
            
        '//APRO_FORM_FALDONI
        DoCmd.OpenForm "LLPP_Frm_MF05_01_FALDONI"
        
        '//*** SOSPESA PER ERRORE QUERY ****
        '//AGGIORNO_DENOMINAZIONE_FALDONI = nella tabella gestione *** sistemare query
        'DoCmd.OpenQuery "LLPP_ATTI_Qry01-06_Gestione_AGGIORNA_DenomFALDONE"

    '//---------------------------------------------------------------------------------------//

Modello_FUNZIONE_N01_PFucnt_Err:
    Exit Function
    MsgBox Error$
    
 
    MsgBox Err.Description & " " & ProceduraMessaggioErrore_s & " - " & ProceduraAttivaEseguita_s & " --> errore Error$ -> : " & Error$
    Debug.Print ProceduraMessaggioErrore_s
    Debug.Print ProceduraAttivaEseguita_s
    Debug.Assert "BLOCCO PROCEDURA -> " & ProceduraAttivaEseguita_s
  
    Stop
    Resume Modello_FUNZIONE_N01_PFucnt_Err
    
End Function
'//'//DENOMINAZIONE---------> Modello_FUNZIONE_N01_PFucnt *** FINE ***
'//=================================================================================================================//



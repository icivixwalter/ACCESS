Attribute VB_Name = "UTIL_DATABASE_Mdl10_OpenDatabase_01_(AproDb_DAO_DB_NON_FUNZIONA)"
Option Compare Text
Option Explicit

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

'//Variabili speciali
Dim ixNRO_ESTR As Integer
Dim lngxID_TERN  As Long


'Gestione parametri comandi
Dim sxParamCmd_3 As String


'Larghezza e numero di colonna
Dim sLarg_Col As String
Dim iNum_Col As Integer


'//ERRORI PROCEDURA_FUNCTION O ROUTINE
Dim ProceduraMessaggioErrore_s As String
Dim ProceduraAttivaEseguita_s As String




'**************************************************************************************************
'   Public Function di APERTURA DI UN DATABASE DAO
'
'
'**************************************************************************************************




Public Function pf_OpenDatabase_DAO(par_dbs, par_rs)
'-----------------------------------
'   Dim OGGETTI, Dabase recordset e le variabili

Dim wrkJet As Workspace
Dim wrkODBC As Workspace

Dim obDBS   As Object
Dim obRS    As Object
Dim dbsDAO As DAO.Database
Dim rsDAO   As DAO.Recordset
Dim Vv1 As Variant
Dim sV2 As String
Dim iV3 As Integer
Dim lV4 As Long
Dim crV5 As Currency

            On Error GoTo Err_pf_OpenDatabase_DAO
    
    
    
        '------------------------------------------------------------------------------------------
        '1.1    APERTURA DEL DABASE
        
            ' assegno la stringa di scelta passata con parametro
                Set obDBS = par_dbs
                
                Set dbsDAO = obDBS
                    
                '..............................................................
                '   OPEN DATABASE DAO
                
                    '   Apro il database con il metodo Dao
                    
                       Set dbsDAO = OpenDatabase(par_dbs)


'----------------------------------------------------------------------------------------------------
'                       FINE FUNCTION E GESTIONE ERRORI

Exit_pf_OpenDatabase_DAO:
Exit Function

Err_pf_OpenDatabase_DAO:
    MsgBox "ERRORE FUNCTION PUBLIC    " & Err.Number & " - " & Err.Description, vbCritical, "pf_OpenDatabase_DAO"
    dbsDAO.Close
    Set dbsDAO = Nothing
    Resume Exit_pf_OpenDatabase_DAO

    
    
End Function


Private Sub APRO_RECORDSET()

     
'//ITERO NEL RS
'//---------------------------------------------------------------------------------------//
 
     '//Inserire tabella o stringa ssql
     sSql = "LLPP_ATTI_Tb01_Gestione"
     'RS ADO
     'Solo l'anno indicato nella variabile
     Set AdodaoRs = CurrentDb.OpenRecordset(sSql)
    
    '//WITH AdodaoRs
    '//............................................................................//
     With AdodaoRs
     
         'Controllo Rs
         If AdodaoRs.EOF = False And AdodaoRs.BOF = False Then
        
             'RESET
             Int1 = 0
             iCount = 0
             iTOTcount = 0
             ixNRO_ESTR = 0
        
             .MoveFirst
        
             'ITERAZIONE RS DI LETTURA
             '----------------------------------------------------------------------
                 While Not AdodaoRs.EOF
        
                 'Controllo windows
                 DoEvents
                 sSql = ""
        
                     'Vado avanti finche la base è uguale
                     '------------------------------------------------
                         'Imposto i codici Base
                         lngxID_TERN = .Fields("ID_TERN")
        
                         Debug.Print
                         Debug.Print "--------------------------------------------"
                         Debug.Print "CONTROLLO " & lngxID_TERN
                         Debug.Print "--------------------------------------------"
                     '............................................
        
                     AdodaoRs.MoveNext
        
                 Wend
             'ITERAZIONE RS DI LETTURA       *** fine ***
             '----------------------------------------------------------------------
             
         End If '//If AdodaoRs.EOF = False And AdodaoRs.BOF = False Then
        
    
     End With
    '//WITH AdodaoRs     *** FINE ***
    '//............................................................................//
     

'//ITERO NEL RS     *** fine ***
'//---------------------------------------------------------------------------------------//
         

End Sub

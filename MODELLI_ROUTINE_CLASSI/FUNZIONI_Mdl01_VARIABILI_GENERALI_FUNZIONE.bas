Attribute VB_Name = "FUNZIONI_Mdl01_VARIABILI_GENERALI_FUNZIONE"
'//++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++//
'//       LE VARIABILI DI MODULO

'//LE VARIABILI DATABASE
'//....................................................................//
    
    '//DATABASE E RECORDSET
    '//-----------------------------------------------------//
    Dim DaoDB As DAO.Database
    Dim DaoWks As DAO.Workspace
    Dim DaoRs As DAO.Recordset

    Dim ADODB As Database
    Dim AdodaoRs As Recordset
    Dim sSql As String                          '//STRINGA SQL
    Dim Path_s As String                        '//la path
    '//-----------------------------------------------------//

'//*** FINE ***
'//LE VARIABILI DATABASE
'//....................................................................//


'//LE VARIABILI GENERICHE E CONTATORI
'//....................................................................//

    '//Contatori
    Dim iCount As Integer
    Dim iTOTcount As Integer
    Dim dbl_count As Double

    'Le variabili generiche
    Dim Str1 As String
    Dim Int1 As Integer
    Dim Int2 As Integer
    Dim Int3 As Integer
    Dim Lng1 As Long
    Dim Dbl1 As Double
    Dim Bln1 As Boolean
    Dim Vv1 As Variant
    Dim obj1 As Object
    
'//*** FINE ***
'//LE VARIABILI GENERICHE E CONTATORI
'//....................................................................//



'//LE VARIABILI DELLE PROCEDURE FUNCTIONE E SUB + COLONNE E RIGHE DATI
'//....................................................................//

    
    '//ERRORI PROCEDURA_FUNCTION O ROUTINE
    Dim ProceduraMessaggioErrore_s As String
    Dim ProceduraAttivaEseguita_s As String



    'Larghezza e numero di colonna
    Dim Larg_Col_s As String
    Dim Num_Col_i As Integer

'//*** FINE ***
'//LE VARIABILI DELLE PROCEDURE FUNCTIONE E SUB + COLONNE E RIGHE DATI
'//....................................................................//
    



'//++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++//





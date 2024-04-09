Attribute VB_Name = "LLPP_IMPEGNI_Mdl01_02_LLPP_ATTI"
Option Compare Database


'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'DATABASE + RECORDSET
'------------------------------------------------
' Dim Variabili di Modulo
Dim DaoDB As DAO.Database
Dim DaoRs As DAO.Recordset

Dim DaoRs_Ordina_dati As DAO.Recordset
Dim DaoRs_Duplicato As DAO.Recordset

'Contatori
Dim iCount As Integer
Dim dbl_count As Double

'Le variabili generiche
Dim sSql As String
Dim sSql_TOT As String
Dim Vv1 As Variant
Dim ssSTR1 As String
Dim sStr2 As String

Dim iInt1 As Integer
Dim lngLong1 As Long

Private Sub CHIAMA_AGGIORNA_FILE_ATTO_pfunct()
AGGIORNA_FILE_ATTO_pfunct "Folium", 1000, "03/03/2017", 2017

End Sub


'//AGGIORNA_FILE_ATTO_pfunct
'//===========================================================:===================================:================================
'//AGGIORNA_ATTI_Mdl_N001_001_Function.001.01__________:Funzione 00 CALCOLA                                :Calcola
'//...........................................:.....................................................................................
'//NOTA PROCEDURA: La Funzione Aggiona il campo FILE_ATTO ricostruendo il nome del file con i parametri passati.

Public Function AGGIORNA_FILE_ATTO_pfunct(par_TipoAtto_s As String, _
                                          par_NroAtto_lng As Long, _
                                          par_DataAtto_d As Date, _
                                          par_AnnoAtto_i As Integer) As String

'//.......................................................
'// DIM VARIABILI
'Dim sScelta_db As String

'//AGGIORNA_FILE_ATTO
'//--------------------------------------------------------------------------//
'//NOTA   : Aggiorna il file atto costruendo il nome del file _
            con i parametri passati ottengo il nome dai seguenti dati _
            FILE_ATTO=TipoAtto+NroAtto+Annoatto a condizione che _
            il NroAtto >0 and DataAtto >"", altrimenti la funzione restituisce _
            il valore nullo.
    '// TipoAtto_s = "" _
    NroAtto_lng = 0 _
    DataAtto_d = "" _
    AnnoAtto_i = 0 _
    AGGIORNA_FILE_ATTO_pfunct TipoAtto_s, NroAtto_lng, DataAtto_d, AnnoAtto_i

'//--------------------------------------------------------------------------//
    
        On Error GoTo Err_AGGIORNA_FILE_ATTO_pfunct
    
    
        '//AGGIORNA_FILE_ATTO
        '//-------------------------------------------:-----------------------------------------------
        '//AGGIORNA_ATTI_Mdl_N001_001_Function.001.02__________:
        '//NOTA   : Aggiorna il file atto costruendo il nome del file con i parametri passati.
        
                                                 
                        '//AGGIORNAMENTO CAMPO NRO ATTO -
                        '//..............................................................
                        '//Aggiorna il nro att and data > 0 allora _
                        creo il nome del file con fileAtto+NroAtto+Annoatto, solo se _
                        nro atto e data sono > 0
                        
                            If par_NroAtto_lng > 0 And par_DataAtto_d > 0 Then
                                Str1 = ""
                                Str1 = par_TipoAtto_s & "_" & par_NroAtto_lng & "_" & par_AnnoAtto_i
                                '//controllo e salvataggio
                                Debug.Print "-------------------------------------------------------"
                                Debug.Print "file creato"
                                Debug.Print Str1
                                Debug.Print "-------------------------------------------------------"
                                AGGIORNA_FILE_ATTO_pfunct = Str1
                                
                                
                            End If
                            
                        '//..............................................................
                
        '//-------------------------------------------:-----------------------------------------------
        






'//--------------------------------------------------------------------------------------------------
'//                       FINE FUNCTION E GESTIONE ERRORI

Exit_AGGIORNA_FILE_ATTO_pfunct:
Exit Function

Err_AGGIORNA_FILE_ATTO_pfunct:
    MsgBox "ERRORE FUNCTION PUBLIC    " & Err.Number & " - " & Err.Description, vbCritical, "AGGIORNA_FILE_ATTO_pfunct"
    Resume Exit_AGGIORNA_FILE_ATTO_pfunct
 
End Function
'//AGGIORNA_FILE_ATTO_pfunct                                              *** FINE ***
'//===========================================================:===================================:================================


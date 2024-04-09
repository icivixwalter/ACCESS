Attribute VB_Name = "UTIL_COSTRUISCI_QUERY_Mdl01_SELECT_IMMOBILI_PER_SOGGETTO"
'********************************************************************************************************
'*                                                                                                      *
'*                         VARIABILI GENERALI                                                           *
'*                                                                                                      *
'*                                                                                                      *
'*NOTE  :                                                                                               *
'*                                                                                                      *
'*                                                                                                      *
'*                                                                                                      *
'********************************************************************************************************


'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>><<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'Option

Option Compare Text
Option Explicit

'Variabili di database
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>><<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

'DAO
Dim DaoDB As DAO.Database
Dim DaoWks As DAO.Workspace
Dim DaoRs As DAO.Recordset

'ADO
Dim ADODB As Database
Dim AdodaoRs As Recordset

'Contatori
Dim iCount As Integer
Dim dbl_count As Double
Dim iField As Integer                                       'nro campi record

'Le variabili generiche
Dim sSql As String                                          ' Stringa di estrazione

'Variabili generali
Dim Str1 As String
Dim Int1 As Integer
Dim Lng1 As Long
Dim Dbl1 As Double
Dim Bln1 As Boolean
Dim Vv1 As Variant



'ITERAZIONE CAMPI DEL RECORD
Private Sub Iterazione_Field()

On Error GoTo Err_Iterazione_Field



            'APRO IL RS
            '------------------------------------------------------------------
                'Inserire nome qry o tabella
                Set DaoRs = CurrentDb.OpenRecordset("UTILVS25_N02_ELENCO_TRACCIATO_IMMOBILI_PER_SOGGETTO")
                     
                    'individuo il totale dei campi
                    iField = DaoRs.RecordCount - 1
                    
                'Specifico il Rs
                With DaoRs
                    'Itero nei campi del rs
                    For iCount = 0 To .Fields.Count - 1
                        'Stampo le proprieta
                        Vv1 = DaoRs.Fields(iCount).Name
                        Debug.Print "Field : " & Vv1 & " - nro field : " & iCount
                    Next
                    .MoveNext
                End With
     




'USCITA  E GESTIONE ERRORI
'...............................................................................................

Exit_Iterazione_Field:
    Exit Sub

Err_Iterazione_Field:
    MsgBox Err.Description
    Resume Exit_Iterazione_Field


End Sub





'COSTRUISCI QUERY SELECT
Private Sub COSTRUISCI_QUERY_SELECT()

On Error GoTo Err_COSTRUISCI_QUERY_SELECT

Dim sxTABELLA As String

            'APRO IL RS
            '------------------------------------------------------------------
                'Esempio 1) di istruzione select da costruire
                '......................................................
                    'SELECT
                    'UTIL_Tb01_N02_TRACCIATO.[TABELLA DI APPLICAZIONE],
                    'UTIL_Tb01_N02_TRACCIATO.[Campo Delle Tabelle]
                    'FROM UTIL_Tb01_N02_TRACCIATO;
                '......................................................

                
                
                'reset
                '...................................................
                    'Tabella interessata dall'istruzione select
                    sxTABELLA = "GESucc04_Tb_QUADRO_FABBRICATI_TMP"
                '...................................................
                
                
                'Inserire nome qry o tabella
                Set DaoRs = CurrentDb.OpenRecordset("UTILVS25_N02_ELENCO_TRACCIATO_IMMOBILI_PER_SOGGETTO")
                        
                        'reset
                        iCount = 0
                        
                        'individuo il totale dei campi
                        iField = DaoRs.RecordCount - 1
                    
                        
                        'Imposto l'istruzione principale
                        'della select
                        Debug.Print "SELECT "
                        
                    If DaoRs.EOF = False And DaoRs.BOF = False Then
                       DaoRs.MoveFirst
                       'Specifico il Rs
                       While Not DaoRs.EOF
                                        
                            'Costruisco il corpo della select
                            If iCount < iField Then
                                'stampo i campi con la virgola di separazione
                                Debug.Print sxTABELLA & ".["; DaoRs.Fields(1).Value & "]"; ","
                                    
                            Else
                                'l'ultimo campo viene stampato senza virgola
                                Debug.Print sxTABELLA & ".["; DaoRs.Fields(1).Value & "]"
                                
                                
                            End If
                           
                                DaoRs.MoveNext
                                iCount = iCount + 1
                       Wend
                       
                        'Imposto l'istruzione finale
                        'della select
                        Debug.Print "FROM " & sxTABELLA & ";"
                        
                    
                    End If




'USCITA  E GESTIONE ERRORI
'...............................................................................................

Exit_COSTRUISCI_QUERY_SELECT:
    Exit Sub

Err_COSTRUISCI_QUERY_SELECT:
    MsgBox Err.Description
    Resume Exit_COSTRUISCI_QUERY_SELECT


End Sub



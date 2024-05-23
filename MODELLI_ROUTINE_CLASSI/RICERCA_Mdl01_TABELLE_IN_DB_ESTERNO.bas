Attribute VB_Name = "RICERCA_Mdl01_TABELLE_IN_DB_ESTERNO"
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
    Dim PATH_s As String                        '//la path


    '//Contatori
    Dim iCount As Integer
    Dim dbl_count As Double
    
   
    'Le variabili generiche
    Dim Vv1 As Variant
    Dim Dbl1 As Double
    Dim Int1 As Integer
    Dim Long1 As Long
    Dim Str1 As String
    
    '//Messaggi di errore
    Dim ProceduraMessaggioErrore_s As String    '//Errore procedura
    Dim ProceduraAttivaEseguita_s As String     '//Errore Attivita eseguita


'....................................................................

'//++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++//





Sub RICERCA_ListTablesInExternalDB()
    Dim dbPath As String
    Dim externalDB As DAO.Database
    Dim tdf As DAO.TableDef
    Dim appAccess As New Access.Application

    ' Percorso del database esterno
    dbPath = "c:\Casa\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MDB\MENU_TB03_OGGETTI_DA_CANCELLARE\MENU_TB03_OGGETTI_DA_CANCELLARE.mdb"

    ' Apri il database esterno
    Set externalDB = DBEngine.Workspaces(0).OpenDatabase(dbPath)
    
        
        


    ' Stampa l'elenco delle tabelle non di sistema
    For Each tdf In externalDB.TableDefs
        ' Esclude le tabelle di sistema
        If Left(tdf.Name, 4) <> "MSys" Then
            Debug.Print tdf.Name
        End If
    Next tdf

    ' Chiudi il database esterno
    externalDB.Close
    Set externalDB = Nothing
End Sub




'// RICERCA DATI NEL RECORD
'//======================================================================================================//
'//Note           : chiamo la funzione private per il recupero dei valori.

'RECUPERO_DATI_NEL_RECORD
'..............................................................................................................
'Tipo           : Routine pubblica.
'Attività'      : Recupero il tipo di Tributo del codice F24
'Note           : Recupero la descrizione del tipo di tributo corrispondente al codice F24.
'Parametro      : par_iAnnoImp = anno di imposta e par_sCodiceTributo = Codice Tributo F24.
'Restituisce    : Il la descrizione del tipo di Tributo.
'Codice         : ITERAZIONE_RECORD_N01_pFunct.01
'

Public Function ITERAZIONE_RECORD_N01_pFunct(par_iAnnoImp As Integer, _
                                            par_sCodiceTributo As String) As String
            
    '....
On Error GoTo Err_ITERAZIONE_RECORD_N01_pFunct


        
        '//Imposto i parametri
        Dim par_AnnoImp_i As Integer
        Dim par_CodiceTributo_s As String

    '//ITERO NELLA TABELLA
    '//.....................................................................................................
    '//Note           : Tramite una Select vengono individuati i valori da restiuire.

        '//RECUPERO PARAMETRO DA TABELLA OGGETTI
        Set DaoRs = CurrentDb.OpenRecordset("Tabella/Query")

        If DaoRs.EOF = False And DaoRs.BOF = False Then

            DaoRs.MoveFirst

            While Not DaoRs.EOF
            If DaoRs.Fields("Campo_01") = par_AnnoImp_i _
            And DaoRs.Fields("Campo_02") = par_CodiceTributo_s Then

                
            End If

            DaoRs.MoveNext

            Wend

            DaoRs.Close
            Set DaoRs = Nothing

        End If

    '//*** fine ***
    '//ITERO NELLA TABELLA
    '//.....................................................................................................

'USCITA  E GESTIONE ERRORI
'..............................................................................................................


Exit_ITERAZIONE_RECORD_N01_pFunct:
    Exit Function

Err_ITERAZIONE_RECORD_N01_pFunct:
    MsgBox Err.Description
    Debug.Print ProceduraMessaggioErrore_s
    Debug.Print ProceduraAttivaEseguita_s
    Stop
    Resume Exit_ITERAZIONE_RECORD_N01_pFunct

End Function
'*** FINE ***
'RECUPERO_DATI_NEL_RECORD
'..............................................................................................................

'// RICERCA DATI NEL RECORD
'//======================================================================================================//




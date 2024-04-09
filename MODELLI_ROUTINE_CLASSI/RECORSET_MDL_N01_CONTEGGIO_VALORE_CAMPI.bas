Attribute VB_Name = "RECORSET_MDL_N01_CONTEGGIO_VALORE_CAMPI"
Option Compare Database

'//Imposto variabili
Dim icount As Integer                                       '//Conta il campo valorizzato
Dim TOT_icount  As Long                                     '//TOTALE CONTEGGI

'//Imposto i parametri
Dim TABELLA_s  As String
Dim CAMPO_s As String


'// RECORSET_CONTEGGIO_VALORE_CAMPI--->PROVA
'//======================================================================================================//
    
 Private Sub PROVA_RECORSET_N01_CONTEGGIO_VALORE_CAMPI_PFunct()

    '//I PARAMETRI
    TABELLA_s = "TB02_QUESITI"
    CAMPO_s = "LAV_b"
    
    '//ESEGUI LA PROVA
    Vv1 = RECORSET_N01_CONTEGGIO_VALORE_CAMPI_PFunct(TABELLA_s, CAMPO_s)

 End Sub

'// RECORSET_CONTEGGIO_VALORE_CAMPI--->PROVA   *** FINE ***
'//======================================================================================================//



'// RECORSET_CONTEGGIO_VALORE_CAMPI
'//======================================================================================================//
'//Note           : chiamo la funzione private per il recupero dei valori.

'//RECORSET_CONTEGGIO_VALORE_CAMPI
'//..............................................................................................................
'//Tipo           : Routine pubblica.
'//Attività'      : Recupero il valore restituito da una query su di una tabella
'//Note           : Recupero IL TOTALE DI UN CAMPO
'//Parametro      : par_TABELLA_s = anno di imposta e par_CAMPO_s = Codice Tributo F24.
'//Restituisce    : Il la descrizione del tipo di Tributo.
'//Codice         : RECORSET_N01_CONTEGGIO_VALORE_CAMPI_PFunct.01
'//

Public Function RECORSET_N01_CONTEGGIO_VALORE_CAMPI_PFunct(par_TABELLA_s As String, _
                                                           par_CAMPO_s As String) As Long
            
    '....
On Error GoTo Err_RECORSET_N01_CONTEGGIO_VALORE_CAMPI_PFunct

'//RESET
icount = 0
TOT_icount = 0

        
    '//ITERO NELLA TABELLA
    '//.....................................................................................................
    '//Note           : Tramite una Select vengono individuati i valori da restiuire.

        '//RECUPERO PARAMETRO DA TABELLA OGGETTI
        Set DaoRs = CurrentDb.OpenRecordset(par_TABELLA_s)

        If DaoRs.EOF = False And DaoRs.BOF = False Then

            DaoRs.MoveFirst

            While Not DaoRs.EOF
            
            If DaoRs.Fields(par_CAMPO_s) = -1 Then
                icount = icount + 1
                TOT_icount = icount
                
              
            End If

            DaoRs.MoveNext

            Wend
                
                  '//Restituisco al Chiamante
                RECORSET_N01_CONTEGGIO_VALORE_CAMPI_PFunct = TOT_icount


            DaoRs.Close
            Set DaoRs = Nothing

        End If

    '//*** fine ***
    '//ITERO NELLA TABELLA
    '//.....................................................................................................

'USCITA  E GESTIONE ERRORI
'..............................................................................................................


Exit_RECORSET_N01_CONTEGGIO_VALORE_CAMPI_PFunct:
    Exit Function

Err_RECORSET_N01_CONTEGGIO_VALORE_CAMPI_PFunct:
    MsgBox Err.Description
    Debug.Print ProceduraMessaggioErrore_s
    Debug.Print ProceduraAttivaEseguita_s
    Stop
    Resume Exit_RECORSET_N01_CONTEGGIO_VALORE_CAMPI_PFunct

End Function
'*** FINE ***
'RECORSET_CONTEGGIO_VALORE_CAMPI
'..............................................................................................................

'// RICERCA DATI NEL RECORD
'//======================================================================================================//




Attribute VB_Name = "MODELLI_DEFINIZIONE_Mdl01_SUB"
Option Compare Database



'//RECUPERO_DATI_NEL_RECORD
'//========================================================================================================================================//
'//Tipo           : Routine pubblica.
'//Attività       : Recupero il tipo di Tributo del codice F24
'//Note           : Recupero la descrizione del tipo di tributo corrispondente al codice F24.
'//Parametro      : par_iAnnoImp = anno di imposta e par_sCodiceTributo = Codice Tributo F24.
'//Codice         : MODELLO_SUB_N01_IterazioneRecord_pSub.01
'//

Public Sub MODELLO_SUB_N01_IterazioneRecord_pSub(par_iAnnoImp As Integer, _
                                                 par_sCodiceTributo As String)

'//MessaggiDiErrore
Dim ProceduraMessaggioErrore_s As String
Dim ProceduraAttivaEseguita_s As String
 
 
'//Campo
Dim CampoCercato_s As String

'//Campi parametri
Dim par_AnnoImp_i As Integer
Dim par_CodiceTributo_s As String

            
    '//....
On Error GoTo Err_MODELLO_SUB_N01_IterazioneRecord_pSub


        
        '//Imposto i parametri
        ProceduraAttivaEseguita_s = "MODELLO_SUB_N01_IterazioneRecord_pSub"
        ProceduraMessaggioErrore_s = "Errore nella procedura"
        
    '//ITERO NELLA TABELLA
    '//.....................................................................................................
    '//Note           : Tramite una Select vengono individuati i valori da restiuire.

        '//RECUPERO PARAMETRO DA TABELLA OGGETTI
        Set DAORs = CurrentDb.OpenRecordset("Tabella/Query")

        If DAORs.EOF = False And DAORs.BOF = False Then

            DAORs.MoveFirst

            While Not DAORs.EOF
            If DAORs.Fields("Campo_01") = par_AnnoImp_i _
            And DAORs.Fields("Campo_02") = par_CodiceTributo_s Then

                '//Salvo nella Variabile
                CampoCercato_s = DAORs.Fields("Campo_01")

            End If

            DAORs.MoveNext

            Wend

            DAORs.Close
            Set DAORs = Nothing

        End If

    '//*** fine ***
    '//ITERO NELLA TABELLA
    '//.....................................................................................................

'//USCITA  E GESTIONE ERRORI
'//..............................................................................................................


Exit_MODELLO_SUB_N01_IterazioneRecord_pSub:
    Exit Sub

Err_MODELLO_SUB_N01_IterazioneRecord_pSub:
    MsgBox Err.Description & " - Errore Messaggio -> : " & ProceduraMessaggioErrore_s & " Procedura -> : " & ProceduraMessaggioErrore_s
    Debug.Print ProceduraMessaggioErrore_s
    Debug.Print ProceduraAttivaEseguita_s
    Stop
    Resume Exit_MODELLO_SUB_N01_IterazioneRecord_pSub

End Sub

'//*** FINE ***
'//RECUPERO_DATI_NEL_RECORD
'//========================================================================================================================================//






'//========================================================================================================================================//
'//MODELLO_ROUTINE
'//MODELLO_SUB_N01_Routine_2_parametri_pSub
'//========================================================================================================================================//
'//DENOMINAZIONE--->:Aggiorno il campo nro progressivo, file di archivio e _
                     la data e l'ora di aggiornamento della tabella GE_CASA_Tb12_POSTA_MOVIM_CARICA_DATI
                                            
'//Tipo------------>:Routine pubblica.
'//Attività-------->:Aggiorno campi _
                     della tabella GE_CASA_Tb12_POSTA_MOVIM_CARICA_DATI
'//NOTE------------>:Aggiorna il campo nro progressivo ed il file di archivio contestualmente alla data ed all'ora _
                     di aggiornamento: Nro progressivo = viene costruito con un contatore mentre il file di archivio = viene _
                     costruito con le seguenti variabili con in questo esempio : _
                     "POSTA_BPOL_" & Year(DaoRs.Fields("DataContabile_d")) & Str1 & ".xls"
'//Parametro------->:par_iInt = tipo integer da utilizzare, _
                     par_Utile01_s = due campi stringa da utilizzare. _
                     par_NomeFileArchivio_s = Nome file xls da passare come parametro e se null _
                     viene lanciato un messaggio di operazione annullata.
'//Restituisce----->:Null _
'//Codice---------->:MODELLO_SUB_N01_Routine_2_parametri_pSub.01 _
'/


Public Sub MODELLO_SUB_N01_Routine_2_parametri_pSub(par_iAnnoImp As Integer, _
                                                    par_sCodiceTributo As String)

'//MessaggiDiErrore
Dim ProceduraMessaggioErrore_s As String
Dim ProceduraAttivaEseguita_s As String
 
 
'//Campo
Dim CampoCercato_s As String

'//Campi parametri
Dim par_AnnoImp_i As Integer
Dim par_CodiceTributo_s As String

            
    '//....
On Error GoTo Err_MODELLO_SUB_N01_Routine_2_parametri_pSub


        
        '//Imposto i parametri
        ProceduraAttivaEseguita_s = "MODELLO_SUB_N01_Routine_2_parametri_pSub"
        ProceduraMessaggioErrore_s = "Errore nella procedura"
        
    '//ITERO NELLA TABELLA
    '//.....................................................................................................
    '//Note           : Tramite una Select vengono individuati i valori da restiuire.

        '//RECUPERO PARAMETRO DA TABELLA OGGETTI
        '//-------------------------------------------------------------------------------
            
    '//-------------------------------------------------------------------------------
    '//*** fine ***
    '//ITERO NELLA TABELLA
    '//.....................................................................................................

'//USCITA  E GESTIONE ERRORI
'//..............................................................................................................


Exit_MODELLO_SUB_N01_Routine_2_parametri_pSub:
    Exit Sub

Err_MODELLO_SUB_N01_Routine_2_parametri_pSub:
 '//-------------------------------------------------------------------------------
    MsgBox Err.Description & " - Errore Messaggio -> : " & ProceduraMessaggioErrore_s & " Procedura -> : " & ProceduraMessaggioErrore_s
    Debug.Print ProceduraMessaggioErrore_s
    Debug.Print ProceduraAttivaEseguita_s
    '//COSTRUISCO IL MESAGGIO DI ERRORE E LO STAMPO CON IL METODO DEBUG.PRINT
    '//.............................................................................................
	Vv1 = "ERRORE NRO : " & Err.Number & " - TIPO DI ERRORE ==>: " & Err.Description & Chr$(13) _
	& " - ROUTINE SUB: " & ROUT_NRO_i & " - " & ROUT_TIPO_MSG_s & " - " & ROUT_ERR_MSG_s
	Debug.Print Vv1
	Str1 = MsgBox(Vv1, vbCritical)

	'//BLOCCO DELLA ROUTINE.
	Stop
	Resume Exit_MODELLO_SUB_N01_Routine_2_parametri_pSub
    
    '//.............................................................................................

    
'//-------------------------------------------------------------------------------
        
End Sub

'//*** FINE ***
'//========================================================================================================================================//
'//MODELLO_SUB_N01_Routine_2_parametri_pSub
'//========================================================================================================================================//




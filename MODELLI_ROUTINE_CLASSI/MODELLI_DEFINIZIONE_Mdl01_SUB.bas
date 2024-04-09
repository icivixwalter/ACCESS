Attribute VB_Name = "MODELLI_DEFINIZIONE_Mdl01_SUB"
Option Compare Database

'//SENZA_PARAMETRI
'//========================================================================================================================================//
'//Tipo           : Routine pubblica.
'//Attività       : Recupero il tipo di Tributo del codice F24
'//Note           : Recupero la descrizione del tipo di tributo corrispondente al codice F24.
'//Parametro      : par_iAnnoImp = anno di imposta e par_sCodiceTributo = Codice Tributo F24.
'//Codice         : MODELLO_SUB_N00_SENZA_PARAMETRI_pSub.01
'//

Public Sub MODELLO_SUB_N00_SENZA_PARAMETRI_pSub()

'//MessaggiDiErrore
Dim ProceduraMessaggioErrore_s As String
Dim ProceduraAttivaEseguita_s As String
 
 
'//Campo
Dim CampoCercato_s As String

'//Campi parametri
Dim par_AnnoImp_i As Integer
Dim par_CodiceTributo_s As String

            
    '//....
On Error GoTo Err_MODELLO_SUB_N00_SENZA_PARAMETRI_pSub


        
        '//Imposto i parametri
        ProceduraAttivaEseguita_s = "MODELLO_SUB_N00_SENZA_PARAMETRI_pSub"
        ProceduraMessaggioErrore_s = "Errore nella procedura"
        
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

                '//Salvo nella Variabile
                CampoCercato_s = DaoRs.Fields("Campo_01")

            End If

            DaoRs.MoveNext

            Wend

            DaoRs.Close
            Set DaoRs = Nothing

        End If

    '//*** fine ***
    '//ITERO NELLA TABELLA
    '//.....................................................................................................

'//USCITA  E GESTIONE ERRORI
'//..............................................................................................................


Exit_MODELLO_SUB_N00_SENZA_PARAMETRI_pSub:
    Exit Sub

Err_MODELLO_SUB_N00_SENZA_PARAMETRI_pSub:
    MsgBox Err.Description & " - Errore Messaggio -> : " & ProceduraMessaggioErrore_s & " Procedura -> : " & ProceduraMessaggioErrore_s
    Debug.Print ProceduraMessaggioErrore_s
    Debug.Print ProceduraAttivaEseguita_s
    Stop
    Resume Exit_MODELLO_SUB_N00_SENZA_PARAMETRI_pSub

End Sub

'//*** FINE ***
'//SENZA_PARAMETRI
'//========================================================================================================================================//



'//RECUPERO_DATI_NEL_RECORD
'//========================================================================================================================================//
'//Tipo           : Routine pubblica.
'//Attività       : Recupero il tipo di Tributo del codice F24
'//Note           : Recupero la descrizione del tipo di tributo corrispondente al codice F24.
'//Parametro      : par_iAnnoImp = anno di imposta e par_sCodiceTributo = Codice Tributo F24.
'//Codice         : MODELLO_SUB_N01_IterazioneRecord_2Parametri_pSub.01
'//

Public Sub MODELLO_SUB_N01_IterazioneRecord_2Parametri_pSub(par_iAnnoImp As Integer, _
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
On Error GoTo Err_MODELLO_SUB_N01_IterazioneRecord_2Parametri_pSub


        
        '//Imposto i parametri
        ProceduraAttivaEseguita_s = "MODELLO_SUB_N01_IterazioneRecord_2Parametri_pSub"
        ProceduraMessaggioErrore_s = "Errore nella procedura"
        
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

                '//Salvo nella Variabile
                CampoCercato_s = DaoRs.Fields("Campo_01")

            End If

            DaoRs.MoveNext

            Wend

            DaoRs.Close
            Set DaoRs = Nothing

        End If

    '//*** fine ***
    '//ITERO NELLA TABELLA
    '//.....................................................................................................

'//USCITA  E GESTIONE ERRORI
'//..............................................................................................................


Exit_MODELLO_SUB_N01_IterazioneRecord_2Parametri_pSub:
    Exit Sub

Err_MODELLO_SUB_N01_IterazioneRecord_2Parametri_pSub:
    MsgBox Err.Description & " - Errore Messaggio -> : " & ProceduraMessaggioErrore_s & " Procedura -> : " & ProceduraMessaggioErrore_s
    Debug.Print ProceduraMessaggioErrore_s
    Debug.Print ProceduraAttivaEseguita_s
    Stop
    Resume Exit_MODELLO_SUB_N01_IterazioneRecord_2Parametri_pSub

End Sub

'//*** FINE ***
'//RECUPERO_DATI_NEL_RECORD
'//========================================================================================================================================//






'//MODELLO_ROUTINE
'//========================================================================================================================================//
'//Tipo           : Routine pubblica.
'//Attività       : Recupero il tipo di Tributo del codice F24
'//Note           : Recupero la descrizione del tipo di tributo corrispondente al codice F24.
'//Parametro      : par_iAnnoImp = anno di imposta e par_sCodiceTributo = Codice Tributo F24.
'//Codice         : MODELLO_SUB_N01_Routine_2_parametri_pSub.01
'//

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
    Stop
    Resume Exit_MODELLO_SUB_N01_Routine_2_parametri_pSub
'//-------------------------------------------------------------------------------
        
End Sub

'//*** FINE ***
'//MODELLO_ROUTINE
'//========================================================================================================================================//




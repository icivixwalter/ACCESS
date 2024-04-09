Attribute VB_Name = "MODELLI_DEFINIZIONE_Mdl15_RICERCA_RECORD"
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'
'                                    MODELLO SUB E FUNZIONI
'
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////




'//MODELLO_01_01_SUB_PRIVATE_SCHELETRO
'//======================================================================================================//

Private Sub Chiama_RicercaPosizioneRECORD_pSub()

'//Chiamo con 5 parametri
RicercaPosizioneRECORD_pSub "A", "B", "C", "D", 1

End Sub


'//Ricerca Record
Private Sub RicercaPosizioneRECORD_pSub(par_sxNomeText As String, _
            par_vxValue As Variant, _
            par_sxDefaultValue As String, _
            par_sxFontName As String, _
            par_ixFontSize As Integer)

'//Dim
Dim ProceduraMessaggioErrore_sx
Dim ProceduraAttivaEseguita_sx

    '....
    
On Error GoTo Err_RicercaPosizioneRECORD_pSub


    'CONTROLLO ESISTENZA DELL'APOSTRO O APICE
    '______________________________________________________________________________

                
                'AGGIORNO TABELLA TTy_ANAG_Tb09_ANAGRAFICA_RECORD_TMP
                '...................................................................
                
                '...................................................................

    'CONTROLLO ESISTENZA DELL'APOSTRO O APICE
    '______________________________________________________________________________



'USCITA  E GESTIONE ERRORI
'..............................................................................................................


Exit_RicercaPosizioneRECORD_pSub:
    Exit Sub

Err_RicercaPosizioneRECORD_pSub:
    MsgBox Err.Description
    Debug.Print ProceduraMessaggioErrore_sx
    Debug.Print ProceduraAttivaEseguita_sx
    Stop
    Resume Exit_RicercaPosizioneRECORD_pSub

End Sub

'//MODELLO_01_01_SUB_PRIVATE_SCHELETRO
'//======================================================================================================//






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

                '//Restituisco al Chiamante
                Recupero_N02_DESCRIZIONE_CODICE_RECORD_DalCodice_pFunct = DaoRs.Fields("Campo_01")

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




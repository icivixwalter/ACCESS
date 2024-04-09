Attribute VB_Name = "MODELLI_FUNCTION"
Option Compare Database

'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'
'                                    MODELLI_FUNCTION
'
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////





'//*****************************************************************************************************************************
'//                                         LE FUNCTION
'//*****************************************************************************************************************************


'//TIPO_FUNCTION_N005_Function_05_PARAMETRI_Funct
'//LA_SUB_RicercaPosizioneRecord_Funct_01.00
'//================================================================================================================//
Private Function RicercaPosizioneRecord_Funct(par_sxNomeText As String, _
                                                par_vxValue As Variant, _
                                                par_sxDefaultValue As String, _
                                                par_sxFontName As String, _
                                                par_ixFontSize As Integer) As Integer

'//Le variabili di controllo della sub
Dim ProceduraMessaggioErrore_s As String
Dim ProceduraAttivaEseguita_s As String
   



    '//....
On Error GoTo Err_RicercaPosizioneRecord_Funct


'//RESET VARIABILI
ProceduraMessaggioErrore_s = "PROCEDURA_ESEGUITA_RicercaPosizioneRecord_Funct"
ProceduraAttivaEseguita_s = "ATTIVITA_DI_CONTROLLO_controllo_record"


'//LA_SUB_RicercaPosizioneRecord_Funct_01.01
'//CONTROLLO_ESISTENZA_APOSTRO_O_APICE
'//______________________________________________________________________________

            '//LA_SUB_RicercaPosizioneRecord_Funct_01.02
            '//AGGIORNO TABELLA TTy_ANAG_Tb09_ANAGRAFICA_RECORD_TMP
            '//...................................................................
            
            '//...................................................................

'//CONTROLLO_ESISTENZA_APOSTRO_O_APICE *** FINE ***
'//______________________________________________________________________________


'//Ritorno_valore_della_FUNZIONE
RicercaPosizioneRecord_Funct = 0



'//USCITA  E GESTIONE ERRORI
'//..............................................................................................................


Exit_RicercaPosizioneRecord_Funct:
    Exit Function

Err_RicercaPosizioneRecord_Funct:
    MsgBox Err.Description
    Debug.Print ProceduraMessaggioErrore_s
    Debug.Print ProceduraAttivaEseguita_s
    Stop
    Resume Exit_RicercaPosizioneRecord_Funct

End Function

'//LA_SUB_RicercaPosizioneRecord_Funct_01.00 *** FINE ***
'//================================================================================================================//




'//*****************************************************************************************************************************
'//                                         LE FUNCTION *** FINE ***
'//*****************************************************************************************************************************


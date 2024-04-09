Attribute VB_Name = "MODELLI_DEFINIZIONE_Mdl02_FUNCTION"
Option Compare Text
Option Explicit


'//*********************************************************************************************************************************************//
'//   Public Function di
'//
'//
'//*********************************************************************************************************************************************//


'//pfFunction_base
'//===========================================================:===================================:================================
'//SEZX_Mdl_n000_000_Function.000.01__________:Funzione 00 CALCOLA                                :Calcola
'//...........................................:.....................................................................................
'//NOTA PROCEDURA: La Funzione attiva un rs per i conteggi dei record e per i calcoli vari.

Public Function pfFunction_base(par_variabile As String)

'//.......................................................
'// DIM SCELTA DATABASE DA APRIRE
Dim sScelta_db As String

    
    
        On Error GoTo Err_pfFunction_base
    
    
        '//DENOM
        '//-------------------------------------------:-----------------------------------------------
        '//SEZX_Mdl_n000_000_Function.000.02__________:
        '//NOTA   :
        
                    
                '//
                '//..............................................................
                '//SEZX_Mdl_n000_000_Function.000.01.01_______:
                '//NOTA   :

                
                    '//   Commento
                        '//scrivi istruzioni
                
        






'//--------------------------------------------------------------------------------------------------
'//                       FINE FUNCTION E GESTIONE ERRORI

Exit_pfFunction_base:
Exit Function

Err_pfFunction_base:
    MsgBox "ERRORE FUNCTION PUBLIC    " & Err.Number & " - " & Err.Description, vbCritical, "pfFunction_base"
    Resume Exit_pfFunction_base
 
End Function
'//pfFunction_base                                              *** FINE ***
'//===========================================================:===================================:================================




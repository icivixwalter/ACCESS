Attribute VB_Name = "ESPORTA_FILE_Mdl01_01_Esporta_File_XLS"
Option Compare Database




'//ESPORTA_FILE_XLS_pFunct
'//*************************************************************************************************//
'//NOTE : ESPORTA i file in formato xls (csv)
'//*************************************************************************************************//
'//ATTIVO LA PROCEDURA
Private Sub Attiva_ESPORTA_FILE_XLS_pFunct()
    Call ESPORTA_FILE_XLS_pFunct
End Sub
'//ESPORTA_FILE_XLS_pFunct()
'//=================================================================================================//
Public Function ESPORTA_FILE_XLS_pFunct()
Dim Path_s As String
Dim Tabella_01_s As String
Dim Tabella_02_s As String

On Error GoTo ESPORTA_FILE_XLS_pFunct_Err


'//ESPORTA I DATI IN FILE XLS
'//........................................................
   'Call ESPORTA_FILE_XLS_pFunct
'//........................................................

'//Reset ed impostazioni
'//........................................................
    Path_s = vbNull
    Path_s = "c:\GESTIONI\GESTIONE_LLPP\25_GESTIONE_LLPP\"
    
    Tabella_01_s = vbNull
    Tabella_01_s = "LLPP_IMPEGNI_Tb01_ATTI_DI_IMPEGNO"
    Tabella_02_s = vbNull
    Tabella_02_s = "LLPP_IMPEGNI_Tb02_ELENCO_DI_SPESA"

'//........................................................

    '//Esporto le tabelle in formato xls
    DoCmd.OutputTo acTable, Tabella_01_s, "MicrosoftExcel(*.xls)", Path_s & "LLPP_IMPEGNI_Tb01_ATTI_DI_IMPEGNO_Coll.xls", False, ""
    DoCmd.OutputTo acTable, Tabella_02_s, "MicrosoftExcel(*.xls)", Path_s & "LLPP_IMPEGNI_Tb02_ELENCO_DI_SPESA_Coll.xls", False, ""
    
    '//Messaggio di fine attivita'
    MsgBox "ESPORTAZIONE DATI IN FILE XLS EFFETTUATI in --> " & Path_s & vbCrLf & Tabella_01_s & "_Coll.xls" & vbCrLf & Tabella_02_s & "_Coll.xls"
    

ESPORTA_FILE_XLS_pFunct_Exit:
    Exit Function

ESPORTA_FILE_XLS_pFunct_Err:
    MsgBox Error$
    Resume ESPORTA_FILE_XLS_pFunct_Exit

End Function
'//ESPORTA_FILE_XLS_pFunct()                    ***FINE***
'//=================================================================================================//

'//ESPORTA_FILE_XLS_pFunct
'//*************************************************************************************************//
'//NOTE : ESPORTA i file in formato xls (csv)   *** FINE ***
'//*************************************************************************************************//


'//Chiamo matrice con Stringhe che rappresentano delle tabelle _
Impegni, Presenze
Private Sub CHIAMA_ESPORTA_FILE_XLS_ConMatriceDati_pFunct()
    
    '//Chiamo la procedura di esportazione indicando il nome delle tabelle da esportare in xls
    ESPORTA_FILE_XLS_ConMatriceDati_pFunct "LLPP_IMPEGNI_Tb01_ATTI_DI_IMPEGNO", "LLPP_IMPEGNI_Tb02_ELENCO_DI_SPESA", _
    "PRES3000_Tb01_Calendario", "PRES3000_Tb02_ElencoGiornaliere", "PRES3000_Tb01_Calendario", "Indirizzi_Tb01_INTESTATARI", _
    "Indirizzi_Tb02_ELENCO"

End Sub


'//ESPORTA_FILE_XLS_ConMatriceDati_pFunct()
'//=================================================================================================//
'//NOTE ------------->  : attenzione esporto la collezione di file ma il primo deve essere nullo ("") per _
                        evitare che venga saltato.

Public Function ESPORTA_FILE_XLS_ConMatriceDati_pFunct(strName As String, ParamArray intScores() As Variant)
Dim Path_s As String
Dim Tabella_01_s As String
Dim NuovaDenominazioneTabella_01_s As String
Dim Tabella_02_s As String
Dim NuovaDenominazioneTabella_02_s As String
Dim intI As Integer
Dim iCount As Integer                   '//il contatore

On Error GoTo ESPORTA_FILE_XLS_ConMatriceDati_pFunct_Err


'//ESPORTA I DATI IN FILE XLS
'//........................................................
 


   '//Chiamo la procedura di esportazione indicando il nome delle tabelle da esportare in xls _
      ATTENZIONE, occorre inserire un record vuoto "", affinche possono essere considerati tutti gli altri _
      file dell'esportazione altrimenti di viene saltato il primo "LLPP_ATTI_Tb01_Gestione", _
                Call ESPORTA_FILE_XLS_ConMatriceDati_pFunct "", _
                "LLPP_ATTI_Tb01_Gestione", _
                "LLPP_ATTI_Tb02_Allegati", _
                "Indirizzi_Tb01_INTESTATARI", _
                "Indirizzi_Tb02_ELENCO", _
                "PRES3000_Tb01_Calendario", _
                "PRES3000_Tb02_ElencoGiornaliere", _
                "LLPP_DF01_CodiceOpera", _
                "LLPP_IMPEGNI_Tb01_ATTI_DI_IMPEGNO", _
                "LLPP_IMPEGNI_Tb02_ELENCO_DI_SPESA"

'//........................................................


'//Reset ed impostazioni
'//........................................................
    Path_s = vbNull
    Path_s = "c:\GESTIONI\GESTIONE_LLPP\25_GESTIONE_LLPP\LPP_ARCHIVI_SALVATAGGI\"
    
    Tabella_01_s = vbNull
    Tabella_01_s = "LLPP_IMPEGNI_Tb01_ATTI_DI_IMPEGNO"

'//........................................................

    Debug.Print
    Debug.Print strName; " Punteggi"
    ' Utilizza la funzione UBound per definire il
    ' limite superiore della matrice.

    '//Dimensione la matrice di file
    
     Dim MatriceFile_s As String
    
    iCount = 0
    For intI = 0 To UBound(intScores())
        Debug.Print "          "; intScores(intI)
        Tabella_01_s = vbNull
        Tabella_01_s = intScores(intI)
        iCount = iCount + 1
    
        NuovaDenominazioneTabella_01_s = intScores(intI) & ".xls"
        MatriceFile_s = MatriceFile_s & iCount & ") " & NuovaDenominazioneTabella_01_s & " " & vbCrLf
        
        '//Esporto le tabelle in formato xls
        DoCmd.OutputTo acTable, Tabella_01_s, "MicrosoftExcel(*.xls)", Path_s & NuovaDenominazioneTabella_01_s, False, ""


    Next intI
    
    
    '//Messaggio di fine attivita'
    MsgBox "ESPORTAZIONE DATI IN FILE XLS EFFETTUATI in --> " & Path_s & vbCrLf & MatriceFile_s & vbCrLf
    

ESPORTA_FILE_XLS_ConMatriceDati_pFunct_Exit:
    Exit Function

ESPORTA_FILE_XLS_ConMatriceDati_pFunct_Err:
    MsgBox Error$
    Resume ESPORTA_FILE_XLS_ConMatriceDati_pFunct_Exit

End Function
'//ESPORTA_FILE_XLS_ConMatriceDati_pFunct()                    ***FINE***
'//=================================================================================================//

'//ESPORTA_FILE_XLS_ConMatriceDati_pFunct
'//*************************************************************************************************//
'//NOTE : ESPORTA i file in formato xls (csv)   *** FINE ***
'//*************************************************************************************************//


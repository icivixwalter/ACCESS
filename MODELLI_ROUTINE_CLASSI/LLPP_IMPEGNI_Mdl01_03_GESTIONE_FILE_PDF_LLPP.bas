Attribute VB_Name = "LLPP_IMPEGNI_Mdl01_03_GESTIONE_FILE_PDF_LLPP"
Option Compare Database
Dim Path_programma_s As String
Dim NameFile_s As String
Dim Stringa1 As String



'//GESTIONE_FILE_PDF_LLPP
'//***********************************************************************************************************************//

'//ESEMPI_DI_GESTIONE_DIRETTA_DELLA_ROUTINE
'//==============================================================================================================//
    
    '//ES_01_CERCA_FILE_PDF
    '//---------------------------------------------------------------------------------------------------//
    '//Note : Routine pubblica che riceve come parametro una stringa, la quale rappresenta
    '//il percoro ed il nome del file da ricercare.
        Private Sub CercaFilePdf_pSub()
            '//Esempio di ricerca file
            Path_programma_s = "c:\GESTIONI\GESTIONE_LLPP\02_SCANNER\ScannerTmp\"
            NameFile_s = "DGC_498_2009.pdf"
            Call ApriFilePdf(Path_programma_s, NameFile_s, 100)
        
        End Sub
    '//---------------------------------------------------------------------------------------------------//
        
    
    '//ES_02_STAMPA_FILE_PDF_ESEMPIO_DIRETTO
    '//---------------------------------------------------------------------------------------------------//
    
        '//Note :Routine di STAMPA DEL FIL PDF di prova per l'esempio indicato.
        Private Sub STAMPA_FILE_Pdf_call()
            '//Esempio di stampa diretta del file passatto come parametro
            Path_programma_s = "c:\GESTIONI\GESTIONE_LLPP\02_SCANNER\ScannerTmp\"
            NameFile_s = "DGC_498_2009.pdf"
            Call ApriFilePdf(Path_programma_s, NameFile_s, 100)
        
            Call StampaFilePdf(Path_programma_s, NameFile_s)
        
        End Sub
    '//---------------------------------------------------------------------------------------------------//


'//ESEMPI_DI_GESTIONE_DIRETTA_DELLA_ROUTINE *** FINE ***
'//==============================================================================================================//







'//STAMPA_FILE_PDF
'//==================================================================================================================//

Public Sub StampaFilePdf(par_Path_programma_s As String, par_NameFile_s As String)


On Error GoTo StampaFilePdf_Err
        
        
        'STAMPO IL FILE PDF SPECIFICO
        '_______________________________________________________________________________________
        'Note       : Apro il file excel speficico indicando come paramentro la path completa
        '           incluso il file exe.
        
            Dim oWshPDF As Object
            Set oWshPDF = CreateObject("WScript.Shell")
            'oWshPDF.Run "AcroRd32.exe" & " /p /h C:\documenti\nomeFile.pdf"
            oWshPDF.Run "AcroRd32.exe" & " /p /h " & par_Path_programma_s
            Set oWshPDF = Nothing
        '_______________________________________________________________________________________
   

StampaFilePdf_Exit:
    Exit Sub

StampaFilePdf_Err:
    MsgBox Error$
    Resume StampaFilePdf_Exit

End Sub


'//STAMPA_FILE_PDF    *** fine ***
'//==================================================================================================================//


'//GESTIONE_FILE_PDF_LLPP        *** FINE ***
'//***********************************************************************************************************************//


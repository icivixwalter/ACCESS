Attribute VB_Name = "GESTIONE_FILE_N01_01_FilePdf"
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




'//APRI_IL_FILE_PDF
'//==================================================================================================================//
'//METODO DI APERTURA DI UN PROGRAMMA ESTERNO O DI UN COMANDO DOS MEDIANTE "WScript.Shell"
'//PARAMETRI        : passo 2 stringhe per parametro, la path e il nome del file.pdf

Public Sub ApriFilePdf(par_Path_programma_s As String, par_NameFile_s As String, par_IDGestione_lng As Long)

Dim Path_programma_s As String
Dim NameFile_s As String
Dim Stringa1 As String

On Error GoTo ApriFilePdf_Err
        
        
                      '//APRO FILE PDF SPECIFICO CON IL COMANDO OGGETTO PDF
                        '//::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
                        '//Note: Apro il file PDF speficico indicando come paramentro la path completa
                        '//     incluso il file PDF.
                                           
                                '//Definisco ed imposto la variabile di ricerca + reset variabili _
                                Dim Path_s As String _
                                Dim Path_programma_s As String _
                                Dim NameFile_s As String _
                                Path_s = "" _
                                Str1 = "" _

                                    '// Path_programma_s = Me.PATH_TXT.Value
                                    '// NameFile_s = Me.FILE_CALEND_Txt.Value
                                    '// Debug.Print "controllo variabili -> Path_programma_s " & Path_programma_s _
                                    Debug.Print "controllo variabili -> NameFile_s " & NameFile_s
                                    
                                    '//chiamo la procedura
                                    '//Call ApriFilePdf(Path_programma_s, NameFile_s)
                                    
                                    '//reset variabili
                                    Path_programma_s = vbNull
                                    NameFile_s = vbNull
                                    Path_s = vbNull

                                
                          '//::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
         
  '//APRO FILE PDF SPECIFICO
  '//_______________________________________________________________________________________
  '//Note       : Apro il file PDF speficico indicando come paramentro la path completa
  '//           incluso il file PDF.
    
    '//La path del file = Unisco la path + il file
    par_Path_programma_s = par_Path_programma_s & par_NameFile_s
           
    '//controllo
    Debug.Print par_Path_programma_s
           
    '//CREO LA SHELL come oggetto, attivo il comando "%comspec% /c start " e gli assegno la path per estesa con il nome del file
    '// perchè Shell lancia un EXE e non il PDF, quindi è necessario costruire la stringa in modo da far lanciare
    '// prima la sessione dos e poi il pdf..
    '// La stringa di comando ("%comspec% /c start ") deve essere unica : comando shell + path + file
    Set WshPDF = CreateObject("wscript.shell")
    
    '// Unisco il comando pdf start + path definitiva
    Stringa1 = "%comspec% /c start " & par_Path_programma_s                               '//Comando di apertura
    '//Attivo il comando
    WshPDF.Run Stringa1
    
    '//libero la memoria dalle variabili e dagli oggetti creati
    WshPDF = Null
    Stringa1 = vbNull
'_______________________________________________________________________________________


      
      
ApriFilePdf_Exit:
    Exit Sub

ApriFilePdf_Err:
    MsgBox Error$
    Resume ApriFilePdf_Exit

End Sub

'//APRI_IL_FILE_PDF             *** fine ***
'//==================================================================================================================//



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



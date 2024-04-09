Attribute VB_Name = "FILE_MDL03_GESTIONE_FILE_APRI_PDF"
Option Compare Database

'//GESTIONE_FILE_PDF_LLPP
'//=======================================================================================================//

Dim Path_programma_s As String
Dim NameFile_s As String

'//ATTIVO LA SUB APRI SINGOLO FILE PDF -
Private Sub APRI_FILE_PDF_pSub()

    '//La Path del programma per esteso
    Path_programma_s = "c:\CASA\PRES3000_07\WALTER_ATTI\"
    '//Il nome del file in pdf
    'NameFile_s = "WALTER_CUD_2015_03_(Cud2015_2014)_(21377_76).pdf"
    NameFile_s = "WALTER_CARTELLINO_2014_06"
    
    '//chiamo la sub con i parametri
    ApriFilePdf Path_programma_s, NameFile_s, 100
    
End Sub

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
                                           
                                '//Definisco ed imposto la variabile di ricerca + reset variabili
                                Dim MyPath_s As String
                                Dim MyFileEstensione_s As String
                                
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
                                   

                                

                                '//TROVA FILE  (INIZIO)_E_RECUPERO_LA_ESTENSIONE
                                '//------------------------------------------------------------------------------------------------
                                    ' Visualizza i nomi in c:\ che rappresentano directory.
                                     MyPath_s = par_Path_programma_s
                                
                                    Debug.Print "                       CONTROLLO FILE TROVATO                      "
                                    Debug.Print "..................................................................."
                                    Debug.Print "Path -> " & MyPath_s
                                            
                                            MyFile_s = Dir(par_Path_programma_s & par_NameFile_s & ".*", vbNormal)
                                            If MyFile_s > "" Then
                                                Debug.Print par_MyFile_s              ' Visualizza la voce solo
                                                
                                                
                                                '//SOSPESO _
                                                 MsgBox "OK_CONTROLLO_FILE=TROVATO--->" & Chr$(13) & MyPath_s & Chr$(13) & MyFile_s, vbInformation, "MSG_BOX_DI_AVVISO"
                                                
                                                
                                                
                                                '//TROVA L'ESTENSIONE _
                                                Note : Esamina gli ulti 4 caratteri del file e li salvo su una variabile _
                                                che sarà usato per completare l'intera path+nomefile da applicare all'oggetto _
                                                di seguito richiamato con la shell. Il salvataggio avviene solo se tra il _
                                                file passato come parametro e quello ricercato mediante dir vie è una differenza. _
                                                Se vi è una differenza allora significa che probabilmente è senza estensione.
                                                If MyFile_s <> MyFile_s Then
                                                    MyFileEstensione_s = Right(MyFile_s, 4)
                                                End If
                                                
                                            Else
                                                
                                                '//Se non trova il file emette il messaggio ed esco dalla routine e scrivo sul campo _
                                                NoteAtto_s = "FILE_NON_TROVATO", mediante la query di aggiornamento basata sull'id del record.
                                                
                                                '//Aggiorno la nota della tabella con la dicitura file non trovato se esiste Id del record
                                                If par_IDGestione_lng > 0 Then
                                                    '//Controllo ed esecuzione cmd_sSql
                                                    sSql = ""
                                                    sSql = sSql & "UPDATE GE_CASA_Tb01_MASTRO SET GE_CASA_Tb01_MASTRO.RICERCA_FileAtto_s = 'FILE TROVATO'"
                                                    sSql = sSql & "WHERE (((GE_CASA_Tb01_MASTRO.ID_Tb01_lng)=" & par_IDGestione_lng & "));"
                                                    Debug.Print sSql
                                                    
                                                    '//SALVATAGGIO SOSPESO
                                                    CurrentDb.Execute sSql
                                                    
                                                End If
                                                
                                                Debug.Print "FILE NON TROVATO"
                                                MsgBox "ATTENZIONE!!! FILE_NON_TROVATO--->" & Chr$(13) & MyPath_s & "/" & par_MyFile_s, vbInformation, "MSG_BOX_DI_AVVISO"
                                                'Me.Recalc
                                                
                                                '//libero la memoria
                                                MyFileEstensione_s = vbNull
                                                MyPath_s = vbNull
                                                par_IDGestione_lng = vbNull
                                                GoTo ApriFilePdf_Exit
                                            End If
                                    
                                    Debug.Print "..................................................................."
                                    
                                   ' MsgBox "FILE_TROVATO--->" & MyPath_s & "/" & par_MyFile_s, vbInformation, "MSG_BOX_DI_AVVISO"
                                     
                                
                                '//TROVA FILE__E_RECUPERO_LA_ESTENSIONE  (***FINE***)
                                '//------------------------------------------------------------------------------------------------

                                
                                
                                
                          '//::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
         
  '//APRO FILE PDF SPECIFICO
  '//_______________________________________________________________________________________
  '//Note       : Apro il file PDF speficico indicando come paramentro la path completa
  '//           incluso il file PDF.
    
    '//La path del file = Unisco la path + il file+estensione sopra individuata
    par_Path_programma_s = par_Path_programma_s & par_NameFile_s & MyFileEstensione_s
           
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
    MyFileEstensione_s = vbNull
    MyPath_s = vbNull
'_______________________________________________________________________________________


      
      
ApriFilePdf_Exit:
    Exit Sub

ApriFilePdf_Err:
    MsgBox Error$
    Resume ApriFilePdf_Exit

End Sub

'//APRI_IL_FILE_PDF         *** FINE ***
'//==================================================================================================================//


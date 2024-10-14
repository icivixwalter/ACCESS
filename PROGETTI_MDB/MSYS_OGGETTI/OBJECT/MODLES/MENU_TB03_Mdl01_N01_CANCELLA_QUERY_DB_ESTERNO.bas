Attribute VB_Name = "MENU_TB03_Mdl01_N01_CANCELLA_QUERY_DB_ESTERNO"
'//MODULO = "MENU_TB03_Mdl01_N01_CANCELLA_QUERY_DB_ESTERNO"

Option Compare Database


Dim dbs As Database
Dim qdfProva As QueryDef
Dim qdfCiclo As QueryDef
Dim prpCiclo As Property



'Parametri Table
Dim sxNomeTable As String
Dim sxCodiceTable As String
Dim sxParametroTable As String
Dim ixLungTable As Integer


'Parametri Query
Dim sxNomeQuery As String
Dim sxCodiceQuery As String
Dim sxParametroQuery As String
Dim ixLungQuery As Integer


Dim iCount  As Integer
Dim i As Integer
Dim iTotOggetti As Integer
Dim Bolean1 As Boolean


Dim sxTipoDatabase As String
Dim sxNomeDatabase As String

Dim sxQueryOrigine As String
Dim sxQueryDestinazione As String

Dim sxMessaggioBox  As String





'-----------------------------------------------------------
'1)
'CANCELLO TUTTE LE QUERY ANAGRAFICHE
'
'------------------------------------------------------------
Function CANCELLA_TUTTE_LeQueryAnagrafiche_pFunct()

On Error GoTo CANCELLA_TUTTE_LeQueryAnagrafiche_pFunct_Err
        
        
        
        'RESET
        '.................................................................................
            sxNomeDatabase = ""
            sxNomeQuery = ""
            sxParametroQuery = ""
            ixLungQuery = 0
        '.................................................................................
        
        'PARAMETRI GENERALI
        '.................................................................................
        'Il database per l'esportazione
            sxNomeDatabase = "c:\Casa\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MDB\PROVA_CANCELLAZIONE_DB_ESTERNO\MENU_TB03_OGGETTI_DA_CANCELLARE.mdb"
            
        '.................................................................................
        
        
        
        'ESPORTAZIONE 1
        '.................................................................................
        'Note   : Esporto tutte le query anagrafiche.
        
            'IMPOSTO LE VARIABILI DI ESPORTAZIONE
            sxParametroQuery = "CONTA_"
            ixLungQuery = 9
            
            'Nome generale delle query
            sxNomeQuery = sxNomeQuery & "Query ANAGRAFICHE : Codice " & sxParametroQuery
    
            'Il messaggio 1
            sxMessaggioBox = "Cancellazione 1 - <<DI TUTTE LE QUERY ANAGRAFICHE>> in GE_ANAGRAFICA effettuato. " _
                             & " Nro Query CANCELLATE : "
       
                
            'chiamo la funzione di Esportazione
            CANCELLO_QUERY_pFunct sxParametroQuery, sxMessaggioBox, sxNomeQuery, _
                                  ixLungQuery, sxNomeDatabase
        '.................................................................................

        'ESPORTAZIONE 2
        '.................................................................................
        'Note   : Esporto tutte le query parsametriche, necessarie per il funzionamento
        '       delle query anagrafiche.
        
            'IMPOSTO LE VARIABILI DI ESPORTAZIONE
            sxParametroQuery = "PARAM_VS01_"
            ixLungQuery = 11
            
            'Nome generale delle query
            sxNomeQuery = sxNomeQuery & "Query PARAMETRI : Codice " & sxParametroQuery
    
            
            'Il messaggio 2
            sxMessaggioBox = "Cancellazione 2 - <<DI TUTTE LE QUERY PARAMETRICHE>> in GE_ANAGRAFICA effettuato. " _
                             & " Nro Query CANCELLATE : "
       
            
            'chiamo la funzione di Esportazione
            CANCELLO_QUERY_pFunct sxParametroQuery, sxMessaggioBox, sxNomeQuery, _
                                  ixLungQuery, sxNomeDatabase
            
        '.................................................................................

        
        
        
        
CANCELLA_TUTTE_LeQueryAnagrafiche_pFunct_Exit:
    Exit Function

CANCELLA_TUTTE_LeQueryAnagrafiche_pFunct_Err:
    MsgBox Error$
    Resume CANCELLA_TUTTE_LeQueryAnagrafiche_pFunct_Exit

End Function










'------------------------------------------------------------
' CANCELLO LE QUERY ANAGRAFICHE
'------------------------------------------------------------
Public Function CANCELLO_QUERY_pFunct(par_sxParametroQuery As String, _
                                               par_sxMessaggioBox As String, _
                                               par_sxNomeQuery As String, _
                                               par_ixLungQuery As Integer, _
                                               par_sxNomeDatabase As String)

        


On Error GoTo CANCELLO_QUERY_pFunct_Err

    
    
        
    'CONTROLLO PREVENTIVO SULLE TABELLE CORRENTI
    '.................................................................................
    'Nota   : Controllo se le TABELLE da cancellare sono le tabelle correnti
    '       recuperando prima il nome del Database corrente, poi viene controllata
    '       la stringa del database per individuare il codice del database corrente
    '       ed infine, dopo aver estratto il Codice database corrente  viene confrontato
    '       con il parametro del database, se uguale viene esclusa la cancellazione con
    '       un messaggio e l'uscita dalla routine.

                                          
                  'SIGLA DATABASE
                  'Individuo le lettere iniziali del database.
                  sxNomeDatabaseCorrente = par_sxNomeDatabase
                       
                        MyPos = 0                                              'Reset valore della posizione
                        SearchString = ""
                        SearchString = sxNomeDatabaseCorrente                  ' Valore in cui eseguire la ricerca.
                          
                        SearchChar = par_sxParametroQuery                      ' Cerca codice stringa .
                           
                          
                        'INDIVIDUO LA POSIZIONE INZIALE DEL CODICE
                        ' Confronto binario a partire dalla posizione 0. Restituisce la posizione inziale del codice.
                        MyPos = InStr(1, SearchString, SearchChar, 0)
                        
                        
                            'Controllo MyPos
                            If MyPos = 0 Then
                                'Se MyPos = 0 significa che la funzione è inserita in un file
                                'diverso da quello originale. Esempio, se la funzione viene attivata per
                                'cancellare le tabelle Millesimi all'interno del file GIUR.mdb, la
                                'variabile sxNomeDatabaseCorrente viene valorizzata con il nome del db corrente
                                'c:\CASA\COND\01_GIUR\GIUR.MDB, le cui iniziali "GIUR" sono diverse dal
                                'codice delle tabelle da cancellare, "MILL" le tabelle Millesimi. Allora la
                                'if che ha valorizzato la variabile MyPos=0 deve reimpostare la variabile
                                'medesima al valore contenuto nella variabile parametro par_ixLungTable. In questo
                                'modo SE LA RICERCA DELLA FUNZIONE Instr da come risultato MyPos = 0,
                                'significa che le tabelle sono cancellabili, in quanto le stessi si trovano
                                'in un file diverso da quello originale. Viceversa la SE LA RICERCA DELLA FUNZIONE Instr
                                'da come risultato MyPos > 1, ciò significa, che ci troviamo nel file originale e quindi,
                                'le tabelle originale non possono essere cancellate, e la routine uscira nel controllo
                                'if <<CONFRONTO CODICE E PARAMETRO>> successivo.
                                
                                MyPos = ixLungQuery
                                '
                            End If
                        
                        
                          'RECUPERO IL CODICE
                          'Reindividuo il codice del database all'interno della stringa
                          sxCodiceDatabaseCorrente = Mid(CurrentDb.Name, MyPos, ixLungQuery)
                      
                          'CONFRONTO CODICE E PARAMETRO
                          'Se Codice database = parametro database
                          If sxCodiceDatabaseCorrente = par_sxParametroQuery Then
                              
                              MsgBox "Nel Database corrente è impossibile cancellare le " & sxNomeQuery & "!. Uscita dalla Routine", vbCritical
                              
                              GoTo CANCELLO_QUERY_pFunct_Exit
                        
                          End If
            
            
    '*** FINE ***
    'CONTROLLO PREVENTIVO SULLE TABELLE CORRENTI
    '.................................................................................


    'CONTROLLO DI ESECUZIONE
    '.................................................................................
        
        If MsgBox("Vuoi effettuare la cancellazione di " & par_sxNomeQuery & " ?", vbYesNo) = vbNo Then
            
            GoTo CANCELLO_QUERY_pFunct_Exit
        
        End If
    '.................................................................................
        

    ' TABELLE MILLESIMI COLLEGATE
    '.................................................................................
    ' COLLEGA LE TABELLE DAL DATABASE c:\Casa\COND\01_GIUR\GIUR.mdb
                
            'esempio funzionante
            '***********************************************************
            
              'DoCmd.TransferDatabase acLink, "Microsoft Access", _
                                   '"c:\Casa\COND\01_GIUR\GIUR.mdb", acTable, "ANAG_DF01_PARAMETRI_INDIV", _
                                   '"ANAG_DF01_PARAMETRI_INDIV", True
                                   
            '***********************************************************
            
    
    
    'OGGETTO docmd :1=TipoTrasferimento (acLink); 2=TipoDatabase ("Microsoft Access");
    '3=NomeDatabase ("c:\Casa\COND\01_GIUR\GIUR.mdb");  4=Tipo Oggetto (acTable);
    '5=Origine ("ANAG_DF01_PARAMETRI_INDIV"); 6= destinazione ("ANAG_DF01_PARAMETRI_INDIV"); 7= Solo Struttura (True)
    'docmd  Metodo..........1........2.................3..................................4.........5............................6............................7
    
    
            'reset
            iCount = 0
            Bolean1 = False

            sxParametroQuery = par_sxParametroQuery
            sxMessaggioBox = par_sxMessaggioBox
            ixLungQuery = par_ixLungQuery
        
        'Oggetto di origine
        sxQueryOrigine = "Nome da recuperare"
        
        'Nome della query di destinazione
        sxQueryDestinazione = "Nome da recuperare"
       
        
        sxTipoDatabase = "Microsoft Access"
        sxNomeDatabase = par_sxNomeDatabase
                          
        'Messaggio di fine salvataggio
        sxMessaggioBox = par_sxMessaggioBox
                            
        

    'INDIVIDUA LE QUERY DEL DB
    '------------------------------------------------------------------------------------
            
            Set dbs = CurrentDb
        
             ' Enumera l'insieme QueryDefs.
                For Each qdfCiclo In dbs.QueryDefs
                                
                    
                    sxNomeQuery = qdfCiclo.Name
                    sxCodiceQuery = Mid(qdfCiclo.Name, 1, ixLungQuery)
                                    
                    
                        '1) ESPORTO LA QUERY NEL DB DI DESTINAZIONE
                        '....................................................................
                                            
                                'Filtro di stampa
                                If sxCodiceQuery = sxParametroQuery Then
                                    Debug.Print iCount & ") La Query : " & sxNomeQuery & " Codice query : " & sxCodiceQuery
                                    Debug.Print
                                                
                        
                                                
                                        'Imposto l'oggetto di origine
                                        sxQueryOrigine = sxNomeQuery
                                                        
                                        'Imposto l'oggetto di destinazione
                                        sxQueryDestinazione = sxNomeQuery
                                    
                                        'CANCELLO LE QUERY
                                        DoCmd.DeleteObject acQuery, sxNomeQuery
                    
                                        'contatore
                                        iCount = iCount + 1
                                        Bolean1 = True
                                        
                                End If
                
                        '....................................................................
                        

                    
                Next qdfCiclo
                            
                                If Bolean1 = True Then
                                    
                                    'messaggio di salvataggio
                                    MsgBox (sxMessaggioBox & " " & iCount)
                            
                                Else
                                    MsgBox "Non sono estate effettuate Cancellazioni!", vbInformation
                                    
                                End If

        
        '------------------------------------------------------------------------------------

        
        


'.................................................................................
CANCELLO_QUERY_pFunct_Exit:
    Exit Function

CANCELLO_QUERY_pFunct_Err:
    MsgBox Error$
    Resume CANCELLO_QUERY_pFunct_Exit

End Function















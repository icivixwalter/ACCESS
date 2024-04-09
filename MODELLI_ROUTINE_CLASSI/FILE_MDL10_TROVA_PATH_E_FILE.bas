Attribute VB_Name = "FILE_MDL10_TROVA_PATH_E_FILE"
'/////////////////////////////////////////////////////////////////////////////////
'//Codice : OPTION.01
'//Note  : Le opzioni di scrittura

    '//OPZIONI
    '//........................................................
    Option Compare Text                     'Le Opzioni di comparazione testo
    Option Explicit                         'Le Opzioni esplicite per le variabili

    '//*** Fine ***
    '//OPZIONI
    '//........................................................
            


'//VARIABILI_DATABASE
'/////////////////////////////////////////////////////////////////////////////////
'//Codice : VARIABILI_DEGLI_INDIRIZZI.01.A
'
  '//LE_VARIABILI_GENERALI
  '//==============================================================================
  '//Codice   :VariabiliGenerali.01
  '//Note     :Le variabili Generali per la gestione dei Database Dao e Ado.

    '//DAO
    '//........................................................
    '//Codice : DimDbDao.01
    '//Note   :Le variabili del database Dao.
        
    Dim DaoDB As DAO.Database                   'Database Dao
    Dim DaoWks As DAO.Workspace
    Dim DaoRs As DAO.Recordset
    Dim DaoRs2 As DAO.Recordset
    
    
    '//*** Fine ***
    '//DAO
    '//........................................................
    
    
    '//ADO
    '//........................................................
    '//Codice : DimdDbAdo.01
    '//Note   :Le variabili del database Ado.
    
        Dim ADODB As Database                       'Database Ado
        Dim AdodaoRs As Recordset
    
        
    '//*** Fine ***
    '//ADO
    '//........................................................
    
    '//LE_VARIABILI_DATABASE_ESTERNO
    '//........................................................
    '//Codice : DimDatabaseEsterno.01
    '//Note   :Le variabili per la ricerca e l'apertura
    '       del database esterno.
    
    
        Dim sPathDbEsterno As String                        'Path del database esterno
        Dim sOption_SEZ As Integer                          'Numero Sezione di Progetto Scelta
        Dim sName_tab As String                             'Nome Tabella da aprire/creare
        Dim Apridbs As Database                             'Apri il database
        Dim intCicloDb As Integer                           'Intero per ciclo lettura oggetti del database
        Dim appAccess As Access.Application                 'Applicazione access
        Dim strDB As String                                 'Stringa per il db
        Dim strReportName As String                         'Stringa per il report.


    '//*** Fine ***
    '//DATABASE_ESTERNO
    '//........................................................
    


    '//LE_VARIABILI_COMUNI
    '//........................................................
    '//Codice : DimVariabiliComuni.01
    '//Note   :Le variabili comuni per la gestione del database
    '       con quelle che rappresentano i tipi visual basic.

    '//Variabili generali
    Dim Str1 As String
    Dim int1 As Integer
    Dim Lng1 As Long
    Dim Dbl1 As Double
    Dim Bln1 As Boolean
    Dim vV1 As Variant

    '//Le variabili di Connessione Al db.
    Dim sSql As String                                              'Stringa sql di estrazione
    Dim sSq2 As String                                              'Stringa sql di estrazione



    '//Contatori                                    'Contatore Integer
    Dim icount As Integer
    Dim dbl_count As Double                             'Contatore Double
    
    
    '//LE_VARIABILI_COMUNI *** FINE ***
    '//........................................................

    
    
    
    '//LE_VARIABILI_DI_GESTIONE_BOLLETTINI
    '//........................................................

    '//utilizzate per l'aggiornamento alla 2 cifra decimale
    Dim blnOrd01 As Double
    Dim blnOrd02 As Double
    Dim blnOrd03 As Double
    Dim blnOrd04 As Double
    Dim blnOrd05 As Double
    Dim blnOrd06 As Double
    Dim blnAff01 As Double
    Dim blnUn01 As Double
    Dim blnStr01 As Double
    Dim blnStr02 As Double
    Dim blnStr03 As Double
    Dim blnStr04 As Double
    Dim blnStr05 As Double
    Dim blnStr06 As Double
    
    '//LE_VARIABILI_DI_GESTIONE_BOLLETTINI  *** FINE ***
    '//........................................................

    
    '//LE_VARIABILI_PER_LA_RICERCA_STRINGA
    '//........................................................
    '//Codice : DimRicercaStringa.01
    '//Note   :Le variabili per la ricerca della stringa e
    '       la sua gestione.
    
    
    Dim SearchString  As String                 'Stringa da ricerca
    Dim SearchChar As String                    'Ricerca il carattere
    Dim MyPos As Integer                        'La posizione.
    Dim MyLen As Integer                                        'La lughezza della stringa
    Dim sStringaIniz As String                                  'Stringa fino all'apostrofo
    Dim MyLenIniz As Integer                                    'La lughezza della stringa Iniziale
    Dim sStringaFin As String                                   'Stringa finale senza apostrofo
    Dim MyLenFin As Integer                                     'La lughezza della stringa Iniziale
    Dim MyLenDiff As Integer                                    'La lughezza rimanente tra (Stringa Iniziale - Stringa Finale = Diff)
    Dim sStringaRicostr As String                               'Stringa ricostruita
    
    '//***Fine***
    '//LE_VARIABILI_PER_LA_RICERCA_STRINGA
    '//........................................................
        
    
    '//LE_VARIABILI_PROCEDURE_ERRORE
    '//........................................................
    '//Codice : DimProcedureErrore.01
    '//Note   :Le variabili per la gestione degli erori e dei
    '       messaggi della procedura.
    
    Dim sxProceduraMessaggioErrore As String            'Messaggio dei errore della procedura
    Dim sxProceduraAttivaEseguita  As String            'Ultima procedura eseguita nell'errore
    
    
    '//***Fine***
    '//LE_VARIABILI_PROCEDURE_ERRORE
    '//........................................................
    
    
    

    '//LE_VARIABILI_DI_GESTIONE_FILE
    '//........................................................

    
        '//LE VARIABILI
        '//La Path, il File da ricercare
        Dim MyPath_s As String
        Dim MyFile_s As String
        Dim MyName_s As String
        Dim IDGestione_lng As Long
        
    
    '//***Fine***
    '//LE_VARIABILI_DI_GESTIONE_FILE
    '//........................................................


    
  '//*** Fine ***
  '//LE_VARIABILI_GENERALI
  '//==============================================================================
    

'//*** Fine ***
'//VARIABILI_DATABASE
'/////////////////////////////////////////////////////////////////////////////////



'//ATTIVO LA FUNZIONE
Private Sub ATTIVA_TROVA_N01_PATH_pFunct()
    MyPath_s = "c:\GESTIONI\GESTIONE_LLPP\02_SCANNER\ScannerTmp\"     ' Imposta il percorso.
    
    TROVA_N01_PATH_pFunct MyPath_s

End Sub

'//TROVA_N01_PATH_pFunct
'//=====================================================================================//
'//Note         :Funzione Trova path
'//par_Path_s   :Parametro in entrata la path da ricercare
'//RESTITUISCE  :Il nome della path individuata - STringa -
Public Function TROVA_N01_PATH_pFunct(par_Path_s As String) As String
    
    '//TROVA_PATH_SU_DISCO  (INIZIO)
    '//------------------------------------------------------------------------------------------------
        ' Visualizza i nomi in c:\ che rappresentano directory.
        MyName_s = Dir(par_Path_s, vbDirectory)                               ' Recupera la prima voce.
        
        
        Debug.Print "                       CONTROLLO PATH                              "
        Debug.Print "..................................................................."
        Debug.Print "Path -> " & par_Path_s
        
        Do While MyName_s <> ""    ' Avvia il ciclo.
            
            'STAMPA SOLO LA DIRECTORY
            ' Ignora la directory corrente e quella di livello superiore.
            If MyName_s <> "." And MyName_s <> ".." Then
                ' Usa il confronto bit per bit per verificare se MyName_s è una directory.
                If (GetAttr(par_Path_s & MyName_s) And vbDirectory) = vbDirectory Then
                    Debug.Print MyName_s      ' Visualizza la voce solo
                End If                      ' se rappresenta una directory.
            End If
            MyName_s = Dir    ' Legge la voce successiva.
        Loop
        
        Debug.Print "..................................................................."
        Debug.Print
        Debug.Print " CONTROLLO PATH DI RICERCA                                         "
        Debug.Print
        Debug.Print "Path -> " & par_Path_s
        
         
        MsgBox "PATH_TROVATA--->" & par_Path_s, vbInformation, "MSG_BOX_DI_AVVISO"
        
    
    '//TROVA_PATH_SU_DISCO  (***FINE***)
    '//------------------------------------------------------------------------------------------------
End Function
'//TROVA_N01_PATH_pFunct
'//=====================================================================================//




'//ATTIVO LA FUNZIONE
Private Sub ATTIVA_TROVA_N02_FILE_SELEZIONATO_pFunct()
    
    IDGestione_lng = 65
    MyPath_s = "c:\GESTIONI\GESTIONE_LLPP\02_SCANNER\ScannerTmp\"     ' Imposta il percorso.
    MyFile_s = "Folium_9406_2015.pdf"                                 ' Recupera la prima voce.
   
    TROVA_N02_FILE_SELEZIONATO_pFunct MyFile_s, IDGestione_lng

End Sub


'//TROVA_N02_FILE_SELEZIONATO_pFunct
'//=====================================================================================//
'//Note         :Funzione Trova path
'//par_Path_s   :Parametro in entrata la path da ricercare
'//RESTITUISCE  :Il nome della path individuata - STringa -
Public Function TROVA_N02_FILE_SELEZIONATO_pFunct(par_MyFile_s As String, par_IDGestione_lng As Long) As String

Dim MyFile_s_PERCORSO_s As String

'//TROVA FILE  (INIZIO)
'//------------------------------------------------------------------------------------------------
    ' Visualizza i nomi in c:\ che rappresentano directory.
    MyPath_s = "c:\GESTIONI\GESTIONE_LLPP\02_SCANNER\ScannerTmp\"     ' Imposta il percorso.
   
    MyFile_s_PERCORSO_s = MyPath_s
    
    Debug.Print "                       CONTROLLO FILE TROVATO                      "
    Debug.Print "..................................................................."
    Debug.Print "Path -> " & MyPath_s
            
            MyFile_s = Dir(MyPath_s & par_MyFile_s & ".*", vbNormal)
            If MyFile_s > "" Then
                Debug.Print par_MyFile_s              ' Visualizza la voce solo
                
                MsgBox "OK_CONTROLLO_FILE=TROVATO--->" & Chr$(13) & MyPath_s & Chr$(13) & MyFile_s, vbInformation, "MSG_BOX_DI_AVVISO"
                
                '//Controllo ed esecuzione cmd_sSql
                sSql = ""
                sSql = sSql & "UPDATE LLPP_ATTI_Tb01_Gestione SET LLPP_ATTI_Tb01_Gestione.RICERCA_FileAtto_s = 'FILE TROVATO'"
                sSql = sSql & "WHERE (((LLPP_ATTI_Tb01_Gestione.IDGestione)=" & par_IDGestione_lng & "));"
                Debug.Print sSql
                CurrentDb.Execute sSql
                
            Else
                Debug.Print "FILE NON TROVATO"
                MsgBox "ATTENZIONE!!! FILE_NON_TROVATO--->" & Chr$(13) & MyPath_s & "/" & par_MyFile_s, vbInformation, "MSG_BOX_DI_AVVISO"
            End If
    
    Debug.Print "..................................................................."
    
   ' MsgBox "FILE_TROVATO--->" & MyPath_s & "/" & par_MyFile_s, vbInformation, "MSG_BOX_DI_AVVISO"
     

'//TROVA FILE  (***FINE***)
'//------------------------------------------------------------------------------------------------


End Function
'//TROVA_N02_FILE_SELEZIONATO_pFunct
'//=====================================================================================//


'//ISTRUZIONE
'Funzione Dir
'Vedere anche     Esempio     Informazioni aggiuntive
'Restituisce un valore String che rappresenta un nome di file, directory, o cartella che corrisponde a un attributo
'o tipo di file specificato o a un'etichetta di volume di un'unità.

'Sintassi
'Dir[(nomepercorso[, attributi])]
'La sintassi della funzione Dir è composta dalle seguenti parti:

'Parte Descrizione

'nomepercorso Facoltativa.
'Espressione stringa che specifica il nome del file. Può includere la directory o cartella e unità.
'Se nomepercorso non viene trovato, la funzione restituisce una stringa di lunghezza zero ("").

'attributi Facoltativa.
'Costante o espressione numerica, la cui somma specifica gli attributi del file.
'Se omessa, vengono restituiti i file che corrispondono a nomepercorso ma che non hanno attributi

'Impostazioni
'Le possibili impostazioni dell'argomento attributi sono:
'Costante       Valore          Descrizione
'vbNormal       0           (Predefinita). Specifica i file senza attributi.
'vbReadOnly         1           Specifica i file di sola lettura oltre ai file senza attributi.
'vbHidden       2           Specifica i file nascosti oltre ai file senza attributi.
'vbSystem       4           Specifica i file di sistema oltre ai file senza attributi. Non disponibile in Macintosh.
'vbVolume       8           Specifica l'etichetta di volume. Viene ignorata se si specifica qualsiasi altro attributo. Non disponibile in Macintosh.
'vbDirectory        16          Specifica le directory o le cartelle oltre ai file senza attributi.
'vbAlias        64          Il nome di file specificato è un alias. Disponibile solo in Macintosh.


'Osservazioni
'In Microsoft Windows Dir supporta i caratteri jolly per più caratteri (*)
'e per singoli caratteri (?) per indicare più file. In Macintosh,
'tali caratteri sono considerati caratteri validi per i nomi di file e non
'possono essere utilizzati come caratteri jolly per indicare più file.

'Poiché in Macintosh i caratteri jolly non sono supportati, è possibile utilizzare il tipo di file
'per identificare gruppi di file. Utilizzare la funzione MacID per specificare il tipo di file
''anziché utilizzare i nomi di file. Ad esempio, la seguente istruzione restituisce il nome del primo
'file di tipo TEXT nella cartella corrente:
'Dir("Percorso", MacID("TEXT"))

'Per eseguire un'iterazione in tutti i file di una cartella, specificare una stringa vuota:
'Dir("")

'Se la funzione MacID viene utilizzata con Dir in Microsoft Windows, viene generato un errore.
'Se all'argomento attributi viene assegnato un valore maggiore di 256, verrà considerato come valore MacID.
'È necessario specificare nomepercorso alla prima chiamata della funzione Dir, altrimenti verrà generato un errore.
'Se si specificano gli attributi del file occorre includere anche nomepercorso.

'Dir restituisce il primo nome di file che corrisponde a quello specificato in nomepercorso.
'Per ottenere i successivi nomi di file corrispondenti a nomepercorso, chiamare di nuovo la
'funzione Dir senza alcun argomento. Se non vengono trovati altri nomi di file corrispondenti,
''Dir restituirà una stringa di lunghezza zero, dopodiché sarà necessario utilizzare di nuovo nomepercorso
'nelle successive chiamate, altrimenti verrà generato un errore.
'È possibile passare a un nuovo nomepercorso senza trovare tutti i nomi di file che corrispondono al nomepercorso corrente.
'Non è possibile tuttavia chiamare la funzione Dir in modo ricorsivo.
'Richiamando Dir con l'attributo vbDirectory non verranno restituite sottodirectory in modo continuo.

'Suggerimento
'Dato che i nomi di file vengono individuati senza rispettare un ordine particolare,
'potrebbe essere utile salvarli in una matrice e quindi ordinarla.










'//ALTRI ESEMPI DA VERIFICARE

'Esempio di funzione Dir
'In questo esempio la funzione Dir viene utilizzata per controllare se esistono
'determinati file e directory.In Macintosh, il nome dell'unità predefinita
'è "HD:" e le parti del percorso sono separate da due punti anziché da una barra rovesciata.
'Inoltre, i caratteri jolly di Microsoft Windows sono considerati come caratteri validi per i nomi di file.
'È tuttavia possibile utilizzare la funzione MacID per specificare gruppi di file.

'Dim MyFile_s, MyPath_s, MyName_s
' Restituisce "WIN.INI" (in Microsoft Windows)se esiste.
'MyFile_s = Dir("C:\WINDOWS\WIN.INI")

' Restituisce il nome dei file con l'estensione
' specificata. Se esistono più file con estensione.ini,
' viene restituito il nome del primo file.
'MyFile_s = Dir("C:\WINDOWS\*.INI")

' Richiama l'istruzione Dir senza argomenti per
' restituire il successivo file con estensione ini
' contenuto nella stessa directory.
'MyFile_s = Dir


' Restituisce il primo file con estensione txt
' impostato come nascosto.
'MyFile_s = Dir("*.TXT", vbHidden)





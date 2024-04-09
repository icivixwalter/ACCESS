Attribute VB_Name = "DIRECTORY_Mdl02_01_CONTROLLO_ESISTENZA"
Option Compare Database
Option Explicit

'//++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++//
'//       LE VARIABILI DI MODULO

'//LE VARIABILI DATABASE
'//....................................................................//
    Dim DaoDB As DAO.Database
    Dim DaoWks As DAO.Workspace
    Dim DaoRs As DAO.Recordset

    Dim ADODB As Database
    Dim AdodaoRs As Recordset
    Dim sSql As String                          '//STRINGA SQL
    Dim Path_s As String                        '//la path


    '//Contatori
    Dim iCount As Integer
    Dim dbl_count As Double

    'Le variabili generiche
    Dim Vv1 As Variant
    Dim Dbl1 As Double
    Dim Int1 As Integer
    Dim Long1 As Long
    Dim Str1 As Long

'....................................................................

'//++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++//


Private Sub CHIAMA_DIR()
'//IMPOSTAZIONE PATH E CONTROLLO ESISTENZA DIRECTORY
'//------------------------------------------------------------------------------//
'//NOTE     : controllo l'esistenza della path definita dai salvataggi se non esiste _
            esco dalla routine.
    Path_s = "c:\CASA\GE_CASA\GE_MARINO\GESTIONE_SPESE\ARCHIVI_XLS\"
    
    Vv1 = DIRECTORY_CONTROLLO_N01_EsistenzaDirectory_pFunct(1, Path_s)
End Sub

'//FUNZIONE------------------->DIRECTORY_CONTROLLO_N01_EsistenzaDirectory_pFunct
'//========================================================================================================================================//
'//Tipo           : Funzione pubblica.
'//Attività       : Controllo sull'esistenza della directory
'//Note           : Individua la directory passata con parametro
'//Parametro      : par_TipoParametro_i = tipo di file o directory vedi specifiche, _
                    par_Directory_s = è la path o la directory
'//Restituisce    : Null
'//Codice         : DIRECTORY_CONTROLLO_N01_EsistenzaDirectory_pFunct.01
'//

Public Function DIRECTORY_CONTROLLO_N01_EsistenzaDirectory_pFunct(par_TipoParametro_i As Integer, _
                                                                  par_Directory_s As String)

'//MessaggiDiErrore
Dim ProceduraMessaggioErrore_s As String
Dim ProceduraAttivaEseguita_s As String
Dim ParametroFile_i As Integer
 
'//Campo
Dim CampoCercato_s As String

'//Campi parametri
Dim par_AnnoImp_i As Integer
Dim par_CodiceTributo_s As String

            
    '//....
On Error GoTo Err_DIRECTORY_CONTROLLO_N01_EsistenzaDirectory_pFunct


        
        '//Imposto i parametri
        ProceduraAttivaEseguita_s = "DIRECTORY_CONTROLLO_N01_EsistenzaDirectory_pFunct"
        ProceduraMessaggioErrore_s = "Errore nella procedura"
        
    '//DIRECTORY_CONTROLLO
    '//.....................................................................................................//
    '//Note           : Tramite una Select vengono individuati i valori da restiuire.

            
            '//IMPOSTAZIONE PATH E CONTROLLO ESISTENZA DIRECTORY
            '//------------------------------------------------------------------------------//
            '//NOTE     : controllo l'esistenza della path definita dai salvataggi se non esiste _
                        esco dalla routine.
                
                '//VALORIZZO I PARAMETRI
                Path_s = par_Directory_s
                ParametroFile_i = par_TipoParametro_i
                
                Dim MyPath, MYNAME As Variant
                'Str1 = Dir(Path_s, 16)
                'Vv1 = Dir("*.TXT", 2)
                'MyPath = "c:\CASA\GE_CASA\GE_MARINO\GESTIONE_SPESE\ARCHIVI_XLS\"    ' Imposta il percorso.
                'MYNAME = Dir(MyPath, vbDirectory)    ' Recupera la prima voce.
                'MYNAME = Dir(par_Directory_s, vbDirectory)    ' Recupera la prima voce.
                Vv1 = Dir(par_Directory_s, vbDirectory)   ' Recupera la prima voce.
                
               ' Vv1 = Dir("c:\", vbDirectory)
                
                If Vv1 = "" Then
                        MsgBox "NON ESISTE LA DIRECTORY ---> " & Path_s & " - USCITA DALLA ROUTINE"
                        GoTo Exit_DIRECTORY_CONTROLLO_N01_EsistenzaDirectory_pFunct
                End If
            '//-------------------------------------------------------------------------------//


          
    '//*** fine ***
    '//DIRECTORY_CONTROLLO
    '//.....................................................................................................//

'//USCITA  E GESTIONE ERRORI
'//.....................................................................................................//.........


Exit_DIRECTORY_CONTROLLO_N01_EsistenzaDirectory_pFunct:
    Exit Function

Err_DIRECTORY_CONTROLLO_N01_EsistenzaDirectory_pFunct:
 '//-------------------------------------------------------------------------------
    MsgBox Err.Description & " - Errore Messaggio -> : " & ProceduraMessaggioErrore_s & " Procedura -> : " & ProceduraMessaggioErrore_s
    Debug.Print ProceduraMessaggioErrore_s
    Debug.Print ProceduraAttivaEseguita_s
    Stop
    Resume Exit_DIRECTORY_CONTROLLO_N01_EsistenzaDirectory_pFunct
'//-------------------------------------------------------------------------------
        
End Function

'//*** FINE ***
'//FUNZIONE------------------->DIRECTORY_CONTROLLO_N01_EsistenzaDirectory_pFunct
'//========================================================================================================================================//






'//NOTE TECNICHE
'//*******************************************************************************************************//
'                                   FUNZIONE DIR  NOTE TECNICHE

'//------------------------------------------------------------------------------------------------------//
'Restituisce un valore String che rappresenta un nome di file, directory, o
'cartella che corrisponde a un attributo o tipo di file specificato o a _
 un'etichetta di volume di un'unità.
'Sintassi
'Dir[(nomepercorso[, attributi])]
'La sintassi della funzione Dir è composta dalle seguenti parti:

'Parte Descrizione
'------>: nomepercorso Facoltativa.
'Espressione stringa che specifica il nome del file.
'Può includere la directory o cartella e unità. Se nomepercorso non viene trovato,
'la funzione restituisce una stringa di lunghezza zero ("").

'------>:attributi Facoltativa.
'Costante o espressione numerica, la cui somma specifica gli
'attributi del file. Se omessa, vengono restituiti i file che corrispondono a nomepercorso ma che non hanno attributi

'Impostazioni
'Le possibili impostazioni dell'argomento attributi sono:

'Costante   Valore  Descrizione
'vbNormal   0   (Predefinita). Specifica i file senza attributi.
'vbReadOnly 1   Specifica i file di sola lettura oltre ai file senza attributi.
'vbHidden   2   Specifica i file nascosti oltre ai file senza attributi.
'vbSystem   4   Specifica i file di sistema oltre ai file senza attributi. Non disponibile in _
                Macintosh.
'vbVolume   8   Specifica l'etichetta di volume. Viene ignorata se si specifica qualsiasi altro attributo. Non disponibile in Macintosh.
'vbDirectory    16  Specifica le directory o le cartelle oltre ai file senza attributi.
'vbAlias    64  Il nome di file specificato è un alias. Disponibile solo in Macintosh.



'Nota   Queste costanti vengono specificate da Visual Basic, Applications Edition e possono _
 essere utilizzate nel codice in sostituzione dei valori effettivi.
'Osservazioni

'In Microsoft Windows Dir supporta i caratteri jolly per più caratteri (*) e per singoli _
 caratteri (?) per indicare più file. In Macintosh, tali caratteri sono considerati caratteri _
 validi per i nomi di file e non possono essere utilizzati come caratteri jolly per indicare _
 più file.
'Poiché in Macintosh i caratteri jolly non sono supportati, è possibile utilizzare il tipo di _
 file per identificare gruppi di file. Utilizzare la funzione MacID per specificare il tipo di _
 file anziché utilizzare i nomi di file. Ad esempio, la seguente istruzione restituisce il nome _
 del primo file di tipo TEXT nella cartella corrente:

'Dir("Percorso", MacID("TEXT"))

'Per eseguire un'iterazione in tutti i file di una cartella, specificare una stringa vuota:
'Dir("")

'Se la funzione MacID viene utilizzata con Dir in Microsoft Windows, viene generato un errore.
'Se all'argomento attributi viene assegnato un valore maggiore di 256, verrà _
 considerato come valore MacID.
'È necessario specificare nomepercorso alla prima chiamata della funzione Dir, _
 altrimenti verrà generato un errore. Se si specificano gli attributi del file occorre _
 includere anche nomepercorso.
'Dir restituisce il primo nome di file che corrisponde a quello specificato in nomepercorso. _
 Per ottenere i successivi nomi di file corrispondenti a nomepercorso, chiamare di nuovo _
 la funzione Dir senza alcun argomento. Se non vengono trovati altri nomi di file corrispondenti, _
 Dir restituirà una stringa di lunghezza zero, dopodiché sarà necessario utilizzare di nuovo _
 nomepercorso nelle successive chiamate, altrimenti verrà generato un errore. È possibile passare _
 a un nuovo nomepercorso senza trovare tutti i nomi di file che corrispondono al nomepercorso _
 corrente. Non è possibile tuttavia chiamare la funzione Dir in modo ricorsivo. Richiamando Dir _
 con l'attributo vbDirectory non verranno restituite sottodirectory in modo continuo.
'Suggerimento   Dato che i nomi di file vengono individuati senza rispettare un ordine _
 particolare, potrebbe essere utile salvarli in una matrice e quindi ordinarla.



'Esempio di funzione Dir
'//==========================================================================================================//
'In questo esempio la funzione Dir viene utilizzata per controllare se esistono determinati file e directory.In Macintosh, il nome dell'unità predefinita è "HD:" e le parti del percorso sono separate da due punti anziché da una barra rovesciata. Inoltre, i caratteri jolly di Microsoft Windows sono considerati come caratteri validi per i nomi di file. È tuttavia possibile utilizzare la funzione MacID per specificare gruppi di file.

'Dim MyFile, MyPath, MyName
' Restituisce "WIN.INI" (in Microsoft Windows)se esiste.
'MyFile = Dir("C:\WINDOWS\WIN.INI")

' Restituisce il nome dei file con l'estensione
' specificata. Se esistono più file con estensione.ini,
' viene restituito il nome del primo file.
'MyFile = Dir("C:\WINDOWS\*.INI")

' Richiama l'istruzione Dir senza argomenti per
' restituire il successivo file con estensione ini
' contenuto nella stessa directory.
'MyFile = Dir

' Restituisce il primo file con estensione txt
' impostato come nascosto.
'MyFile = Dir("*.TXT", vbHidden) ATTENZIONE AL POSTO DI vbHidden INSERIRE LA VARIABILE NUMERICA ES. 2

' Visualizza i nomi in c:\ che rappresentano directory.
'MyPath = "c:\"    ' Imposta il percorso.
'MyName = Dir(MyPath, vbDirectory)    ' Recupera la prima voce.
'Do While MyName <> ""    ' Avvia il ciclo.
    ' Ignora la directory corrente e quella di livello superiore.
'    If MyName <> "." And MyName <> ".." Then
        ' Usa il confronto bit per bit per verificare se MyName è una directory.
'        If (GetAttr(MyPath & MyName) And vbDirectory) = vbDirectory Then
'            Debug.Print MyName    ' Visualizza la voce solo
'        End If    ' se rappresenta una directory.
'    End If
'    MyName = Dir    ' Legge la voce successiva.
'Loop

'//*******************************************************************************************************//






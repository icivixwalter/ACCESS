Attribute VB_Name = "FUNZIONI_Mdl10_FILE_(TROVA_FILE)"
Option Compare Database

Public Sub TrovaSOLO_DIRECTORY_pSub()
    
'//TROVA FILE  (INIZIO)
'//------------------------------------------------------------------------------------------------
    ' Visualizza i nomi in c:\ che rappresentano directory.
    MyPath = "c:\GESTIONI\GESTIONE_LLPP\02_SCANNER\ScannerTmp\"     ' Imposta il percorso.
    MYNAME = Dir(MyPath, vbDirectory)                               ' Recupera la prima voce.
    
    
    Debug.Print "                       CONTROLLO PATH                              "
    Debug.Print "..................................................................."
    Debug.Print "Path -> " & MyPath
    
    Do While MYNAME <> ""    ' Avvia il ciclo.
        
        'STAMPA SOLO LA DIRECTORY
        ' Ignora la directory corrente e quella di livello superiore.
        If MYNAME <> "." And MYNAME <> ".." Then
            ' Usa il confronto bit per bit per verificare se MyName è una directory.
            If (GetAttr(MyPath & MYNAME) And vbDirectory) = vbDirectory Then
                Debug.Print MYNAME      ' Visualizza la voce solo
            End If                      ' se rappresenta una directory.
        End If
        MYNAME = Dir    ' Legge la voce successiva.
    Loop
    
    Debug.Print "..................................................................."
    Debug.Print
    Debug.Print " CONTROLLO PATH DI RICERCA                                         "
    Debug.Print
    Debug.Print "Path -> " & MyPath
    
     

'//TROVA FILE  (***FINE***)
'//------------------------------------------------------------------------------------------------


End Sub


Public Sub TrovaSOLO_FILE_pSub()
Dim iCount As Integer

'//RESET VARIABILI
iCount = 0

'//TROVA FILE  (INIZIO)
'//------------------------------------------------------------------------------------------------
    ' Visualizza i nomi in c:\ che rappresentano directory.
    MyPath = "c:\GESTIONI\GESTIONE_LLPP\02_SCANNER\ScannerTmp\"     ' Imposta il percorso.
    MYNAME = Dir(MyPath, vbNormal)                      ' Recupera la prima voce.
    
    Debug.Print "                       CONTROLLO PATH                              "
    Debug.Print "..................................................................."
    Debug.Print "Path -> " & MyPath
    
    Do While MYNAME <> ""    ' Avvia il ciclo.
        
        'STAMPA SOLO LA DIRECTORY
        ' Ignora la directory corrente e quella di livello superiore.
        If MYNAME <> "." And MYNAME <> ".." Then
            ' Usa il confronto bit per bit per verificare se MyName è una directory.
            If (GetAttr(MyPath & MYNAME) And vbNormal) = vbNormal Then
                Debug.Print MYNAME              '//Visualizza la voce solo se rappresenta una un file.
                iCount = iCount + 1             '//conta i file
            End If
            
        End If
        MYNAME = Dir    ' Legge la voce successiva.
    Loop
    
        Debug.Print "..................................................................."
        Debug.Print "TOTALE FILE INDIVIDUATI---> " & iCount
        Debug.Print "PATH DI RICERCA-----------> " & MyPath
     

'//TROVA FILE  (***FINE***)
'//------------------------------------------------------------------------------------------------


End Sub


Public Sub TrovaSOLO_FILE_SELEZIONATO_pSub()
    
'//TROVA FILE  (INIZIO)
'//------------------------------------------------------------------------------------------------
    ' Visualizza i nomi in c:\ che rappresentano directory.
    MyPath = "c:\GESTIONI\GESTIONE_LLPP\02_SCANNER\ScannerTmp\"     ' Imposta il percorso.
    MYNAME = "Folium_7885_2015.pdf"                                 ' Recupera la prima voce.
    
    
    
    Debug.Print "                       CONTROLLO FILE TROVATO                      "
    Debug.Print "..................................................................."
    Debug.Print "Path -> " & MyPath
            
            MyFile = Dir(MyPath & MYNAME, vbNormal)
            If MyFile > "" Then
                Debug.Print MYNAME              ' Visualizza la voce solo
            Else
                Debug.Print "FILE NON TROVATO"
            End If
    
    Debug.Print "..................................................................."
        Debug.Print
     

'//TROVA FILE  (***FINE***)
'//------------------------------------------------------------------------------------------------


End Sub


'//ALTRI ESEMPI DA VERIFICARE

'Esempio di funzione Dir
'In questo esempio la funzione Dir viene utilizzata per controllare se esistono
'determinati file e directory.In Macintosh, il nome dell'unità predefinita
'è "HD:" e le parti del percorso sono separate da due punti anziché da una barra rovesciata.
'Inoltre, i caratteri jolly di Microsoft Windows sono considerati come caratteri validi per i nomi di file.
'È tuttavia possibile utilizzare la funzione MacID per specificare gruppi di file.

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
'MyFile = Dir("*.TXT", vbHidden)




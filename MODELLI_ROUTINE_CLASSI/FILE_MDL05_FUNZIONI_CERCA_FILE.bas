Attribute VB_Name = "FILE_MDL05_FUNZIONI_CERCA_FILE"
Option Compare Database

'//FUNZIONE STRING$
'//@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@//
    '//Funzione String _
    Restituisce un valore Variant (String) che contiene una stringa di caratteri _
    ripetuti della lunghezza specificata. _
    sintassi _
    String(number, character)
    
    '//La sintassi della funzione String è composta dai seguenti argomenti predefiniti: _
    Parte Descrizione _
    number Obbligatoria. Long. Lunghezza della stringa restituita. Se number contiene Null verrà restituito Null. _
    character Obbligatoria. Variant. Codice di carattere che specifica il carattere o _
    l 'espressione stringa il cui primo carattere viene utilizzato per costruire la stringa restituita. _
    Se character contiene Null verrà restituito Null.
    
    '//Osservazioni _
    Se per character viene specificato un numero maggiore di 255, _
    la funzione String converte il numero in un codice di carattere valido tramite la formula: _
    character Mod 256

'//FUNZIONE STRING$     *** FINE ***
'//@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@//


 '//OPZIONI
    '//........................................................
   ' Option Compare Text                     'Le Opzioni di comparazione testo
    Option Explicit                         'Le Opzioni esplicite per le variabili

    '//*** Fine ***
    '//OPZIONI
    '//........................................................
            
'//DICHIARAZIONE DELLA LIBRERIA DLL
Private Declare Function apiSearchTreeForFile Lib "ImageHlp.dll" Alias _
        "SearchTreeForFile" (ByVal lpRoot As String, ByVal lpInPath As String, _
                             ByVal lpOutPath As String) As Long


'//PROVA FUNZIONE
Private Sub CHIAMA_fReturnFilePath()
Dim Str1 As String
    
    '//Funzione CERCA FILE
    '//...................................................................//
    '//ATTIVITA         : Ricerco il file passato con parametro e restituisco il percorso completo _
                        unitamente al nome sole se la ricerca nel percorso path è positiva, altrimenti _
                        viene restituito il valore NULL con l'avviso che il file non è stato individuato.
    '//CODICE           :fSearchFile_PFunct.01.CHIAMA
        Str1 = fSearchFile_PFunct("GE_CASA_TB90_SALVATAGGI_ARCHIVI.mdb", _
                                  "c:\CASA\GE_CASA\GE_MARINO\GESTIONE_SPESE\GESTIONE_PROCEDURE\GE_CASA_MDB\")
    '//...................................................................//

End Sub

'//Funzione CERCA FILE
'//****************************************************************************************************//
'//NRO_FUNZIONE     : 112
'//ATTIVITA         : Ricerco il file passato con parametro e restituisco il percorso completo _
                      unitamente al nome sole se la ricerca nel percorso path è positiva, altrimenti _
                      viene restituito il valore NULL con l'avviso che il file non è stato individuato.
'//Parametri        : Per riferime nto _
                    1) strFilename = NOME FILE _
                    2) strSearchPath = PATH FILE per la ricerca.
'//RESTITUISCE      : il nome ed il percorso completo del file solo se esiste, altrimenta da un avviso _
                    mediante messaggio del file non trovato.
'//CODICE           :fSearchFile_PFunct.01
Function fSearchFile_PFunct(ByVal par_strFilename As String, _
                            ByVal par_strSearchPath As String) As String
    
    
    
    'Returns the first match found
    Dim lpBuffer_s As String                            '//la lunghezza del Buffer   = stringa di 1024 byte
    Dim lngResult As Long                               '//Risultato della ricerca 1 = trovato 0 = non trovato
    
    '//RESET
    fSearchFile_PFunct = ""
    lpBuffer_s = String$(1024, 0)                       '//String$ = Vedi sopra spiegazione della funzione.
    
    
    
    '//CERCO IL FILE NELLA DIRECTORY PASSATA CON I PARAMETRI
    '//===============================================================================//
    '//CODICE           :fSearchFile_PFunct.01.02
    '//Restituisce un valore >0 = File Trovato; se = 0 FILE NON TROVATO
        lngResult = apiSearchTreeForFile(par_strSearchPath, par_strFilename, lpBuffer_s)
        If lngResult <> 0 Then
            If InStr(lpBuffer_s, vbNullChar) > 0 Then
                
                '//CONTROLLO FILE
                '//Restituisco il nome del file al chiamante.
                fSearchFile_PFunct = Left$(lpBuffer_s, InStr(lpBuffer_s, vbNullChar) - 1)
                Debug.Print
                Debug.Print "==========================================================="
                Debug.Print "           file TROVATO :                                  "
                Debug.Print fSearchFile_PFunct
                Debug.Print
                Debug.Print "==========================================================="
                            
            End If
            
         Else
                '//SE IL FILE NON E' STATO TROVATO
                '//Restituisco il valore nullo se la ricerca è FALLITA.
                fSearchFile_PFunct = ""
                MsgBox "MSG_112_ATTENZIONE, il File -> " & Chr$(13) & par_strFilename _
                       & Chr$(13) & " nella path -> " & par_strSearchPath & Chr$(13) _
                       & " NON E' TROVATO!", vbExclamation
        End If
        
    '//CERCO IL FILE NELLA DIRECTORY PASSATA CON I PARAMETRI *** FINE ***
    '//===============================================================================//
    
    
End Function
'//Funzione CERCA FILE      *** fine ***
'//****************************************************************************************************//







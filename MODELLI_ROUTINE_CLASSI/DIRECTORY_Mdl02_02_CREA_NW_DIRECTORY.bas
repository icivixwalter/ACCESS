Attribute VB_Name = "DIRECTORY_Mdl02_02_CREA_NW_DIRECTORY"
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

'Ciao mptd,
'per creare una directory:
'    MkDir ("C:\TuaCartella\" & NomeCampo)
'per verificare se esiste:
'If Len(Dir("C:\TuaCartella\" & NomeCampo, vbDirectory)) > 0 Then
'     MsgBox ("C:\TuaCartella\" & NomeCampo & " esiste già!")
'End If



'DIRECOTRY_CREAZIONE_N01_CreoNuovaDirectory_pFunct



Private Sub CHIAMA_DIR()
'//IMPOSTAZIONE PATH E CONTROLLO ESISTENZA DIRECTORY
'//------------------------------------------------------------------------------//
'//NOTE     : controllo l'esistenza della path definita dai salvataggi se non esiste _
            esco dalla routine.
    Path_s = "c:\CASA\GE_CASA\GE_MARINO\GESTIONE_SPESE\ARCHIVI_XLS\"
    
    Vv1 = DIRECOTRY_CREAZIONE_N01_CreoNuovaDirectory_pFunct(1, Path_s)
End Sub

'//FUNZIONE------------------->DIRECOTRY_CREAZIONE_N01_CreoNuovaDirectory_pFunct
'//========================================================================================================================================//
'//Tipo           : Funzione pubblica.
'//Attività       : Controllo prima dell'esistenza della directory e successivamente se non esiste creo _
                    una nuova directory
'//Note           : Individua la directory passata con parametro
'//Parametro      : par_TipoParametro_i = tipo di file o directory vedi specifiche, _
                    par_Directory_s = è la path o la directory
'//Restituisce    : Null
'//Codice         : DIRECOTRY_CREAZIONE_N01_CreoNuovaDirectory_pFunct.01
'//

Public Function DIRECOTRY_CREAZIONE_N01_CreoNuovaDirectory_pFunct(par_TipoParametro_i As Integer, _
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
On Error GoTo Err_DIRECOTRY_CREAZIONE_N01_CreoNuovaDirectory_pFunct


        
        '//Imposto i parametri
        ProceduraAttivaEseguita_s = "DIRECOTRY_CREAZIONE_N01_CreoNuovaDirectory_pFunct"
        ProceduraMessaggioErrore_s = "Errore nella procedura"
        
    '//DIRECTORY_CONTROLLO
    '//.....................................................................................................
    '//Note           : Tramite una Select vengono individuati i valori da restiuire.

            
            '//IMPOSTAZIONE PATH E CONTROLLO ESISTENZA DIRECTORY
            '//------------------------------------------------------------------------------//
            '//NOTE     : controllo l'esistenza della path definita dai salvataggi se non esiste _
                        creo la directory.
                
                '//VALORIZZO I PARAMETRI
                Path_s = par_Directory_s
                ParametroFile_i = par_TipoParametro_i
                
                '//Variabili della directory
                Dim MyPath_v, MYNAME_v, MyValue_v As Variant
                
                '//Variabili della funzione Imput box
                Dim Message_s As String, Title_s As String, Default_i  As Integer
                '//Titolo della casella di messaggio di imput
                Title_s = "MESSAGGIO DELLA CASELLA DI IMPUT"
                
                Vv1 = Dir(par_Directory_s, vbDirectory)   ' Recupera la prima voce.
                
               ' Vv1 = Dir("c:\", vbDirectory)
                
                If Vv1 = "" Then
                        MsgBox "NON ESISTE LA DIRECTORY ---> " & Path_s & " - DA CREARE COME NUOVA", vbExclamation
                        'GoTo Exit_DIRECOTRY_CREAZIONE_N01_CreoNuovaDirectory_pFunct
                        
                        ' Display message, title, and default value.
                        MyValue_v = InputBox(Message_s, Title_s, Default_i)
                            'per creare una directory:
                                MkDir (Path_s)
                            'per verificare se esiste:
                            If Len(Dir(Path_s, vbDirectory)) > 0 Then
                                 MsgBox (Path_s & " esiste già!")
                            End If


                End If
            '//-------------------------------------------------------------------------------//


          
    '//*** fine ***
    '//DIRECTORY_CONTROLLO
    '//.....................................................................................................

'//USCITA  E GESTIONE ERRORI
'//..............................................................................................................


Exit_DIRECOTRY_CREAZIONE_N01_CreoNuovaDirectory_pFunct:
    Exit Function

Err_DIRECOTRY_CREAZIONE_N01_CreoNuovaDirectory_pFunct:
 '//-------------------------------------------------------------------------------
    MsgBox Err.Description & " - Errore Messaggio -> : " & ProceduraMessaggioErrore_s & " Procedura -> : " & ProceduraMessaggioErrore_s
    Debug.Print ProceduraMessaggioErrore_s
    Debug.Print ProceduraAttivaEseguita_s
    Stop
    Resume Exit_DIRECOTRY_CREAZIONE_N01_CreoNuovaDirectory_pFunct
'//-------------------------------------------------------------------------------
        
End Function

'//*** FINE ***
'//FUNZIONE------------------->DIRECOTRY_CREAZIONE_N01_CreoNuovaDirectory_pFunct
'//========================================================================================================================================//



VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_UTILITA_MsysDbEst_Frm20_IMPORTA_PROGETTO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************************//
'*  CLASSE DELLA FORM: UTILITA_MsysDbEst_Frm20_IMPORTA_PROGETTO
'*
'*  ATTIVITA        :questa classe della @form@IMPORTA@PROGETTO viene utilizzata per importare
'*                  tutti gli oggetti (tabelle, query, form, report ecc..) che in questo caso appartengono
'*                  al progetto MSYSDbEst_GE_OBJECT.mdb ma puo essere adattata per importare
'*                  altri oggetti che appartengono ad al progetti access.
'*                  VANTAGGIO: è dato dal fatto che gli oggetti tabelle query ecc... NON SONO SALVATI
'*                  in un tabella anchessa da importare ma sono salvati all'interno della classe in
'*                  in diverse collezioni precaricate. In questo puo spostare tutta form master
'*                  in qualsiasi progetto mdb senza preoccuparti di portare dietro tabelle oggetti, query
'*                  ed altre form di gestione. L'importazione degli oggetti RISULTA SEMPLICATA.
'*
'*  FUNZIONAMENTO   : Il funzionamento si basa su queste collezione di oggetti precaricati che  possono
'*                  essere sostituite:
'*
'*                  @PARAMETRO_01 = LA PATH ED IL FILE.MDB DOVE SI TROVANO GLI OGGETTI
'*
'*                  @PARAMETRO_02= COLLECTION DELLE FORMS DA IMPORTARE
'*                              I° COLLEZIONE : OGGETTI FORM PRECARICATI PRONTI PER L'IMPORTAZIONE
'*                              VAI AL CODICE: '//CODICE ----> COLLECTION_FORM_PRECARICATE
'*
'*  APRI ANALISI    : puoi aprire l'analisi di questa classe con questo comando
'*                  START "APRI ANALISI GESTIONE DB ESTERNO " call "c:\Casa\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MDB\MSYS_OBJECT_DB_EST\Project_MSYSDbEst_GE_OBJECT_ANALISI.sublime-project"
'*
'*
'*  ANALISI ACCESS  : L'analisi generale di access si apre con questo comando
'*                      START "APRI ANALISI GENERALE DI ACCESS" call "c:\Casa\LINGUAGGI\ACCESS\Project_ACCESS.sublime-project"
'*  FAQ             :
'*                  COME IMPORTARE GLI OGGETTI TABLE, QUERY, FORM IN UN PROGETTO?
'*                  CLASSE IMPORTA PROGETTI IN UN DB ESTERNO
'*                  @FORM@IMPORTA@OGGETTI_(Questa form permette l'importazione degli oggetti table, query ecc. in modo autonomo)
'*
'*
'*****************************************************************************************************//
Option Compare Database
    
'//DIM LE VARIABILI
'//-----------------------------------------------------------------------------------//
    '//APRI FILE
   
' Dim objectTables As Collection
   
    '//DATABASE PATH SORGENTE, CONTATORI
    Dim db As DAO.Database
    Dim ws As DAO.Workspace
    Dim sourceDB As DAO.Database
    'DAO
    Dim daoDB As DAO.Database
    Dim dbCurrent As DAO.Database

    Dim DaoWks As DAO.Workspace
    Dim daoRS As DAO.Recordset
    Dim DaoRs_Lett As DAO.Recordset
    Dim tdf As DAO.TableDef

    
    'ADO
    Dim ADODB As Database
    Dim AdodaoRs As Recordset
    
    'Contatori
    Dim icount As Integer
    Dim iTOTcount As Integer
    Dim dbl_count As Double
    

    '//PATH E FILE DOVE IMPORTARE
    Dim sourceDBPath_s As String
    
    '//DIM input box
    Dim Message As Variant
    Dim Title As Variant
    Dim Default As Variant
    Dim MyValue As Variant
    Dim MyMsg_V As Variant              '//risultato se 0 exit 1 continua procedura se annulla stringa = ""
    
    
    'Variabili generali
    Dim i As Integer
    Dim str1 As String
    Dim Int1 As Integer
    Dim Int2 As Integer
    Dim Int3 As Integer
    Dim Lng1 As Long
    Dim Dbl1 As Double
    Dim Bln1 As Boolean
    Dim Vv1 As Variant
    Dim obj1 As Object
    
    'Gestione parametri comandi
    Dim sxParamCmd_3 As String
    
    
    'Larghezza e numero di colonna
    Dim sLarg_Col As String
    Dim iNum_Col As Integer
    
    
    '//ERRORI PROCEDURA_FUNCTION O ROUTINE
    Dim ProceduraMessaggioErrore_s As String
    Dim ProceduraAttivaEseguita_s As String
    
    
    
    '//RECUPERO PARAMETRO DA TABELLA OGGETTI
    '//File da ricercare e la path
    Dim MyFile_s As String
    Dim MyPath_s As String
            
    
    


    
   '//@DIM@COLLECTION@TABLE COLLEZIONI TABELLE
   '//...................................
      
    '//DIM IMPORTA TABELLE
    Dim tableName As String
    Dim tblCiclo As TableDef
    Dim mytableName_s As String
    Dim objectTables As Collection              '//todo: doppia collezione di tabelle eliminare objectTables
    Dim tablesCollection As Collection
    
    Dim varItem As Variant
    '//i contatori delle tabelle
    Dim tableExistsInCollectionTables_Funct_b As Boolean
    Dim importedTablesCount_i As Integer
    Dim deletedTablesCount_i As Integer
    Dim deletedTables_s As String
       
   

    
   
    '//...................................

   '//@DIM@COLLECTION@FORM COLLEZIONI FORM
   '//...................................
     
        '//DIM IMPORTA FORM
        Dim myformName_s As String                      '//nome form da importare
        Dim formsToImport As Variant                    '//calcolo form da importare
        '//DIM della collezione form da importare
        Dim formsCollection As Collection                '//collezione delle form prestabilita
        
        '// i contatori delle form
        Dim deletedForms_s As String
        Dim deletetFormsCount_i As Integer
        Dim importedForms_s As String
        Dim importedFormsCount_i As Integer
    
        
   '//...................................


   '//@DIM@COLLECTION@QUERY COLLEZIONI QUERY
   '//...................................
     
        '//DIM IMPORTA QUERY
        Dim myNameQuery_s As String                         '//nome query da importare
        Dim queryesToImport As Variant                      '//calcolo query da importare
        '//DIM della collezione form da importare
        Dim queryesCollection As Collection                   '//collezione delle query prestabilita
        
        '// i contatori delle form
        Dim deleteNameQueryes_s As String
        Dim deleteQueryesCount_i As Integer
        Dim importedNameQueryes_s As String
        Dim importedQueryesCount_i As Integer
    
        
   '//...................................




   '//@DIM@COLLECTION@REPORT COLLEZIONI REPORTS
   '//...................................
     
        '//DIM IMPORTA QUERY
        Dim myreportName_s As String                         '//nome query da importare
        Dim reportsToImport As Variant                      '//calcolo REPORT da importare
        '//DIM della collezione form da importare
        Dim reportsCollection As Collection                   '//collezione delle REPORT prestabilita
        
        '// i contatori delle form
        Dim deletedReports_s As String
        Dim deletetReportesCount_i As Integer
        Dim importedReports_s As String
        Dim importedReportesCount_i As Integer
    
        
   '//...................................







    '//@DIM@COLLECTION@MACRO COLLEZIONI MACRO
    '//...................................
      
         '//DIM IMPORTA MACRO
         Dim myMacroName_s As String                        '//nome MACRO da importare
         Dim macrosToImport As Variant                      '//calcolo MACRO da importare
         '//DIM della collezione MACRO da importare
         Dim macrosCollection As Collection                 '//collezione delle MACRO prestabilita
         
         '// i contatori delle MACRO
         Dim deletedMacro_s As String
         Dim deletetMacroCount_i As Integer
         Dim importedMacro_s As String
         Dim importedMacroCount_i As Integer
     
         
    '//...................................

    













    '//@DIM@COLLECTION@MODULI COLLEZIONI MODULES
    '//...................................
      
         '//DIM IMPORTA MODULI
         Dim mymoduloName_s As String                         '//nome MODULI da importare
         Dim modulesToImport As Variant                      '//calcolo MODULI da importare
         '//DIM della collezione form da importare
         Dim modulesCollection As Collection                   '//collezione delle MODULI prestabilita
         
         '// i contatori delle form
         Dim deletedModules_s As String
         Dim deletetModulesCount_i As Integer
         Dim importedModules_s As String
         Dim importedModulesCount_i As Integer
     
         
    '//...................................

        
        
        
        
        'Le variabili generiche
        Dim sSql As String                                          ' Stringa di estrazione
        Dim sSql_Lett As String                                     ' Stringa di estrazione



'//DIM LE VARIABILI  *** FINE ***
'//-----------------------------------------------------------------------------------//







'//PARAMETRO PATH E FILE DOVE SI TROVANO GLI OGGETTI
'//-----------------------------------------------------------------------------------//

Public Function PathFile_s_pFunct() As String

On Error GoTo CollectionForm_PFunct_Err


        '//@PARAMETRO_01 = LA PATH ED IL FILE.MDB DOVE SI TROVANO GLI OGGETTI
        ' Percorso del database di origine
        sourceDBPath_s = "c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MDB\MSYS_OBJECT_DB_EST\MSYSDbEst_GE_OBJECT.mdb"
                            
                 
        PathFile_s_pFunct = sourceDBPath_s


'//USCITA ED ERRORI
'//--------------------------------------------------------------------------------------//

CollectionForm_PFunct:
    Exit Function

CollectionForm_PFunct_Err:
    MsgBox Error$
    Resume CollectionForm_PFunct

End Function
'//-----------------------------------------------------------------------------------//



'//************************************************************************************************************************//
'//***                              I PARAMETRI DELLE COLLECTION,
'//***                              COLLECTION OBJECT   *** INIZIO ***
'//************************************************************************************************************************//



'//LA COLLEZIONE DELLE FORM DA IMPORTARE  *** INIZIO ***
'//-----------------------------------------------------------------------------------//
'//NOTE: in questa funzione precarico la collezione di oggetti form _
    @PARAMETRO_02= COLLECTION DELLE FORMS DA IMPORTARE
    
'//PRECARICO LA COLLECTION FORM DA IMPORTARE
Public Function CollectionForm_PFunct() As Collection       '//restituisco una collection
    

On Error GoTo CollectionForm_PFunct_Err
        


    Dim frm As Variant             '//OGGETTO FORM DELLA COLLEZIONE PRECARICATA


    '//CODICE ----> COLLECTION_FORM_PRECARICATE
    '//CREO LA COLLEZIONE E LA POPOLO
    '//I° COLLEZIONE : OGGETTI FORM   @collection@oggetti_(@form la collezione delle form)
    '//------------------------------------------------------------------------------//
    '//NOTE: questa collection, rappresenta gli OGGETTI FORM  CHE APPARTENGONO al progetto _
            MSYSDbEst_GE_OBJECT.mdb e che vengono precaricate in un collezione. Quindi tali _
            oggetti devono esere eliminati o aggiunti a seconda che le form del progetto variano. _
            VANTAGGIO: il vantaggio di utilizzare la collezione facilita l'utilizzo della form di _
            importazione, in quanto le form non sono impostate una tabella esterna che dovrebbe essere _
            spostata insime alla form di importazione. MA ESSENDO PRECARICATE IN UNA COLLECTION si ha _
            sicurezza di spostare tutta la classe FORM con i dati relativi in questo caso agli oggetti FORM. _


    '//@OGGETTI@FORM@PRECARICATI
    '//popola la collection delle FORM di DA IMPORTARE
    ' Aggiungi i nomi delle FORM o ELIMINA O COMMENTA LE FORM DA ESCLUDERE. _
      SE VUOI UTILIZZARE LA COLLEZIONE anche per oggetti form DI PROGETTI DIVERSI inserisci i nome delle form _
      del progetto da importare. N.B. Rircordati di cambiare anche il parametro path+ file per rendere flessibile _
      questa collezione.
        Set objectForms = New Collection
        '//Aggiungo gli elementi della collezione
        
        '//@popolo@collection@form@system_(delle form di sistema, inserisco nella @collezione)
        objectForms.Add "MsysDbEstTb01Frm01_Mts01_}-------------------------------------@"
        objectForms.Add "MsysDbEstTb01Frm01_Mts01_GE_OBJECT"
        objectForms.Add "MsysDbEstTb01Frm01_Stf01_OBJECT"
        objectForms.Add "MsysDbEstTb01Frm01_Stf01_OBJECT_TMP"
        objectForms.Add "MsysDbEstTb05Frm01_Mt00_}--------------------------------------@"
        objectForms.Add "MsysDbEstTb05Frm01_Stf01_DB_ESTERNO"
        objectForms.Add "UTILITA_MsysDbEst_Frm12_}--------------------------------------@"
        objectForms.Add "UTILITA_MsysDbEst_Frm12_SERVIZI_TUTTI"
    
    
        ' Inizializza le collezioni per FORM COLLEGATE
        Set connectedFORMS = New Collection
      
    '// *** FINE ****
    '//I° COLLEZIONE DI FORM DI SISTEMA E LA POPOLO
    '//------------------------------------------------------------------------------//
    
        '//CONTROLLO DELLE FORM PRECARICATE NELLA COLLEZIONE
        '//...............................................................................//
        '//NOTE: queste sono le form da carica nel nuovo progetto.
                
                    
                    '//RIEPILOGO COLLEZIONI
                    '//---------------------------------------------------------//
                           
                            ' Stampa le FORM fisiche
                            Debug.Print
                            Debug.Print
                            Debug.Print "FORM DA IMPORTARE:"
                            '//iterazione nella collezione per il controllo
                            For Each frm In objectForms
                                Debug.Print "la form da importare nel progetto: " & frm
                            Next frm
        
                    '//---------------------------------------------------------//
                    
                        
                        
                        
                        ' Restituisci la collezione
                        Set CollectionForm_PFunct = objectForms
               



        
        '//...............................................................................//
        

'//USCITA ED ERRORI
'//--------------------------------------------------------------------------------------//

CollectionForm_PFunct_Exit:
    Exit Function

CollectionForm_PFunct_Err:
    MsgBox Error$
    Resume CollectionForm_PFunct_Exit

End Function



'//LA COLLEZIONE DELLE FORM DA IMPORTARE  *** FINE ***
'//-----------------------------------------------------------------------------------//




'//LA COLLEZIONE DELLE QUERY DA IMPORTARE  *** INIZIO ***
'//-----------------------------------------------------------------------------------//
'//NOTE: in questa funzione precarico la collezione di oggetti QUERY _
    @PARAMETRO_02= COLLECTION DELLE QUERYS DA IMPORTARE
    
'//PRECARICO LA COLLECTION QUERY DA IMPORTARE
Public Function CollectionQuery_PFunct() As Collection       '//restituisco una collection
    

On Error GoTo CollectionQuery_PFunct_Err
        


    Dim frm As Variant             '//OGGETTO QUERY DELLA COLLEZIONE PRECARICATA


    '//CODICE ----> COLLECTION_QUERY_PRECARICATE
    '//CREO LA COLLEZIONE E LA POPOLO
    '//I° COLLEZIONE : OGGETTI QUERY   @collection@oggetti_(@quert la collezione delle query)
    '//------------------------------------------------------------------------------//
    '//NOTE: questa collection, rappresenta gli OGGETTI QUERY  CHE APPARTENGONO al progetto _
            MSYSDbEst_GE_OBJECT.mdb e che vengono precaricate in un collezione. Quindi tali _
            oggetti devono esere eliminati o aggiunti a seconda che le QUERY del progetto variano. _
            VANTAGGIO: il vantaggio di utilizzare la collezione facilita l'utilizzo della QUERY di _
            importazione, in quanto le QUERY non sono impostate una tabella esterna che dovrebbe essere _
            spostata insieme alla QUERY di importazione. MA ESSENDO PRECARICATE IN UNA COLLECTION si ha _
            sicurezza di spostare tutta la classe QUERY con i dati relativi in questo caso agli oggetti QUERY. _


    '//@OGGETTI@QUERY@PRECARICATI
    '//popola la collection delle QUERY di DA IMPORTARE
    ' Aggiungi i nomi delle QUERY o ELIMINA O COMMENTA LE QUERY DA ESCLUDERE. _
      SE VUOI UTILIZZARE LA COLLEZIONE anche per oggetti QUERY DI PROGETTI DIVERSI inserisci i nome delle QUERY _
      del progetto da importare. N.B. Rircordati di cambiare anche il parametro path+ file per rendere flessibile _
      questa collezione.
        Set objectquery = New Collection
        '//Aggiugno gli elementi della collezione
        
        '//@popolo@collection@QUERY@system_(delle QUERY di sistema, inserisco nella @collezione)
        objectquery.Add "MSSys_DF02_Qry01_01_SELECT_TIPO_OGGETTO"
        objectquery.Add "MsysDbEstTb01Qry01_Dlt00_}-------------------------------------@"
        objectquery.Add "MsysDbEstTb01Qry01_Dlt01_OBJECT"
        objectquery.Add "MsysDbEstTb01Qry01_Dlt01_OBJECT_TMP"
        objectquery.Add "MsysDbEstTb01Qry01_Ins00_}-------------------------------------@"
        objectquery.Add "MsysDbEstTb01Qry01_Ins01_OBJECT_Tmp_To_Archivio"
        objectquery.Add "MsysDbEstTb01Qry01_Slt00_}-------------------------------------@"
        objectquery.Add "MsysDbEstTb01Qry01_Slt01_OBJECT"
        objectquery.Add "MsysDbEstTb01Qry01_Slt01_OBJECT_TMP"
        objectquery.Add "MsysDbEstTb05Qry01_Slt00_}-------------------------------------@"
        objectquery.Add "MsysDbEstTb05Qry01_Slt01_DB_ESTERNO"
        
    
    
    '// *** FINE ****
    '//I° COLLEZIONE DI QUERY DI SISTEMA E LA POPOLO
    '//------------------------------------------------------------------------------//
    
        '//CONTROLLO DELLE QUERY PRECARICATE NELLA COLLEZIONE
        '//...............................................................................//
        '//NOTE: queste sono le QUERY da carica nel nuovo progetto.
                
                    
                    '//RIEPILOGO COLLEZIONI
                    '//---------------------------------------------------------//
                           
                            ' Stampa le QUERY fisiche
                            Debug.Print
                            Debug.Print
                            Debug.Print "QUERY DA IMPORTARE:"
                            '//iterazione nella collezione per il controllo
                            For Each qry In objectquery
                                Debug.Print "la QUERY da importare nel progetto: " & qry
                            Next qry
        
                    '//---------------------------------------------------------//
                    
                        
                        
                        
                        ' Restituisci la collezione
                        Set CollectionQuery_PFunct = objectquery
               



        
        '//...............................................................................//
        

'//USCITA ED ERRORI
'//--------------------------------------------------------------------------------------//

CollectionQuery_PFunct_Exit:
    Exit Function

CollectionQuery_PFunct_Err:
    MsgBox Error$
    Resume CollectionQuery_PFunct_Exit

End Function



'//LA COLLEZIONE DELLE TABLE DA IMPORTARE  *** INIZIO ***
'//-----------------------------------------------------------------------------------//


'//PRECARICO LA COLLECTION TABELLE DA IMPORTARE
Public Function CollectionTable_PFunct() As Collection          '//restituisco una collection
    

On Error GoTo CollectionTable_PFunct_Err
        
    Dim frm As Variant             '//OGGETTO TABELLE DELLA COLLEZIONE PRECARICATA


    '//CODICE ----> COLLECTION_TABELLE_PRECARICATE
    '//CREO LA COLLEZIONE E LA POPOLO
    '//I° COLLEZIONE : OGGETTI TABELLE
    '//------------------------------------------------------------------------------//
    '//NOTE: questa collection, rappresenta gli OGGETTI TABELLE  CHE APPARTENGONO al progetto _
            MSYSDbEst_GE_OBJECT.mdb e che vengono precaricate in un collezione. Quindi tali _
            oggetti devono esere eliminati o aggiunti a seconda che le TABELLE del progetto variano. _
            VANTAGGIO: il vantaggio di utilizzare la collezione facilita l'utilizzo della TABELLE di _
            importazione, in quanto le TABELLE non sono impostate una tabella esterna che dovrebbe essere _
            spostata insime alla TABELLE di importazione. MA ESSENDO PRECARICATE IN UNA COLLECTION si ha _
            sicurezza di spostare tutta la classe TABELLE con i dati relativi in questo caso agli oggetti TABELLE. _


    '//@OGGETTI@TABELLE@PRECARICATI
    '//popola la collection delle TABELLE di DA IMPORTARE
    ' Aggiungi i nomi delle TABELLE o ELIMINA O COMMENTA LE TABELLE DA ESCLUDERE. _
      SE VUOI UTILIZZARE LA COLLEZIONE anche per oggetti TABELLE DI PROGETTI DIVERSI inserisci i nome delle TABELLE _
      del progetto da importare. N.B. Rircordati di cambiare anche il parametro path+ file per rendere flessibile _
      questa collezione.
        
        Set objectTables = New Collection
        '//Aggiugno gli elementi della collezione
        
        '//@popolo@collection@TABELLE@system_(delle TABELLE di sistema, inserisco nella @collezione)
        objectTables.Add "COLL_MSysTb06_IMPORTA_PROGETTO"
        objectTables.Add "MSys_{@===================================================@}_Tab"
        'objectTables.Add "MSys_DF00_}----------------------------------------------------@"
        objectTables.Add "MSys_DF10_DATABASE"
        objectTables.Add "Msys_DF11_PROGETTI"
        objectTables.Add "Msys_DF12_GE_ObjProgetti"
        objectTables.Add "Msys_DF12_TIPO_OGGETTI"
        objectTables.Add "MSys_ELENCO_ObjProgetti"
        objectTables.Add "MSys_F_}-------------------------------------------------------@"
        objectTables.Add "MsysDbEstTb00_}------------------------------------------------@"
        objectTables.Add "MsysDbEstTb01_OBJECT"
        objectTables.Add "MsysDbEstTb01_OBJECT_TMP"
        objectTables.Add "MSysTb05_DB_EST"
        objectTables.Add "MSysTb06_IMPORTA_PROGETTO"
        
    
        ' Inizializza le collezioni per TABELLE COLLEGATE
        Set connectedTABELLES = New Collection
        
              '//ATTIVO LA COLLEZIONE
              '//------------------------------------------------------------------------//
              '//Note: CHIAMO LA FUNZIONE COLLECTION TABLE per ottenere la collezione di TABELLE
                
                '//Set tablesCollection = CollectionTable_PFunct()
                
              '//------------------------------------------------------------------------//
        
      
    '// *** FINE ****
    '//I° COLLEZIONE DI TABELLE DI SISTEMA E LA POPOLO
    '//------------------------------------------------------------------------------//
    
        '//CONTROLLO DELLE TABELLE PRECARICATE NELLA COLLEZIONE
        '//...............................................................................//
        '//NOTE: queste sono le TABELLE da carica nel nuovo progetto.
                
                    
                    '//RIEPILOGO COLLEZIONI
                    '//---------------------------------------------------------//
                           
                            ' Stampa le TABELLE fisiche
                            Debug.Print
                            Debug.Print
                            Debug.Print "TABELLE DA IMPORTARE:"
                            '//iterazione nella collezione per il controllo
                            For Each frm In objectTables
                                Debug.Print "la TABELLE da importare nel progetto: " & frm
                            Next frm
        
                    '//---------------------------------------------------------//
                    
                     ' Restituisci la collezione
                        Set CollectionTable_PFunct = objectTables
               



        
        '//...............................................................................//
        

'//USCITA ED ERRORI
'//--------------------------------------------------------------------------------------//

CollectionTable_PFunct_Exit:
    Exit Function

CollectionTable_PFunct_Err:
    MsgBox Error$
    Resume CollectionTable_PFunct_Exit

End Function


'//LA COLLEZIONE DELLE TABLE DA IMPORTARE  *** FINE ***
'//-----------------------------------------------------------------------------------//






'//LA COLLEZIONE DELLE MACRO DA IMPORTARE  *** INIZIO ***
'//-----------------------------------------------------------------------------------//


'//PRECARICO LA COLLECTION MACRO DA IMPORTARE
Public Function CollectionMacro_PFunct() As Collection          '//restituisco una collection
    

On Error GoTo CollectionMacro_PFunct_Err
        
    Dim frm As Variant             '//OGGETTO MACRO DELLA COLLEZIONE PRECARICATA


    '//CODICE ----> COLLECTION_MACRO_PRECARICATE
    '//CREO LA COLLEZIONE E LA POPOLO
    '//I° COLLEZIONE : OGGETTI MACRO
    '//------------------------------------------------------------------------------//
    '//NOTE: questa collection, rappresenta gli OGGETTI MACRO  CHE APPARTENGONO al progetto _
            MSYSDbEst_GE_OBJECT.mdb e che vengono precaricate in un collezione. Quindi tali _
            oggetti devono esere eliminati o aggiunti a seconda che le MACRO del progetto variano. _
            VANTAGGIO: il vantaggio di utilizzare la collezione facilita l'utilizzo della MACRO di _
            importazione, in quanto le MACRO non sono impostate una tabella esterna che dovrebbe essere _
            spostata insime alla MACRO di importazione. MA ESSENDO PRECARICATE IN UNA COLLECTION si ha _
            sicurezza di spostare tutta la classe MACRO con i dati relativi in questo caso agli oggetti MACRO. _


    '//@OGGETTI@MACRO@PRECARICATI
    '//popola la collection dei MACRO di DA IMPORTARE
    ' Aggiungi i nomi dei MACRO o ELIMINA O COMMENTA I MACRO DA ESCLUDERE. _
      SE VUOI UTILIZZARE LA COLLEZIONE anche per oggetti MACRO DI PROGETTI DIVERSI inserisci i nome dei MACRO _
      del progetto da importare. N.B. Rircordati di cambiare anche il parametro path+ file per rendere flessibile _
      questa collezione.
        
        Set objectMacro = New Collection
        '//Aggiugno gli elementi della collezione
        
        '//@popolo@collection@MACRO@system_(dei MACRO di sistema, inserisco nella @collezione)
        objectMacro.Add "MsysDbEstTb00Mcr00_}--------------------------------------@Macro"
        objectMacro.Add "MsysDbEstTb00Mcr01_N01_ESPORTA_XLS"
        objectMacro.Add "MsysDbEstTb00Mcr01_N02_COLLEGA_XLS"




    
        ' Inizializza le collezioni per MACRO COLLEGATE
        Set connectedMACRO = New Collection
        
              '//ATTIVO LA COLLEZIONE
              '//------------------------------------------------------------------------//
              '//Note: CHIAMO LA FUNZIONE COLLECTION MACRO per ottenere la collezione di MACRO
                
                '//Set macroCollection = CollectionMacro_PFunct()
                
              '//------------------------------------------------------------------------//
        
      
    '// *** FINE ****
    '//I° COLLEZIONE DI MACRO DI SISTEMA E LA POPOLO
    '//------------------------------------------------------------------------------//
    
        '//CONTROLLO DELI MACRO PRECARICATE NELLA COLLEZIONE
        '//...............................................................................//
        '//NOTE: queste sono le MACRO da carica nel nuovo progetto.
                
                    
                    '//RIEPILOGO COLLEZIONI
                    '//---------------------------------------------------------//
                           
                            ' Stampa le MACRO fisiche
                            Debug.Print
                            Debug.Print
                            Debug.Print "MACRO DA IMPORTARE:"
                            '//iterazione nella collezione per il controllo
                            For Each frm In objectMacro
                                Debug.Print "la MACRO da importare nel progetto: " & frm
                            Next frm
        
                    '//---------------------------------------------------------//
                    
                     ' Restituisci la collezione
                        Set CollectionMacro_PFunct = objectMacro
               



        
        '//...............................................................................//
        

'//USCITA ED ERRORI
'//--------------------------------------------------------------------------------------//

CollectionMacro_PFunct_Exit:
    Exit Function

CollectionMacro_PFunct_Err:
    MsgBox Error$
    Resume CollectionMacro_PFunct_Exit

End Function


'//LA COLLEZIONE DELLE MACRO DA IMPORTARE  *** FINE ***
'//-----------------------------------------------------------------------------------//










'//LA COLLEZIONE DELLE MODULI DA IMPORTARE  *** INIZIO ***
'//-----------------------------------------------------------------------------------//


'//PRECARICO LA COLLECTION MODULI DA IMPORTARE
Public Function CollectionModuli_PFunct() As Collection          '//restituisco una collection
    

On Error GoTo CollectionModuli_PFunct_Err
        
    Dim frm As Variant             '//OGGETTO MODULI DELLA COLLEZIONE PRECARICATA


    '//CODICE ----> COLLECTION_MODULI_PRECARICATE
    '//CREO LA COLLEZIONE E LA POPOLO
    '//I° COLLEZIONE : OGGETTI MODULI
    '//------------------------------------------------------------------------------//
    '//NOTE: questa collection, rappresenta gli OGGETTI MODULI  CHE APPARTENGONO al progetto _
            MSYSDbEst_GE_OBJECT.mdb e che vengono precaricate in un collezione. Quindi tali _
            oggetti devono esere eliminati o aggiunti a seconda che le MODULI del progetto variano. _
            VANTAGGIO: il vantaggio di utilizzare la collezione facilita l'utilizzo della MODULI di _
            importazione, in quanto le MODULI non sono impostate una tabella esterna che dovrebbe essere _
            spostata insime alla MODULI di importazione. MA ESSENDO PRECARICATE IN UNA COLLECTION si ha _
            sicurezza di spostare tutta la classe MODULI con i dati relativi in questo caso agli oggetti MODULI. _


    '//@OGGETTI@MODULI@PRECARICATI
    '//popola la collection dei MODULI di DA IMPORTARE
    ' Aggiungi i nomi dei MODULI o ELIMINA O COMMENTA I MODULI DA ESCLUDERE. _
      SE VUOI UTILIZZARE LA COLLEZIONE anche per oggetti MODULI DI PROGETTI DIVERSI inserisci i nome dei MODULI _
      del progetto da importare. N.B. Rircordati di cambiare anche il parametro path+ file per rendere flessibile _
      questa collezione.
        
        Set objectModules = New Collection
        '//Aggiugno gli elementi della collezione
        
        '//@popolo@collection@MODULI@system_(dei MODULI di sistema, inserisco nella @collezione)
        objectModules.Add "APPLICATION_Mdl00_}--------------------------------------------@"
        objectModules.Add "APPLICATION_Mdl01_CANCELLA_FORM_DB_ESTERNO"
        'objectModules.Add "MSys_DF00_}----------------------------------------------------@"
        objectModules.Add "APPLICATION_Mdl02_CANCELLA_REPORT_DB_ESTERNO"
        objectModules.Add "APPLICATION_Mdl03_CANCELLA_TABELLA_IN_DB_ESTERNO"
        objectModules.Add "APPLICATION_Mdl03_CONTROLLO_TABLE_DB_ESTERNO"
        objectModules.Add "APPLICATION_Mdl03_RICERCA_TABLE_IN_DB_ESTERNO"
        objectModules.Add "CURRENTB_Mdl00_}-----------------------------------------------@"
        objectModules.Add "CURRENTB_Mdl03_CANCELLA_QUERY_INTERNE_(da_finire)"
        objectModules.Add "GESTIONE_Mdl10_}-----------------------------------------------@"
        objectModules.Add "GESTIONE_Mdl10_}-----------------------------------------------@"
        objectModules.Add "MODELLO_Mdl00_}-----------------------------------------------@"
        objectModules.Add "MODELLO_Mdl01_SUB"
        objectModules.Add "MsysDbEst_Mdl01_MODULO_PROVA_FUNZIONE_E_ROUTINE"
        objectModules.Add "MsysDbEstTb01Mdl00_}-------------------------------------------@"
        objectModules.Add "MsysDbEstTb01Mdl01_Num01_CONTROLLO_DB_ESTERNO"
        
        
    
        ' Inizializza le collezioni per MODULI COLLEGATE
        Set connectedMODULIS = New Collection
        
              '//ATTIVO LA COLLEZIONE
              '//------------------------------------------------------------------------//
              '//Note: CHIAMO LA FUNZIONE COLLECTION MODULI per ottenere la collezione di MODULI
                
                '//Set tablesCollection = CollectionModuli_PFunct()
                
              '//------------------------------------------------------------------------//
        
      
    '// *** FINE ****
    '//I° COLLEZIONE DI MODULI DI SISTEMA E LA POPOLO
    '//------------------------------------------------------------------------------//
    
        '//CONTROLLO DELI MODULI PRECARICATE NELLA COLLEZIONE
        '//...............................................................................//
        '//NOTE: queste sono le MODULI da carica nel nuovo progetto.
                
                    
                    '//RIEPILOGO COLLEZIONI
                    '//---------------------------------------------------------//
                           
                            ' Stampa le MODULI fisiche
                            Debug.Print
                            Debug.Print
                            Debug.Print "MODULI DA IMPORTARE:"
                            '//iterazione nella collezione per il controllo
                            For Each frm In objectModules
                                Debug.Print "la MODULI da importare nel progetto: " & frm
                            Next frm
        
                    '//---------------------------------------------------------//
                    
                     ' Restituisci la collezione
                        Set CollectionModuli_PFunct = objectModules
               



        
        '//...............................................................................//
        

'//USCITA ED ERRORI
'//--------------------------------------------------------------------------------------//

CollectionModuli_PFunct_Exit:
    Exit Function

CollectionModuli_PFunct_Err:
    MsgBox Error$
    Resume CollectionModuli_PFunct_Exit

End Function


'//LA COLLEZIONE DELLE MODULI DA IMPORTARE  *** FINE ***
'//-----------------------------------------------------------------------------------//











'//LA COLLEZIONE DEI REPORT DA IMPORTARE  *** INIZIO ***
'//-----------------------------------------------------------------------------------//


'//PRECARICO LA COLLECTION REPORT DA IMPORTARE
Public Function CollectionReport_PFunct() As Collection          '//restituisco una collection
    

On Error GoTo CollectionReport_PFunct_Err
        
    Dim frm As Variant             '//OGGETTO REPORT DELLA COLLEZIONE PRECARICATA


    '//CODICE ----> COLLECTION_REPORT_PRECARICATE
    '//CREO LA COLLEZIONE E LA POPOLO
    '//I° COLLEZIONE : OGGETTI REPORT   @collection@oggetti_(@report la collezione delle Report)
    '//------------------------------------------------------------------------------//
    '//NOTE: questa collection, rappresenta gli OGGETTI REPORT  CHE APPARTENGONO al progetto _
            MSYSDbEst_GE_OBJECT.mdb e che vengono precaricate in un collezione. Quindi tali _
            oggetti devono esere eliminati o aggiunti a seconda che le REPORT del progetto variano. _
            VANTAGGIO: il vantaggio di utilizzare la collezione facilita l'utilizzo delI REPORT di _
            importazione, in quanto le REPORT non sono impostate una tabella esterna che dovrebbe essere _
            spostata insime alI REPORT di importazione. MA ESSENDO PRECARICATE IN UNA COLLECTION si ha _
            sicurezza di spostare tutta la classe REPORT con i dati relativi in questo caso agli oggetti REPORT. _


    '//@OGGETTI@REPORT@PRECARICATI
    '//popola la collection delle REPORT di DA IMPORTARE
    ' Aggiungi i nomi delle REPORT o ELIMINA O COMMENTA LE REPORT DA ESCLUDERE. _
      SE VUOI UTILIZZARE LA COLLEZIONE anche per oggetti REPORT DI PROGETTI DIVERSI inserisci i nome delle REPORT _
      del progetto da importare. N.B. Rircordati di cambiare anche il parametro path+ file per rendere flessibile _
      questa collezione.
        
        Set objectreports = New Collection
        '//Aggiugno gli elementi della collezione
        
        '//@popolo@collection@REPORT@system_(delle REPORT di sistema, inserisco nella @collezione)
        objectreports.Add "MsysDbEstTb01Rpt00_}-------------------------------------------@"
        objectreports.Add "MsysDbEstTb01Rpt01_S01_OBJECT"
        objectreports.Add "MsysDbEstTb01Rpt01_S01_OBJECT_TMP"
        
        ' Inizializza le collezioni per REPORT COLLEGATE
        Set connectedREPORTS = New Collection
        
              '//ATTIVO LA COLLEZIONE
              '//------------------------------------------------------------------------//
              '//Note: CHIAMO LA FUNZIONE COLLECTION TABLE per ottenere la collezione di REPORT
                
                '//Set tablesCollection = CollectionReport_PFunct()
                
              '//------------------------------------------------------------------------//
        
      
    '// *** FINE ****
    '//I° COLLEZIONE DI REPORT DI SISTEMA E LA POPOLO
    '//------------------------------------------------------------------------------//
    
        '//CONTROLLO DELLE REPORT PRECARICATE NELLA COLLEZIONE
        '//...............................................................................//
        '//NOTE: queste sono le REPORT da carica nel nuovo progetto.
                
                    
                    '//RIEPILOGO COLLEZIONI
                    '//---------------------------------------------------------//
                           
                            ' Stampa le REPORT fisiche
                            Debug.Print
                            Debug.Print
                            Debug.Print "REPORT DA IMPORTARE:"
                            '//iterazione nella collezione per il controllo
                            For Each frm In objectreports
                                Debug.Print "I REPORT da importare nel progetto: " & frm
                            Next frm
        
                    '//---------------------------------------------------------//
                    
                     ' Restituisci la collezione
                        Set CollectionReport_PFunct = objectreports
               



        
        '//...............................................................................//
        

'//USCITA ED ERRORI
'//--------------------------------------------------------------------------------------//

CollectionReport_PFunct_Exit:
    Exit Function

CollectionReport_PFunct_Err:
    MsgBox Error$
    Resume CollectionReport_PFunct_Exit

End Function

'//CHIAMO LA FUNZIONE DI CONTROLLO OGGETTI DI SISTEMA ESTERNI
Private Sub Cmd_CONTROLLO_OGGETTI_ESTERNI_Click()
    ListExternalDBObjects
End Sub



'//LA COLLEZIONE DEI REPORT DA IMPORTARE  *** FINE ***
'//-----------------------------------------------------------------------------------//





'//=========================================================================================//
'//
'//                 GESTIONE DELLE QUERY ESTERNE O INTERNE   *** inizio ***
'//@GESTIONE@DELLE@QUERY                    @GESTONE@QUERY
'//=========================================================================================//



'//CANCELLA GLI OGGETTI QUERY *** INIZIO *** '//DELETE QUERIES
'//==========================================================================================================//

Private Sub Cmd_Delete_QUERY_Click()

    Dim dbCurrent As DAO.Database
    Dim qry As DAO.QueryDef
    Dim objectTables As Collection
    Dim varItem As Variant
    Dim MATRICE_queryNames_s() As String             'matrice di Query
    Dim tempCount As Integer
    Dim queryCount As Integer
    Dim qdf  As QueryDef
    

    
'//TODO: controllare se tale routine funziona???
On Error GoTo Err_Cmd_Delete_QUERY_Click
    
    
    '//RESET
    icount = 0
    deleteQueryesCount_i = 0
    deletedTablesCount_i = 0
    deletedTables_s = ""
   
        
              '//ATTIVO LA COLLEZIONE
              '//------------------------------------------------------------------------//
              '//Note: CHIAMO LA FUNZIONE COLLECTION QUERY per ottenere la collezione di QUERY
                
                
                
                '//
                Set queryesCollection = CollectionQuery_PFunct()
                
              '//------------------------------------------------------------------------//
                    
                    ' Loop attraverso tutte I REPORT nel database corrente
                     Set dbCurrent = CurrentDb
                     
                     '//CALCOLO INDICE QUERY
                     icount = dbCurrent.QueryDefs.Count
                   
                    Set dbCurrent = CurrentDb

                        
                    '//se ci sono query da cancellare
                    If icount > 0 Then
                    
                        '//reimposto la matrice sulle query esistenti nel db corrente
                        ReDim MATRICE_queryNames_s(icount - 1)
                        
                        ' Costruisci un array con i nomi dei Query
                        tempCount = 0
                        
                        '//CARICO LA MATRICE: ciclo for per caricare MATRICE_queryNames_s(tempCount)
                        For Each qdf In dbCurrent.QueryDefs
                            MATRICE_queryNames_s(tempCount) = qdf.Name
                            tempCount = tempCount + 1
                        Next qdf
                        
                        ' Cancella i Query usando l'array dei nomi
                        For i = 0 To tempCount - 1
                            On Error Resume Next  ' Aggiungi gestione errori specifica per ogni iterazione
                            Debug.Print "Cancellazione Query: " & MATRICE_queryNames_s(i)
                            
                            '//CONTROLLO MATRICE = COLLECTION
                            Bool1 = InCollectionQueryes_Funct_b(queryesCollection, MATRICE_queryNames_s(i))
                            
                            '//cancello solo se la matrice (i) esiste nella collection
                            If Bool1 = True Then
                                    Debug.Print
                                    Debug.Print "cancello la query trovata sia nella collezione che nella matrice : " & MATRICE_queryNames_s(i)
                                    DoCmd.DeleteObject acQuery, MATRICE_queryNames_s(i)
                                    
                                    If Err.Number <> 0 Then
                                        Debug.Print "Errore durante la cancellazione del Query: " & MATRICE_queryNames_s(i) & " - " & Err.Description
                                        Err.Clear  ' Pulisci l'errore per la prossima iterazione
                                    Else
                                        deleteQueryesCount_i = deleteQueryesCount_i + 1
                                    End If
                                    
                            End If '//If Bool1 = True Then
                            
                            On Error GoTo 0  ' Reimposta gestione errori
                        Next i
                    End If
    
            ' Messaggio finale di riepilogo SE ABILITATO CON TRUE - FALSE DISABILITATO
            If Me.ctr_AVVISO_ABILITATO_TXT = True Then
                MsgBox "LE QUERY CANCELLATA DA QUESTO DATABASE SONO :  " & deleteQueryesCount_i, vbExclamation, "MSG CHIAMATO LA FUNZIONE PATHFILE"
            End If





'//USCITA  E GESTIONE ERRORI
'//..............................................................................................................

Exit_Cmd_Delete_QUERY_Click:
    Exit Sub

Err_Cmd_Delete_QUERY_Click:
        
    Set dbCurrent = Nothing
    
    MsgBox Err.Description & " - Errore Messaggio -> : " & ProceduraMessaggioErrore_s & " Procedura -> : " & ProceduraMessaggioErrore_s
    Debug.Print ProceduraMessaggioErrore_s
    Debug.Print ProceduraAttivaEseguita_s
    Stop
    Resume Exit_Cmd_Delete_QUERY_Click

End Sub

'//*** FINE ***
'//SENZA_PARAMETRI
'//========================================================================================================================================//



' Funzione per verificare se un elemento esiste nella collezione
Function InCollectionQueryes_Funct_b(col As Collection, key As String) As Boolean
    Dim varItem As Variant
    On Error Resume Next
    '//ITERNO NELLA COLLECTION PASSATA COME PARAMETRO
    For Each varItem In col
        
        Debug.Print
        Debug.Print "CONTROLLO COLLEZIONE = MATRICE (I)"
        Debug.Print "collezione: " & varItem & " = " & "matrice key: " & key
        Debug.Print
        If varItem = key Then
            InCollectionQueryes_Funct_b = True
            Debug.Print "trovato restituisco : " & InCollectionQueryes_Funct_b
            Exit Function
        End If
    Next varItem
    InCollectionQueryes_Funct_b = False
    On Error GoTo 0
End Function




'//DELETE QUERIES        ***FINE ***
'//==========================================================================================================//







'//=========================================================================================//
'//             *** FINE ***
'//                 GESTIONE DELLE QUERY ESTERNE O INTERNE
'//@GESTIONE@DELLE@QUERY    (@MODELLO@DELETE DELLE QUERY DA APPLICARE AGLI ALTRI OGGETTI)
'//=========================================================================================//



'//************************************************************************************************************************//
'//***
'//***                              COLLECTION OBJECT   *** FINE ***
'//************************************************************************************************************************//



'//IMPORTA PROGETTO
'//=========================================================================================//

'//IMPORTA TUTTI GLI OGGETTI FORM DEL DB MSYSDBEST
'//----------------------------------------------------------------------------------------//
Private Sub Cmd_Importa_FORM_Click()
 

 '//DIM della collezione form da importare
    Dim formsCollection As Collection
    Dim formName_v As Variant



    On Error GoTo ErrorHandler
    
    ' Inizializza i contatori e le stringhe di riepilogo
    importedCount_i = 0
    deletedCount_i = 0
    deletedForms_s = ""
    importedForms_s = ""
    
    
       
    
    '//RECUPERO LA COLLEZIONE FORM DA IMPORTARE E FACCIO IL CONTROLLO
    '//---------------------------------------------------------------------------//
    '//NOte Chiamo la funzione in cui vengo precariate le form da importare in una _
            collection che viene restituiscta con oggetto dalla funzione e impostata _
            nella variabile oggetto formsCollection.
        ' Chiama la funzione per ottenere la collezione di form
        Set formsCollection = CollectionForm_PFunct()
        
        '//ciclo nella collezzione per il controllo delle form precaricate
        For Each formName_v In formsCollection
            Debug.Print "Form nella collezione: " & formName_v
            ' Aggiungi qui il codice per lavorare con ciascun form nella collezione
               '//conteggio form DA IMPORTARE
                importedCount_i = importedCount_i + 1
                importedForms_s = importedForms_s & formName_v & vbCrLf
    
        Next formName_v
        
                    
            
            ' Messaggio finale di riepilogo SE ABILITATO CON TRUE - FALSE DISABILITATO             ' Messaggio finale di riepilogo
            If Me.ctr_AVVISO_ABILITATO_TXT = True Then
       
            MsgBox "CONTROLLO FORM PRECARICATE DA IMPORTARE:" & vbCrLf & _
                   "Forms importate: " & importedCount & vbCrLf & importedForms_s & _
                   "Forms esistenti cancellate: " & deletedCount & vbCrLf & deletedForms, vbInformation
    
            End If
    '//---------------------------------------------------------------------------//
    
    
    '//RECUPERO LA PATH ED IL FILE.MDB PER LE IMPORTAZIONI
    '//---------------------------------------------------------------------------//
    '//NOTE: cheiamo la funzione e recupero la path ed il file.mdb facendo un controllo _
            se la variabile è nulla msg di avviso ed uscita dalla routine.
            
        '//reset e recupero la path e ed il file
        sourceDBPath = ""
        
        sourceDBPath = PathFile_s_pFunct
        
        If sourceDBPath = "" Then
            MsgBox "ATTENZIONE ERRORE path e file sono null uscita dalla routine", vbCritical, "MSG CONTROLLO PATH E FILE VALORIZZATI"
              If MyValue = 0 Then Exit Sub
        End If
        
            ' Messaggio finale di riepilogo SE ABILITATO CON TRUE - FALSE DISABILITATO             ' Messaggio finale di riepilogo
            If Me.ctr_AVVISO_ABILITATO_TXT = True Then
       
                MsgBox "IMPORTO LE FORM DA QUESTO DATABASE; path e del file : " & sourceDBPath, vbExclamation, "MSG CHIAMATO LA FUNZIONE PATHFILE"
                
            End If
    '//---------------------------------------------------------------------------//

    
    
    
   '//CONTROLLO PROGETTO        **** fine ***
   '//----------------------------------------------------------------------------------------------------//
    

    
            '//=====================================================================================================================//
            '//                             STEP_02 = CONTROLLO PROGETTO
            '//CHIAMO IL CONTROLLO DEL PROGETTO  - 01 E 02) CONTROLLO - *** INIZIO ***
            '//=====================================================================================================================//
            '//NOTE: chiamo la funzione che controlla se il progetto corrente è il progetto originale, infatti _
                     non si possono caricare gli oggetti nel PROGETTO MSYSDbEst_GE_OBJECT.mdb perhè è originale _
                     e vi è una distruzione degli stessi. LA FUNZIONE CONTROLLO SE E' UN PROGETTO ESTERNO AMMISSIBILE _
                     per caricare gli oggetti table, query , form ecc. da importare. _
                     01) CONTROLLO = MSG DI AVVISO CON POSSIBILITA DI ANNULLARE _
                     02) CONTROLLO = CONTROLLO SE STIAMO NEL PROGETTO ORIGINALE  E QUINDI ESCLUSIONE.
                     
                        
                        
                   '//01) PRIMO CONTROLLO _
                     MESSAGGIO DI IMPORTAZIONE 0 = ANNULLA  1 = CONTINUA
                   '//..........................................................//
            
                         '//Attenzione L'AVVISO produce il messaggio solo se il _
                            parametro è TRUE altrimenti il messaggio viene escluso.
                         MyValue = CheckAvviso_i_PFunct(Me.ctr_AVVISO_ABILITATO_TXT)
                         
                         If MyValue = 0 Then Exit Sub
                     '//..........................................................//
            
                            '***************************************************************************************
                            '//todo : attenzione sistemare questo II controllo ???
                            '***************************************************************************************
                                    '//CONTROLLO PROGETTO
                                    '//----------------------------------------------------------------------------------------------------//
                                    '//Nota: controllo il progetto per vedere se ci troviamo in quello originale su cui NON E' POSSIBILE _
                                         fare importazioni. Vi è un 01) PRIMO CONTROLLO, imput se = 0 annulla la routine; ed _
                                             02) SECONDO CONTROLLO controllo data dalla chiamata alla funzione che controlla in quale progetto ci troviamo.
                                             
                                     ' Messaggio finale di riepilogo SE ABILITATO CON TRUE - FALSE DISABILITATO             ' Messaggio finale di riepilogo
                                    If Me.ctr_AVVISO_ABILITATO_TXT = True Then
                                         
                                     '//01) SECONDO CONTROLLO
                                     Message = "ATTENZIONE QUESTO COMANDO DI IMPORTAZIONE VALE SOLO PER II PROGETTI DIVERSI  " & Chr$(13) _
                                             & " PER ANNULLARE INSERIRE 0"   ' Imposta il messaggio.
                                             Title = "Demo InputBox"    ' Imposta il titolo.
                                             Default = "1"    ' Imposta il valore predefinito.
                                         
                                             ' Visualizza la finestra di dialogo nella posizione 100,100
                                             MyValue = InputBox(Message, Title, Default, 100, 100)
                                             
                                             If MyValue = 0 Then Exit Sub
                                    
                                    End If
                                    
                                    
                            
                            '***************************************************************************************
                            '//todo : attenzione sistemare questo II controllo ???  *** fine ***
                            '***************************************************************************************

                   
                   '//02) SECONDO CONTROLLO
                   '//CHIAMO IL CONTROLLO DEL PROGETTO
                   '//..........................................................//
                   '//Note  : True = siamo nel progetto originale non possiamo _
                              importare nessun oggetto perche distruggiamo il db corrente. _
                              Fale = siamo in un progetto esterno possiamo caricare tutto.
                                  
                                  
                       '//SE VIENE RESTITUITO TRUE è FUORI PROGETTO PUO CARICARE, SE FALSE EXIT PERCHE PROGETTO ORIGINALE
                       Bool1 = CheckProjectName_b_PFunct
                       
                       If Bool1 = True Then Exit Sub
                   '//..........................................................//
                   
             
            '//=====================================================================================================================//
            '//                             STEP_02 = CONTROLLO PROGETTO
            '//CHIAMO IL CONTROLLO DEL PROGETTO  - 01 E 02) CONTROLLO - *** FINE ***
            '//=====================================================================================================================//
            

            
        '//CONTROLLO FORM DA IMPORTARE SE SIAMO SU UN PROGETTO ESTERNO : Bool1 = False      **** INIZIO ***
        '//----------------------------------------------------------------------------------------------------//
             
             '//se false attivo la procedura, posso caricare nel progetto esterno se True siamo _
                 nel progetto originale e non si puo caricar nulla.
             If Bool1 = False Then
         
                 
                     ' Percorso del database di origine
                     sourceDBPath_s = "c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MDB\MSYS_OBJECT_DB_EST\MSYSDbEst_GE_OBJECT.mdb"   '//...sostituire con funzione che recupera path+file
                     
                     
                     
                     '//todo: come caricare gli oggetti form da una tabella esterna ??
                     ' Elenco dei nomi delle form da importare
                     'formsToImport = Array("MsysDbEstTb01Frm01_Mts01_}-------------------------------------@", _
                     '                      "MsysDbEstTb01Frm01_Mts01_GE_OBJECT", _
                     '                      "MsysDbEstTb01Frm01_Stf01_OBJECT", _
                     '                      "MsysDbEstTb01Frm01_Stf01_OBJECT_TMP")
                     '..... sostituiti
                     
                     ' Apri il database di origine e il database di destinazione
                     Set ws = DBEngine.Workspaces(0)
                     Set sourceDB = ws.OpenDatabase(sourceDBPath_s)
                     Set db = CurrentDb
                     
                     ' Importa ogni form
                     'For i = LBound(formsToImport) To UBound(formsToImport) '//SOSTITUIRE CON LA COLLECTION
                      For Each formName_v In formsCollection
                      

                         'Dim formName As String  .... sostituito con myformName_s
                         'formName = formsToImport(i)
                         myformName_s = formName_v
                         
                         ' Controlla se il modulo esiste nel database di destinazione ObjectExists = TRUE
                         'If ObjectExists(db, acForm, formName) Then    ... sostituito
                          If ObjectExists(db, acForm, myformName_s) Then
                            
                             ' Se esiste, cancella il modulo esistente
                             'DoCmd.DeleteObject acForm, formName   ... sostituito
                             DoCmd.DeleteObject acForm, myformName_s
                             
                             '//contatore form cancellate
                             deletedCount_i = deletedCount_i + 1
                             'deletedForms_s = deletedForms_s & formName & vbCrLf    ... sostituito
                             
                             '//raggruppo nella variabile il nome delle form cancellate + ritorno a capo
                             deletedForms_s = deletedForms_s & myformName_s & vbCrLf
                             
                         End If
                         
                         ' Controlla se il modulo esiste nel database di origine
                         'If ObjectExists(sourceDB, acForm, formName) Then   .... sostitutito
                         If ObjectExists(sourceDB, acForm, myformName_s) Then
                         
                             ' Importa il modulo dal database di origine
                             'DoCmd.TransferDatabase acImport, "Microsoft Access", sourceDBPath_s, acForm, formName, formName  ...sostituito
                             DoCmd.TransferDatabase acImport, "Microsoft Access", sourceDBPath_s, acForm, myformName_s, myformName_s
                             
                             '//contatore form importate
                             importedCount_i = importedCount_i + 1
                             '//raggruppa le form importate + ritorno a capo
                             
                             'importedForms_s = importedForms_s & formName & vbCrLf  .... SOSTITUTO
                             importedForms_s = importedForms_s & myformName_s & vbCrLf
                             
                         Else
                             '//AVVISO DI NON ESISTENZA DELLA FORM
                             'MsgBox "La form '" & formName & "' non esiste nel database di origine.", vbExclamation   ... sostituito
                             
                             MsgBox "La form '" & myformName_s & "' non esiste nel database di origine.", vbExclamation
                             
                         End If
                     'Next i '//sostituisco da collectio
                         Next formName_v

                             
                             
                             
                    '//TODO: MANCA IL CONTROLLO SULLA TABELLA DI SISTEMA DEL DB ESTERNO MSysObjects DA CREARE VEDI FUNZIONE Cmd_Importa_MODULI_Click
                    '//----------------------------------------------------------------------------------------------------------------------//
                    
                        '//TODO: adattare il controllo sulla tabella di sistema esterna ... vai a Cmd_Importa_MODULI_Click
            
                    '//----------------------------------------------------------------------------------------------------------------------//
                 
                     ' Chiudi il database di origine
                     sourceDB.Close
                     Set sourceDB = Nothing
                     Set ws = Nothing
                 
                     
                     ' Messaggio finale di riepilogo SE ABILITATO CON TRUE - FALSE DISABILITATO             ' Messaggio finale di riepilogo
                     If Me.ctr_AVVISO_ABILITATO_TXT = True Then
                        MsgBox "Importazione completata." & vbCrLf & _
                            "Forms importate: " & importedCount_i & vbCrLf & importedForms_s & _
                            "Forms esistenti cancellate: " & deletedCount_i & vbCrLf & deletedForms_s, vbInformation
                            
                    End If
                    
                                     
                                    
                        '//ciclo nella collezione per il controllo delle QUERY precaricate per SELEZIONARE LA PRIMA
                        '//--------------------------------------------------------------------------------------//
                            For Each formName_v In formsCollection
                                
                                '//vado all'ultima query
                                DoCmd.SelectObject acForm, formName_v, True
                                
                            Next formName_v
                                        
                            '//PRIMA DI USCIRE RILASCIO GLI OGGETTI
                            Set dbCurrent = Nothing
                            Set formsCollection = Nothing
                            
                        '//--------------------------------------------------------------------------------------//
                        
                            
                            
                            
             
             End If   '//If Bool1 = False Then
         '//CONTROLLO FORM DA IMPORTARE SE SIAMO SU UN PROGETTO ESTERNO : Bool1 = False      **** fine ***
        '//----------------------------------------------------------------------------------------------------//
        
    Exit Sub

ErrorHandler:
    MsgBox "Errore durante l'importazione: " & Err.Description, vbCritical
    If Not sourceDB Is Nothing Then
        sourceDB.Close
    End If
    Set sourceDB = Nothing
    Set ws = Nothing
End Sub

'//FUNZIONE DI CONTROLLO SE ESISTE L'OGGETTO
Function ObjectExists(db As DAO.Database, objType As Integer, objName As String) As Boolean
    Dim obj As DAO.Document
    On Error Resume Next
    Set obj = db.Containers("Forms").Documents(objName)
    ObjectExists = (Err.Number = 0)
    Set obj = Nothing
    On Error GoTo 0
End Function

'//IMPORTA TUTTI GLI OGGETTI FORM DEL DB MSYSDBEST *** FINE ***
'//----------------------------------------------------------------------------------------//









'//CANCELLA GLI OGGETTI FORM *** INIZIO *** DELETE_FORM
'//======================================================================================================//
'//@GESTIONE@DELLE@FORM    (@MODELLO@DELETE@OGGETTI)

'//NOTA: visto che la cancellazione riguarda il database corrente in cui è stata importata _
        la classe di IMPORTA PROGETTO, la routine cancella solo se le FORM esistono _
        nel database corrente mediante il cofronto tra : _
        LA COLLECTION FORM = GLI OGGETTI FORM  che sono attivi nel database ESTERNO. _
        Prima di cancellare viene effettuato il confronto di esistenza _
        @cancellazione@form@caricate_(cancellazione delle FORM gia caricate nel db esterno) _
        @cancella@TABLE@esterne_(attenzione puoi cancellare le FORM Nel @db@esterno)
            
'//FAQ : _
        DOVE SI CANCELLANO LE FORM PRECARICATE _
        COME CANCELLARE LE FORM ESTERNE _
        LE FORM ESTERNE DA CANCELLARE


'// PARAMETRI DA CAMBIARE : _
    -----> FORM _
    -----> DELLE@FORM _
    -----> DELETE_FORM _
    -----> CollectionForm_PFunct _
    -----> formsCollection _
    -----> objectFomr _
    -----> MATRICE_FORM_Names_s _
    -----> formsCount_i _
    -----> FormDef _
    -----> frm _


Private Sub Cmd_Delete_FORM_Click()

    Dim dbCurrent As DAO.Database
    Dim frm As Form
    Dim formsCollection As Collection
    Dim varItem As Variant
    Dim MATRICE_FORM_Names_s() As String             'matrice di FORM
    Dim tempCount As Integer
    Dim formsCount_i As Integer
    Dim fdf As AccessObject
    Dim icount As Integer
    Dim i As Integer
    Dim Bool1 As Boolean
    
    On Error GoTo Err_Cmd_Delete_FORM_Click

    '//RESET
    icount = 0
    deletetFormsCount_i = 0
    deletedForms_s = ""

    '//ATTIVO LA COLLEZIONE
    '//------------------------------------------------------------------------//
    '//Note: CHIAMO LA FUNZIONE COLLECTION FORM per ottenere la collezione di FORM

    Set formsCollection = CollectionForm_PFunct()
    '//------------------------------------------------------------------------//

    ' Loop attraverso tutte le FORM nel database corrente
    Set dbCurrent = CurrentDb

    '//CALCOLO INDICE FORM
    icount = Application.CurrentProject.AllForms.Count

    '//se ci sono FORM da cancellare
    If icount > 0 Then

        '//reimposto la matrice sulle FORM esistenti nel db corrente
        ReDim MATRICE_FORM_Names_s(icount - 1)

        ' Costruisci un array con i nomi dei FORM
        tempCount = 0

        '//CARICO LA MATRICE: ciclo for per caricare MATRICE_FORM_Names_s(tempCount)
        For i = 0 To icount - 1
            Set fdf = Application.CurrentProject.AllForms(i)
            MATRICE_FORM_Names_s(tempCount) = fdf.Name
            tempCount = tempCount + 1
        Next i

        ' Cancella i FORM usando l'array dei nomi
        For i = 0 To tempCount - 1
            On Error Resume Next  ' Aggiungi gestione errori specifica per ogni iterazione
            Debug.Print "Cancellazione FORM: " & MATRICE_FORM_Names_s(i)

            '//CONTROLLO MATRICE = COLLECTION
            Bool1 = InCollectionFORMes_Funct_b(formsCollection, MATRICE_FORM_Names_s(i))

            '//cancello solo se la matrice (i) esiste nella collection
            If Bool1 = True Then
                Debug.Print
                Debug.Print "cancello la FORM trovata sia nella collezione che nella matrice : " & MATRICE_FORM_Names_s(i)
                DoCmd.DeleteObject acForm, MATRICE_FORM_Names_s(i)

                If Err.Number <> 0 Then
                    Debug.Print "Errore durante la cancellazione del FORM: " & MATRICE_FORM_Names_s(i) & " - " & Err.Description
                    Err.Clear  ' Pulisci l'errore per la prossima iterazione
                Else
                    deletetFormsCount_i = deletetFormsCount_i + 1
                End If

            End If '//If Bool1 = True Then

            On Error GoTo 0  ' Reimposta gestione errori
        Next i
    End If

    ' Messaggio finale di riepilogo SE ABILITATO CON TRUE - FALSE DISABILITATO
    If Me.ctr_AVVISO_ABILITATO_TXT = True Then
        MsgBox "LE FORM CANCELLATA DA QUESTO DATABASE SONO :  " & deletetFormsCount_i, vbExclamation, "MSG CHIAMATO LA FUNZIONE PATHFILE"
    End If

'//USCITA  E GESTIONE ERRORI
'//..............................................................................................................

Exit_Cmd_Delete_FORM_Click:
    Exit Sub

Err_Cmd_Delete_FORM_Click:
    Set dbCurrent = Nothing
    MsgBox Err.Description & " - Errore Messaggio -> : " & ProceduraMessaggioErrore_s & " Procedura -> : " & ProceduraMessaggioErrore_s
    Debug.Print ProceduraMessaggioErrore_s
    Debug.Print ProceduraAttivaEseguita_s
    Stop
    Resume Exit_Cmd_Delete_FORM_Click

End Sub

' Funzione per verificare se un elemento esiste nella collezione
Function InCollectionFORMes_Funct_b(col As Collection, key As String) As Boolean
    Dim varItem As Variant
    On Error Resume Next
    '//ITERNO NELLA COLLECTION PASSATA COME PARAMETRO
    For Each varItem In col
        Debug.Print
        Debug.Print "CONTROLLO COLLEZIONE = MATRICE (I)"
        Debug.Print "collezione: " & varItem & " = " & "matrice key: " & key
        Debug.Print
        If varItem = key Then
            InCollectionFORMes_Funct_b = True
            Debug.Print "trovato restituisco : " & InCollectionFORMes_Funct_b
            Exit Function
        End If
    Next varItem
    InCollectionFORMes_Funct_b = False
    On Error GoTo 0
End Function




'//CANCELLA GLI OGGETTI FORM *** FINE *** DELETE_FORM
'//=========================================================================================================










'//DELTE DI TUTTI GLI OGGETTI FORM DEL DB MSYSDBEST *** FINE ***
'//----------------------------------------------------------------------------------------//

'//HELP DELETE FORM
Private Sub Cmd_Delete_FORM_HELP_Click()
                 
 '//APRO FILE PDF SPECIFICO CON IL COMANDO OGGETTO PDF - @APRI.FILE
'//--------------------------------------------------------------------------------//--------//
'//NOTE                 -> Apro il file di tipo doc, zip o pdf
'//CODICE               -> Function ApriFilePdf.01.01
'//PARAMETRI            -> par_Path_s         = PATH _
                        -> par_NameFile_s               = NOME FILE _
                        -> par_IDGestione_lng           = ID FILE DA RICERCARE per futuri utilizzi
                        
  '//IMPOSTO LE VARIABILI PATH E FILE
  '//
  Dim MyFile_s As String
  '//
  Dim MyPath_s  As String
  
        '//imposto i parametri - LA PATH
        '
        MyPath_s = "c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MDB\OBJECT\HELP\"
        
        '//IL FILE = attenzione al file ho lasciato un spazio perche a volte non funziona senza
        MyFile_s = "HELP_Cmd_Delete_FORM.pdf "
            
            
                     
         '//chiamo la funzione INTERNA DI APERTURA DEL FILE con i parametri =   'CALL (ApriFilePdf Path_programma_s, NameFile_s, IDGestione_lng) _
            Il vantaggio della procedura interna è quello di evitare di portare dietro al progetto IMPORTA OGGETTI incluso in questa form _
            anche i moduli e classi che potrebbero essere oggetto di duplicazione con l'importazione e cancellazione in caso di resettamento _
            totale.
            Call ApriFilePdf_Interna_s_PFunct(MyPath_s, MyFile_s, 0)

             
        
'//--------------------------------------------------------------------------------//--------//

End Sub














'//=========================================================================================//
'//
'//                 GESTIONE DELLE TABELLE ESTERNE O INTERNE   *** INIZIO ***
'//@GESTIONE@DELLE@TABELLE                    @GESTONE@TABELLE
'//=========================================================================================//
'//FAQ: _
        @HELP@TABELLE    @HELP@IMPORTA@TABELLE  @HELP@CANCELLA@TABELLE _
             @LA@GESTIONE@DELLE@TABELLE _
        @GESTIONE@TABELLE  @IMPORTA@TABELLE @CANCELLA@TABELLE




Private Sub Cmd_Importa_FORM_HELP_Click()
 '//APRO FILE PDF SPECIFICO CON IL COMANDO OGGETTO PDF - @APRI.FILE
'//--------------------------------------------------------------------------------//--------//
'//NOTE                 -> Apro il file di tipo doc, zip o pdf
'//CODICE               -> Function ApriFilePdf.01.01
'//PARAMETRI            -> par_Path_s         = PATH _
                        -> par_NameFile_s               = NOME FILE _
                        -> par_IDGestione_lng           = ID FILE DA RICERCARE per futuri utilizzi
                        
  '//IMPOSTO LE VARIABILI PATH E FILE
  '//
  Dim MyFile_s As String
  '//
  Dim MyPath_s  As String
  
        '//imposto i parametri - LA PATH
        '
        MyPath_s = "c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MDB\OBJECT\HELP\"
        
        '//IL FILE = attenzione al file ho lasciato un spazio perche a volte non funziona senza
        MyFile_s = "HELP_Cmd_Importa_FORM.pdf "
            
            
         '//chiamo la funzione INTERNA DI APERTURA DEL FILE con i parametri =   'CALL (ApriFilePdf Path_programma_s, NameFile_s, IDGestione_lng) _
            Il vantaggio della procedura interna è quello di evitare di portare dietro al progetto IMPORTA OGGETTI incluso in questa form _
            anche i moduli e classi che potrebbero essere oggetto di duplicazione con l'importazione e cancellazione in caso di resettamento _
            totale.
         Call ApriFilePdf_Interna_s_PFunct(MyPath_s, MyFile_s, 0)
        
             
        
'//--------------------------------------------------------------------------------//--------//

End Sub


'//**************************************************************************************************************************************//
'//*
'//*
'//*                                    GESTIONE DELLE MARCRO   *** INIZIO ***
'//*
'//*
'//**************************************************************************************************************************************//


'//=====================================================================================================================//
'//                            IMPORTO LE MACRO *** INIZIO ***
'//
'//=====================================================================================================================//



'//IMPORTA GLI OGGETTI MACRO DEL DB ESTERNO IN TABELLA *** INIZIO ***
'//----------------------------------------------------------------------------------------//
'//NOTE:   dal database di origine importa solo gli oggetti MACRO controllando prima se _
            nel db corrente di destinazione già esiste la MACRO per effettuare la cancellazione _
            prima dell'importazione.


'//PARAMETRI_DA_CAMBIARE _
        ----> Cmd_Importa_MACRO_Click _
        ----> Dim qdf As DAO.MacroDef _
        ----> Dim strSourceDbPath As String _
        ----> Dim objectMacros As Collection _
        ----> Dim objectExistsInCollectionMacros_Funct_b As Boolean _
        ----> MACRO _
        ----> MACRO _
        ----> COLLECTION MACRO _
        ----> importedMACROesCount_i _
        ----> deleteMACROCount_i _
        ----> deleteNameMACRO_s _


Private Sub Cmd_Importa_MACRO_Click()

    Dim dbCurrent As DAO.Database
    Dim dbSource As DAO.Database
    Dim dbs As Object
    Dim obj As AccessObject

    Dim qdf As DAO.QueryDef
    Dim rs As DAO.Recordset
    Dim strSQL As String
    Dim strSourceDbPath As String
    Dim objectMacros As Collection
    Dim varItem As Variant
    Dim objectExistsInCollectionMacros_Funct_b As Boolean
    Dim importedMACROesCount_i As Integer
    Dim deleteMACROCount_i As Integer
    Dim deleteNameMACRO_s As String
    Dim importedNameMACRO_s As String
    Dim myNameMACRO_s As String
    Dim Bool1 As Boolean
    Dim MyValue As Integer

    On Error GoTo ErrorHandler

    ' Inizializza le variabili
    importedMACROesCount_i = 0
    deleteMACROCount_i = 0
    deleteNameMACRO_s = ""
    importedNameMACRO_s = ""

    ' ATTIVA LA COLLEZIONE DI MACRO
    '//------------------------------------------------------------------------//

        Set objectMacros = CollectionMacro_PFunct()
        
    '//------------------------------------------------------------------------//


    ' STEP 01: RECUPERO PATH E FILE DI ORIGINE
    strSourceDbPath = PathFile_s_pFunct()
    If strSourceDbPath = "" Then
        MsgBox "ATTENZIONE: Path e nome del file di origine non specificati.", vbCritical, "Errore di importazione"
        Exit Sub
    End If

    ' STEP 02: APERTURA DEI DATABASE
    Set dbCurrent = CurrentDb
    Set dbSource = DBEngine.Workspaces(0).OpenDatabase(strSourceDbPath)

    ' STEP 03: RECUPERO DELLE MACRO DAL DATABASE DI ORIGINE
    strSQL = "SELECT Name FROM MSysObjects WHERE Type=-32766;" ' Tipo -32764 corrisponde alle macro
    Set qdf = dbSource.CreateQueryDef("", strSQL)
    Set rs = qdf.OpenRecordset()

    ' Loop attraverso le macro trovate nel database di origine
    Do While Not rs.EOF
        myNameMACRO_s = rs.Fields("Name").Value
        rs.MoveNext

        ' Controlla se la macro è nella collezione di macro
        objectExistsInCollectionMacros_Funct_b = InCollectionMacros_Funct_b(objectMacros, myNameMACRO_s)

        ' Se la macro esiste nella collezione, procedi con l'importazione
        If objectExistsInCollectionMacros_Funct_b Then
            ' Prova a eliminare la macro esistente nel database corrente
            On Error Resume Next
            
            DoCmd.DeleteObject acMacro, myNameMACRO_s
            
            deleteMACROCount_i = deleteMACROCount_i + 1
            deleteNameMACRO_s = deleteNameMACRO_s & myNameMACRO_s & vbCrLf
            On Error GoTo 0

            ' Importa la macro dal database di origine al database corrente
            DoCmd.TransferDatabase acImport, "Microsoft Access", strSourceDbPath, acMacro, myNameMACRO_s, myNameMACRO_s
            importedMACROesCount_i = importedMACROesCount_i + 1
            importedNameMACRO_s = importedNameMACRO_s & myNameMACRO_s & vbCrLf
        End If
    Loop

    ' Messaggio di riepilogo
    If Me.ctr_AVVISO_ABILITATO_TXT = True Then
        MsgBox "Importazione completata:" & vbCrLf & _
               "Macro importate: " & importedMACROesCount_i & vbCrLf & _
               "Macro eliminate: " & deleteMACROCount_i & vbCrLf & _
               "Elenco macro eliminate: " & vbCrLf & deleteNameMACRO_s, vbInformation
    End If
    
    
                '//ciclo nella collezione per il controllo DELLE MACRO per SELEZIONARE LA PRIMA
                '//--------------------------------------------------------------------------------------//
                     
                     '//itero nelle macro per la selezione dell'ultima
                    Set dbs = Application.CurrentProject
            
                     For Each obj In dbs.AllMacros
                     '//vado all'ultima query
                        DoCmd.SelectObject acMacro, obj.Name, True
                    
                     Next obj
                     
                     
            

                '//--------------------------------------------------------------------------------------//


    ' Libera le risorse
    rs.Close
    Set rs = Nothing
    Set qdf = Nothing
    Set dbSource = Nothing
    Set dbCurrent = Nothing

    Exit Sub

ErrorHandler:
    MsgBox "Errore durante l'importazione: " & Err.Description, vbCritical
    Debug.Print "Errore durante l'importazione: " & Err.Description
    If Not rs Is Nothing Then rs.Close
    If Not qdf Is Nothing Then Set qdf = Nothing
    If Not dbSource Is Nothing Then Set dbSource = Nothing
    If Not dbCurrent Is Nothing Then Set dbCurrent = Nothing
End Sub

Function InCollectionMacros_Funct_b(col As Collection, key As String) As Boolean
    Dim varItem As Variant
    On Error Resume Next
    For Each varItem In col
        If varItem = key Then
            InCollectionMacros_Funct_b = True
            Exit Function
        End If
    Next varItem
    InCollectionMacros_Funct_b = False
    On Error GoTo 0
End Function


'//=====================================================================================================================//
'//                            STEP_03 = IMPORTO LE TABELLE E LE MACRO *** fine ***
'//
'//=====================================================================================================================//




'//=========================================================================================//
'//
'//                 GESTIONE DELLE MACRO   *** INIZIO ***
'//@GESTIONE@DELLE@MACRO                    @GESTONE@MACRO
'//=========================================================================================//
'//FAQ: _
        @HELP@MACRO    @HELP@IMPORTA@MACRO  @HELP@CANCELLA@MACRO _
             @LA@GESTIONE@DELLE@MACRO _
        @GESTIONE@MACRO  @IMPORTA@MACRO @CANCELLA@MACRO

'//NOTA: visto che la cancellazione riguarda il database corrente in cui è stata importata _
        la classe di IMPORTA PROGETTO, la routine cancella solo se le MACRO esistono _
        nel database corrente mediante il confronto tra : _
        LA COLLECTION MACRO = GLI OGGETTI MACRO che sono attivi nel database ESTERNO. _
        Prima di cancellare viene effettuato il confronto di esistenza _
        @cancellazione@macro@caricati_(cancellazione delle MACRO già caricate nel db esterno) _
        @cancella@MACRO@esterne_(attenzione puoi cancellare le MACRO nel @db@esterno)

'//FAQ : _
        DOVE SI CANCELLANO LE MACRO PRECARICATE _
        COME CANCELLARE LE MACRO ESTERNE _
        LE MACRO ESTERNE DA CANCELLARE
 
'// PARAMETRI DA CAMBIARE : _
    -----> MACRO _
    -----> DELLE@MACRO _
    -----> DELETE_MACRO _
    -----> CollectionMacro_PFunct _
    -----> macrosCollection _
    -----> objectMacro _
    -----> MATRICE_MACRO_Names_s _
    -----> macrosCount_i _
    -----> MacroDef _
    -----> mdf _

Private Sub Cmd_Delete_MACRO_Click()

    Dim dbCurrent As DAO.Database
    Dim dbs As Object

    Dim obj As AccessObject
    
    Dim macrosCollection As Collection
    Dim varItem As Variant
    Dim MATRICE_MACRO_Names_s() As String             'matrice di MACRO
    Dim tempCount As Integer
    Dim macrosCount_i As Integer
    Dim od As AccessObject
    Dim icount As Integer
    Dim i As Integer
    Dim Bool1 As Boolean
    
    On Error GoTo Err_Cmd_Delete_MACRO_Click

    '//RESET
    icount = 0
    deletetMacrosCount_i = 0
    deletedMacros_s = ""

    '//ATTIVO LA COLLEZIONE
    '//------------------------------------------------------------------------//
    '//Note: CHIAMO LA FUNZIONE COLLECTION MACRO per ottenere la collezione di MACRO

    Set macrosCollection = CollectionMacro_PFunct()
    '//------------------------------------------------------------------------//

    ' Loop attraverso tutte le MACRO nel database corrente
    Set dbCurrent = CurrentDb

    '//CALCOLO INDICE MACRO
    icount = CurrentProject.AllMacros.Count
    

    '//se ci sono MACRO da cancellare
    If icount > 0 Then

        '//reimposto la matrice sulle MACRO esistenti nel db corrente
        ReDim MATRICE_MACRO_Names_s(icount - 1)

        ' Costruisci un array con i nomi delle MACRO
        tempCount = 0

            '//CARICO LA MATRICE: ciclo for per caricare MATRICE_MACRO_Names_s(tempCount)
            Set dbs = Application.CurrentProject
            ' Ricerca oggetti AccessObject aperti in insieme AllMacros.
            For Each obj In dbs.AllMacros
                    ' Stampa nome di obj.
                    Debug.Print obj.Name
                    MATRICE_MACRO_Names_s(tempCount) = obj.Name
                    tempCount = tempCount + 1
            Next obj
    
    
        

        ' Cancella le MACRO usando l'array dei nomi
        For i = 0 To tempCount - 1
            On Error Resume Next  ' Aggiungi gestione errori specifica per ogni iterazione
            Debug.Print "Cancellazione MACRO: " & MATRICE_MACRO_Names_s(i)

            '//CONTROLLO MATRICE = COLLECTION
            Bool1 = InCollectionMacrosDELETE_Funct_b(macrosCollection, MATRICE_MACRO_Names_s(i))

            '//cancello solo se la matrice (i) esiste nella collection
            If Bool1 = True Then
                Debug.Print
                Debug.Print "cancello la MACRO trovata sia nella collezione che nella matrice : " & MATRICE_MACRO_Names_s(i)
                DoCmd.DeleteObject acMacro, MATRICE_MACRO_Names_s(i)

                If Err.Number <> 0 Then
                    Debug.Print "Errore durante la cancellazione della MACRO: " & MATRICE_MACRO_Names_s(i) & " - " & Err.Description
                    Err.Clear  ' Pulisci l'errore per la prossima iterazione
                Else
                    deletetMacrosCount_i = deletetMacrosCount_i + 1
                End If

            End If '//If Bool1 = True Then

            On Error GoTo 0  ' Reimposta gestione errori
        Next i
    End If

    ' Messaggio finale di riepilogo SE ABILITATO CON TRUE - FALSE DISABILITATO
    If Me.ctr_AVVISO_ABILITATO_TXT = True Then
        MsgBox "Le MACRO CANCELLATE DA QUESTO DATABASE SONO :  " & deletetMacrosCount_i, vbExclamation, "MSG CHIAMATO LA FUNZIONE PATHFILE"
    End If

'//USCITA E GESTIONE ERRORI
'//..............................................................................................................

Exit_Cmd_Delete_MACRO_Click:
    Exit Sub

Err_Cmd_Delete_MACRO_Click:
    Set dbCurrent = Nothing
    MsgBox Err.Description & " - Errore Messaggio -> : " & ProceduraMessaggioErrore_s & " Procedura -> : " & ProceduraMessaggioErrore_s
    Debug.Print ProceduraMessaggioErrore_s
    Debug.Print ProceduraAttivaEseguita_s
    Stop
    Resume Exit_Cmd_Delete_MACRO_Click

End Sub

' Funzione per verificare se un elemento esiste nella collezione di MACRO
Function InCollectionMacrosDELETE_Funct_b(col As Collection, key As String) As Boolean
    Dim varItem As Variant
    On Error Resume Next
    '//ITERNO NELLA COLLECTION PASSATA COME PARAMETRO
    For Each varItem In col
        Debug.Print
        Debug.Print "CONTROLLO COLLEZIONE = MATRICE (I)"
        Debug.Print "collezione: " & varItem & " = " & "matrice key: " & key
        Debug.Print
        If varItem = key Then
            InCollectionMacrosDELETE_Funct_b = True
            Debug.Print "trovato restituisco : " & InCollectionMacrosDELETE_Funct_b
            Exit Function
        End If
    Next varItem
    InCollectionMacrosDELETE_Funct_b = False
    On Error GoTo 0
End Function

'//CANCELLA GLI OGGETTI MACRO *** FINE *** DELETE_MACRO
'//=========================================================================================================



Private Sub Cmd_Delete_MACRO_HELP_Click()

    '//APRO FILE PDF SPECIFICO CON IL COMANDO OGGETTO PDF - @APRI.FILE
    '//--------------------------------------------------------------------------------//--------//
    '//NOTE                 -> Apro il file di tipo doc, zip o pdf
    '//CODICE               -> Function ApriFilePdf.01.01
    '//PARAMETRI            -> par_Path_s         = PATH _
                            -> par_NameFile_s               = NOME FILE _
                            -> par_IDGestione_lng           = ID FILE DA RICERCARE per futuri utilizzi
                            
      '//IMPOSTO LE VARIABILI PATH E FILE
      '//
      Dim MyFile_s As String
      '//
      Dim MyPath_s  As String
      
            '//imposto i parametri - LA PATH
            '
            MyPath_s = "c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MDB\OBJECT\HELP\"
            '//IL FILE = attenzione al file ho lasciato un spazio perche a volte non funziona senza
            '
            MyFile_s = "HELP_Cmd_Delete_MACRO.pdf "
                
                
                       
         '//chiamo la funzione INTERNA DI APERTURA DEL FILE con i parametri =   'CALL (ApriFilePdf Path_programma_s, NameFile_s, IDGestione_lng) _
            Il vantaggio della procedura interna è quello di evitare di portare dietro al progetto DELETE OGGETTI incluso in questa MODULI _
            anche i moduli e classi che potrebbero essere oggetto di duplicazione con l'DELETEzione e cancellazione in caso di resettamento _
            totale.
            Call ApriFilePdf_Interna_s_PFunct(MyPath_s, MyFile_s, 0)
      
            
    '//--------------------------------------------------------------------------------//--------//


End Sub



'//**************************************************************************************************************************************//
'//*
'//*
'//*                                    GESTIONE DELLE MARCRO   *** INIZIO ***
'//*
'//*
'//**************************************************************************************************************************************//







'//IMPORTA MODULI HELP *** INIZIO ***
'//-----------------------------------------------------------------------------------//

Private Sub Cmd_IMPORTA_MODULI_HELP_Click()

    '//APRO FILE PDF SPECIFICO CON IL COMANDO OGGETTO PDF - @APRI.FILE
    '//--------------------------------------------------------------------------------//--------//
    '//NOTE                 -> Apro il file di tipo doc, zip o pdf
    '//CODICE               -> Function ApriFilePdf.01.01
    '//PARAMETRI            -> par_Path_s         = PATH _
                            -> par_NameFile_s               = NOME FILE _
                            -> par_IDGestione_lng           = ID FILE DA RICERCARE per futuri utilizzi
                            
      '//IMPOSTO LE VARIABILI PATH E FILE
      '//
      Dim MyFile_s As String
      '//
      Dim MyPath_s  As String
      
            '//imposto i parametri - LA PATH
            '
            MyPath_s = "c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MDB\OBJECT\HELP\"
            '//IL FILE = attenzione al file ho lasciato un spazio perche a volte non funziona senza
            '
            MyFile_s = "HELP_Cmd_Importa_MODULI.pdf "
                
                
                       
         '//chiamo la funzione INTERNA DI APERTURA DEL FILE con i parametri =   'CALL (ApriFilePdf Path_programma_s, NameFile_s, IDGestione_lng) _
            Il vantaggio della procedura interna è quello di evitare di portare dietro al progetto DELETE OGGETTI incluso in questa MODULI _
            anche i moduli e classi che potrebbero essere oggetto di duplicazione con l'DELETEzione e cancellazione in caso di resettamento _
            totale.
            Call ApriFilePdf_Interna_s_PFunct(MyPath_s, MyFile_s, 0)
      
            
    '//--------------------------------------------------------------------------------//--------//

End Sub

'//IMPORTA MODULI HELP *** FINE ***
'//-----------------------------------------------------------------------------------//


'//IMPORTA PROGETTO
'//=========================================================================================//

'//IMPORTA TUTTI GLI OGGETTI REPORT DEL DB MSYSDBEST
'//----------------------------------------------------------------------------------------//
 
 
 

 
 
'//==========================================================================================================//
'// IMPORTA GLI OGGETTI REPORT *** INIZIO *** '@IMPORT@REPORT
'//==========================================================================================================//

Private Sub Cmd_Importa_REPORT_Click()
    Dim obj As AccessObject
    Dim dbs As Object
    Dim ProceduraMessaggioErrore_s As String
    Dim reportsCollection As Collection
    Dim reportExists As Boolean
    Dim sourceDBPath As String
    Dim externalDB As DAO.Database
    Dim rs As DAO.Recordset
    
    On Error GoTo Err_Cmd_Import_REPORT_Click
   
    ' Reset counts and strings
    deletetReportesCount_i = 0
    importedReportesCount_i = 0
    importedReports_s = ""
    
    ' Recupero la collezione di report da importare
    Set reportsCollection = CollectionReport_PFunct()
    
    ' Recupero il percorso e il nome del file del database esterno
    sourceDBPath = PathFile_s_pFunct
    
    If sourceDBPath = "" Then
        MsgBox "ATTENZIONE ERRORE: il percorso del file del database esterno non è stato specificato. Uscita dalla routine.", vbCritical, "Errore: Percorso del File Non Specificato"
        Exit Sub
    End If
    
    If Me.ctr_AVVISO_ABILITATO_TXT = True Then
        MsgBox "IMPORTO I REPORT DA QUESTO DATABASE; Percorso del file: " & sourceDBPath, vbExclamation, "Informazione: Chiamata Funzione PathFile"
    End If
    
    
    ' Apro il database esterno
    Set externalDB = DBEngine.Workspaces(0).OpenDatabase(sourceDBPath)
    
    ' Ciclo nella collezione per il controllo dei report precaricati
    For Each reportName In reportsCollection
        Debug.Print "Report nella collezione: " & myreportName_s
        
        ' Verifica se il report esiste nel database corrente
        reportExists = False
        Set dbs = Application.CurrentProject
        For Each obj In dbs.AllReports
            If obj.Name = reportName Then
                reportExists = True
                Exit For
            End If
        Next obj
        
        If reportExists Then
            ' Se il report esiste nel database corrente, cancellalo
            DoCmd.DeleteObject acReport, reportName
            deletetReportesCount_i = deletetReportesCount_i + 1
            Debug.Print "Report eliminato: " & reportName
        End If
        
        '//CONTROLLO SULLA TABELLA DI SISTEMA DEL DB ESTERNO MSysObjects
        '//----------------------------------------------------------------------------------------------------------------------//
        ' Verifica se il modulo esiste nel database esterno
        ' La riga SELECT Name FROM MSysObjects WHERE Type=-32761 è una query SQL utilizzata per selezionare i
        ' nomi degli oggetti di tipo modulo nel database di Access. Ecco una spiegazione dettagliata:

        ' MSysObjects: è una tabella di sistema in Microsoft Access che contiene informazioni su tutti gli oggetti nel database, _
          come tabelle, query, moduli, report, ecc. Type=-32761: questo filtro nella clausola WHERE specifica il tipo di oggetto _
          da selezionare. In Access, ogni tipo di oggetto ha un valore numerico associato nella tabella MSysObjects. _
          Il valore -32761 è il codice che rappresenta i moduli (VBA Modules).
         'Quindi, la query SELECT Name FROM MSysObjects WHERE Type=-32761 seleziona tutti i nomi degli oggetti nella tabella _
         di sistema MSysObjects che sono identificati come moduli VBA.
         ' Verifica se il report esiste nel database esterno
        ' Utilizzo una query SQL per verificare l'esistenza del report nel database esterno
        ' il codice del report è : -32764 controllato con la funzione codici esterni.
        

        reportExists = False
        Set rs = externalDB.OpenRecordset("SELECT Name FROM MSysObjects WHERE Type= -32764 AND Name='" & reportName & "'")
            If Not rs.EOF Then
                reportExists = True
            End If
                rs.Close
            
            If reportExists Then
                ' Se il report esiste nel database esterno, importalo
                DoCmd.TransferDatabase acImport, "Microsoft Access", sourceDBPath, acReport, reportName, reportName
                importedReportesCount_i = importedReportesCount_i + 1
                importedReports_s = importedReports_s & reportName & vbCrLf
                Debug.Print "Report importato: " & reportName
            Else
                ' Report non trovato nel database esterno
                Debug.Print "Report non trovato nel database esterno: " & reportName
            End If
        Next reportName
        
            ' Chiudo il database esterno
            externalDB.Close
        '//----------------------------------------------------------------------------------------------------------------------//
        
    ' Messaggio di riepilogo finale
    If Me.ctr_AVVISO_ABILITATO_TXT = True Then
        MsgBox "CONTROLLO REPORT PRECARICATI DA IMPORTARE:" & vbCrLf & _
               "Report importati: " & importedReportesCount_i & vbCrLf & importedReports_s & _
               "Report esistenti cancellati: " & deletetReportesCount_i, vbInformation
    End If
    
    
                '//ciclo nella collezione per il controllo delle QUERY precaricate per SELEZIONARE LA PRIMA
                '//--------------------------------------------------------------------------------------//
                     For Each obj In dbs.AllReports
                            '//vado all'ultima query
                        DoCmd.SelectObject acReport, obj.Name, True
                    
                     Next obj
                '//--------------------------------------------------------------------------------------//


Exit_Cmd_Import_REPORT_Click:
    Exit Sub

Err_Cmd_Import_REPORT_Click:
    MsgBox Err.Description & " - Errore: " & ProceduraMessaggioErrore_s
    Debug.Print ProceduraMessaggioErrore_s
    Stop
    Resume Exit_Cmd_Import_REPORT_Click

End Sub

'//==========================================================================================================//
'// IMPORTA GLI OGGETTI REPORT *** FINE *** '@IMPORT@REPORT
'//==========================================================================================================//


'//FUNZIONE DI CONTROLLO SE ESISTE L'OGGETTO = todo: modificata
Function ObjectReportExists(db As DAO.Database, objType As Integer, objName As String) As Boolean
    'Dim obj As DAO.Document
    'On Error Resume Next
    'Set obj = db.Containers("Reports").Documents(objName)
    'ObjectReportExists = (Err.Number = 0)
    'Set obj = Nothing
    'On Error GoTo 0
    
    Dim obj As AccessObject
    On Error Resume Next
    For Each obj In CurrentProject.AllReports
        Debug.Print "report corrente controllato :" & obj.Name
        If obj.Name = objName Then
            ReportExistsInCurrentDb = True
            Exit Function
        End If
    Next obj
    ReportExistsInCurrentDb = False
    On Error GoTo 0
    
    
    
End Function


'//IMPORTA TUTTI GLI OGGETTI REPORT DEL DB MSYSDBEST *** FINE ***
'//----------------------------------------------------------------------------------------//





'//CANCELLA GLI OGGETTI REPORT *** INIZIO *** DELETE_REPORT
'//======================================================================================================//
'//@GESTIONE@DEI@REPORT    (@MODELLO@DELETE@OGGETTI)

'//NOTA: visto che la cancellazione riguarda il database corrente in cui è stata importata _
        la classe di IMPORTA PROGETTO, la routine cancella solo se i REPORT esistono _
        nel database corrente mediante il confronto tra : _
        LA COLLECTION REPORT = GLI OGGETTI REPORT che sono attivi nel database ESTERNO. _
        Prima di cancellare viene effettuato il confronto di esistenza _
        @cancellazione@report@caricati_(cancellazione dei REPORT già caricati nel db esterno) _
        @cancella@TABLE@esterne_(attenzione puoi cancellare i REPORT nel @db@esterno)

'//FAQ : _
        DOVE SI CANCELLANO I REPORT PRECARICATI _
        COME CANCELLARE I REPORT ESTERNI _
        I REPORT ESTERNI DA CANCELLARE

'// PARAMETRI DA CAMBIARE : _
    -----> REPORT _
    -----> DEI@REPORT _
    -----> DELETE_REPORT _
    -----> CollectionReport_PFunct _
    -----> reportsCollection _
    -----> objectReport _
    -----> MATRICE_REPORT_Names_s _
    -----> reportsCount_i _
    -----> ReportDef _
    -----> rpt _

Private Sub Cmd_Delete_REPORT_Click()

    Dim dbCurrent As DAO.Database
    Dim rpt As Report
    Dim reportsCollection As Collection
    Dim varItem As Variant
    Dim MATRICE_REPORT_Names_s() As String             'matrice di REPORT
    Dim tempCount As Integer
    Dim reportsCount_i As Integer
    Dim rdf As AccessObject
    Dim icount As Integer
    Dim i As Integer
    Dim Bool1 As Boolean
    
    On Error GoTo Err_Cmd_Delete_REPORT_Click

    '//RESET
    icount = 0
    deletetReportsCount_i = 0
    deletedReports_s = ""

    '//ATTIVO LA COLLEZIONE
    '//------------------------------------------------------------------------//
    '//Note: CHIAMO LA FUNZIONE COLLECTION REPORT per ottenere la collezione di REPORT

    Set reportsCollection = CollectionReport_PFunct()
    '//------------------------------------------------------------------------//

    ' Loop attraverso tutti i REPORT nel database corrente
    Set dbCurrent = CurrentDb

    '//CALCOLO INDICE REPORT
    icount = Application.CurrentProject.AllReports.Count

    '//se ci sono REPORT da cancellare
    If icount > 0 Then

        '//reimposto la matrice sui REPORT esistenti nel db corrente
        ReDim MATRICE_REPORT_Names_s(icount - 1)

        ' Costruisci un array con i nomi dei REPORT
        tempCount = 0

        '//CARICO LA MATRICE: ciclo for per caricare MATRICE_REPORT_Names_s(tempCount)
        For i = 0 To icount - 1
            Set rdf = Application.CurrentProject.AllReports(i)
            MATRICE_REPORT_Names_s(tempCount) = rdf.Name
            tempCount = tempCount + 1
        Next i

        ' Cancella i REPORT usando l'array dei nomi
        For i = 0 To tempCount - 1
            On Error Resume Next  ' Aggiungi gestione errori specifica per ogni iterazione
            Debug.Print "Cancellazione REPORT: " & MATRICE_REPORT_Names_s(i)

            '//CONTROLLO MATRICE = COLLECTION
            Bool1 = InCollectionREPORTes_Funct_b(reportsCollection, MATRICE_REPORT_Names_s(i))

            '//cancello solo se la matrice (i) esiste nella collection
            If Bool1 = True Then
                Debug.Print
                Debug.Print "cancello il REPORT trovato sia nella collezione che nella matrice : " & MATRICE_REPORT_Names_s(i)
                DoCmd.DeleteObject acReport, MATRICE_REPORT_Names_s(i)

                If Err.Number <> 0 Then
                    Debug.Print "Errore durante la cancellazione del REPORT: " & MATRICE_REPORT_Names_s(i) & " - " & Err.Description
                    Err.Clear  ' Pulisci l'errore per la prossima iterazione
                Else
                    deletetReportsCount_i = deletetReportsCount_i + 1
                End If

            End If '//If Bool1 = True Then

            On Error GoTo 0  ' Reimposta gestione errori
        Next i
    End If

    ' Messaggio finale di riepilogo SE ABILITATO CON TRUE - FALSE DISABILITATO
    If Me.ctr_AVVISO_ABILITATO_TXT = True Then
        MsgBox "I REPORT CANCELLATI DA QUESTO DATABASE SONO :  " & deletetReportsCount_i, vbExclamation, "MSG CHIAMATO LA FUNZIONE PATHFILE"
    End If

'//USCITA E GESTIONE ERRORI
'//..............................................................................................................

Exit_Cmd_Delete_REPORT_Click:
    Exit Sub

Err_Cmd_Delete_REPORT_Click:
    Set dbCurrent = Nothing
    MsgBox Err.Description & " - Errore Messaggio -> : " & ProceduraMessaggioErrore_s & " Procedura -> : " & ProceduraMessaggioErrore_s
    Debug.Print ProceduraMessaggioErrore_s
    Debug.Print ProceduraAttivaEseguita_s
    Stop
    Resume Exit_Cmd_Delete_REPORT_Click

End Sub

' Funzione per verificare se un elemento esiste nella collezione
Function InCollectionREPORTes_Funct_b(col As Collection, key As String) As Boolean
    Dim varItem As Variant
    On Error Resume Next
    '//ITERNO NELLA COLLECTION PASSATA COME PARAMETRO
    For Each varItem In col
        Debug.Print
        Debug.Print "CONTROLLO COLLEZIONE = MATRICE (I)"
        Debug.Print "collezione: " & varItem & " = " & "matrice key: " & key
        Debug.Print
        If varItem = key Then
            InCollectionREPORTes_Funct_b = True
            Debug.Print "trovato restituisco : " & InCollectionREPORTes_Funct_b
            Exit Function
        End If
    Next varItem
    InCollectionREPORTes_Funct_b = False
    On Error GoTo 0
End Function


'//CANCELLA GLI OGGETTI REPORT *** FINE *** DELETE_REPORT
'//=========================================================================================================


'//**************************************************************************************//
'//                     EVENTI FORM  @GLI@EVENTI@FORM
'//**************************************************************************************//

'//FORM@SU@CARICAMENTO
Private Sub Form_Load()

    '//imposto il flag avviso ad abilitato per visualizzare il messaggio di proseguimento
    Me.ctr_AVVISO_ABILITATO_TXT = True
    
    
End Sub






'//**************************************************************************************//
'//                     EVENTI FORM  *** FINE ***
'//**************************************************************************************//




'//=========================================================================================//
'//
'//                 GESTIONE DELLE MACRO ESTERNE O INTERNE   *** INIZIO ***
'//@GESTIONE@DELLE@MACRO                    @GESTONE@MACRO
'//=========================================================================================//


'//IMPORTA GLI OGGETTI MACRO DEL DB ESTERNO IN TABELLA *** INIZIO ***
'//----------------------------------------------------------------------------------------//
'//NOTE:   dal database di origine importa solo gli oggetti MACRO controllando prima se _
            nel db corrente di destinazione già esiste la MACRO per effettuare la cancellazione _
            prima dell'importazione.

'//IMPORTA GLI OGGETTI MACRO DEL DB ESTERNO IN TABELLA *** FINE ***
'//----------------------------------------------------------------------------------------//

'//=========================================================================================//
'//
'//                 GESTIONE DELLE MACRO ESTERNE O INTERNE   *** FINE ***
'//@GESTIONE@DELLE@MACRO                    @GESTONE@MACRO
'//=========================================================================================//







'//=========================================================================================//
'//
'//                 GESTIONE DELLE QUERY ESTERNE O INTERNE   *** INIZIO ***
'//@GESTIONE@DELLE@QUERY                    @GESTONE@QUERY
'//=========================================================================================//


'//IMPORTA GLI OGGETTI TABLE E QUERY DEL DB ESTERNO IN TABELLA *** INIZIO ***
'//----------------------------------------------------------------------------------------//
'//NOTE:   dal database di origine importa solo gli oggetti Table e Query controllando prima se _
            nel db corrente di destinazione già esiste la Table o Query per effettuare la cancellazione _
            prima dell'importazione.


'//PARAMETRI_DA_CAMBIARE _
        ----> Cmd_Importa_QUERY_Click _
        ----> Dim qdf As DAO.QueryDef _
        ----> Dim strSourceDbPath As String _
        ----> Dim objectQueries As Collection _
        ----> Dim objectExistsInCollectionQueries_Funct_b As Boolean _
        ----> TABELLE E QUERY _
        ----> QUERY _
        ----> Dim objectExistsInCollectionQueries_Funct_b As Boolean _
        ----> COLLECTION QUERY _
        ----> importedQueryesCount_i _
        ----> deleteQueryesCount_i _
        ----> deleteNameQueryes_s _

Private Sub Cmd_Importa_QUERY_Click()

Dim dbCurrent As DAO.Database
Dim dbSource As DAO.Database
Dim qdf As DAO.QueryDef
Dim strSourceDbPath As String
Dim strObjectName As String
Dim objectQueries As Collection
Dim varItem As Variant
Dim objectExistsInCollectionQueries_Funct_b As Boolean

On Error GoTo ErrorHandler

'//=====================================================================================================================//
'//                            STEP_01 = RECUPERO PATH E COLLECTION
'//RECUPERO LA COLLEZIONE TABELLE E QUERY DA IMPORTARE E FACCIO IL CONTROLLO
'//=====================================================================================================================//

    '//RESET
    'Inizializza i contatori e le stringhe di riepilogo
    icount = 0
    importedQueryesCount_i = 0
    importedObjects_s = ""
    deleteQueryesCount_i = 0
    deleteNameQueryes_s = ""
              
    '//ATTIVO LA COLLEZIONE
    '//------------------------------------------------------------------------//
    '//Note: CHIAMO LA FUNZIONE COLLECTION QUERY per ottenere la collezione di TABELLE E QUERY
    Set objectQueries = CollectionQuery_PFunct()
    '//------------------------------------------------------------------------//

    '//I° CONTROLLO SULLA COLLECTION QUERY RECUPERATA
    '//ciclo nella collezione per il controllo delle TABELLE E QUERY precaricate
    For Each objCollection_objectName_v In objectQueries
        Debug.Print "Oggetto nella collezione: " & objCollection_objectName_v
        '//conteggio oggetti DA IMPORTARE
        icount = icount + 1
        importedNameQueryes_s = importedNameQueryes_s & objCollection_objectName_v & vbCrLf
    Next objCollection_objectName_v

    ' Messaggio finale di riepilogo SE ABILITATO CON TRUE - FALSE DISABILITATO
    If Me.ctr_AVVISO_ABILITATO_TXT = True Then
        MsgBox "CONTROLLO OGGETTI COLLECTION PRECARICATI DA IMPORTARE:" & vbCrLf & _
               "Oggetti COLLECTION PRECARICATA DA IMPORTARE: " & icount & vbCrLf & importedNameQueryes_s & _
               "Oggetti esistenti cancellati: " & deleteQueryesCount_i & vbCrLf & deleteNameQueryes_s, vbInformation
    End If
    
            '//RESET
            icount = 0
             
    

'//---------------------------------------------------------------------------//

'//RECUPERO LA PATH ED IL FILE.MDB PER LE IMPORTAZIONI
'//---------------------------------------------------------------------------//
'//NOTE: cheiamo la funzione e recupero la path ed il file.mdb facendo un controllo _
        se la variabile è nulla msg di avviso ed uscita dalla routine.
    '//reset
    sourceDBPath_s = ""

    '//@RECUPERO@PATH_(ed il file
    sourceDBPath_s = PathFile_s_pFunct

    If sourceDBPath_s = "" Then
        MsgBox "ATTENZIONE ERRORE path e file sono null uscita dalla routine", vbCritical, "MSG CONTROLLO PATH E FILE VALORIZZATI"
        If MyValue = 0 Then Exit Sub
    End If

      ' Messaggio finale di riepilogo SE ABILITATO CON TRUE - FALSE DISABILITATO
    If Me.ctr_AVVISO_ABILITATO_TXT = True Then
        MsgBox "IMPORTO GLI OGGETTI DA QUESTO DATABASE; path e del file : " & sourceDBPath_s, vbExclamation, "MSG CHIAMATO LA FUNZIONE PATHFILE"
    End If

'//=====================================================================================================================//
'//                            STEP_01 = RECUPERO PATH E COLLECTION
'//RECUPERO LA COLLEZIONE FORM DA IMPORTARE E FACCIO IL CONTROLLO          **** FINE ***
'//=====================================================================================================================//

'//=====================================================================================================================//
'//                             STEP_02 = CONTROLLO PROGETTO
'//CHIAMO IL CONTROLLO DEL PROGETTO  - 01 E 02) CONTROLLO - *** INIZIO ***
'//=====================================================================================================================//
'//NOTE: chiamo la funzione che controlla se il progetto corrente è il progetto originale, infatti _
         non si possono caricare gli oggetti nel PROGETTO MSYSDbEst_GE_OBJECT.mdb perhè è originale _
         e vi è una distruzione degli stessi. LA FUNZIONE CONTROLLO SE E' UN PROGETTO ESTERNO AMMISSIBILE _
         per caricare gli oggetti table, query , form ecc. da importare. _
         01) CONTROLLO = MSG DI AVVISO CON POSSIBILITA DI ANNULLARE _
         02) CONTROLLO = CONTROLLO SE STIAMO NEL PROGETTO ORIGINALE  E QUINDI ESCLUSIONE.
                 
           '//01) PRIMO CONTROLLO _
             MESSAGGIO DI IMPORTAZIONE 0 = ANNULLA  1 = CONTINUA
           '//..........................................................//
             '//Attenzione L'AVVISO produce il messaggio solo se il _
                parametro è TRUE altrimenti il messaggio viene escluso.
             MyValue = CheckAvviso_i_PFunct(Me.ctr_AVVISO_ABILITATO_TXT)
             
             If MyValue = 0 Then Exit Sub
           '//..........................................................//
                 
           '//02) SECONDO CONTROLLO
           '//CHIAMO IL CONTROLLO DEL PROGETTO
           '//..........................................................//
           '//Note  : True = siamo nel progetto originale non possiamo _
                      importare nessun oggetto perche distruggiamo il db corrente. _
                      False = siamo in un progetto esterno possiamo caricare tutto.
                          
               '//SE VIENE RESTITUITO TRUE è FUORI PROGETTO PUO CARICARE, SE FALSE EXIT PERCHE PROGETTO ORIGINALE
               Bool1 = CheckProjectName_b_PFunct
               
               If Bool1 = True Then Exit Sub
           '//..........................................................//

'//=====================================================================================================================//
'//                             STEP_02 = CONTROLLO PROGETTO
'//CHIAMO IL CONTROLLO DEL PROGETTO  - 01 E 02) CONTROLLO - *** FINE ***
'//=====================================================================================================================//

'//=====================================================================================================================//
'//                            STEP_03 = IMPORTO LE TABELLE E LE QUERY
'//
'//=====================================================================================================================//
    
    '//RESET
    '//Inizializza i contatori e le stringhe di riepilogo
    
    
    importedQueriesCount_i = 0
    deleteQueryesCount_i = 0
    deleteNameQueryes_s = ""
    importedNameQueryes_s = ""
    

    ' Apri il database corrente
    Set dbCurrent = CurrentDb

    ' Apri il database di origine
    Set dbSource = DBEngine.Workspaces(0).OpenDatabase(sourceDBPath_s)
    
    ' Loop attraverso tutte le tabelle nel database di origine
    For Each qdf In dbSource.QueryDefs
        ' Salta le tabelle di sistema (i cui nomi iniziano con "MSys") aggiungo come parametro( la collezione table + nome tabella)
        
        If Left(qdf.Name, 3) <> "~sq" Or (Left(qdf.Name, 3) = "~sq" And InCollectionQueries_Funct_b(objectQueries, qdf.Name)) Then
            ' Nome della tabella nel database di origine
            myNameQuery_s = qdf.Name
            
            ' Controlla se la tabella è nella collezione
            objectExistsInCollectionQueries_Funct_b = False
            
            For Each varItem In objectQueries
                If myNameQuery_s = varItem Then
                    objectExistsInCollectionQueries_Funct_b = True
                    Exit For
                End If
            Next varItem
            
            ' Se la tabella è nella collezione, importala
            If objectExistsInCollectionQueries_Funct_b Then
                ' Elimina la tabella se esiste già nel database corrente
                On Error Resume Next
                
                dbCurrent.QueryDefs.Delete myNameQuery_s
                deleteQueryesCount_i = deleteQueryesCount_i + 1
                deleteNameQueryes_s = deleteNameQueryes_s & myNameQuery_s & vbCrLf
                
                On Error GoTo 0
                
                ' Copia la tabella dal database di origine al database corrente
                DoCmd.TransferDatabase acImport, "Microsoft Access", sourceDBPath_s, acQuery, myNameQuery_s, myNameQuery_s
                
                '//conteggio TABELLE IMPORTATE
                
                importedQueryesCount_i = importedQueryesCount_i + 1
                importedNameQueryes_s = importedNameQueryes_s & myNameQuery_s & vbCrLf
                
            End If
        End If
    Next qdf
                
        '//TODO: MANCA IL CONTROLLO SULLA TABELLA DI SISTEMA DEL DB ESTERNO MSysObjects DA CREARE VEDI FUNZIONE Cmd_Importa_MODULI_Click
        '//----------------------------------------------------------------------------------------------------------------------//
        
            '//TODO: adattare il controllo sulla tabella di sistema esterna ... vai a Cmd_Importa_MODULI_Click

        '//----------------------------------------------------------------------------------------------------------------------//
                
    ' Messaggio finale di riepilogo SE ABILITATO CON TRUE - FALSE DISABILITATO
    If Me.ctr_AVVISO_ABILITATO_TXT = True Then
        MsgBox "CONTROLLO OGGETTI PRECARICATI DA IMPORTARE:" & vbCrLf & _
               "Oggetti importati: " & importedQueryesCount_i & vbCrLf & importedNameQueryes_s & vbCrLf & _
               "Oggetti esistenti cancellati: " & deleteQueryesCount_i & vbCrLf & deleteNameQueryes_s, vbInformation
    End If
    
                '//ciclo nella collezione per il controllo delle QUERY precaricate per SELEZIONARE LA PRIMA
                '//--------------------------------------------------------------------------------------//
                    For Each objCollection_objectName_v In objectQueries
                        
                        '//vado all'ultima query
                        DoCmd.SelectObject acQuery, objCollection_objectName_v, True
                        
                    Next objCollection_objectName_v
                                
                    '//PRIMA DI USCIRE RILASCIO GLI OGGETTI
                    Set dbCurrent = Nothing
                    Set objectQueries = Nothing
                    
                '//--------------------------------------------------------------------------------------//

'//=====================================================================================================================//
'//                            STEP_03 = IMPORTO LE TABELLE E LE QUERY *** fine ***
'//
'//=====================================================================================================================//

'// CHIUSURA E CONTROLLO ERRORI
'//------------------------------------------------------------------------------------------------//
Exit Sub

ErrorHandler:
    MsgBox "Errore durante l'importazione: " & Err.Description, vbCritical
    Debug.Print "Errore durante l'importazione: " & Err.Description, vbCritical
    If Not dbSource Is Nothing Then
        dbSource.Close
    End If
    Set dbSource = Nothing
    Set dbCurrent = Nothing
End Sub   '//Private Sub Cmd_Importa_QUERY_Click()


' Funzione per verificare se un elemento esiste nella collezione
Function InCollectionQueries_Funct_b(col As Collection, key As String) As Boolean
    Dim varItem As Variant
    On Error Resume Next
    For Each varItem In col
        If varItem = key Then
            InCollectionQueries_Funct_b = True
            Exit Function
        End If
    Next varItem
    InCollectionQueries_Funct_b = False
    On Error GoTo 0
End Function

'//IMPORTA GLI OGGETTI TABLE E QUERY DEL DB ESTERNO IN TABELLA  *** FINE ***
'//----------------------------------------------------------------------------------------//



'//IMPORTA QUERY HELP *** INIZIO ***
'//-----------------------------------------------------------------------------------//

Private Sub Cmd_Importa_QUERY_HELP_Click()

    '//APRO FILE PDF SPECIFICO CON IL COMANDO OGGETTO PDF - @APRI.FILE
    '//--------------------------------------------------------------------------------//--------//
    '//NOTE                 -> Apro il file di tipo doc, zip o pdf
    '//CODICE               -> Function ApriFilePdf.01.01
    '//PARAMETRI            -> par_Path_s         = PATH _
                            -> par_NameFile_s               = NOME FILE _
                            -> par_IDGestione_lng           = ID FILE DA RICERCARE per futuri utilizzi
                            
      '//IMPOSTO LE VARIABILI PATH E FILE
      '//
      Dim MyFile_s As String
      '//
      Dim MyPath_s  As String
      
            '//imposto i parametri - LA PATH
            '
            MyPath_s = "c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MDB\OBJECT\HELP\"
            '//IL FILE = attenzione al file ho lasciato un spazio perche a volte non funziona senza
            '
            MyFile_s = "Help_IMPORTA_QUERY.pdf "
                
                
                       
         '//chiamo la funzione INTERNA DI APERTURA DEL FILE con i parametri =   'CALL (ApriFilePdf Path_programma_s, NameFile_s, IDGestione_lng) _
            Il vantaggio della procedura interna è quello di evitare di portare dietro al progetto IMPORTA OGGETTI incluso in questa form _
            anche i moduli e classi che potrebbero essere oggetto di duplicazione con l'importazione e cancellazione in caso di resettamento _
            totale.
            Call ApriFilePdf_Interna_s_PFunct(MyPath_s, MyFile_s, 0)
      
            
    '//--------------------------------------------------------------------------------//--------//

End Sub

'//IMPORTA QUERY HELP *** FINE ***
'//-----------------------------------------------------------------------------------//







'//DELETE REPORT HELP *** INIZIO ***
'//-----------------------------------------------------------------------------------//

Private Sub Cmd_Delete_REPORT_HELP_Click()

    '//APRO FILE PDF SPECIFICO CON IL COMANDO OGGETTO PDF - @APRI.FILE
    '//--------------------------------------------------------------------------------//--------//
    '//NOTE                 -> Apro il file di tipo doc, zip o pdf
    '//CODICE               -> Function ApriFilePdf.01.01
    '//PARAMETRI            -> par_Path_s         = PATH _
                            -> par_NameFile_s               = NOME FILE _
                            -> par_IDGestione_lng           = ID FILE DA RICERCARE per futuri utilizzi
                            
      '//IMPOSTO LE VARIABILI PATH E FILE
      '//
      Dim MyFile_s As String
      '//
      Dim MyPath_s  As String
      
            '//imposto i parametri - LA PATH
            '
            MyPath_s = "c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MDB\OBJECT\HELP\"
            '//IL FILE = attenzione al file ho lasciato un spazio perche a volte non funziona senza
            '
            MyFile_s = "Help_DELETE_REPORT.pdf "
                
                
                       
         '//chiamo la funzione INTERNA DI APERTURA DEL FILE con i parametri =   'CALL (ApriFilePdf Path_programma_s, NameFile_s, IDGestione_lng) _
            Il vantaggio della procedura interna è quello di evitare di portare dietro al progetto DELETE OGGETTI incluso in questa REPORT _
            anche i moduli e classi che potrebbero essere oggetto di duplicazione con l'DELETEzione e cancellazione in caso di resettamento _
            totale.
            Call ApriFilePdf_Interna_s_PFunct(MyPath_s, MyFile_s, 0)
      
            
    '//--------------------------------------------------------------------------------//--------//

End Sub

'//DELETE REPORT HELP *** FINE ***
'//-----------------------------------------------------------------------------------//



'//IMPORTA REPORT HELP *** INIZIO ***
'//-----------------------------------------------------------------------------------//

Private Sub Cmd_Importa_REPORT_HELP_Click()

    '//APRO FILE PDF SPECIFICO CON IL COMANDO OGGETTO PDF - @APRI.FILE
    '//--------------------------------------------------------------------------------//--------//
    '//NOTE                 -> Apro il file di tipo doc, zip o pdf
    '//CODICE               -> Function ApriFilePdf.01.01
    '//PARAMETRI            -> par_Path_s         = PATH _
                            -> par_NameFile_s               = NOME FILE _
                            -> par_IDGestione_lng           = ID FILE DA RICERCARE per futuri utilizzi
                            
      '//IMPOSTO LE VARIABILI PATH E FILE
      '//
      Dim MyFile_s As String
      '//
      Dim MyPath_s  As String
      
            '//imposto i parametri - LA PATH
            '
            MyPath_s = "c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MDB\OBJECT\HELP\"
            '//IL FILE = attenzione al file ho lasciato un spazio perche a volte non funziona senza
            '
            MyFile_s = "Help_IMPORTA_REPORT.pdf "
                
                
                       
         '//chiamo la funzione INTERNA DI APERTURA DEL FILE con i parametri =   'CALL (ApriFilePdf Path_programma_s, NameFile_s, IDGestione_lng) _
            Il vantaggio della procedura interna è quello di evitare di portare dietro al progetto IMPORTA OGGETTI incluso in questa REPORT _
            anche i moduli e classi che potrebbero essere oggetto di duplicazione con l'importazione e cancellazione in caso di resettamento _
            totale.
            Call ApriFilePdf_Interna_s_PFunct(MyPath_s, MyFile_s, 0)
      
            
    '//--------------------------------------------------------------------------------//--------//

End Sub

'//IMPORTA REPORT HELP *** FINE ***
'//-----------------------------------------------------------------------------------//




'//DELETE QUERY HELP *** INIZIO ***
'//-----------------------------------------------------------------------------------//

Private Sub Cmd_Delete_QUERY_HELP_Click()

    '//APRO FILE PDF SPECIFICO CON IL COMANDO OGGETTO PDF - @APRI.FILE
    '//--------------------------------------------------------------------------------//--------//
    '//NOTE                 -> Apro il file di tipo doc, zip o pdf
    '//CODICE               -> Function ApriFilePdf.01.01
    '//PARAMETRI            -> par_Path_s         = PATH _
                            -> par_NameFile_s               = NOME FILE _
                            -> par_IDGestione_lng           = ID FILE DA RICERCARE per futuri utilizzi
                            
      '//IMPOSTO LE VARIABILI PATH E FILE
      '//
      Dim MyFile_s As String
      '//
      Dim MyPath_s  As String
      
            '//imposto i parametri - LA PATH
            '
            MyPath_s = "c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MDB\OBJECT\HELP\"
            '//IL FILE = attenzione al file ho lasciato un spazio perche a volte non funziona senza
            '
            MyFile_s = "Help_delete_QUERY.pdf "
                
                
                       
         '//chiamo la funzione INTERNA DI APERTURA DEL FILE con i parametri =   'CALL (ApriFilePdf Path_programma_s, NameFile_s, IDGestione_lng) _
            Il vantaggio della procedura interna è quello di evitare di portare dietro al progetto IMPORTA OGGETTI incluso in questa form _
            anche i moduli e classi che potrebbero essere oggetto di duplicazione con l'importazione e cancellazione in caso di resettamento _
            totale.
            Call ApriFilePdf_Interna_s_PFunct(MyPath_s, MyFile_s, 0)
      
            
    '//--------------------------------------------------------------------------------//--------//

End Sub

'//DELETE QUERY HELP *** FINE ***
'//-----------------------------------------------------------------------------------//

'//=========================================================================================//
'//
'//                 GESTIONE DELLE QUERY ESTERNE O INTERNE   *** FINE ***
'//@GESTIONE@DELLE@QUERY
'//=========================================================================================//





    






'//=========================================================================================//
'//
'//                 GESTIONE DELLE TABELLE ESTERNE O INTERNE   *** INIZIO ***
'//@GESTIONE@DELLE@TABELLE                   @GESTONE@TABELLE
'//=========================================================================================//






'//IMPORTA GLI OGGETTI TABLE DEL DB ESTERNO IN TABELLA *** INIZIO ***
'//----------------------------------------------------------------------------------------//
'//NOTE:   dal database di origine importa solo gli oggetti Table controllando prima se _
            nel db corrente di destinazione già esiste la Table per effettuare la cancellazione _
            prima dell'importazione.


Private Sub Cmd_Importa_TABLE_Click()

Dim dbCurrent As DAO.Database
 Dim dbSource As DAO.Database
 Dim tdf As DAO.TableDef
 Dim strSourceDbPath As String
 Dim strTableName As String
' Dim objectTables As Collection
 Dim connectedTABELLES As Collection
 Dim varItem As Variant
 Dim tableExistsInCollectionTables_Funct_b As Boolean
         


    On Error GoTo ErrorHandler



    '//=====================================================================================================================//
    '//                            STEP_01 = RECUPERO PATH E COLLECTION
    '//RECUPERO LA COLLEZIONE TABELLE DA IMPORTARE E FACCIO IL CONTROLLO
    '//=====================================================================================================================//
    '//NOte Chiamo la funzione in cui vengo precariate le form da importare in una _
            collection che viene restituiscta con oggetto dalla funzione e impostata _
            nella variabile oggetto formsCollection.
        
    
                '//RESET
                'Inizializza i contatori e le stringhe di riepilogo
                importedTablesCount_i = 0
                importedTables_s = ""
                deletedTablesCount_i = 0
                
                                  
              '//ATTIVO LA COLLEZIONE
              '//------------------------------------------------------------------------//
              '//Note: CHIAMO LA FUNZIONE COLLECTION TABLE per ottenere la collezione di TABELLE
                
                '//
                    Set tablesCollection = CollectionTable_PFunct()
                
              '//------------------------------------------------------------------------//

                
        '//ciclo nella collezzione per il controllo delle TABELLE precaricate
        For Each objCollection_tableName_v In tablesCollection
            Debug.Print "Form nella collezione: " & objCollection_tableName_v
            ' Aggiungi qui il codice per lavorare con ciascuna TABELLA nella collezione
               '//conteggio form DA IMPORTARE
                importedTablesCount_i = importedTablesCount_i + 1
                importedTables_s = importedTables_s & objCollection_tableName_v & vbCrLf
    
        Next objCollection_tableName_v
        
            
            ' Messaggio finale di riepilogo SE ABILITATO CON TRUE - FALSE DISABILITATO
            If Me.ctr_AVVISO_ABILITATO_TXT = True Then
                MsgBox "CONTROLLO TABELLE PRECARICATE DA IMPORTARE:" & vbCrLf & _
                       "Tables importate: " & importedTABLECount & vbCrLf & importedTables_s & _
                       "Tables esistenti cancellate: " & deletedTABLECount & vbCrLf & deletedTables, vbInformation
            End If

    '//---------------------------------------------------------------------------//
    
    
    '//RECUPERO LA PATH ED IL FILE.MDB PER LE IMPORTAZIONI
    '//---------------------------------------------------------------------------//
    '//NOTE: cheiamo la funzione e recupero la path ed il file.mdb facendo un controllo _
            se la variabile è nulla msg di avviso ed uscita dalla routine.
            
        '//reset
        sourceDBPath_s = ""
        
        '//@RECUPERO@PATH_(ed il file
        sourceDBPath_s = PathFile_s_pFunct
        
        If sourceDBPath_s = "" Then
            MsgBox "ATTENZIONE ERRORE path e file sono null uscita dalla routine", vbCritical, "MSG CONTROLLO PATH E FILE VALORIZZATI"
              If MyValue = 0 Then Exit Sub
        End If
        
          ' Messaggio finale di riepilogo SE ABILITATO CON TRUE - FALSE DISABILITATO
            If Me.ctr_AVVISO_ABILITATO_TXT = True Then
                MsgBox "IMPORTO LE TABELLE DA QUESTO DATABASE; path e del file : " & sourceDBPath_s, vbExclamation, "MSG CHIAMATO LA FUNZIONE PATHFILE"
            End If
    '//=====================================================================================================================//
    '//                            STEP_01 = RECUPERO PATH E COLLECTION
    '//RECUPERO LA COLLEZIONE FORM DA IMPORTARE E FACCIO IL CONTROLLO          **** FINE ***
    '//=====================================================================================================================//
    
            '//=====================================================================================================================//
            '//                             STEP_02 = CONTROLLO PROGETTO
            '//CHIAMO IL CONTROLLO DEL PROGETTO  - 01 E 02) CONTROLLO - *** INIZIO ***
            '//=====================================================================================================================//
            '//NOTE: chiamo la funzione che controlla se il progetto corrente è il progetto originale, infatti _
                     non si possono caricare gli oggetti nel PROGETTO MSYSDbEst_GE_OBJECT.mdb perhè è originale _
                     e vi è una distruzione degli stessi. LA FUNZIONE CONTROLLO SE E' UN PROGETTO ESTERNO AMMISSIBILE _
                     per caricare gli oggetti table, query , form ecc. da importare. _
                     01) CONTROLLO = MSG DI AVVISO CON POSSIBILITA DI ANNULLARE _
                     02) CONTROLLO = CONTROLLO SE STIAMO NEL PROGETTO ORIGINALE  E QUINDI ESCLUSIONE.
                     
                        
                        
                   '//01) PRIMO CONTROLLO _
                     MESSAGGIO DI IMPORTAZIONE 0 = ANNULLA  1 = CONTINUA
                   '//..........................................................//
            
                         '//Attenzione L'AVVISO produce il messaggio solo se il _
                            parametro è TRUE altrimenti il messaggio viene escluso.
                         MyValue = CheckAvviso_i_PFunct(Me.ctr_AVVISO_ABILITATO_TXT)
                         
                         If MyValue = 0 Then Exit Sub
                     '//..........................................................//
            
                         
                 
                   
                   '//02) SECONDO CONTROLLO
                   '//CHIAMO IL CONTROLLO DEL PROGETTO
                   '//..........................................................//
                   '//Note  : True = siamo nel progetto originale non possiamo _
                              importare nessun oggetto perche distruggiamo il db corrente. _
                              Fale = siamo in un progetto esterno possiamo caricare tutto.
                                  
                                  
                       '//SE VIENE RESTITUITO TRUE è FUORI PROGETTO PUO CARICARE, SE FALSE EXIT PERCHE PROGETTO ORIGINALE
                       Bool1 = CheckProjectName_b_PFunct
                       
                       If Bool1 = True Then Exit Sub
                   '//..........................................................//
                   
             
            '//=====================================================================================================================//
            '//                             STEP_02 = CONTROLLO PROGETTO
            '//CHIAMO IL CONTROLLO DEL PROGETTO  - 01 E 02) CONTROLLO - *** FINE ***
            '//=====================================================================================================================//
            
    
       
       
   
    
    '//=====================================================================================================================//
    '//                            STEP_03 = IMPORTO LE TABELLE
    '//
    '//=====================================================================================================================//
            
            
         
    
                '//RESET
                '//Inizializza i contatori e le stringhe di riepilogo
                importedTablesCount_i = 0
                deletedTablesCount_i = 0
                deletedTables_s = ""
                importedTables_s = ""
                
          
            ' Apri il database corrente
            Set dbCurrent = CurrentDb
            
            ' Apri il database di origine
            Set dbSource = DBEngine.Workspaces(0).OpenDatabase(sourceDBPath_s)
            
            ' Loop attraverso tutte le tabelle nel database di origine
            For Each tdf In dbSource.TableDefs
                
                ' Salta le tabelle di sistema (i cui nomi iniziano con "MSys") aggiungo come parametro( la collezione table + nome tabella)
                If Left(tdf.Name, 4) <> "MSys" Or (Left(tdf.Name, 4) = "MSys" And InCollectionTables_Funct_b(objectTables, tdf.Name)) Then
                    ' Nome della tabella nel database di origine
                    strTableName = tdf.Name
                    
                    ' Controlla se la tabella è nella collezione
                    tableExistsInCollectionTables_Funct_b = False
                    For Each varItem In objectTables
                        If strTableName = varItem Then
                            tableExistsInCollectionTables_Funct_b = True
                            Exit For
                        End If
                    Next varItem
                    
                    ' Se la tabella è nella collezione, importala
                    If tableExistsInCollectionTables_Funct_b Then
                        ' Elimina la tabella se esiste già nel database corrente
                        On Error Resume Next
                        dbCurrent.TableDefs.Delete strTableName
                        deletedTablesCount_i = deletedTablesCount_i + 1
                        
                        On Error GoTo 0
                        
                        ' Copia la tabella dal database di origine al database corrente
                        DoCmd.TransferDatabase acImport, "Microsoft Access", sourceDBPath_s, acTable, strTableName, strTableName
                        
                           '//conteggio TABELLE IMPORTATE
                            importedTablesCount_i = importedTablesCount_i + 1
                            importedTables_s = importedTables_s & strTableName & vbCrLf
                
                        
                    End If
                End If
            Next tdf
            
                    
        '//TODO: MANCA IL CONTROLLO SULLA TABELLA DI SISTEMA DEL DB ESTERNO MSysObjects DA CREARE VEDI FUNZIONE Cmd_Importa_MODULI_Click
        '//----------------------------------------------------------------------------------------------------------------------//
        
            '//TODO: adattare il controllo sulla tabella di sistema esterna ... vai a Cmd_Importa_MODULI_Click

        '//----------------------------------------------------------------------------------------------------------------------//
                    
                      ' Messaggio finale di riepilogo SE ABILITATO CON TRUE - FALSE DISABILITATO
                        If Me.ctr_AVVISO_ABILITATO_TXT = True Then
                            MsgBox "CONTROLLO TABELLE PRECARICATE DA IMPORTARE:" & vbCrLf & _
                                   "Tables importate: " & importedTablesCount_i & vbCrLf & DELETETableCount_i & _
                                   "Tables esistenti cancellate: " & deletedTablesCount_i & vbCrLf & deletedTables_s, vbInformation
                        End If
            
                    
                    
                    Set dbCurrent = Nothing
                    
                '//ciclo nella collezione per il controllo delle TABELLE precaricate per SELEZIONARE LA PRIMA
                '//--------------------------------------------------------------------------------------//
                    For Each objCollection_objectName_v In tablesCollection
                    
                        
                        '//vado all'ultima query
                        DoCmd.SelectObject acTable, objCollection_tableName_v, True
                        
                    Next objCollection_objectName_v
                                
                    '//PRIMA DI USCIRE RILASCIO GLI OGGETTI
                    Set dbCurrent = Nothing
                    Set objectQueries = Nothing
                    
                '//--------------------------------------------------------------------------------------//
                
                
                    
  

                    
            
    '//=====================================================================================================================//
    '//                            STEP_03 = IMPORTO LE TABELLE *** fine ***
    '//
    '//=====================================================================================================================//


'// CHIUSURA E CONTROLLO ERRORI
'//------------------------------------------------------------------------------------------------//
Exit Sub

ErrorHandler:
    MsgBox "Errore durante l'importazione: " & Err.Description, vbCritical
    Debug.Print "Errore durante l'importazione: " & Err.Description, vbCritical
    If Not sourceDB Is Nothing Then
        sourceDB.Close
    End If
    Set sourceDB = Nothing
    Set ws = Nothing
End Sub   '//Private Sub Cmd_Importa_TABLE_Click()


' Funzione per verificare se un elemento esiste nella collezione
Function InCollectionTables_Funct_b(col As Collection, key As String) As Boolean
    Dim varItem As Variant
    On Error Resume Next
    For Each varItem In col
        If varItem = key Then
            InCollectionTables_Funct_b = True
            Exit Function
        End If
    Next varItem
    InCollectionTables_Funct_b = False
    On Error GoTo 0
End Function



'//IMPORTA GLI OGGETTI TABLE DEL DB ESTERNO IN TABELLA  *** FINE ***
'//----------------------------------------------------------------------------------------//




'//CANCELLA GLI OGGETTI TABLE *** INIZIO *** '//DELETE TABLES
'//==========================================================================================================//

Private Sub Cmd_Delete_TABLE_Click()

    Dim dbCurrent As DAO.Database
    Dim tdf As DAO.TableDef
    Dim objectTables As Collection
    Dim varItem As Variant
    
'//TODO: controllare se tale routine funziona???
On Error GoTo Err_Cmd_Delete_TABLE_Click
    
    
    '//RESET
    deletedTablesCount_i = 0
    deletedTables_s = ""
   
        
              '//ATTIVO LA COLLEZIONE
              '//------------------------------------------------------------------------//
              '//Note: CHIAMO LA FUNZIONE COLLECTION TABLE per ottenere la collezione di TABELLE
                
                '//
                Set tablesCollection = CollectionTable_PFunct()
                
              '//------------------------------------------------------------------------//
        
    
    ' Apri il database corrente
    Set dbCurrent = CurrentDb
    
    ' Loop attraverso tutte le tabelle nel database corrente
    For Each tdf In dbCurrent.TableDefs
        ' Controlla se la tabella è nella collezione e passo tutta la collezione objectTables + il nome della tabella da controllare
        If InCollectionTables_Funct_b(tablesCollection, tdf.Name) Then
            ' Elimina la tabella
            On Error Resume Next
                deletedTablesCount_i = deletedTablesCount_i + 1
                DoCmd.DeleteObject acTable, tdf.Name
            On Error GoTo 0
        End If
    Next tdf
    
    
            ' Messaggio finale di riepilogo SE ABILITATO CON TRUE - FALSE DISABILITATO
            If Me.ctr_AVVISO_ABILITATO_TXT = True Then
                MsgBox "LE TABELLE CANCELLATA DA QUESTO DATABASE SONO :  " & deletedTablesCount_i, vbExclamation, "MSG CHIAMATO LA FUNZIONE PATHFILE"
            End If


                '




'//USCITA  E GESTIONE ERRORI
'//..............................................................................................................

Exit_Cmd_Delete_TABLE_Click:
    Exit Sub

Err_Cmd_Delete_TABLE_Click:
        
    Set dbCurrent = Nothing
    
    MsgBox Err.Description & " - Errore Messaggio -> : " & ProceduraMessaggioErrore_s & " Procedura -> : " & ProceduraMessaggioErrore_s
    Debug.Print ProceduraMessaggioErrore_s
    Debug.Print ProceduraAttivaEseguita_s
    Stop
    Resume Exit_Cmd_Delete_TABLE_Click

End Sub




'//*** FINE ***
'//SENZA_PARAMETRI
'//========================================================================================================================================//


'//DELETE TABLES        ***FINE ***
'//==========================================================================================================//










'//@HELP@IMPORTA@TABELLE
Private Sub Importa_TABLE_HELP_Click()



    '//APRO FILE PDF SPECIFICO CON IL COMANDO OGGETTO PDF - @APRI.FILE
    '//--------------------------------------------------------------------------------//--------//
    '//NOTE                 -> Apro il file di tipo doc, zip o pdf
    '//CODICE               -> Function ApriFilePdf.01.01
    '//PARAMETRI            -> par_Path_s         = PATH _
                            -> par_NameFile_s               = NOME FILE _
                            -> par_IDGestione_lng           = ID FILE DA RICERCARE per futuri utilizzi
                            
      '//IMPOSTO LE VARIABILI PATH E FILE
      '//
      Dim MyFile_s As String
      '//
      Dim MyPath_s  As String
      
            '//imposto i parametri - LA PATH
            '
            MyPath_s = "c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MDB\OBJECT\HELP\"
            '//IL FILE = attenzione al file ho lasciato un spazio perche a volte non funziona senza
            '
            MyFile_s = "Help_Importa_TABLE.pdf "
                
                
            
            
         '//chiamo la funzione INTERNA DI APERTURA DEL FILE con i parametri =   'CALL (ApriFilePdf Path_programma_s, NameFile_s, IDGestione_lng) _
            Il vantaggio della procedura interna è quello di evitare di portare dietro al progetto IMPORTA OGGETTI incluso in questa form _
            anche i moduli e classi che potrebbero essere oggetto di duplicazione con l'importazione e cancellazione in caso di resettamento _
            totale.
            Call ApriFilePdf_Interna_s_PFunct(MyPath_s, MyFile_s, 0)
             
            
    '//--------------------------------------------------------------------------------//--------//

      
End Sub




'// TODO: da utilizzare, procedura non attiva
Private Sub Help_Importa_TABLE_GotFocus()
 
'//APRO FILE PDF SPECIFICO CON IL COMANDO OGGETTO PDF
'//--------------------------------------------------------------------------------//--------//
'//NOTE                 -> Apro il file di tipo doc, zip o pdf
'//CODICE               -> Function ApriFilePdf.01.01
'//PARAMETRI            -> par_Path_s         = PATH _
                        -> par_NameFile_s               = NOME FILE _
                        -> par_IDGestione_lng           = ID FILE DA RICERCARE per futuri utilizzi
        
        '//IMPOSTO LE VARIABILI PATH E FILE
        '//
        Dim MyFile_s As String
        '//
        Dim MyPath_s  As String
        
        
        '//imposto i parametri - LA PATH
        MyPath_s = "c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MDB\OBJECT\HELP\"
        
        '//IL FILE = attenzione al file ho lasciato un spazio perche a volte non funziona senza
        MyFile_s = "Help_Importa_TABLE.pdf "
            
                                 

         '//chiamo la sub con i parametri =   'CALL (ApriFilePdf Path_programma_s, NameFile_s, IDGestione_lng)
            Call ApriFilePdf(MyPath_s, MyFile_s, 0)
             
        
'//--------------------------------------------------------------------------------//--------//
                        
    
End Sub






'//CANCELLA GLI OGGETTI TABELLE *** INIZIO *** DELETE TABLE
'//==========================================================================================================//
'//@GESTIONE@DELLE@TABELLE    (@MODELLO@DELETE@OGGETTI)

'// PARAMETRI DA CAMBIARE : _
    -----> TABELLE _
    -----> DELLE@TABELLE _
    -----> DELETE TABLE _
    -----> CollectionTable_PFunct _
    -----> objectTable _
    -----> MATRICE_TABLE_Names_s _
    -----> tableiCount _
    -----> TableDef _
    -----> tdf _


Private Sub Cmd_Delete_TABELLE_Click()

    Dim dbCurrent As DAO.Database
    Dim qry As DAO.TableDef
    Dim objectQueries As Collection
    Dim varItem As Variant
    Dim MATRICE_TABLE_Names_s() As String             'matrice di TABELLE
    Dim tempCount As Integer
    Dim tableiCount As Integer
    Dim tdf As TableDef



'//TODO: controllare se tale routine funziona???
On Error GoTo Err_Cmd_Delete_TABELLE_Click


    '//RESET
    icount = 0
    deleteTABELLEesCount_i = 0
    deletedTablesCount_i = 0
    deletedTables_s = ""


          '//ATTIVO LA COLLEZIONE
         '//------------------------------------------------------------------------//
          '//Note: CHIAMO LA FUNZIONE COLLECTION TABELLE per ottenere la collezione di TABELLE



        '//
        Set tablesCollection = CollectionTable_PFunct()

          '//------------------------------------------------------------------------//

            ' Loop attraverso tutte I REPORT nel database corrente
             Set dbCurrent = CurrentDb

             '//CALCOLO INDICE TABELLE
             icount = dbCurrent.TableDefs.Count

            Set dbCurrent = CurrentDb


            '//se ci sono TABELLE da cancellare
            If icount > 0 Then

            '//reimposto la matrice sulle TABELLE esistenti nel db corrente
            ReDim MATRICE_TABLE_Names_s(icount - 1)

            ' Costruisci un array con i nomi dei TABELLE
            tempCount = 0

            '//CARICO LA MATRICE: ciclo for per caricare MATRICE_TABLE_Names_s(tempCount)
            For Each tdf In dbCurrent.TableDefs
                MATRICE_TABLE_Names_s(tempCount) = tdf.Name
                tempCount = tempCount + 1
            Next tdf

            ' Cancella i TABELLE usando l'array dei nomi
            For i = 0 To tempCount - 1
                On Error Resume Next  ' Aggiungi gestione errori specifica per ogni iterazione
                Debug.Print "Cancellazione TABELLE: " & MATRICE_TABLE_Names_s(i)

                '//CONTROLLO MATRICE = COLLECTION
                Bool1 = InCollectionTABELLEes_Funct_b(tablesCollection, MATRICE_TABLE_Names_s(i))

                '//cancello solo se la matrice (i) esiste nella collection
                If Bool1 = True Then
                    Debug.Print
                    Debug.Print "cancello la TABELLE trovata sia nella collezione che nella matrice : " & MATRICE_TABLE_Names_s(i)
                    DoCmd.DeleteObject acTABELLE, MATRICE_TABLE_Names_s(i)

                    If Err.Number <> 0 Then
                    Debug.Print "Errore durante la cancellazione del TABELLE: " & MATRICE_TABLE_Names_s(i) & " - " & Err.Description
                    Err.Clear  ' Pulisci l'errore per la prossima iterazione
                    Else
                    deleteTABELLEesCount_i = deleteTABELLEesCount_i + 1
                    End If

                End If '//If Bool1 = True Then

                On Error GoTo 0  ' Reimposta gestione errori
            Next i
            End If

        ' Messaggio finale di riepilogo SE ABILITATO CON TRUE - FALSE DISABILITATO
        If Me.ctr_AVVISO_ABILITATO_TXT = True Then
        MsgBox "LE TABELLE CANCELLATA DA QUESTO DATABASE SONO :  " & deleteTABELLEesCount_i, vbExclamation, "MSG CHIAMATO LA FUNZIONE PATHFILE"
        End If





'//USCITA  E GESTIONE ERRORI
'//..............................................................................................................

Exit_Cmd_Delete_TABELLE_Click:
    Exit Sub

Err_Cmd_Delete_TABELLE_Click:

    Set dbCurrent = Nothing

    MsgBox Err.Description & " - Errore Messaggio -> : " & ProceduraMessaggioErrore_s & " Procedura -> : " & ProceduraMessaggioErrore_s
    Debug.Print ProceduraMessaggioErrore_s
    Debug.Print ProceduraAttivaEseguita_s
    Stop
    Resume Exit_Cmd_Delete_TABELLE_Click

End Sub

'//*** FINE ***
'//SENZA_PARAMETRI
'//========================================================================================================================================//



' Funzione per verificare se un elemento esiste nella collezione
Function InCollectionTABELLEes_Funct_b(col As Collection, key As String) As Boolean
    Dim varItem As Variant
    On Error Resume Next
    '//ITERNO NELLA COLLECTION PASSATA COME PARAMETRO
    For Each varItem In col

    Debug.Print
    Debug.Print "CONTROLLO COLLEZIONE = MATRICE (I)"
    Debug.Print "collezione: " & varItem & " = " & "matrice key: " & key
    Debug.Print
    If varItem = key Then
        InCollectionTABELLEes_Funct_b = True
        Debug.Print "trovato restituisco : " & InCollectionTABELLEes_Funct_b
        Exit Function
    End If
    Next varItem
    InCollectionTABELLEes_Funct_b = False
    On Error GoTo 0
End Function



'//CANCELLA GLI OGGETTI TABELLE *** FINE *** DELETE TABLE
'//=========================================================================================================






'//=========================================================================================//
'//
'//                 GESTIONE DELLE TABELLE ESTERNE O INTERNE   *** FINE ***
'//
'//=========================================================================================//





'//=========================================================================================//
'//
'//                 GESTIONE DELI MODULI ESTERNE O INTERNE   *** inizio ***
'//@GESTIONE@DELLE@MODULI                    @GESTONE@MODULI
'//=========================================================================================//


'//IMPORTA GLI OGGETTI MODULI *** INIZIO *** '@IMPORT@MODULI
'//==========================================================================================================//
Private Sub Cmd_Importa_MODULI_Click()
    Dim obj As AccessObject
    Dim dbs As Object
    Dim ProceduraMessaggioErrore_s As String
    Dim modulesCollection As Collection
    Dim modulesCollection_v As Variant
    Dim moduleExists As Boolean
    Dim deletedModulesCount_i As Integer
    Dim deletetModulesCount_i As Integer
    Dim deletedModules_s As String
    Dim importedCount As Integer
    Dim importedModules_s As String
    Dim sourceDBPath As String
    Dim externalDB As DAO.Database
    Dim rs As DAO.Recordset
    
    On Error GoTo Err_Cmd_Import_MODULI_Click
   
    ' Reset count
    deletedModulesCount_i = 0
    deletetModulesCount_i = 0
    deletedModules_s = ""
    importedCount = 0
    importedModules_s = ""
    
    ' Recupero la collezione moduli da importare
    Set modulesCollection = CollectionModuli_PFunct()
    
    ' Recupero la path e il file dal database esterno
    sourceDBPath = PathFile_s_pFunct
    
    If sourceDBPath = "" Then
        MsgBox "ATTENZIONE ERRORE: path e file sono null, uscita dalla routine", vbCritical, "MSG CONTROLLO PATH E FILE VALORIZZATI"
        Exit Sub
    End If
    
    If Me.ctr_AVVISO_ABILITATO_TXT = True Then
        MsgBox "IMPORTO LE FORM DA QUESTO DATABASE; path e del file: " & sourceDBPath, vbExclamation, "MSG CHIAMATO LA FUNZIONE PATHFILE"
    End If
    
    ' Apro il database esterno
    Set externalDB = DBEngine.Workspaces(0).OpenDatabase(sourceDBPath)
    
    ' Ciclo nella collezione per il controllo dei moduli precaricati
    For Each modulesCollection_v In modulesCollection
        Debug.Print "I MODULI nella collezione: " & modulesCollection_v
        
        ' Verifica se il modulo esiste nel database corrente
        moduleExists = False
        Set dbs = Application.CurrentProject
        For Each obj In dbs.AllModules
            If obj.Name = modulesCollection_v Then
                moduleExists = True
                Exit For
            End If
        Next obj
        
        If moduleExists Then
            ' Se il modulo esiste nel database corrente, cancellalo
            DoCmd.DeleteObject acModule, modulesCollection_v
            deletetModulesCount_i = deletetModulesCount_i + 1
            deletedModules_s = deletedModules_s & modulesCollection_v & vbCrLf
            Debug.Print "Modulo eliminato: " & modulesCollection_v
        End If
        
        
        
            '//CONTROLLO SULLA TABELLA DI SISTEMA DEL DB ESTERNO MSysObjects
            '//----------------------------------------------------------------------------------------------------------------------//
            ' Verifica se il modulo esiste nel database esterno
            ' La riga SELECT Name FROM MSysObjects WHERE Type=-32761 è una query SQL utilizzata per selezionare i
            ' nomi degli oggetti di tipo modulo nel database di Access. Ecco una spiegazione dettagliata:
    
            ' MSysObjects: è una tabella di sistema in Microsoft Access che contiene informazioni su tutti gli oggetti nel database, _
              come tabelle, query, moduli, report, ecc. Type=-32761: questo filtro nella clausola WHERE specifica il tipo di oggetto _
              da selezionare. In Access, ogni tipo di oggetto ha un valore numerico associato nella tabella MSysObjects. _
              Il valore -32761 è il codice che rappresenta i moduli (VBA Modules).
             'Quindi, la query SELECT Name FROM MSysObjects WHERE Type=-32761 seleziona tutti i nomi degli oggetti nella tabella _
             di sistema MSysObjects che sono identificati come moduli VBA.
            
            
                moduleExists = False
                Set rs = externalDB.OpenRecordset("SELECT Name FROM MSysObjects WHERE Type=-32761 AND Name='" & modulesCollection_v & "'")
                If Not rs.EOF Then
                    moduleExists = True
                End If
                rs.Close
                
                If moduleExists Then
                    ' Se il modulo esiste nel database esterno, importalo
                    DoCmd.TransferDatabase acImport, "Microsoft Access", sourceDBPath, acModule, modulesCollection_v, modulesCollection_v
                    importedCount = importedCount + 1
                    importedModules_s = importedModules_s & modulesCollection_v & vbCrLf
                    Debug.Print "Modulo importato: " & modulesCollection_v
                Else
                    ' Modulo non trovato nel database esterno
                    Debug.Print "Modulo non trovato nel database esterno: " & modulesCollection_v
                End If
            '//----------------------------------------------------------------------------------------------------------------------//
        
    Next modulesCollection_v
    
    ' Chiudo il database esterno
    externalDB.Close
    
    ' Messaggio finale di riepilogo
    If Me.ctr_AVVISO_ABILITATO_TXT = True Then
        MsgBox "CONTROLLO FORM PRECARICATE DA IMPORTARE:" & vbCrLf & _
               "Modules importate: " & importedCount & vbCrLf & importedModules_s & _
               "Modules esistenti cancellate: " & deletetModulesCount_i & vbCrLf & deletedModules_s, vbInformation
    End If
    
                  '//ciclo nella collezione per il controllo delle MODULI precaricate per SELEZIONARE LA PRIMA
                '//--------------------------------------------------------------------------------------//
                     For Each obj In dbs.AllModules
                            '//vado all'ultima query
                        DoCmd.SelectObject acModule, obj.Name, True
                    
                     Next obj
                '//--------------------------------------------------------------------------------------//

    
    

Exit_Cmd_Import_MODULI_Click:
    Exit Sub

Err_Cmd_Import_MODULI_Click:
    MsgBox Err.Description & " - Error Message -> : " & ProceduraMessaggioErrore_s
    Debug.Print ProceduraMessaggioErrore_s
    Stop
    Resume Exit_Cmd_Import_MODULI_Click

End Sub


'//IMPORTA GLI OGGETTI MODULI *** INIZIO *** '@IMPORT@MODULI
'//==========================================================================================================//




'//CANCELLA GLI OGGETTI MODULI *** INIZIO *** '//@DELETE MODULI
'//==========================================================================================================//


Private Sub Cmd_Delete_MODULI_Click()
    Dim obj As AccessObject
    Dim dbs As Object
    Dim ProceduraMessaggioErrore_s As String
    Dim modulesCollection As Collection
    'Dim modulesCollection_v As Variant
    Dim moduleExists As Boolean
    'Dim deletedModulesCount_i As Integer
    'Dim deletetModulesCount_i As Integer
    'Dim deletedModules_s As String
    'Dim importedModules_s As String
    
    On Error GoTo Err_Cmd_Delete_MODULI_Click
   
    ' Reset count
    deletedModulesCount_i = 0
    deletetModulesCount_i = 0
    deletedModules_s = ""
    importedCount = 0
    importedModules_s = ""
     
    ' Recupero la collezione moduli da importare
    Set modulesCollection = CollectionModuli_PFunct()
                        
    ' Ciclo nella collezione per il controllo delle form precaricate
    For Each modulesCollection_v In modulesCollection
        Debug.Print "I MODULI nella collezione: " & modulesCollection_v
        
        ' Verifica se il modulo esiste nel database
        moduleExists = False
        Set dbs = Application.CurrentProject
        
        '//itero nell'insime ALLMODULES
        For Each obj In dbs.AllModules
        
            '//CONFRONTA SE ESISTE IMPOSTA TRUE
            If obj.Name = modulesCollection_v Then
                moduleExists = True
                Exit For
            End If
        Next obj
        
        '//TRUE= ESISTE E CANCELLO IL MODULO
        If moduleExists Then
            ' Se il modulo esiste, cancellalo
            DoCmd.DeleteObject acModule, modulesCollection_v
            deletetModulesCount_i = deletetModulesCount_i + 1
            deletedModules_s = deletedModules_s & modulesCollection_v & vbCrLf
            Debug.Print "Modulo eliminato: " & modulesCollection_v
        Else
            ' Se il modulo non esiste, considera per l'importazione
            importedModulesCount_i = importedModulesCount_i + 1
            importedModules_s = importedModules_s & modulesCollection_v & vbCrLf
        End If
    Next modulesCollection_v
                        
    ' Messaggio finale di riepilogo
    If Me.ctr_AVVISO_ABILITATO_TXT = True Then
        MsgBox "CONTROLLO MODULI PRECARICATE DA IMPORTARE:" & vbCrLf & _
               "Modules importate: " & importedModulesCount_i & vbCrLf & importedModules_s & _
               "Modules esistenti cancellate: " & deletetModulesCount_i & vbCrLf & deletedModules_s, vbInformation
    End If
    
     ' Messaggio finale di riepilogo
    If Me.ctr_AVVISO_ABILITATO_TXT = True Then
   
        ' Optional: display a message indicating that the modules have been deleted
        MsgBox "Tutti i moduli standard sono stati cancellati. Totale moduli cancellati: " & deletetModulesCount_i, vbInformation, "Moduli Cancellati"
    
    End If

Exit_Cmd_Delete_MODULI_Click:
    Exit Sub

Err_Cmd_Delete_MODULI_Click:
    MsgBox Err.Description & " - Error Message -> : " & ProceduraMessaggioErrore_s
    Debug.Print ProceduraMessaggioErrore_s
    Stop
    Resume Exit_Cmd_Delete_MODULI_Click

End Sub

'//DELETE MODULI        ***FINE ***
'//==========================================================================================================//


Private Sub Cmd_DELETE_MODULI_HELP_Click()

    '//APRO FILE PDF SPECIFICO CON IL COMANDO OGGETTO PDF - @APRI.FILE
    '//--------------------------------------------------------------------------------//--------//
    '//NOTE                 -> Apro il file di tipo doc, zip o pdf
    '//CODICE               -> Function ApriFilePdf.01.01
    '//PARAMETRI            -> par_Path_s         = PATH _
                            -> par_NameFile_s               = NOME FILE _
                            -> par_IDGestione_lng           = ID FILE DA RICERCARE per futuri utilizzi
                           
      '//IMPOSTO LE VARIABILI PATH E FILE
      '//
      Dim MyFile_s As String
      '//
      Dim MyPath_s  As String
     
            '//imposto i parametri - LA PATH
            '
            MyPath_s = "c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MDB\OBJECT\HELP\"
            '//IL FILE = attenzione al file ho lasciato un spazio perche a volte non funziona senza
            '
            MyFile_s = "HELP_Cmd_Delete_MODULI.pdf "
               
               
             '//CHIAMO LA FUNZIONE INTERNA PER L'APERTURA DEI FILE
                Call ApriFilePdf_Interna_s_PFunct(MyPath_s, MyFile_s, 0)
                 
           
    '//--------------------------------------------------------------------------------//--------//




End Sub

'//MODULI DELETE  *** FINE ***
'//-----------------------------------------------------------------------------------//







'//************************************************************************************************************************//
'//***
'//                 GESTIONE DELI MODULI ESTERNE O INTERNE   *** FINE ***
'//************************************************************************************************************************//



'//=========================================================================================//
'//
'//                 GESTIONE DEI MODULI ESTERNI O INTERNI   *** FINE ***
'//
'//=========================================================================================//

















'//*********************************************************************************************//
'//***
'//***                                            VARIE FUNZIONI
'//*********************************************************************************************//


'//CONTROLLO DEL PROGETTO CORRENTE *** INIZIO ***
'//----------------------------------------------------------------------------------------//
'//NOTE: QUESTA ROUTINE  controlla se il progetto in cui si trova la form carica oggetti, _
    è quello di origine ossia il database MSYSDbEst_GE_OBJECT.mdb. Interroganto la _
    proprita name se ottengo il nome del DB DI ORIGINE non eseguo l'importazione perchè altrimenti _
    distrutto il db di origine. QUESTE FUNZIONI SI POSSONO UTILIZZARE SOLO IN DB DIVERSI DA _
    QUELLO DI ORIGINE.

Function CheckProjectName_b_PFunct() As Boolean
    
    '//CHIAMO IL CONTROLLO DEL PROGETTO
    '//..........................................................//
    '//Note  : True = siamo nel progetto originale non possiamo _
               importare nessun oggetto perche distruggiamo il db corrente. _
               Fale = siamo in un progetto esterno possiamo caricare tutto.
                   
        '//solo se FALSE attivo la procedura
        'Bool1 = CheckProjectName_b_PFunct
    '//..........................................................//
    
        
    ' Variabile per memorizzare il nome del progetto corrente
    Dim projectName As String
    
    ' Ottieni il nome del progetto corrente
    projectName = Application.CurrentProject.Name
    
    ' Controlla se il nome del DATABASE QUELLO ORIGINARIO OSSIA "MSYSDbEst_GE_OBJECT.mdb"
    If projectName = "MSYSDbEst_GE_OBJECT.mdb" Then
        MsgBox "USCITA DALLA FUNZIONE: ATTENZIONE CI TROVIAMO NEL PROGETTO ORIGINALE E NON POSSIAMO CARICARE NESSUN OGGETTO SUL PROGGETTO DI PARTENZA DENOMINATO : " & projectName, vbCritical
            Debug.Print "ATTENZIONE IL PROGETTO E' QUELLO ORIGINALE  : " & projectName & " NON POSSIAMO CARICARE NESSUNO OGGETTO!!"
            '//TRUE = progetto originale  NON SI PUO CARICARE NESSU OGGETTO
            CheckProjectName_b_PFunct = True
    Else
            '//FALSE = progetto diverso dall'originale  possiamo importare gli oggetti."
            CheckProjectName_b_PFunct = False
            Debug.Print "Il progetto si chiama diversamente: " & projectName & " POSSIAMO CARICARE GLI OGGETTI !!"
    End If
    
    
End Function

'//CONTROLLO DEL PROGETTO CORRENTE *** FINE ***
'//----------------------------------------------------------------------------------------//

'//AVVISO DI CONTINUAZIONE *** INIZIO ***
'//----------------------------------------------------------------------------------------//

Function CheckAvviso_i_PFunct(par_Bool As Boolean) As Integer
    
    '//il messaggio viene attivato solo se il parametro è TRUE
    If par_Bool = True Then
      '//AVVISO  - @MSGBOX
      Message = "ATTENZIONE QUESTO COMANDO DI IMPORTAZIONE VALE SOLO PER II PROGETTI DIVERSI  " & Chr$(13) _
                & " PER ANNULLARE INSERIRE 0"               ' Imposta il messaggio.
                Title = "AVVISO DI CONTINUAZIONE"           ' Imposta il titolo.
                Default = "1"                               ' Imposta il valore predefinito.
            
                ' Visualizza la finestra di dialogo nella posizione 100,100
                
                      MyMsg_V = InputBox(Message, Title, Default, 100, 100)
            
                
                
                
                '//CONTROLLO SE LA SCELTA = 1 O SCELTA = 0
                If MyMsg_V <> "" Or MyMsg_V <> Null Then
                    '//se scelta =0 oppure = 1 restisce un valore
                    CheckAvviso_i_PFunct = MyMsg_V
                Else
                    '//se annulla restituisce stringa = null e quindi restituisce 0 ANNULLATA
                    CheckAvviso_i_PFunct = 0
                    MyMsg_V = 0             ' 0 = SCELTA ANNULLATA ANCHE SE HAI FATTO ANNULLA
                End If
                
                
                '//AVVISO DI USCITA PERCHE SCELTO 0
                If MyMsg_V = 0 Then
                
                        '//AVVISO  DI USCITA- @MSGBOX
                        
                        MsgBox "USCITA DALLA ROUTINE", vbInformation, "AVVISO DI USCITA"
                        
   
                End If
                       
                
                
    Else
            '//CON IL MESSAGGIO DISABILITATO SI INTENDE SEMPRE 1 = ATTIVARE LA PROCEURA
             MyMsg_i = 1
            CheckAvviso_i_PFunct = MyMsg_i
        
    End If
            
End Function

'//AVVISO DI CONTINUAZIONE *** FINE ***
'//----------------------------------------------------------------------------------------//



'//APRI_IL_FILE_PDF - FUNZIONE INTERNA DI APERTURA DEI FILE -
'//==================================================================================================================//
'//METODO DI APERTURA DI UN PROGRAMMA ESTERNO O DI UN COMANDO DOS MEDIANTE "WScript.Shell"
'//PARAMETRI            -> passo 2 stringhe per parametro, la path e il nome del file.pdf
'//VALORE_DI_RITORNO    -> Nulla
'//NOTE                 -> Apro il file di tipo doc, zip o pdf
'//CODICE               -> Function ApriFilePdf_Interna_s_PFunct.01.00

Public Function ApriFilePdf_Interna_s_PFunct(par_Path_s As String, par_NameFile_s As String, par_IDGestione_lng As Long) As String

Dim Path_programma_s As String
Dim NameFile_s As String
Dim Stringa1 As String

On Error GoTo ApriFilePdf_Interna_s_PFunct_Err
        
    
'//APRO FILE PDF SPECIFICO CON IL COMANDO OGGETTO PDF - @APRI.FILE
'//--------------------------------------------------------------------------------//--------//
'//NOTE                 -> Apro il file di tipo doc, zip o pdf
'//CODICE               -> Function ApriFilePdf_Interna_s_PFunct.01.01
'//PARAMETRI            -> par_Path_s         = PATH _
                        -> par_NameFile_s               = NOME FILE _
                        -> par_IDGestione_lng           = ID FILE DA RICERCARE per futuri utilizzi
                        
  '//IMPOSTO LE VARIABILI PATH E FILE
  '//Dim MyFile_s As String
  '//Dim MyPath_s  As String
  
        '//imposto i parametri - LA PATH
        'MyPath_s = "c:\Casa\LINGUAGGI\ACCESS\PROGETTI_MDB\"
        '//IL FILE = attenzione al file ho lasciato un spazio perche a volte non funziona senza
        'MyFile_s = "Project_PROGETTI_MDB.sublime-project "
            
            
         '//chiamo la sub con i parametri =   'CALL (ApriFilePdf_Interna_s_PFunct Path_programma_s, NameFile_s, IDGestione_lng)
            'Call ApriFilePdf_Interna_s_PFunct(MyPath_s, MyFile_s, 0)
             
        
'//--------------------------------------------------------------------------------//--------//
                        
         
                        '//APRO FILE PDF SPECIFICO
                        '//>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>//
                        '//Note       : Apro il file PDF speficico indicando come paramentro la path completa
                        '//           incluso il file PDF.
                          
                          '//La path del file = Unisco la path + il file+estensione sopra individuata
                          par_Path_s = par_Path_s & par_NameFile_s & MyFileEstensione_s
                                 
                          '//controllo
                          Debug.Print par_Path_s
                                 
                          '//CREO LA SHELL come oggetto, attivo il comando "%comspec% /c start " e gli assegno la path per estesa con il nome del file
                          '// perchè Shell lancia un EXE e non il PDF, quindi è necessario costruire la stringa in modo da far lanciare
                          '// prima la sessione dos e poi il pdf..
                          '// La stringa di comando ("%comspec% /c start ") deve essere unica : comando shell + path + file
                          Set WshPDF = CreateObject("wscript.shell")
                          
                          '// Unisco il comando pdf start + path definitiva
                          Stringa1 = "%comspec% /c start " & par_Path_s                               '//Comando di apertura
                          '//Attivo il comando
                          WshPDF.Run Stringa1
                                  
                                  
                          '//libero la memoria dalle variabili e dagli oggetti creati
                          WshPDF = Null
                          Stringa1 = vbNull
                          MyFileEstensione_s = vbNull
                          MyPath_s = vbNull
                            
                            '//sistemare ??? invio tasti ???
                            Lng = 0
                            For Lng = 1 To 90000000 Step 1
                            
                            Next Lng
                            
                          'SendKeys "{ESC}"
                          'SendKeys "{ESC}"
                          'SendKeys "{ESC}"
                          
                            
                        '//>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>//

       '//CREO LA SHELL come oggetto, attivo il comando "%comspec% /c start " e gli assegno la path per estesa con il nome del file
            '// perchè Shell lancia un EXE e non il PDF, quindi è necessario costruire la stringa in modo da far lanciare
            '// prima la sessione dos e poi il pdf..
            '// La stringa di comando ("%comspec% /c start ") deve essere unica : comando shell + path + file
            'Set WshPDF = CreateObject("wscript.shell")
            
            '// Unisco il comando pdf start + path definitiva
            'Stringa1 = "%comspec% /c start " & MyPath_s & MyFile_s                               '//Comando di apertura
                '//Attivo il comando ed aggiungo @exit per @chiudere@il@terminale occorre lo spazio per evitare l'errore & " ^exit"
            'WshPDF.Run Stringa1 & " ^exit"
      
      
ApriFilePdf_Interna_s_PFunct_Exit:
    Exit Function

ApriFilePdf_Interna_s_PFunct_Err:
    MsgBox Error$
    Resume ApriFilePdf_Interna_s_PFunct_Exit

End Function

'//APRI_IL_FILE_PDF - FUNZIONE INTERNA DI APERTURA DEI FILE -         *** FINE ***
'//==================================================================================================================//





'//*** INIZIO ***
'//APRI FILE @APRO@ANALISI@SUBLIME.TEXT  @CMD@03@CMD.03@CMD.3
'//========================================================================================================================================//

Private Sub Cmd_HELP_Click()

On Error GoTo Cmd_HELP_Err
    
      
      
      
                     
        
        '//APRO SORGENTE FORM FILTRO 66 CAMPI
        '//.....................................................................................................
        '//Note           : Tramite una Select vengono individuati i valori da restiuire.
                   
            '//reset
            '//Imposto i parametri
            ProceduraAttivaEseguita_s = ""
            ProceduraMessaggioErrore_s = ""
            
            MyFile_s = ""
            MyPath_s = ""
            
            '//Imposto i parametri
            ProceduraAttivaEseguita_s = ""
            ProceduraMessaggioErrore_s = "Errore nella procedura - apri FILE GESTIONE """ & MyPath_s & MyFile_s
            
            
        
                    
                '//APRO FILE PDF SPECIFICO CON IL COMANDO OGGETTO PDF
                '//--------------------------------------------------------------------------------//--------//
                '//NOTE                 -> Apro il file di tipo doc, zip o pdf
                '//CODICE               -> Function ApriFilePdf.01.01
                '//PARAMETRI            -> par_Path_s         = PATH _
                                        -> par_NameFile_s               = NOME FILE _
                                        -> par_IDGestione_lng           = ID FILE DA RICERCARE per futuri utilizzi
                        
                        '//imposto i parametri - LA PATH
                        MyPath_s = "c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MDB\MSYS_OBJECT_DB_EST\"
                        '//IL FILE = attenzione al file ho lasciato un spazio perche a volte non funziona senza
                        MyFile_s = "Project_MSYSDbEst_GE_OBJECT_ANALISI.sublime-project "
                
                         '//chiamo la sub con i parametri
                            Call ApriFilePdf(MyPath_s, MyFile_s, 0)
                        
                '//--------------------------------------------------------------------------------//--------//
                        
                 
          
 

           
        '//*** fine ***
        '//APRO IL FILE PASSANDO COME PARAMETRO PATH E NOME
        '//.....................................................................................................
 
                    







        

'//*** fine ***
'//APRI FILE @APRO@ANALISI@SUBLIME.TEXT
'//###########################################################################################################################//


        
    '//........................................................//


'USCITA  E GESTIONE ERRORI
'..............................................................................................................

Cmd_HELP_Exit:
    Exit Sub

Cmd_HELP_Err:
    MsgBox Error$
    Resume Cmd_HELP_Exit

End Sub


'//*** FINE ***
'//APRI PATH
'//========================================================================================================================================//







'//ESEGUI QUERY SULLA TABELLA DI SISTEMA MSysObjects PER INDIVIDUARE GLI OGGETTI ATTIVI (TABELLE, QUERY FORM REPORT)
'//========================================================================================================================================//
'//Recupero del percorso del database esterno: La funzione PathFile_s_pFunct viene chiamata per ottenere il percorso del database esterno. _
    Se il percorso è vuoto, la funzione mostra un messaggio di errore e esce. Apertura del database esterno: Il database esterno viene aperto _
    utilizzando DBEngine.Workspaces(0).OpenDatabase. Esecuzione della query su MSysObjects: Viene eseguita una query su MSysObjects per _
    ottenere tutti gli oggetti che non sono nascosti (Flags=0) e che sono di tipo tabella, query, form, modulo, report o macro. _
    Ciclo attraverso i risultati: La funzione cicla attraverso i risultati della query e stampa il nome e il tipo di ciascun _
    oggetto nella finestra di debug. La variabile objectTypeDescription viene utilizzata per fornire una descrizione leggibile del tipo di oggetto. _
    Chiusura del recordset e del database: Il recordset e il database esterno vengono chiusi al termine della funzione. _
    Messaggio informativo: Viene visualizzato un messaggio informativo per indicare che la lista degli oggetti è stata stampata nella finestra di debug.


Private Sub ListExternalDBObjects()
    Dim sourceDBPath As String
    Dim externalDB As DAO.Database
    Dim rs As DAO.Recordset
    Dim objectType As Integer
    Dim objectName As String
    Dim objectTypeDescription As String
    
    ' Liste per ogni tipo di oggetto
    Dim tables As String
    Dim queries As String
    Dim forms As String
    Dim modules As String
    Dim reports As String
    Dim macros As String
    Dim others As String
    
    ' Recupero la path e il file dal database esterno
    sourceDBPath = PathFile_s_pFunct
    
    If sourceDBPath = "" Then
        MsgBox "ATTENZIONE ERRORE: path e file sono null, uscita dalla routine", vbCritical, "MSG CONTROLLO PATH E FILE VALORIZZATI"
        Exit Sub
    End If
    
    ' RESET
    Me.TXT_OBJECT_ESTERNI.Value = ""
    Me.TXT_OBJECT_ESTERNI.Requery
    
    ' Apro il database esterno
    Set externalDB = DBEngine.Workspaces(0).OpenDatabase(sourceDBPath)
    
    ' Eseguo la query su MSysObjects per ottenere tutti gli oggetti
    Set rs = externalDB.OpenRecordset("SELECT Name, Type FROM MSysObjects WHERE Flags=0 AND Type IN (1, 4, 5, -32761, -32764, -32766)")
    
    ' Ciclo attraverso i risultati e raccolgo gli oggetti nelle liste appropriate
    Do While Not rs.EOF
    
        objectName = rs!Name
        objectType = rs!Type
        
        
        '// Chr$(13) =  è una funzione in VBA che restituisce il carattere corrispondente al codice ASCII specificato. _
            In questo caso, 13 è il codice ASCII per il carattere di ritorno a capo (Carriage Return, CR), che viene spesso _
            utilizzato in combinazione con Chr$(10) (line feed, LF) per creare una nuova linea su molte piattaforme. Tuttavia, per creare una _
            nuova linea in modo più robusto e compatibile, è meglio usare vbCrLf, che combina entrambi i caratteri di controllo (Chr$(13) & Chr$(10)). _
            vbCrLf è una costante predefinita in VBA che rappresenta una nuova linea e garantisce che il testo appaia correttamente formattato _
            su tutte le piattaforme. _
            Chr$(10) = Il Line Feed (LF) è un carattere di controllo utilizzato per indicare la fine di una linea di testo e l'inizio di una nuova. _
            Nel sistema ASCII, il codice per il Line Feed è 10. Questo carattere è comunemente utilizzato nei sistemi Unix e Linux _
            per terminare le linee di testo. Quando un Line Feed viene incontrato in un testo, il cursore o il puntatore si _
            sposta all 'inizio della riga successiva. Tuttavia, il Line Feed da solo non riporta il cursore all'inizio della riga corrente, _
            per questo motivo spesso viene combinato con il Carriage Return (CR, codice ASCII 13), che invece riporta il _
            cursore all'inizio della riga corrente.0
            
        ' Descrivo il tipo di oggetto
        Select Case objectType
            Case 1
                objectTypeDescription = "Tabella"
                tables = tables & "Nome: " & objectName & ", Tipo: " & objectTypeDescription & " - codice : " & objectType & vbCrLf
            Case 4
                objectTypeDescription = "Query"
                queries = queries & "Nome: " & objectName & ", Tipo: " & objectTypeDescription & " - codice : " & objectType & vbCrLf
            Case 5
                objectTypeDescription = "Form"
                forms = forms & "Nome: " & objectName & ", Tipo: " & objectTypeDescription & " - codice : " & objectType & vbCrLf
            Case -32761
                objectTypeDescription = "Modulo"
                modules = modules & "Nome: " & objectName & ", Tipo: " & objectTypeDescription & " - codice : " & objectType & vbCrLf
            Case -32764
                objectTypeDescription = "Report"
                reports = reports & "Nome: " & objectName & ", Tipo: " & objectTypeDescription & " - codice : " & objectType & vbCrLf
            Case -32766
                objectTypeDescription = "Macro"
                macros = macros & "Nome: " & objectName & ", Tipo: " & objectTypeDescription & " - codice : " & objectType & vbCrLf
            Case Else
                objectTypeDescription = "Altro"
                others = others & "Nome: " & objectName & ", Tipo: " & objectTypeDescription & " - codice : " & objectType & vbCrLf
        End Select
        
        rs.MoveNext
    Loop
    
    ' Stampo le liste in ordine
    Debug.Print "Oggetti nel database esterno (" & sourceDBPath & "):"
    Debug.Print "-----------------------------------------------"
    Debug.Print
    
    If tables <> "" Then Debug.Print tables
    Debug.Print
    If queries <> "" Then Debug.Print queries
    Debug.Print
    If forms <> "" Then Debug.Print forms
    Debug.Print
    If modules <> "" Then Debug.Print modules
    Debug.Print
    If reports <> "" Then Debug.Print reports
    Debug.Print
    If macros <> "" Then Debug.Print macros
    Debug.Print
    If others <> "" Then Debug.Print others
    Debug.Print
    
    ' Imposto il valore del controllo di testo
    Me.TXT_OBJECT_ESTERNI.Value = tables & vbCrLf & queries & vbCrLf & forms & vbCrLf & modules & vbCrLf & reports & vbCrLf & macros & vbCrLf & others
    
    ' Chiudo il recordset e il database
    rs.Close
    externalDB.Close
    
    MsgBox "Lista degli oggetti nel database esterno stampata nella finestra di debug.", vbInformation, "Completato"
End Sub


'//*** FINE ***
'//ESEGUI QUERY SULLA TABELLA DI SISTEMA MSysObjects PER INDIVIDUARE GLI OGGETTI ATTIVI (TABELLE, QUERY FORM REPORT)
'//========================================================================================================================================//



'//CANCELLA TUTTO
Private Sub Cmd_DELETE_TUTTO_Click()
Cmd_Delete_FORM_Click
Cmd_Delete_MODULI_Click
Cmd_Delete_QUERY_Click
Cmd_Delete_REPORT_Click
Cmd_Delete_TABLE_Click
Cmd_Delete_MACRO_Click
End Sub



Private Sub Cmd_IMPORTA_TUTTO_Click()

Cmd_Importa_FORM_Click
Cmd_Importa_MODULI_Click
Cmd_Importa_QUERY_Click
Cmd_Importa_REPORT_Click
Cmd_Importa_TABLE_Click
Cmd_Importa_MACRO_Click

End Sub


'//*********************************************************************************************//
'//***
'//***                                            VARIE FUNZIONI *** FINE ***
'//*********************************************************************************************//




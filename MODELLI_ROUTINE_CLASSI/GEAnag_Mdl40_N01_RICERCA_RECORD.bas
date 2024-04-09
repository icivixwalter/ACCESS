Attribute VB_Name = "GEAnag_Mdl40_N01_RICERCA_RECORD"
'********************************************************************************************************
'*                                                                                                      *
'*                            INIZIO FORM  VISUALIZZA    :SEZ1_Frm002                                   *
'*                         LE VARIABILI GLOBALI DEL PROGETTO                                            *
'*                                                                                                      *
'*NOTE  :Visualizza la tabella gestione condomini                                                       *
'*                                                                                                      *
'*                                                                                                      *
'*                                                                                                      *
'********************************************************************************************************


'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>:>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'Option

Option Compare Text
Option Explicit

'Variabili di database
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>:>>>>>>>>>>>>>>>>>>>>>>>>>>>>

'DAO
Dim DaoDB As DAO.Database
Dim DaoWks As DAO.Workspace
Dim DaoRs As DAO.Recordset
'rs di scrittura
Dim DaoRs_Write As DAO.Recordset

'ADO
Dim ADODB As Database
Dim AdodaoRs As Recordset

'Contatori
Dim iCount As Integer
Dim dbl_count As Double

'Le variabili generiche
Dim sSql As String                                          ' Stringa di estrazione

'Variabili generali
Dim Str1 As String
Dim Int1 As Integer
Dim Lng1 As Long
Dim Dbl1 As Double
Dim Bln1 As Boolean
Dim Vv1 As Variant


'Gestione parametri condominio
Dim sxCODCOND As String
Dim ixANNOESERC As Integer
Dim sxDATAINIZIO As String
Dim sxDATAFINE As String
Dim sxGESTIONE As String


'Larghezza e numero di colonna
Dim sLarg_Col As String
Dim iNum_Col As Integer

'Ricerche
Dim SearchString  As String
Dim SearchChar As String
Dim MyPos As Integer


'Le variabili della Anagrafica
Dim sCodice  As String
Dim sCodice_Fiscale  As String
Dim sRagione_Sociale  As String
Dim sComune_Residenza  As String
Dim sIndirizzo_Residenza  As String
Dim sCivico  As String
Dim sCODPRAT As String
Dim sNRO_PRATICA As String
Dim sTrim As String
Dim sFASC As String
Dim sFALD As String

Dim sComune_Nascita  As String
Dim sDataNasc As String
Dim vDataNasc As Variant
Dim sDataMorte As String
Dim dtDATAMORTE As Date
Dim iGIORNO  As Integer
Dim iMESE  As Integer
Dim iANNO  As Integer



'Le proprieta delle Label
    
Dim sxNomeLabel As String
Dim sxNomeText  As String
Dim sxCaption As String
Dim sxFontName As String
Dim ixFontSize As Integer

'Le proprieta delle Text
Dim sxDefaultValue As String
Dim vxValue  As Variant


'Le proprieta delle LabelTIT
Dim sxNomeLabelTit As String
Dim lngxBackColorTit As Long


'Variabili DELLA FORM
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>:>>>>>>>>>>>>>>>>>>>>>>>>>>>>

Dim sxFrm_CHIAMANTE_GENERALE As String                  '... Form CHIAMANTE generale
Dim sxFrm_CHIAMANTE_GENERALE_CORRENTE As String         '... Form CHIAMANTE generale OLD
Dim sxFrm_CHIAMANTE_GENERALE_PRECEDENTE As String       '... Form CHIAMANTE precedente
Dim sxProceduraMessaggioErrore  As String
Dim sxProceduraAttivaEseguita As String
Dim bFormAperta  As Boolean
'


    'Variabili della Casella combinata
    Dim sxCmb_01  As String
    Dim sxCmb_02  As String
    Dim sxCmb_03  As String
    Dim sxCmb_04  As String


    'Variabili dei Controlli Button
    '......................................................
    'Variabile di una solo click alla volta Es. bxCmd_01=False
    'pulsante non attivato; bxCmd_01=True pulsante attivato
    'Aspettare la fine del processo in atto dove la variabile boolena sarà
    'settata a false.
    
    Dim bxCmd_01  As Boolean
    Dim bxCmd_02  As Boolean
    Dim bxCmd_03  As Boolean
    Dim bxCmd_04  As Boolean
    Dim bxCmd_05  As Boolean
    Dim bxCmd_06  As Boolean
    Dim bxCmd_07  As Boolean
    Dim bxCmd_08  As Boolean
    Dim bxCmd_09  As Boolean
    Dim bxCmd_10  As Boolean
    Dim bxCmd_11  As Boolean
    '......................................................
    


'COMANDI FORM
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::


    'GLI OGGETTI COMANDI FORM
    '--------------------------------------------------------------------
    'Oggetti generali della form
    Dim sxCODOGGETTO As String                                          'Il codice form
    Dim sxCODOGGETTO_2 As String                                        'Il codice form
    
    'ATTIVITA CMD_02
    Dim sxOpenFORM_01_Cmd_02 As String                                   'Stringa per l'oggetto Docmd Apri form per il Cmd_02
    Dim sxOpenREPORT_01_Cmd_02 As String                                 'Stringa per l'oggetto Docmd Apri Report per il Cmd_02
    
    Dim sxOpenQUERY_01_Cmd_02 As String                                  'Stringa per l'oggetto Docmd Apri Query 1 per il Cmd_02
    Dim sxOpenQUERY_02_Cmd_02 As String                                  'Stringa per l'oggetto Docmd Apri Query 2 per il Cmd_02
    Dim sxOpenQUERY_03_Cmd_02 As String                                  'Stringa per l'oggetto Docmd Apri Query 3 per il Cmd_02
    
    Dim sxEseguiRoutine_01_Cmd_02 As String                              'Stringa per l'oggetto Docmd Esegui Macro 1 per il Cmd_02
    
    
    
    'ATTIVITA CMD_03
    Dim sxOpenFORM_01_Cmd_03 As String                                   'Stringa per l'oggetto Docmd Apri form per il Cmd_03
    Dim sxOpenREPORT_01_Cmd_03 As String                                 'Stringa per l'oggetto Docmd Apri Report per il Cmd_03
    
    Dim sxOpenQUERY_01_Cmd_03 As String                                  'Stringa per l'oggetto Docmd Apri Query 1 per il Cmd_03
    Dim sxOpenQUERY_02_Cmd_03 As String                                  'Stringa per l'oggetto Docmd Apri Query 2 per il Cmd_03
    Dim sxOpenQUERY_03_Cmd_03 As String                                  'Stringa per l'oggetto Docmd Apri Query 3 per il Cmd_03
    
    Dim sxEseguiRoutine_01_Cmd_03 As String                              'Stringa per l'oggetto Docmd Esegui Macro 1 per il Cmd_03
    
    'ATTIVITA CMD_04
    Dim sxOpenFORM_01_Cmd_04 As String                                   'Stringa per l'oggetto Docmd Apri form per il Cmd_4
    Dim sxOpenREPORT_01_Cmd_04 As String                                 'Stringa per l'oggetto Docmd Apri Report per il Cmd_4
    
    Dim sxOpenQUERY_01_Cmd_04 As String                                  'Stringa per l'oggetto Docmd Apri Query 1 per il Cmd_4
    Dim sxOpenQUERY_02_Cmd_04 As String                                  'Stringa per l'oggetto Docmd Apri Query 2 per il Cmd_4
    Dim sxOpenQUERY_03_Cmd_04 As String                                  'Stringa per l'oggetto Docmd Apri Query 3 per il Cmd_4
    
    Dim sxEseguiRoutine_01_Cmd_04 As String                              'Stringa per l'oggetto Docmd Esegui Macro 1 per il Cmd_4
    
    
    'ATTIVITA CMD_05
    Dim sxOpenFORM_01_Cmd_05 As String                                   'Stringa per l'oggetto Docmd Apri form per il Cmd_5
    Dim sxOpenREPORT_01_Cmd_05 As String                                 'Stringa per l'oggetto Docmd Apri Report per il Cmd_5
    
    Dim sxOpenQUERY_01_Cmd_05 As String                                  'Stringa per l'oggetto Docmd Apri Query 1 per il Cmd_5
    Dim sxOpenQUERY_02_Cmd_05 As String                                  'Stringa per l'oggetto Docmd Apri Query 2 per il Cmd_5
    Dim sxOpenQUERY_03_Cmd_05 As String                                  'Stringa per l'oggetto Docmd Apri Query 3 per il Cmd_5
    
    Dim sxEseguiRoutine_01_Cmd_05 As String                              'Stringa per l'oggetto Docmd Esegui Macro 1 per il Cmd_5
    
    'ATTIVITA Cmd_6
    Dim sxOpenFORM_01_Cmd_06 As String                                   'Stringa per l'oggetto Docmd Apri form per il Cmd_6
    Dim sxOpenREPORT_01_Cmd_06 As String                                 'Stringa per l'oggetto Docmd Apri Report per il Cmd_6
    
    Dim sxOpenQUERY_01_Cmd_06 As String                                  'Stringa per l'oggetto Docmd Apri Query 1 per il Cmd_6
    Dim sxOpenQUERY_02_Cmd_06 As String                                  'Stringa per l'oggetto Docmd Apri Query 2 per il Cmd_6
    Dim sxOpenQUERY_03_Cmd_06 As String                                  'Stringa per l'oggetto Docmd Apri Query 3 per il Cmd_6
    
    Dim sxEseguiRoutine_01_Cmd_06 As String                              'Stringa per l'oggetto Docmd Esegui Macro 1 per il Cmd_6
    
    
    'ATTIVITA Cmd_7
    Dim sxOpenFORM_01_Cmd_07 As String                                   'Stringa per l'oggetto Docmd Apri form per il Cmd_7
    Dim sxOpenREPORT_01_Cmd_07 As String                                 'Stringa per l'oggetto Docmd Apri Report per il Cmd_7
    
    Dim sxOpenQUERY_01_Cmd_07 As String                                  'Stringa per l'oggetto Docmd Apri Query 1 per il Cmd_7
    Dim sxOpenQUERY_02_Cmd_07 As String                                  'Stringa per l'oggetto Docmd Apri Query 2 per il Cmd_7
    Dim sxOpenQUERY_03_Cmd_07 As String                                  'Stringa per l'oggetto Docmd Apri Query 3 per il Cmd_7
    
    Dim sxEseguiRoutine_01_Cmd_07 As String                              'Stringa per l'oggetto Docmd Esegui Macro 1 per il Cmd_7
    
    'ATTIVITA Cmd_8
    Dim sxOpenFORM_01_Cmd_08 As String                                   'Stringa per l'oggetto Docmd Apri form per il Cmd_8
    Dim sxOpenREPORT_01_Cmd_08 As String                                 'Stringa per l'oggetto Docmd Apri Report per il Cmd_8
    
    Dim sxOpenQUERY_01_Cmd_08 As String                                  'Stringa per l'oggetto Docmd Apri Query 1 per il Cmd_8
    Dim sxOpenQUERY_02_Cmd_08 As String                                  'Stringa per l'oggetto Docmd Apri Query 2 per il Cmd_8
    Dim sxOpenQUERY_03_Cmd_08 As String                                  'Stringa per l'oggetto Docmd Apri Query 3 per il Cmd_8
    
    Dim sxEseguiRoutine_01_Cmd_08 As String                               'Stringa per l'oggetto Docmd Esegui Macro 1 per il Cmd_8
    
    
    'ATTIVITA Cmd_9
    Dim sxOpenFORM_01_Cmd_09 As String                                   'Stringa per l'oggetto Docmd Apri form per il Cmd_9
    Dim sxOpenREPORT_01_Cmd_09 As String                                 'Stringa per l'oggetto Docmd Apri Report per il Cmd_9
    
    
    Dim sxOpenQUERY_01_Cmd_09 As String                                  'Stringa per l'oggetto Docmd Apri Query 1 per il Cmd_9
    Dim sxOpenQUERY_02_Cmd_09 As String                                  'Stringa per l'oggetto Docmd Apri Query 2 per il Cmd_9
    Dim sxOpenQUERY_03_Cmd_09 As String                                  'Stringa per l'oggetto Docmd Apri Query 3 per il Cmd_9
    
    Dim sxEseguiRoutine_01_Cmd_09 As String                              'Stringa per l'oggetto Docmd Esegui Macro 1 per il Cmd_9
    
    'ATTIVITA Cmd_10
    Dim sxOpenFORM_01_Cmd_10 As String                                  'Stringa per l'oggetto Docmd Apri form per il Cmd_10
    Dim sxOpenREPORT_01_Cmd_10 As String                                'Stringa per l'oggetto Docmd Apri Report per il Cmd_10
    
    Dim sxOpenQUERY_01_Cmd_10 As String                                 'Stringa per l'oggetto Docmd Apri Query 1 per il Cmd_10
    Dim sxOpenQUERY_02_Cmd_10 As String                                 'Stringa per l'oggetto Docmd Apri Query 2 per il Cmd_10
    Dim sxOpenQUERY_03_Cmd_10 As String                                 'Stringa per l'oggetto Docmd Apri Query 3 per il Cmd_10
    
    Dim sxEseguiRoutine_01_Cmd_10 As String                             'Stringa per l'oggetto Docmd Esegui Macro 1 per il Cmd_10
    
    'ATTIVITA Cmd_11
    Dim sxOpenFORM_01_Cmd_11 As String                                  'Stringa per l'oggetto Docmd Apri form per il Cmd_11
    Dim sxOpenREPORT_01_Cmd_11 As String                                'Stringa per l'oggetto Docmd Apri Report per il Cmd_11
    
    Dim sxOpenQUERY_01_Cmd_11 As String                                 'Stringa per l'oggetto Docmd Apri Query 1 per il Cmd_11
    Dim sxOpenQUERY_02_Cmd_11 As String                                 'Stringa per l'oggetto Docmd Apri Query 2 per il Cmd_11
    Dim sxOpenQUERY_03_Cmd_11 As String                                 'Stringa per l'oggetto Docmd Apri Query 3 per il Cmd_11
    
    Dim sxEseguiRoutine_01_Cmd_11 As String                               'Stringa per l'oggetto Docmd Esegui Macro 1 per il Cmd_11

'--------------------------------------------------------------------


'GESTIONE PROGETTI
'--------------------------------------------------------------------


    Dim sxTIPOGGETTO As String
    Dim sxTIPOGGETTO_2 As String
    
    Dim sxNOMEOGGETTO As String
    Dim sxNOMEOGGETTO_2 As String
    
    Dim sxPROPRIETA As String
    Dim sxPROPRIETA_2 As String
    
    Dim sxVALOREPROPRIETA As String
    Dim sxVALOREPROPRIETA_2 As String
    
    
    
    Dim sxMETODO    As String
    Dim sxMETODO_2    As String
    
    
    Dim sxCODATTIVITA As String
    Dim sxCODATTIVITA_2 As String
    
    Dim sxATTIVITAEVENTO As String
    
    
    Dim sxCRITERIO As String
    Dim sxCRITERIO_2 As String
    
    Dim stDocName As String                             'nome form da aprire con il comando
    Dim stLinkCriteria As String                        'criterio di ricerca record nella form da aprire con il comando
    
    
    'PAGINE TAB CONTROLL
    '-----------------------------------------------
    Dim ixPagCorrente As Integer
    Dim ixPage As Integer
    
    
    'Etichetta della pagina
    Dim sxPagCorrente_Label  As String
    'Indice pagina
    Dim ixPagCorrente_Index  As Integer
    'Nome elemento
    Dim sxPagCorrente_Name  As String
    
    'Nome elemento - Etichetta
    Dim sxPagCorrente_Caption  As String

    'I parametri del tab controll
    Dim ixCmb As Integer                            'Casella combinata : 0= Cmb_01, 1 = Cmb_02 ecc.
    Dim sxItemText  As String                       'Colonna zero casella combinata
    Dim ixCommand As Integer                        'Colonna 1 della casella combinata
    Dim vxArgument As Variant                       'Colonna 2 della casella combinata
    
    '-----------------------------------------------
    
'*** FINE ***
'Variabili DELLA FORM
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>:>>>>>>>>>>>>>>>>>>>>>>>>>>>>

'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'                                   GESTIONE MENU CON STAMPE E FUNZIONI
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<


Private Sub ProvaRoutine()

Dim sxParametro As String

        sxParametro = "RSSVTR65"
        
        RicercaRecord_pSub (sxParametro)
End Sub






'Codice     :RicercaRecord_pSub.01
Public Sub RicercaRecord_pSub(par_sxDenominazione As String)

'Modello
'....................................................
'RicercaRecord_pSub(par_sxDenominazione)



        'RICERCA SOGGETTO IN TABELLA ANAGRAFICA
        '=================================================================================================
        'Note   : Faccio una copia del rs corrente, ed eseguo una ricerca
        '       del codice del soggetto nel rs clonato, medianto il metodo
        'Codice : RicercaRecord_pSub.01.A01
                
                sSql = ""
                sSql = sSql & "SELECT TTy_ANAG_VS01_N00_ANAGRAFICA_CampiTabella.[Ragione Sociale], " & Chr$(13)
                sSql = sSql & "TTy_ANAG_VS01_N00_ANAGRAFICA_CampiTabella.[Codice Fiscale] , " & Chr$(13)
                sSql = sSql & "TTy_ANAG_VS01_N00_ANAGRAFICA_CampiTabella.Codice ," & Chr$(13)
                sSql = sSql & "TTy_ANAG_VS01_N00_ANAGRAFICA_CampiTabella.[Comune residenza] ," & Chr$(13)
                sSql = sSql & "TTy_ANAG_VS01_N00_ANAGRAFICA_CampiTabella.Prov ," & Chr$(13)
                sSql = sSql & "TTy_ANAG_VS01_N00_ANAGRAFICA_CampiTabella.[Indirizzo residenza] ," & Chr$(13)
                sSql = sSql & "TTy_ANAG_VS01_N00_ANAGRAFICA_CampiTabella.Civico ," & Chr$(13)
                sSql = sSql & "TTy_ANAG_VS01_N00_ANAGRAFICA_CampiTabella.Successioni ," & Chr$(13)
                sSql = sSql & "TTy_ANAG_VS01_N00_ANAGRAFICA_CampiTabella.Lettera ," & Chr$(13)
                sSql = sSql & "TTy_ANAG_VS01_N00_ANAGRAFICA_CampiTabella.Interno ," & Chr$(13)
                sSql = sSql & "TTy_ANAG_VS01_N00_ANAGRAFICA_CampiTabella.CAP ," & Chr$(13)
                sSql = sSql & "TTy_ANAG_VS01_N00_ANAGRAFICA_CampiTabella.[Note Has] ," & Chr$(13)
                sSql = sSql & "TTy_ANAG_VS01_N00_ANAGRAFICA_CampiTabella.[Sosp ICI] ," & Chr$(13)
                sSql = sSql & "TTy_ANAG_VS01_N00_ANAGRAFICA_CampiTabella.SospRSU ," & Chr$(13)
                sSql = sSql & "TTy_ANAG_VS01_N00_ANAGRAFICA_CampiTabella.[Sosp ICP] ," & Chr$(13)
                sSql = sSql & "TTy_ANAG_VS01_N00_ANAGRAFICA_CampiTabella.[Sosp TOSAP] ," & Chr$(13)
                sSql = sSql & "TTy_ANAG_VS01_N00_ANAGRAFICA_CampiTabella.Succ ," & Chr$(13)
                sSql = sSql & "TTy_ANAG_VS01_N00_ANAGRAFICA_CampiTabella.[Comune di nascita] ," & Chr$(13)
                sSql = sSql & "TTy_ANAG_VS01_N00_ANAGRAFICA_CampiTabella.[Partita Iva] , " & Chr$(13)
                sSql = sSql & "TTy_ANAG_VS01_N00_ANAGRAFICA_CampiTabella.Password," & Chr$(13)
                sSql = sSql & "TTy_ANAG_VS01_N00_ANAGRAFICA_CampiTabella.[Tipo Persona] ," & Chr$(13)
                sSql = sSql & "TTy_ANAG_VS01_N00_ANAGRAFICA_CampiTabella.[Cognome / Rag Sociale] ," & Chr$(13)
                sSql = sSql & "TTy_ANAG_VS01_N00_ANAGRAFICA_CampiTabella.Nome ," & Chr$(13)
                sSql = sSql & "TTy_ANAG_VS01_N00_ANAGRAFICA_CampiTabella.Sesso ," & Chr$(13)
                sSql = sSql & "TTy_ANAG_VS01_N00_ANAGRAFICA_CampiTabella.[Data Nascita] ," & Chr$(13)
                sSql = sSql & "TTy_ANAG_VS01_N00_ANAGRAFICA_CampiTabella.[Tipo Residenza] ," & Chr$(13)
                sSql = sSql & "TTy_ANAG_VS01_N00_ANAGRAFICA_CampiTabella.Scala ," & Chr$(13)
                sSql = sSql & "TTy_ANAG_VS01_N00_ANAGRAFICA_CampiTabella.Piano ," & Chr$(13)
                sSql = sSql & "TTy_ANAG_VS01_N00_ANAGRAFICA_CampiTabella.Telefono ," & Chr$(13)
                sSql = sSql & "TTy_ANAG_VS01_N00_ANAGRAFICA_CampiTabella.Fax ," & Chr$(13)
                sSql = sSql & "TTy_ANAG_VS01_N00_ANAGRAFICA_CampiTabella.[E-mail] ," & Chr$(13)
                sSql = sSql & "TTy_ANAG_VS01_N00_ANAGRAFICA_CampiTabella.[Matricola Anagrafe] ," & Chr$(13)
                sSql = sSql & "TTy_ANAG_VS01_N00_ANAGRAFICA_CampiTabella.[Codice esattoria] ," & Chr$(13)
                sSql = sSql & "TTy_ANAG_VS01_N00_ANAGRAFICA_CampiTabella.Contr ," & Chr$(13)
                sSql = sSql & "TTy_ANAG_VS01_N00_ANAGRAFICA_CampiTabella.[Codice Famiglia] ," & Chr$(13)
                sSql = sSql & "TTy_ANAG_VS01_N00_ANAGRAFICA_CampiTabella.[N Componenti] ," & Chr$(13)
                sSql = sSql & "TTy_ANAG_VS01_N00_ANAGRAFICA_CampiTabella.[Cod attivita econom] ," & Chr$(13)
                sSql = sSql & "TTy_ANAG_VS01_N00_ANAGRAFICA_CampiTabella.[Attivita economica prevalente]," & Chr$(13)
                sSql = sSql & "TTy_ANAG_VS01_N00_ANAGRAFICA_CampiTabella.[Data inizio attivita]," & Chr$(13)
                sSql = sSql & "TTy_ANAG_VS01_N00_ANAGRAFICA_CampiTabella.[Data fine attivita]," & Chr$(13)
                
                sSql = sSql & "TTy_ANAG_VS01_N00_ANAGRAFICA_CampiTabella.[FILLER_DI_GESTIONE]," & Chr$(13)
                
                sSql = sSql & "TTy_ANAG_VS01_N00_ANAGRAFICA_CampiTabella.[CodBelfiore]," & Chr$(13)
                sSql = sSql & "TTy_ANAG_VS01_N00_ANAGRAFICA_CampiTabella.[Note estese]," & Chr$(13)
                
                sSql = sSql & "TTy_ANAG_VS01_N00_ANAGRAFICA_CampiTabella.[Filler Agg]," & Chr$(13)
                sSql = sSql & "TTy_ANAG_VS01_N00_ANAGRAFICA_CampiTabella.[Data Ins]," & Chr$(13)
                sSql = sSql & "TTy_ANAG_VS01_N00_ANAGRAFICA_CampiTabella.[Data Agg]," & Chr$(13)
                sSql = sSql & "TTy_ANAG_VS01_N00_ANAGRAFICA_CampiTabella.[Time]," & Chr$(13)
                
           
                sSql = sSql & "[Ragione Sociale]  & ' ' & [Codice Fiscale] & ' ' & [Codice] AS DENOMINAZIONE "
                sSql = sSql & "FROM TTy_ANAG_VS01_N00_ANAGRAFICA_CampiTabella "
                sSql = sSql & "WHERE ((([Ragione Sociale]  & ' ' & [Codice Fiscale] & ' ' & [Codice]) Like  '*" & par_sxDenominazione & "*' ));"
                
                'controllo
                Debug.Print sSql
                                
                'RECUPERO L'ULTIMO CODICE LAVORATO
                '......................................................................
                
                        Set DaoRs = CurrentDb.OpenRecordset(sSql)
                        
                        Set DaoRs_Write = CurrentDb.OpenRecordset("TTy_ANAG_VS01_N02_ANAGRAFICA_TMP")
                                                
                        'svuoto la tabella tmp
                        DoCmd.OpenQuery ("CANCELLA_N01_ANAGRAFICA_TMP")
                        
                                                
                        'Se il rs pieno recupero l'ultimo soggetto
                        If DaoRs.EOF = False And DaoRs.BOF = False Then
                        
                        
                                                
                              
                  
                                Do While Not DaoRs.EOF
                                        
                                        'controllo window
                                        DoEvents
                                                           
                                        'trova il codice contribuente
                                        'DaoRs.FindLast "codice ='" & sCodice & "'"
                                         
                                        Debug.Print DaoRs.Fields("Codice").Value
                                        Debug.Print DaoRs.Fields("Ragione Sociale").Value
                                        Debug.Print DaoRs.Fields("Codice Fiscale").Value
                                             
                                            'INSERIMENTO DATI NELLA TABELLA TMP
                                            '.................................................................
                                                DaoRs_Write.AddNew
                                                    
                                                    DaoRs_Write.Fields("Codice").Value = DaoRs.Fields("Codice").Value
                                                    DaoRs_Write.Fields("Ragione Sociale").Value = DaoRs.Fields("Ragione Sociale").Value
                                                    DaoRs_Write.Fields("Codice Fiscale").Value = DaoRs.Fields("Codice Fiscale").Value
                                                    
                                                    'Codice comune
                                                    DaoRs_Write.Fields("CodBelfiore").Value = DaoRs.Fields("CodBelfiore").Value
                                                    
                                                    'Aggiornamento
                                                    DaoRs_Write.Fields("Data_Ins").Value = Date
                                                    DaoRs_Write.Fields("TIMEoper").Value = Now
                                                    
                                                
                                                DaoRs_Write.Update
                                            
                                            '.................................................................
                                             
                                        DaoRs.MoveNext
                                        
                                        
                                        
                                Loop
                                    
                                    
                                

                    End If
                    
                            DaoRs.Close
                            Set DaoRs = Nothing
                            
                'POSIZIONE IL CURSORE SUL RS CORRENTE
                '......................................................................

End Sub

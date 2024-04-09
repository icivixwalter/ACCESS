Attribute VB_Name = "UTIL_Mdl03_N01_CREA_TABELLA_COLLEGATA"
Option Compare Text
Option Explicit

'................................
' DIM OGGETTI E PARAMETRI
Dim obTipoDB As Object

'................................
' DIM SCELTA DATABASE DA APRIRE
Dim sScelta_db As String
Dim sOption_SEZ As Integer          'Numero Sezione di Progetto Scelta
Dim sName_Tab As String             'Nome Tabella da creare

Dim dbs As Database
Dim rs As Recordset

'................................
' DIM TABELLA DA ESPORTARE
Dim ExportTable1 As String
Dim ExportTable2 As String

'................................
' DIM CREAZIONE NUOVE TABELLE
Dim TableDefNuovo    As TableDef
Dim FieldNuovo  As Field
Dim idxNuovo_Contatore_Univoco  As Index
Dim idxNuovo_Testo_Univoco  As Index
Dim idx1    As Index                            ' per il passaggio dei parametri


'................................
' DIM CONTROLLO OGGETTI
Dim blnFlag_Oggetti As Boolean                  ' per il controllo oggetti da cancellare o creare


'................................
' DIM TABELLE COLLEGATE
Dim TableDefCollegata    As TableDef
Dim sStringConn As String

'................................
' DIM VARIABILI APERTURA DB DA AGGIORNARE
Dim ReturnValue As Variant


'................................
' DIM LE OPZIONI
Dim Opzione_1 As Integer
Dim Opzione_2 As Integer

Dim iOperazione As Integer
Dim Flag As Boolean
        
'................................
' DIM LE VARIABILI LOCALI GENERALI
Dim Vv1 As Variant
Dim sS1 As String
Dim iInt1    As Integer
Dim dblDBL1   As Double



Private Sub CREA_TABELLA_COLLEGATA()




            ' Definisco la stringa di connessione
            sStringConn = "Excel 8.0;HDR=YES;IMEX=2;DATABASE=C:\GESTIONI\GESTIONE_LLPP\25_GESTIONE_LLPP\ARCHIVIO_ELETTRONICO_PROVVISORIO.xls;TABLE='2010$'"
                              'Excel 5.0;HDR=NO;IMEX=2;DATABASE=C:\CASA\CDM\Fant2005\FANT2005.xls;TABLE='GESTIONE CORRENTE-txt$'
                            'Excel 8.0;HDR=YES;IMEX=2;DATABASE=C:\GESTIONI\GESTIONE_LLPP\25_GESTIONE_LLPP\ARCHIVIO_ELETTRONICO_PROVVISORIO.xls;TABLE='2010$'
            'Set dbs = OpenDatabase("c:\GESTIONI\GESTIONE_LLPP\25_GESTIONE_LLPP\LLPP_ARCHVI_MDB\GE_MODELLI_DbBase_COLLEGA_TABELLE.mdb")
            
            '//DATABASE CORRENTE = se il db a cui collegare la tabella è quello corrente occorre inserire la dichiarazione Set dbs = CurrentDb.
            Set dbs = CurrentDb
            
            'DATABASE=C:\CASA\LTT\LTT_AGG+TMP.mdb;TABLE=LTT_ORDINA_DATI_TMP
            ' Crea oggetto
            Set TableDefCollegata = dbs.CreateTableDef("2010_COLLEGATA")
            TableDefCollegata.Connect = sStringConn
            TableDefCollegata.SourceTableName = "'2010$'"
             
            '------------------------------------------------------
            '   Accoda la tabella al database
                dbs.TableDefs.Append TableDefCollegata
                dbs.TableDefs.Refresh
                
                dbs.Close

End Sub


Private Sub ESEMPIO_COLLEGAMENTO_CARTELLA_XLS_NEL_DATABASE_CORRENTE()

'//IL NOME DELLA TABELLA SORGENTE
Dim SourcePathNameTable_s As String

'//INDIRIZZO CORREENTE DELL'INDIRIZZO
Dim SoucerceNameTable_s As String

'//LA PAT DEL SORGENTE
SourcePathNameTable_s = "C:\GESTIONI\GESTIONE_LLPP\25_GESTIONE_LLPP\"

        
        '// Definisco la stringa di connessione per excel 2010 Excel 8.0
        '//
        sStringConn = "Excel 8.0;HDR=YES;IMEX=2;DATABASE=c:\CASA\GE_CASA\GE_MARINO\GESTIONE_SPESE\GE_CASA_SALVATAGGI\GE_CASA_Tb03_MASTRO_PARAM;TABLE='" & SourcePathNameTable_s & "$'"
        'sStringConn = "Excel 5.0;HDR=YES;IMEX=2;DATABASE=C:" & SourcePathNameTable_s & "Table = " & SourcePathNameTable_s & " '" & SourcePathNameTable_s & "$'"
        
           
        '//DATABASE CORRENTE = se il db a cui collegare la tabella è quello corrente occorre inserire la dichiarazione Set dbs = CurrentDb.
        Set dbs = CurrentDb
        
        
        ' Crea oggetto
        Set TableDefCollegata = dbs.CreateTableDef("GE_CASA_Tb03_MASTRO_PARAM")
        TableDefCollegata.Connect = sStringConn
        TableDefCollegata.SourceTableName = "'GE_CASA_Tb03_MASTRO_PARAM$'"
           
        
        
        '------------------------------------------------------
        '   Accoda la tabella al database
        dbs.TableDefs.Append TableDefCollegata
        dbs.TableDefs.Refresh
        
        
        dbs.Close
                

End Sub



'//CREA UNA TABELLA COLLEGATA SUL DATABASE CORRENTE SU EXCEL
'//PARAMETRI            ---->:  parTipoExcel 			= Tipo di excel a cui collegarsi es. Excel 8.0 oppure Excel 5.0; _
                                parDatabaseExcel        	= indirizzo completo con path e file xls. _
                                parNameCreateTableDef_s 	= Nome finale della tabella dopo il collegamento, _
                                par_SourceTableName_s   	= Nome di origine a cui fare il collegamento.
                                
Private Sub CREA_TABELLA_COLLEGATA_DbCorrente(parTipoExcel_s As String, parDatabaseExcel_s As String, _
                                              parNameCreateTableDef_s As String, par_SourceTableName_s As String)
	
	
	select case parTipoExcel_s 
	
	case "Excel 5.0"
	
		'// Definisco la stringa di connessione per excel 2010 Excel 5.0
		sStringConn ="Excel 5.0;HDR=YES;IMEX=2;DATABASE=C:\CASA\GE_CASA\GE_MARINO\GESTIONE_SPESE\GESTIONE_PROCEDURE\GE_CASA_SALVATAGGIO_ARCHIVI_XLS\GE_CASA_DF08_CODICI_BANCA.xls;TABLE=GE_CASA_DF08_CODICI_BANCA$"

	case  "parTipoExcel_s"
		'// Definisco la stringa di connessione per excel 2010 Excel 8.0
		sStringConn = "Excel 8.0;HDR=YES;IMEX=2;DATABASE=C:\GESTIONI\GESTIONE_LLPP\25_GESTIONE_LLPP\ARCHIVIO_ELETTRONICO_PROVVISORIO.xls;TABLE='2010$'"

        case Else
        	msgbox "COLLEGAMENTO EXCEL NON ESEGUITO"
        case End	
        
        '//DATABASE CORRENTE = se il db a cui collegare la tabella è quello corrente occorre inserire la dichiarazione Set dbs = CurrentDb.
        Set dbs = CurrentDb
        
        
        ' Crea oggetto
        Set TableDefCollegata = dbs.CreateTableDef("2010_COLLEGATA")
        TableDefCollegata.Connect = sStringConn
        TableDefCollegata.SourceTableName = "'2010$'"
           
        
        
        '------------------------------------------------------
        '   Accoda la tabella al database
        dbs.TableDefs.Append TableDefCollegata
        dbs.TableDefs.Refresh
        
        
        dbs.Close
                

End Sub




'Nell 'esempio riportato di seguito viene indicato come collegare la tabella Autori del database ODBC al database corrente.

'DoCmd.TransferDatabase acLink, "Database ODBC", _
'    "ODBC;DSN=OrigineDati1;UID=Utente2;PWD=www;LANGUAGE=italiano;" _
'    & "DATABASE=pub", acTable, "Autori", "dboAutori"



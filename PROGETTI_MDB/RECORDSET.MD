RECORDSET
	Note
		per aprire un recordset utilizzare queste funzioni:


	SCHEMA






					Apro il progetto con TUTTI i progetti con il @DOS
							START "APRI TUTORIAL ACCESS CON IL DOS " CALL c:\\CASA\\LINGUAGGI\\ACCESS\\ACCESS_PROGETTI_MDB\\GE_ACCESS.sublime-project"
		




				




'//COMANDO 03 APRIRE UN RECORSET PER LA VISUALIZZAZIONE DI UN FILE
	'//========================================================================================================================================//
	'//@sub@di@apertura@di@un@rs_(QUESTA routine apre un recordset con @un@filtro @SQL che permetta di estrarre solo il @record@corrente)
		Private Sub Cmd_03_Click()
		
		'//@come@apri@UN@recordset_(IN ACCESS @esempio di routine per @aprire@un@rs con parametri)




		
		'//ERRORI PROCEDURA_FUNCTION O ROUTINE sono variabile in caso di errore della procedura
		Dim ProceduraMessaggioErrore_s As String
		Dim ProceduraAttivaEseguita_s As String
		
		On Error GoTo Cmd_03_Click_Err
		
		ProceduraAttivaEseguita_s = ""
		ProceduraAttivaEseguita_s = "ESTRAI LA PATH ED FILE DA APRIRE"
		ProceduraMessaggioErrore_s = "ERRORE nel comando Cmd_03_Click"
		
		'//Apro un @area@di@lavoro@dao del db corrente ed apro un recorset
		'//RECUPERO IL FILE DA APRIRE NEL RS DELLA TABELLA GEST_DF01_TUTORIAL_SERVIZI
		
		'//@stringa@sql_(con filtro where per aprire un @recorset)'
		sSql = ""
		sSql = sSql & "SELECT  "
		sSql = sSql & "GEST_DF01_TUTORIAL_SERVIZI.NOMEOGGETTO_s, "
		sSql = sSql & "GEST_DF01_TUTORIAL_SERVIZI.NRO_OGGETTO_i, "
		sSql = sSql & "GEST_DF01_TUTORIAL_SERVIZI.FILE_s, "
		sSql = sSql & "GEST_DF01_TUTORIAL_SERVIZI.PATH_m, "
		sSql = sSql & "GEST_DF01_TUTORIAL_SERVIZI.COD_PROGETTO_s "
		sSql = sSql & "FROM "
		sSql = sSql & "GEST_DF01_TUTORIAL_SERVIZI "
		sSql = sSql & "WHERE "
		sSql = sSql & "(((GEST_DF01_TUTORIAL_SERVIZI.NOMEOGGETTO_s)='" & Me.Cmd_03.Caption & "')) "
		sSql = sSql & "WITH  "
		sSql = sSql & "OWNERACCESS OPTION;"
		
		'//imposto una variabile oggetto DB che apre un'area di lavoro in @DAO
		Set DaoDB = DBEngine.Workspaces(0).Databases(0)
		
		'//controllo ed apertura del rs con parametro
		Debug.Print
		Debug.Print sSql
		
		'//Apro un Recordset dal parametro ssql
		Set DaoRs = DaoDB.OpenRecordset(sSql)
		
		'//controllo se il rs è popolato
		If DaoRs.EOF = False And DaoRs.BOF = False Then
		'//Posizione Primo record
		DaoRs.MoveFirst
		
		
	    	'//ITERAZIONE_RECORSET_(@ESEGUO l'iterazione di un rs)
	    	'//.....................................................................................................//
	    	'//Codice---------->: MODELLO_SUB_N01_IterazioneRecord_pSub.01.01
	    	'//Note------------>: Tramite una Select vengono individuati i valori da restiuire.
		
		
	            	While Not DaoRs.EOF
	              	'//Blocco iterazione
	             	DoEvents
		
	                	'//POSIZIONE_ATTIVITA
	                	ProceduraAttivaEseguita_s = ProceduraAttivaEseguita_s & "Cmd_03_Click"
		
	                	'//IF DI CONTROLLO CAMPI
	                	'//_______________________________________________________________________
	                	'//NOTE :
		
	                	If DaoRs.Fields("NOMEOGGETTO_s") = Me.Cmd_03.Caption Then
	    	
	                    	'//CHIAMA LA PROCEDURA PRI FILE - @APRI.FILE
	                    	'//-----------------------------------------------------------------------------------------------//
	                    	'//apri file par1 = file; par2 = path
	                    	    '//se trovo il rs recupero i parametri del file e della path  e chiamo la funzione _
                                    controllo i due parametri
                                Debug.Print "file : " & DaoRs.Fields("FILE_s")
                                Debug.Print "path : " & DaoRs.Fields("PATH_m")
                            
                              
	                        
	                    			
						            '//CHIAMA LA PROCEDURA PRI FILE - @APRI.FILE
						            '//-----------------------------------------------------------------------------------------------//
						            '//apri file par1 = file; par2 = path
						                '//
						                ApriFile_Sub DaoRs.Fields("FILE_s"), _
						                DaoRs.Fields("PATH_m")
						            
						            
						            '//-----------------------------------------------------------------------------------------------//
						   
	                    	'//-----------------------------------------------------------------------------------------------//
	            	
		
	                	End If
	                	'//_______________________________________________________________________
		
	                	'//Record Successivo
	                	DaoRs.MoveNext
		
	            	Wend
	      	'//** FINE **
	      	'//ITERAZIONE_RECORSET
	      	'//.....................................................................................................//
		
	        	'//Uscita Rs e chiusura oggetti
	        	DaoRs.Close
	        	Set DaoRs = Nothing
		
		End If  '//If DAORs.EOF = False And DAORs.BOF = False Then
		
		
		
		
		'USCITA  E GESTIONE ERRORI
		'..............................................................................................................
		
		Cmd_03_Click_Exit:
	    	Exit Sub
		
		Cmd_03_Click_Err:
	    	MsgBox Error$
	    	Debug.Print ProceduraAttivaEseguita_s = "ESTRAI LA PATH ED FILE DA APRIRE"
	    	Debug.Print ProceduraMessaggioErrore_s = "ERRORE nel comando Cmd_03_Click"
		
	    	Resume Cmd_03_Click_Exit
		
		End Sub
		
		'//COMANDO 03 APRIRE UN RECORSET PER LA VISUALIZZAZIONE DI UN FILE  *** FINE ***
		'//========================================================================================================================================//
		

APRI_FILE.MD


	Nota
		procedura per aprire qualsiasi file a condizione che si passi 2 parametri, la path
		ed il file da aprire





	FUNZIONE APRI FILE CON 2 PARAMETRI

		
		'//APRI FILE   @APRI.FILE  @APRO.IL.FILE  *** INIZIO ***
			'//========================================================================================================================================//
			'//NOTE: @FUNZIONE@APRI@FILE_(che utilizza 2 parametri la @path Ed il nome del file per aprirlo)'
			
				Private Sub ApriFile_Sub(par_MyFile_s As String, par_MyPath_s As String)
				
				On Error GoTo Cmd_03_Err
			    	
			    	
			    	
			            	'//CHIAMA LA PROCEDURA PRI FILE - @APRI.FILE
			            	'//-----------------------------------------------------------------------------------------------//
			            	'//apri file par1 = file; par2 = path
			                	'//ApriFile_Sub "CIVILIA_SUBLIME_TEXT_ANALISI.sublime-project", _
			                	"c:\GESTIONI\GESTIONE_LLPP\25_GESTIONE_LLPP\LLPP_ARCHIVI_MDB\CIVILIA_GESTIONE\CIVILIA_TEXT_ANALISI\CIVILIA_ANALISI_SUBLIME_TEXT\PROGETTO_SUBLIME_TEXT\"
			            	
			            	
			            	'//-----------------------------------------------------------------------------------------------//
			            	
			            	
			            	
			             	
			            	'//Imposto i parametri
			            	ProceduraAttivaEseguita_s = "ApriFile_Sub"
			            	ProceduraMessaggioErrore_s = "Errore nella procedura - apri FILE GESTIONE """ & MyPath_s & MyFile_s
			            	
			    	
			    	
			        	'//AGGIORNAMENTO_DEL_15_07_2020
			        	'//APRO LA PATH DEGLI ATTI DI GESTIONE -
			        	'//###########################################################################################################################//
			        	'//NOTE: apro il file .mdb _
			        	non ci sono parametri _
			        	AGGIORNAMENTO DEL = 13.07.2020
			        	'//CODICE----------->APRO_PATH_Tb01_ATTI_DI_GESTIONE_pSub
			        	
			        	
			        	
			        	'//reset
			            	'//Imposto i parametri
			            	ProceduraAttivaEseguita_s = ""
			            	ProceduraMessaggioErrore_s = ""
			            	
			            	MyFile_s = ""
			            	MyPath_s = ""
			            	
			        	
			                    	
			        	
			         	
			        	
			        	'//APRO SORGENTE FORM FILTRO 66 CAMPI
			        	'//.....................................................................................................
			        	'//Note           : Tramite una Select vengono individuati i valori da restiuire.
			           	
			            	MyFile_s = par_MyFile_s
			            	'//ATTENZIONE occorre inserire uno spazio finale altrimenti non riconosce il percorso. Se invece è _
			               	un programma oppure un eseguibile non occorre lo spazio.
			            	MyPath_s = par_MyPath_s
			           	
			            	'//Imposto i parametri
			            	ProceduraAttivaEseguita_s = "ApriFile_Sub"
			            	ProceduraMessaggioErrore_s = "Errore nella procedura - apri FILE GESTIONE """ & MyPath_s & MyFile_s
			            	
			            	
			        	
			          	
			             	
			            	
			            	'//CREO LA SHELL come oggetto, attivo il comando "%comspec% /c start " e gli assegno la path per estesa con il nome del file
			            	'// perchè Shell lancia un EXE e non il PDF, quindi è necessario costruire la stringa in modo da far lanciare
			            	'// prima la sessione dos e poi il pdf..
			            	'// La stringa di comando ("%comspec% /c start ") deve essere unica : comando shell + path + file
			            	Set WshPDF = CreateObject("wscript.shell")
			            	
			            	'// Unisco il comando pdf start + path definitiva
			            	Stringa1 = "%comspec% /c start " & MyPath_s & MyFile_s                               '//Comando di apertura
			                	'//Attivo il comando ed aggiungo @exit per @chiudere@il@terminale occorre lo spazio per evitare l'errore & " ^exit"
			            	WshPDF.Run Stringa1 & " ^exit"
			        	
			        	'//*** fine ***
			        	'//APRO SORGENTE FORM FILTRO 66 CAMPI
			        	'//.....................................................................................................
			 	
			                    	
				
				
									
									
									
									
									
								        	
									
									'//*** fine ***
									'//APRO LA PATH DEGLI ATTI DI GESTIONE -
									'//###########################################################################################################################//
									
									
								        	
								    	'//........................................................//
									
									
							'USCITA  E GESTIONE ERRORI
							'..............................................................................................................
								
								Cmd_03_Exit:
							    	Exit Sub
								
								Cmd_03_Err:
							    	MsgBox Error$
							    	Resume Cmd_03_Exit
								
								End Sub
								
								
							'//*** FINE ***
							'//APRI FILE
							'//========================================================================================================================================//
							

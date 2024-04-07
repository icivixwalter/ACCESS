FUNZIONI

    CANCELLA TUTTI GLI OGGETTI MSYS
				Note
					2022
					Cancello tutti gli oggetti tramite con l'attivazione di una serie di query
					deputate alla cancellazione di ogni singolo oggetto msys.

					la funzione che attiva la cancellazione colletiva sono i seguenti:
					@CANCELLO.TUTTO 
					@CANCELLO.TABELLE+QUERY+FORM+REPORT+MODULI, 
					@Cmd_CANCELLA_TUTTO_TO_Msys_DF12
					@oggetti.cancello.tutto
					@CANCELLAZIONE.COLLETTIVA.OGGETTI
					@msys.oggetti.cancella.tutto

					Vengono usati tutti questi comandi collettivi per cancellare tutti gli oggetti con questo comando della casella combinata:
						02_CANCELLA_TUTTO_TABELLE+FORM+QUERY+REPORT+MODULI
						@CANCELLAZIONE.COLLETTIVA.OGGETTI, @msys.oggetti.cancella.tutto. @cancellazione.multipla.oggetti

					 DoCmd.OpenQuery "MSys_FORMS_Qry01_81_DELETE_TUTTE"
			                 DoCmd.OpenQuery "MSys_MODULI_Qry01_81_DELETE_TUTTE"
			                 DoCmd.OpenQuery "MSys_QUERY_Qry01_81_DELETE_TUTTE"
			                 DoCmd.OpenQuery "MSys_TABELLE_Qry01_81_DELETE_TUTTE"
			                 DoCmd.OpenQuery "MSys_REPORT_Qry01_81_DELETE_TUTTE"
			                     



    IMPORTA TUTTI GLI OGGETTI
				Note
					con la presente matrice accodo tutti gli oggetti nella tabella df12 
					pronti per l'esportazione.


					    Dim sTipo_Oggetto(5) As String
    
   						 sTipo_Oggetto (0)= "TABLE" 
   						 sTipo_Oggetto (1)= "QUERY" 
   						 sTipo_Oggetto (2)= "MODULI" 
   						 sTipo_Oggetto (3)= "FORMS" 
   						 sTipo_Oggetto (4)= "REPORTS" 
   						 sTipo_Oggetto (5)= "MACRO" 
   						 

    ACCODO LE QUERY IN DF12
    			Note
    				Accodo solo gli oggetti query della tabella oggetti MSys_QUERY nella tabella df12, esegue solo un comando singolo.
    				codici:
    					@accoda.elenco.progetti  @cmd.04_ACCODA_QUERY_TO_Msys_DF12 @comando.cmd.04_ACCODA_QUERY_TO_Msys_DF12 _
                       			@accoda.DF12, @DF12.cancella @DF12.delete _
                       			@accoda.query.to.DF12,
                    

    ACCODO I MODULI IN DF12, _
		               Note
		               		Accodo solo gli oggetti MODULI della tabella MSys_MODULI come singolo comando
		               		nella tabella DF12

		                       @accoda.elenco.progetti  @cmd.04_ACCODA_MODULI_TO_Msys_DF12 @comando.cmd.04_ACCODA_MODULI_TO_Msys_DF12 _
		                       @accoda.DF12, @DF12.cancella @DF12.delete _
		                       @accoda.MODULI.to.DF12
		                    
   
   IMPORTA I MODULI NELLA TABELLA
   				Note accodo i moduli nella tabella oggetti moduli, vedi :
   					@IMPORTA.MODULI

   ACCODO TUTTO IN DF12
   				Note
   					Viene eseguito il comando di accodamento di tutti i moduli, query, tabelle, form, report, macro  nella tabella dell'elenco degli oggetti DF12 con questo codice:
   					@Cmd_Accoda_TUTTO_TO_Msys_DF12



   IMPORTA MACRO
   				Note	
   					Importo le macro nella tabella oggetti apposita, vedi:
   					il comando
   						04_ACCODA_MACRO_TO_Msys_DF12


   CARICA PAGINE TAB CONTROLL
   				Note
   					All'apertura della form MASTER	attivo la ricerca del codice progetto
   					scelto da caricare.

   						@carica.pagine.tab.controll, @tab.controll.carica.pagine

					Salvo il codice progetto in una variabile myCOD_PROGETTO_s da utilizzare come filtro per le pagine assegnate al progetto. Una volta individuato il 
					progetto scelto viene aperto un II rs con la query di filtraggio con il 
					comando CurrentDb.OpenRecordset(sSql)
						@PAGINE.FILTRATE.DA.CARICARE
					
							sSql = ""
							sSql = sSql & "SELECT "
							sSql = sSql & "MSys_TUTTI_PROG_VIS_PAGINE.COD_PROGETTO_s, "
							sSql = sSql & "MSys_TUTTI_PROG_VIS_PAGINE.NRO_OGGETTO_i, "
							sSql = sSql & "MSys_TUTTI_PROG_VIS_PAGINE.TIPOGGETTO_s, "
							sSql = sSql & "MSys_TUTTI_PROG_VIS_PAGINE.NOMEOGGETTO_s, "
							sSql = sSql & "MSys_TUTTI_PROG_VIS_PAGINE.NOMEpagina_s, "
							sSql = sSql & "MSys_TUTTI_PROG_VIS_PAGINE.VISIBLE_b, "
							sSql = sSql & "MSys_TUTTI_PROG_VIS_PAGINE.NOTA_OGGETTO_s, "
							sSql = sSql & "MSys_TUTTI_PROG_VIS_PAGINE.NOTEex_s, "
							sSql = sSql & "MSys_TUTTI_PROG_VIS_PAGINE.ID "
							sSql = sSql & "FROM MSys_TUTTI_PROG_VIS_PAGINE "
							sSql = sSql & "WHERE (((MSys_TUTTI_PROG_VIS_PAGINE.COD_PROGETTO_s)='" & myCOD_PROGETTO_s & "'))"
							sSql = sSql & "ORDER BY MSys_TUTTI_PROG_VIS_PAGINE.NRO_OGGETTO_i "
							sSql = sSql & "WITH OWNERACCESS OPTION; "



						@TABELLA.CARICA.PAGINE, @PAGINE.DA.CARICARE
						@attivo.il.caricamento.pagine

						FAQ
							QUALE PROGETTO VIENE CARICATO
							come viene scelto il progetto da caricare
							progetto da caricare se il campo = true
							@pagine.filtrate.da.caricare



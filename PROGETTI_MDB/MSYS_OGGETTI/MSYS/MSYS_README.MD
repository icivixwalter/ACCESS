MSYS_REDME

	PROGETTO_MSYS_REDME
				Note
					il progetto MySys si riferisce al progetto relativo a tutti gli oggetti Access
					utili per gli altri progetti.
					Il progetto Msys si trova in questa path e file:
						c:\CASA\CDM\GE_CDM\MSYS\MSYS_MDB\MSYS_N01_TABELLE.mdb
						@PROGETTO@MSYS



					ATTENZIONE PATH DEL PROGETTO IN UFFICIO è QUESTA:
						c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MSYS\


				DB
					MSYS_N01_TABELLE.mdb
						si trova in questa path:
					c:\CASA\CDM\GE_CDM\MSYS\MSYS_MDB\



				FORM MASTER
					MSys_FORM_M01_GESTIONE_OGGETTI_MSys
						Questa è la form per la gestione di tutti gli oggetti msys e viene utilizzata
						per gli altri progetti.
						Permette di caricare i vari oggetti necessari al progetto in questione;

							Per tutti i dettagli fare riferimento a questo codice: @FORM@MASTER@MSYS

							Attenzione se cambi il nome alla form master questa procedura garantisce
							il suo funzionamento:
								@cambio@denominazione@form@master_(se cambi denominazione alla form master viene utilizzata ugualmente)

							Attenzione 


							'//@ERRORE@RECORD@VUOTO, se esiste il progetto da caricare, MA NON HA FORM DA IMPLEMENTARE allora il msg di errore _
	                                            + uscita dalla routine,
	                                        DoCmd.OpenQuery "MSys_TUTTI_PROG_SELECT_Qry01_01_PROGETTI", acViewNormal
	                                        

				GESTIONE_PROGETTO
					TODO: provo ad eliminare tutti gli oggetti senza filtro in db mdb esterno

					

					TODO
					sono da fare i seguenti miglioramenti per il progetto:
						1) Eliminare tutte le form  che non hanno il codice MSys_ perche appartengono 
							agli altri progetti. 
								DONE
									2023.02.11 FORM ELIMINATE = sono state eliminate tutte le form che 
												non appartegono al progetto Mys

						2) Eliminare tutte le QUERY  che non hanno il codice MSys_ perche appartengono 
							agli altri progetti.

								DONE
									2023.02.11 QUERY ELIMINATE = sono state eliminate tutte le QUERY che 
												non appartegono al progetto Mys

						3) Eliminare tutti I MODULI  he non hanno il codice MSys_ perche appartengono 
							agli altri progetti.
		


		FORM_CARICAMENTO_AUTOMATICO
			Il progetto FORM_M01_MASTER_base_(CaricaSottoform).mdb contiene la form MASTER per
			il caricamento automatico delle sottoform in numero max 12

		FORM
			MSys_FORM_M01_}------------------------------------------------@
			MSys_FORM_M01_GESTIONE_OGGETTI_MSys
			AA_MSys_FORM_M01_GESTIONE_OGGETTI_MASTER
				Note
					form a caricamento automatico di 12 pagine
					Le due form sono le stesse solo la master dovrebbe essere quella che puo essere utilizzata
					per tutti i progetti, e quindi potrebbe rimanere mentre la GESTIONE_OGGETTI_Msys verrebbe
					cancellata appena finito la stesura del progetto.




						@form@caricamento@automatico_(La form @MASTER per caricare automaticamente @12@pagine)



				FUNZIONAMENTO
					01)FORM -> EVENTO @FORM@LOAD@PAGINE@TUTTE
						All'apertura della form carico tutto le pagine del tab controll per un  max di 12

						@RESET@PAGINE_(sempre all'apertura vengono resettare le sottoform le pagine del tab controll)

						@ATTIVO@FUNZIONE@CARICA@PAGINE_(viene attivata la funzione che carica le pagine delle sottoform)

						@FUNZIONE@CARICA@PAGINE@TAB@CONTROLL_(eseguo un ciclo sulla tabella MSys_TUTTI_PROG per trovare il progetto selezionato a True da caricare)


						@FUNZIONE@CARICA@PAGINE@TAB@CONTROLL_(recupero i parametri CODICE PROGETTO TIPO OGGETTO NOME OGGETTO NOME PAGINA VIBILE e chiamo la funzione imposta le sottoform)

						@CHIAMO@LA@ROUTIN@IMPOSTA@LE@SOTTOFORM_(per impostare le sottoform e passo 4 parametri)


						@ROUTINE@IMPOSTA@LE@SINGOLE@SOTTOFORM_(con la ROUTINE imposta le sottoform MA attivo la classe ricerca oggetto nel db)



						@ROUTINE@IMPOSTA@LE@SINGOLE@SOTTOFORM_(ma prima controllo il Parametro Vible = True e poi attivo la classe di controllo esistenza sottoform, infine solo se la form esiste la carico nella pagina scelta)




						@faq@come@caricare_(Le@pagine delle 12 sottoform sono a @caricamento@automatico
							)


					02) @PROGETTO@NULLO
						Nel caso che nessun progetto viene selezionato allora viene emesso un avviso di 
						progetto nullo, vedi codice
							@PROGETTO@NON@SELEZIONATO@ ...


					03) COMANDI O MENU
						Nella form MASTER, i menu devono cambiare nel cambio del progetto.
						Per ottenere questo risultato occorre quindi:
						a) creare una tabella tmp dei comandi denominata 	
								MSys_DF01_COMANDI_TMP
							In questa tabella devono essere salvati i comandi temporanei che attengono al progetto e poi caricati e visualizzati nella form master.
							Prima modifico la tabella originale e tmp con il codice progetto inserendo questo
							campo.
								COD_PROGETTO_s
							HO anche modificato le chiavi delle tabelle comandi e comandi tmp con questi due
							campi:
								COMANDO_s
								COD_PROGETTO_s
								PROGETTO_s

							- Ho creato la query di RAGGRUPPAMENTO per visualizzare i comandi assegnati ai
								progetti:
									MSys_DF01_COMANDI_Qry01_01_GROUP_TUTTI
									codice: comandi

							- Ho creato la query di accodamento dei comandi tmp in base al codice progetto.
									codice: comandi
								


						B)
							@2024_03_11_CREATO_INSERIMENTO_CODICI_PROGETTO_IN_TMP

							ho creato anche la tabella tmp relativo a i PROGETTI tmp dove vengono salvati
							solo 
								MSys_TUTTI_PROG_TMP

							Ho aggiunto e creato una query di accodamento del progetto scelto denominata
								MSys_TUTTI_PROG_INSERT_Qry01_01_PROGETTI_in_TMP

							questa query di accodamento è un base per l'operazione di cancellazione del progetto nella tabella
							tmp ed il successivo accodamento nella tabella
								MSys_TUTTI_PROG_TMP
							dei 3 codice del progetto scelto e questo è un esempio:
								NRO_OGGETTO_i	COD_PROGETTO_s				DENOM_PROGETTO_s
								5				FORM_AUTOCARICA	PROGETTO 	FORM AUTOCARICANTE


							vedi codice :
								@INSERTI@CODICI@PROGETTO@TMP   ...


						C)
							dopo aver inserito i codice del progetto tmp occorre creare:

								Cmb.01 = MODIFICO IL NOME DELLA COMBINATA 01 IN Cmb_01_TXT
									Ho cambiato la path dei comandi ....

									@cmbd_01.casella.combinata.evento_(APRO GRUPPO COMANDI CASELLA COMBINATA 01)
								TODO: fare una query che estrae i menu con il codice progetto tmp

								TODO: aggiornare il campo Cmb_01
									   CON IL MENU  del codice selezioto ...



	   REGISTRAZIONE_DELLE_LIBRERIE_DLL_(cmd Dos)

	   		Per caricare  la libreria OCX  che si trova qui:
		   	c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MDB\DLL_LIBRERIE\MSCOMCTL.OCX

		   			@DLL@CMD@DOS_(comando di apertura della shell del @registro@di@sistema)
		   			@REGISTRO@DLL@CMD@DOS_(comando DOS PER LA REGISTRAZIONE DELLE LIBRERIE)
		   		eseguire il comando
				regSvr32 con il file .ocx

				per registrare manualmente la ddl ocorre 
				1) aprire il cmd come amministratore 
				2) cd.. radice posizionando in 
					c:\Windows\SysWOW64\
				3) eseguire il comando:
					regSvr32 MSCOMCTL.OCX


	GORNALE_OPERAZIONI

			start "apri GIORNALE DELLE OPERAZIONI" CALL "c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MSYS\OBJECT\TUTORIAL\APRI_GIORNALE_OPERAZIONI_PROGETTO_MSYS.bat"

		2024
				2024.07.15=inserito nuova tabella MysDLL

							Msys_DLL_Qry01_81-}--------------------------------@
							Msys_DLL_Qry01_81_DELETE

				2024.07.27=CREATO_TABELLE_DF_DLL+QUERY
						creato le tabelle per la gestione delle dll dei vari progetti. Creato Query
						per la gestione delle librerie.
						Tutti gli oggetti possono essere estratti con il seguente :
								codice ----> @DLL_(codice di estrazione oggetti per la gestione delle DLL)
								
							MSys_DF13_}----------------------------------------------------@
							Msys_DF13_DLL_PROGETTO = qui vengono salvati il codice delle dll ed i
													progetti interessati

							Msys_DF14_DLL_LIBRERIE = questa tabella salva tutte le librerie attive e che
													interssano il progetto base della precedente tabella.
							Msys_DLL 				= tabella dove vengono salvate le librerie del progetto
													esaminato e che dovrebbero integrare la tabella DF13 E
													DF14 con query di accodamento e di aggiornamento da creare.
							Msys_DF13Qry01_00_}-----------------------------------------@DLL
							Msys_DLL_Qry01_01_DF13_DLL_LIBRERIE
													= creato la query che estrae i progetti e le libreria collegate. codice ----> @QUERY@DF13@DLL@PROGETTI


							Msys_DF14Qry01_01_DLL_LIBRERIE
													= creato la query che estrae tutte le librerie DLL
													legate ad un progetto.
													codice ----> @QUERY@DF14@DLL@LIBRERIE


							Msys_DLLQry01_00_}---------------------------------------------@	
							Msys_DLLQry01_01_LIBRERIE_ATTIVE
													= creato la query che estrae le libreria DLL SALVATE nel progetto corrente
													e quindi rappresenta le LIBRERIE ATTIVE.
													codice ----> @QUERY@Msys_DLL

				2024.07.28=CREATO_MACRO_(esporto la macro di controllo delle librerie DLL)
						UTILITA_Msys_Mcr01_STAMPA_LIBRERIE_ATTIVE_DLL
							Questa è la macro per LA STAMPA DELLE LIBRERIE del database corrente.

				2024.08.02=MODIFICA_FORM_IMPORTAZIONE_PROGETTI_(Adattato la form all'importazione delle librerie DLL)

							CMB_02 = creata una combinata cmb_02 che permette la scelta 
								degli oggetti da importare.
							CMB_01 = modificato la combinata 1 puoi sceglere il tipo di 
								progetto identificato dal file.mdb
								La combinata 01 deve gestire gli oggetti del progetto scelto
								e deve attivare i button necessari per l'importazione e la cancellazione. L'importazione è interdetta se ci troviamo nel progetto originario. 
								Ammessa nei progetti esterni.



								OPERAZIONI:
									01) CHIAMO LA COLLECTION TABLE :
									 Case "IMPORTA_OGGETTI_MSYS_DLL"	
									 	Se nella combinata 2 vengono scelte le librerie DLL
									 	la select carica nella collection gli oggetti tabella
									 	inerenti le librerie DLL

									 	DEVE CHIAMARE CollectionTable_PFunct

								 	02 = Attivo la funzione chiama la collection

								 		bol=CollectionTable_PFunct

								 		per utilizzare il cmd Cmd_Importa_TABLE_Click devo effettuare una
								 		modifica.
								 		Questo comando funziona per la cmb_01 e la cmb_02


						 			03 = CREO HELP DELLA COMBINATA 02

 				2024.08.03 = COMBINATA_01 
 						e COMBINATA_la casella combinata 
 						Cmb_01_txt contiene anch'essa il progetto 
 						02 aggiornamento
 				2024.08.03_Cmb_01_txt_AfterUpdate 
 					= inserito il case ELENCO PROGETTI
 					in caso di scelta con ELENCO PROGETTI disattiva tuti i button

 				2024.08.03_Abilito_sourceDBPath_Sub 
 					= creato l'abilitazione della path

 				2024.08.03_Disabilito_sourceDBPath_Sub 
 					= creato la Disabilitazione della path

 				2024.08.03_Abilito_Table+Report+Macro+Query+Moduli 
 					= creato routine per 
 					l'abilitazione delle query, macro, report, table ecc.

 				2024.08.03_DisabilitoAbilito_Table+Report+Macro+Query+Moduli = creato routine per 
 					la DISABILITAZIONE  delle query, macro, report, table ecc.

 				2024.08.03_MODIFICA_CMB_02 = Abilito soltanto i button e non eseguo nessun comando.
 											l'ambitazione differenziata permette l'utilizzo del singolo comando oppure di quello collettivo per i comandi abilitati.

				2024.08.03_MODIFICA_CMB_DELETE_TUTTO = la modifica riguarda la routine di attivazione 
													di tutti i comandi. Gli stessi sono eseguiti solo se 
													abilitati es. 
													la chiamata sequenziale dei comandi come di seguito 
													indidata :
														Cmd_Delete_FORM_Click
														Cmd_Delete_MODULI_Click
														Cmd_Delete_QUERY_Click
														Cmd_Delete_REPORT_Click
														Cmd_Delete_TABLE_Click
														Cmd_Delete_MACRO_Click

													viene modificata con il controllo per ogni comando se 
													abilitato oppure no es.



													If Me.Cmd_Delete_FORM.Enabled = True Then
													    Cmd_Delete_FORM_Click
							                        End If
							                        If Me.Cmd_Delete_MODULI.Enabled = True Then
							                            Cmd_Delete_MODULI_Click
							                        End If

							                        ecc..
                        

           	2024.08.03_MODIFICA_CMB_IMPORTA_TUTTO = la modifica riguarda la routine di attivazione     
                        							stessa cosa per il comando collettivo IMPORTA TUTTO.
                        							Le if di abilitazioni = true consentono l'attivazione.

        		2024.08.03_Select_Case_Me.Cmb_02_txt = SOSTITUITO IL VALORE CON LA CMB_02 al posto della CMB_01 _
                    Le collection viene scelta in base combinata 2 che individua il progetto _
                    al posto della combinata 1 e la sostituzione riguarda tutte le collection.

            2024.08.03_CollectionQuery_PFunct = SE LA COLLECTION E' VUOTA ESCI DALLA ROUTINE
                		faccio un controllo sul numero degli oggetti della collection se = 0 allora
                		non ci sono oggetti da caricare per cui si esce dalla routine.

            2024.08.03_CollectionForm_PFunct = SE LA COLLECTION E' VUOTA ESCI DALLA ROUTINE _
                		idem

	        	2024.08.03_CollectionMacro_PFunct 
	        				= SE LA COLLECTION E' VUOTA ESCI DALLA ROUTINE idem

				2024.08.03_CollectionModuli_PFunct
							  = SE LA COLLECTION E' VUOTA ESCI DALLA ROUTINE _

				2024.08.03_CollectionReport_PFunct  
							= SE LA COLLECTION E' VUOTA ESCI DALLA ROUTINE _
							idem

				2024.08.10 =COLLECTION=0- 
									prima di uscire imposto la collezione per evitare di restiuire null e poi esco dalla routine.
									E' stata applicata a tutte le routine delete.

				2024.08.10 =CHECKPROJECT_MODIFICA_CONTROLLO_CON_RAFFORZAMENTO
							Sulla funzione di controllo del progetto CheckProjectName_b_PFunct il controllo se il progetto
							gestito è quello degli oggetti msys viene rafforzato perche viene controllata :
								1) se il progetto è MSYS_N01_OGGETTI.mdb
								2) se la casella combinata Cmb_01_txt contiene anch'essa il progetto MSYS_N01_OGGETTI.mdb
							In questo modo evitiamo le importazioni o le cancellazioni degli oggetti Msys ma sono ammesse
							tutte le altre.


				2024.08.05_ATTIVA_LIBRERIE_FORM_UTILITA_Msys_Frm21_IMPORTA_OGGETTI
						FORM   -->  UTILITA_Msys_Frm21_IMPORTA_OGGETTI
							UTILITA_Msys_Frm21_}-------------------------------------------@
							UTILITA_Msys_Frm21_IMPORTA_OGGETTI
						La form importa oggetti viene utilizzata per importare gli oggetti del progetto Msys nei
						database esterni non in quello di origine. Bisogna esportare la form in nuovo progetto ed 
						attivare le procedure di importazione. I progetti gestiti dalla form importa oggetti sono:

							Se il progetto scelto è :   	MSYS_N01_OGGETTI.mdb		

						Per attivare le librerie DLL per LA FORM DI IMPORTAZIONE DEGLI gli oggetti msys è stato inserita nella collezione
						delle macro anche la macro che le attiva. La macro corregge l'errore di mancanza della libreria Dll per il suo funzionamento.
							2024.08.05_@macro@attiva@librerie@dll_(inserita la macro da importare per attivare le librerie FORM IMPORTA OGGETTI)

				2024.08.06_RIATTIVAZIONE_TREEVIEW = con il codice "TREEVIEW.mdb
					todo: riattivo la treeview nel menu di scelta E LE MODIFICHE PER LA RIATTIVAZIONE SONO:
					sulla combinata cmb_01 = riattivato il comando con il codice "TREEVIEW.mdb"
					sulla combinata cmb_02 = inserito il nuovo codice "IMPORTA_OGGETTI_TREEVIEW"
					sulla collectionFORM   = 2024.08.06_CORREZIONE_VOCE_(corretto errori della collection DA objectFORMS.Add "TreView_FILE" A objectFORMS.Add) "TreeView_FILE"

					sul button Cmd_Importa_FORM_Click  = 2024.08.06_EFFICIENTAMENTO_(efficientata la procedura con l'interrogazione sulla tabella MSysObjects per il controllo delle form esterne)

				2024.08.06_RIATTIVAZIONE_ESPROPRI = con il codice GE_ESPROPRI.mdb
					creato il codice per gli espropri con "IMPORTA_OGGETTI_ESPROPRI" e creato le seguenti collection:
					2024.08.06_COLLECTION@FORM@ESPROPRI_(inserito la collezione MACRO per gli espropri)
					reimpostato le collection:

					collection MODULI = 2024.08.06_COLLECTION@MODULI@ESPROPRI_(inserito la collezione MODULI per gli espropri)
					collection MACRO = 2024.08.06_COLLECTION@MACRO@ESPROPRI_(inserito la collezione MACRO per gli espropri)
					collection REPORT = 2024.08.06_COLLECTION@REPORT@ESPROPRI_(inserito la collezione REPORT per gli espropri)

				2024.08.05_IMPORTAZIONE_SEMPLICE
					MODULO PER L'IMPORTAZIONE SEMPLICE: modifico il modulo UTILITA_Msys_Mdl05_IMPORTAZIONE_SEMPLICE_DI_OGGETTI. Qui effetto l'importazione degli oggetti
					in modo semplice cioè senza controllo del DATABASE ESTERNO.
					Il controllo è solo SUGLI OGGETTI INTERNI.

						@importazione@semplice_(importa tabella o collega tabelle SENZA CONTROLLO ESTERNO)

          	2024.08.09_EFFICIENTAMENTO_CONTROLLO_PROGETTO_ORIGINALE
             		Il controllo del progetto originale viene spostato sulla combinata 01 in modo da verifiare
             		se ci troviamo nel progetto MSYS_N01_OGGETTI.mdb già al momento in cui dalla combinata_01
             		viene scelto il progetto MSYS_N01_OGGETTI.mdb. In questo modo evitiamo di ripetere il 
             		controllo che peraltro già esiste sui button di importazione ed esportazione. L'avviso
             		scatta al momento della scelta del progetto originale.
                         '//CONTROLLO SE SIAMO NEL PROGETTO ORIGINALE MSYS ED INIBISCO IMPORT E DELETE
     
				2024.08.07_MIGLIORATA_RICERCA_TABELLE
						MODULO-->UTILITA_Msys_Mdl04_IMPORTA_LINK_TABELLE_Msys_DLL
							Importo le tabelle di gestione delle librerie DLL in collegamento. Effettua il @link 
							dal dabase di origine a quello corrente. Per attivare la funzione di collegamento
							denominata ImportaLINKTabelle_Msys_PFunct() usare la macro TREEVIEW_Mcr04_IMPORTA_LINK_TABELLE_Msys_DLL_PROGETTO.
								@link@Tabelle@dll_(per fare il link delle tabelle Dll utilizzare il modulo Msys Dll e la macro associata)
						FAQ_LINK
							Come fare il link delle tabelle dll?
								Msys_DF13_DLL_PROGETTO, Msys_DF14_DLL_LIBRERIE e Msys_DLL_LIBRERIE
							Link delle tabelle Dll come fare?
								utilizzare il modulo link delle tabelle msy DLL e la macro.
								Vedi codice  @link@Tabelle@dll

						MODULO-->UTIL_MDL40_N09_CONTROLLO_ESISTENZA_OGGETTO_DB AGGIORNAMENTI
						La procedura permette ora la ricerca delle tabelle con efficienza.
						@RICERCA@ESISTENZA@TABELLE_(controllo se la tabella esiste nel db corrente)

				2024.08.07_AGGIORNAMENTO
						MODULO-->UTILITA_Msys_Mdl06_COLLEGA_TABELLE_LINK
						Controllo se effettua il collegamento delle tabelle.

				2024_08_17=CREO_TABELLE_Msys_ESPORTA_DB_EST+Msys_ESPORTA_DB_EST
					Le tabelle Esporta Db e esporta tabelle oggetti sono state create il 17 agosto 2024 al fine di gestire 
					l'esportazione degli oggetti dal database corrente ad un database esterno. La tabella ESPORTA DB EST si 
					occupa del database e della path guida per l'esportazione.
					ESPORTA ELENCO OGGETTI è la tabella collegata alla prima mediante le chiave PATH e PROGETTO che elenca gli 
					oggetti msy che possono essere spostati nel db esterni.
					I codice per le due tabelle sono :
						@ESPORTA@OGGETTI_(la tabella  Msys_ESPORTA_DB_EST)
						@ESPORTA@ELENCO@OGGETTI_(la tabella  Msys_ESPORTA_DB_EST)

				2024_08_17=Query_ESPORTA_OGGETTI+GE_FORM_ESPORTA_OGGETTI
					Ho creato nuovi oggetti table, query e form per gestire con efficiacia la esportazione degli
					oggetti che fino a questo momento non era presa in considerazione. Sono state create tabelle,
					query e form per raggiungere questo risultato. Ho pure adottato la copia della master GE_MSYS 
					per adattarla a MINI FORM MASTER e funziona egregiamente in modo da poter inserire in una master 
					PRINCIPALE  anche mini Master che hanno lo stesso funzionamento delle form a caricamento automatico.
					Creato 2 tabelle;
					creato le 8 query di esportazione degli oggetti.
					Creato le 8 form per le esportazione degli oggetti
					Creato UNA MINI FORM MASTER GE_ESPORTAZIONE OGGETTI.

					Le 8 query per gestire la nuova esportazione 
					degli oggetti. La db est serve per la gestione del database, quelle degli
					oggetti per le esportazione degli oggetti legate al database.

						Msys_Qry01_01_ESPORTA_DB_}-------------------------------------@
						Msys_Qry01_01_ESPORTA_DB_EST
						Msys_Qry01_01_ESPORTA_OGGETT_}---------------------------------@
						Msys_Qry01_01_ESPORTA_OGGETTI
						Msys_Qry01_01_ESPORTA_OGGETTI_Form
						Msys_Qry01_01_ESPORTA_OGGETTI_Macro
						Msys_Qry01_01_ESPORTA_OGGETTI_Moduli
						Msys_Qry01_01_ESPORTA_OGGETTI_Query
						Msys_Qry01_01_ESPORTA_OGGETTI_Report
						Msys_Qry01_01_ESPORTA_OGGETTI_Table

					Le 8 sub form che si basano sulle query sopra indicate sono:
						Msys_Frm01_S01_{===============@ESPORTA@=======================}
						Msys_Frm01_S01_ESPORTA_DB_GEST
						Msys_Frm01_S01_ESPORTA_DB_}------------------------------------@
						Msys_Frm01_S01_ESPORTA_DB_EST
							sottoform per la gestione dei database in cui esportare gli oggetti
							msysA

						Msys_Frm01_S01_ESPORTA_OGGETT_}--------------------------------@
						
						Msys_Frm01_S01_ESPORTA_OGGETTI
							serve a gestire tutti gli oggetti msys da esportare.

						Msys_Frm01_S01_ESPORTA_OGGETTI_Form
							la form gestisce solo gli oggetti form
						
						Msys_Frm01_S01_ESPORTA_OGGETTI_Macro
							la form gestisce solo gli oggetti macro

						Msys_Frm01_S01_ESPORTA_OGGETTI_Moduli
							la form gestisce solo gli oggetti moduli

						Msys_Frm01_S01_ESPORTA_OGGETTI_Query
							la form gestisce solo gli oggetti query

						Msys_Frm01_S01_ESPORTA_OGGETTI_Report
							la form gestisce solo gli oggetti report
						
						Msys_Frm01_S01_ESPORTA_OGGETTI_Table

							la form gestisce solo gli oggetti Table



						HO CREATO ANCHE UNA MINI FORM MASTER che permette il caricamento automatico delle form ed è
						esportabile ma ancora da migliorare.


						Ho creato anche i report
							Msys_Rpt01_01_ESPORTA_DB_}-------------------------------------@
							Msys_Rpt01_01_ESPORTA_DB_EST


							Msys_Rpt01_01_ESPORTA_OGGETT_}---------------------------------@
							Msys_Rpt01_01_ESPORTA_OGGETTI


						Ho creato le macro:
							MSys_Mcr00_}---------------------------------------@ESPORTAMSSys
							MSys_Mcr01_ATTIVA_MACRO_ESPORTA_OGGETTI_MSSys






			2024_08_18=EventoDoppioclick_CampoCOD_PROGETTO_s_Txt
						Ho creato nella form Msys_Frm01_S01_ESPORTA_DB_EST, esporta database l'evento doppio click per l'aggiornamento della tabella Tmp  Msys_ESPORTA_DB_EST_TMP creata appositamente per salvare nella tabella tmp la path ed il database scelto con il doppio click in modo che le altre sottoform relative agli oggetti si aggiornino automaticamente


						ATTIVA Le query che leggono dalla tabella tmp Msys_ESPORTA_DB_EST_TMP

						CREATO ANCHE LE FORM ESPORTA OGGETTI TMP
						Msys_Frm01_S01_ESPORTA_OGGETTI_Form_tmp
						Msys_Frm01_S01_ESPORTA_OGGETTI_Macro_TMP
						Msys_Frm01_S01_ESPORTA_OGGETTI_Moduli_TMP
						Msys_Frm01_S01_ESPORTA_OGGETTI_Query_TMP
						Msys_Frm01_S01_ESPORTA_OGGETTI_Report_TMP
						Msys_Frm01_S01_ESPORTA_OGGETTI_Table_TMP

						per ogni form una query Tmp che estrae gli oggetti da esportare dalle query
						Tmp


			2024_08_20=CreataNuovaForm_Msys_Frm01_S01_ESPORTA_OGGETTI_TMP
					per effettuare l'estrazione dei soli oggetti relativi alla path e al database scelto _
					ho dovuto creare una nuova form che filtra solo gli oggetti msy da esportare ma con  _
					il codice progetto salvato in tmp (altrimenti estrae tutti gli oggetti anche quelli _
					di altre procedure). 

						- ho creato una query select tmp : Msys_Qry01_01_ESPORTA_OGGETTI_TMP

						- la form creata si chiama Msys_Frm01_S01_ESPORTA_OGGETTI_TMP




			2024_08_21=DEFINIZIONE_TABELLA_PROGETTO
				IL codice che estrae tutti gli elementi di un progetto msys è :
					@CODICE----> PROGETTI_Msy

				viene ridefinita la tabella DEI PROGETTI che preleva almeno gli elementi fondamentali
				dalle tabelle DF02 E DF11.
				La denominazione  è 

						PROGETTI_MSYS_{@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@}_Tab
						
						MSYS_PROGETTI_TB02_}-------------------------------------------@
						PROGETTI_Msys_TB01_PROJECT
						
						MSYS_PROGETTI_TB02_}-------------------------------------------@
						PROGETTI_Msys_TB02_TIPO_OGGETTI
							l'elenco degli oggetti del progetto
							
						MSYS_PROGETTI_TB03_}-------------------------------------------@
						MSYS_PROGETTI_TB03_ELENCO_ObjProgetti

			2024_08_21=QUERY_DI_PROGETTO
				sempre con il codice @CODICE----> PROGETTI_Msy	 
				sono state create le seguenti query:

						PROGETTI_Msys_TB01Qry01_01_PROJECT
							query di estrazione dei progetti definiti

						PROGETTI_Msys_TB02Qry01_01_TIPO_OGGETTI
							query tipo di oggetti

						PROGETTI_Msys_TB03Qry01_01_ELENCO_ObjProgetti
							query elenco object


				Le tre tabelle sono collegate a cascata.  Ma è stata creata quella di gestione:

					PROGETTI_Msys_{@==============GESTITI_DA_MSYS============@}_Tb00

								PROGETTI_Msys_Tb00FrmM01_GE_PROGETTI
									questa è la form MASTER che gestisce tutti i progetti con il tab
									controll da 12 pagine

								PROGETTI_Msys_TB01Qry01_01_PROJECT
									form per i progetti con gli oggetti con il codice progetti

										Master : COD_PROGETTO_s --- detail : COD_PROGETTO_s
										
									PROGETTI_Msys_TB02Qry01_01_TIPO_OGGETTI

										Master : COD_PROGETTO_s;TIPOGGETTO_s --- detail : COD_PROGETTO_s;TIPOGGETTO_s
											PROGETTI_Msys_TB03Qry01_01_ELENCO_ObjProgetti

			2024_08_23_DF11_PROGETTI_RIDEFINITO
				La tabella df11 relativa alla definizione dei progetti viene ridefinita come attivita. Rimane
				basata esclusivamente come definizione dei progetti nuovi o la gestione delle definizioni dei
				procedenti escludendo il collegamento con la definizione degli oggetti. Anche la form ad essa
				collegata viene ridefinita come attivita esclusiva alla gestione dei progetti.

			2024_08_23_Msys_DF12_TIPO_OGGETTI_CANCELLATA
				Questa tabella viene cancellata perche un duplicato in quanto il tipo di oggetti
				della vecchia tabella MSys_DF02_TIPO_OGGETTO è stata sostituita con la 
				tabella gia esistente MSys_DF02_TIPO_OGGETTO deputata solo alla definizione degli oggetti.



			2024_08_23_Msys_DF12_Qry01_01_TIPO_OGGETTI_CANCELLATA
				anche la query Df12 viene cancellata e sostituita con la MSys_DF02_TIPO_OGGETTO
				creata come
							Msys_DF02_Qry01_}----------------------------------------------@
							Msys_DF02_Qry01_01_TIPO_OGGETTI


							Msys_DF10_Qry01_}----------------------------------------------@



			2024_08_23_Msys_DF02_Frm01_01_TIPO_OGGETTI
				creata la nuova form che gestisce solo la definizione degli oggetti,
					Msys_DF02_Frm02_}------------------------------------------@


			2024_08_23_Msys_DF11_PROGETTI_Frm01_M01_Obj_RIDENOMINATA
				la form che contiene la definizione dei progetto è stata ridenominata da 
				Msys_DF11_PROGETTI_Frm01_M01_Obj ----> Msys_DF11_PROGETTI_Frm01_M01_Elenco

			
			2024_08_23_CORREZIONE_ERRORE_PAG_12 = corretto errore della sottoform non caricata
				Nella pagina 12 se non è precaricata la sottoform 12 da errore oggetto chiuso o eliminato nel 
				comando Me.SottoForm_12.Form.Refresh. Non esistendo la sottoform non si puo fare il refresh.
				corretto con l'errore con il controllo del recordsource dell'oggetto sottoform 12 con questa if:

										If Me.SottoForm_12.SourceObject > "" Then
                                            Me.SottoForm_12.Form.Refresh
                                        End If
              	IL controllo errore viene esteso a tutte le altre sottoform con il codice:

                      2024_08_23_CORREZIONE_ERRORE_PAG_01-12 = corretto errore delle sottoform non caricate con recordsource =



			2024_08_23_PROJECT_TMP
				TODO: devo creare una evento di salvataggio del codice  progetto nella tabella tmp
					PROGETTI_Msys_TB01_PROJECT_TMP che deve essere creata.

					CREATA LA TABELLA TMP dove salvare il codice progetto tmp.. Appunto per individuare le pagine del progetto.

					creo la form per il valore del progetto tmp : PROGETTI_Msys_TB01Qry01_01_PROJECT_TMP
					creo la query base sui valori tmp : PROGETTI_Msys_TB01Qry01_01_PROJECT_TMP


			2024_09_04_SOTTOFORM_MASTER_SENZA_TITOLO = CREATA LA SOTTOFORM MASTER A CARICAMENTO AUTOMATICO DI 12 PAGINE DEL TAB CONTROLL
				creato una sottoform Master abilitata al caricamento automatico di altre sottoform. La peculiarita che la 
				sottoform master senza titolo è quella di poter caricare 12 sottoform e quindi rendere il caricamento 
				automatico senza obbligo di codice statico.
				FUNZIONAMENTO
					il funzionamento della SOTTOFORM MASTER è il seguente:

					- creato 2 tabelle per la gestione delle sottoform denominate
						SottoFormMaster_}------------------------------------------@Msys
						SottoFormMaster_GE_Msys
						SottoFormMaster_Vis_Pagine_Msys

					- le QUERY per gestire le tabelle sono :
						SottoFormMaster_Qry00_}-----------------------------------@Msys
						SottoFormMaster_Qry01_GE_Msys
						SottoFormMaster_Qry01_Vis_Pagine_Msys

					Puoi inserire un progetto con un codice sottoform master che identifica la sottoform che contiene 12 pagine.

					Il prototipo per il CARICAMENTO AUTOMATICO  si chiama:
						SottoFormMaster_Frm00_CARICAMENTO_AUTOMATICO_{SenzaTitolo}  = contiene 12 pagine

						SottoFormMaster_FrmM01_GE_Msys  							= gestisce il progetto
						SottoFormMaster_FrmS01_Vis_Pagine_Msys  					= imposta le pagine da caricare nel tab controll


					Puo essere utilizzata per altre sottoform con tab controll di 12 elementi.

					Non va in contratto o in conflitto con il gestore di tutti i progetti in quanto si basa su tabelle separate pur
					utilizando lo stesso principio.

					Non ha titolo in quanto è a sua volta una sottofom MASTER perche gestisce un tab controll complesso

			2024_09_13_SOTTOFORM_MASTER_SENZA_TITOLO = CORREZIONE ERRORE IMPEDISCE L'INSERIMENTO DI NUOVI RECORD

				Per attivare l'inserimento di nuovi record bisogna verificare se il codice chiave viene aggiornato.
				Prima di individuare le chiavi da correggere ho modificato il nome delle due tabelle con il codice
				tb01 per essere distinte dagli altri oggetti.
				SottoFormMaster_Tb01_GE_Msys



				Creato anche le tabelle tmp 
					SottoFormMaster_Tb01_GE_Msys_TMP
					SottoFormMaster_Tb02_Vis_Pagine_Msys_TMP
				per l'utilizzo dei dati provvisori.

				
			2024_09_24_CREATA_NUOVA_ESTRAZIONE = Estraggo solo le sottoform da visualizzare _
				Le sottoform da visualizzare vengono salvate con una stringa sql nella tabella _
				tmp denominata SottoFormMaster_Tb02_Vis_Pagine_Msys_TMP che viene gestite anche dalle _
				query dedicate. In base al codice del progetto scelto e clicacato attivando l'evento _
				doppio click del campo testo COD_PROGETTO_s_Txt, vengono salvate i nomi delle sottoform _
				dagli archivi di base alla tabbella tmp. _
 _


			2024_09_24_SOTTOFORM_MASTER_NON_FUNZIONANTE = CORREZIONE ERRORE IMPEDISCE IL CARICAMENTO DEL PROGETTO SCELTO
				la Sottoform MASTER SENZA TITOLO presenta ancora problemi di gestione delle sottomaschere; non carica il progetto scelto ma rimane impostata sempre a
				quello originario scelto dalla MASTER ORIGINALE.

				

				creo una query Update che salva i dati dalla tabella 
					SottoFormMaster_Tb02_Vis_Pagine_Msys    ----->  SottoFormMaster_Tb02_Vis_Pagine_Msys_TMP
				solo con i dati uguali.

				La query verrà chiamata
					SottoFormMaster_Tb01Qry01_01_GE_Msys_TMP_UPDATE_pagine

					schema  = da inserire in archivio query ***
							La query preleva i dati facendo il confronto tra le due tabelle quella dei codice Progetto
							e quella delle pagine da visualizzare tra la TB01 TMP  e la TB02 archivio:
								SottoFormMaster_Tb01_GE_Msys_TMP INNER = SottoFormMaster_Tb02_Vis_Pagine_Msys 
							e salvo le pagine da visualizzare.


						INSERT INTO 
							SottoFormMaster_Tb02_Vis_Pagine_Msys_TMP 
							( NRO_PROGETTO_i, COD_PROGETTO_s, DENOM_PROGETTO_s, 
							NRO_OGGETTO_i, TIPOGGETTO_s, COD_FORMS_s, 
							NOMEOGGETTO_s, NOMEpagina_s, VISIBLE_b, NOTA_OGGETTO_s, NOTEex_s )
							SELECT 
							SottoFormMaster_Tb02_Vis_Pagine_Msys.NRO_PROGETTO_i, 
							SottoFormMaster_Tb02_Vis_Pagine_Msys.COD_PROGETTO_s, 
							SottoFormMaster_Tb02_Vis_Pagine_Msys.DENOM_PROGETTO_s, 
							SottoFormMaster_Tb02_Vis_Pagine_Msys.NRO_OGGETTO_i, 
							SottoFormMaster_Tb02_Vis_Pagine_Msys.TIPOGGETTO_s, 
							SottoFormMaster_Tb02_Vis_Pagine_Msys.COD_FORMS_s, 
							SottoFormMaster_Tb02_Vis_Pagine_Msys.NOMEOGGETTO_s, 
							SottoFormMaster_Tb02_Vis_Pagine_Msys.NOMEpagina_s, 
							SottoFormMaster_Tb02_Vis_Pagine_Msys.VISIBLE_b, 
							SottoFormMaster_Tb02_Vis_Pagine_Msys.NOTA_OGGETTO_s, 
							SottoFormMaster_Tb02_Vis_Pagine_Msys.NOTEex_s
							FROM 
							SottoFormMaster_Tb01_GE_Msys_TMP 
							INNER JOIN SottoFormMaster_Tb02_Vis_Pagine_Msys 
							ON (SottoFormMaster_Tb01_GE_Msys_TMP.COD_PROGETTO_s = SottoFormMaster_Tb02_Vis_Pagine_Msys.COD_PROGETTO_s) 
							AND (SottoFormMaster_Tb01_GE_Msys_TMP.DENOM_PROGETTO_s = SottoFormMaster_Tb02_Vis_Pagine_Msys.DENOM_PROGETTO_s)
							WITH OWNERACCESS OPTION;


				
				
						@2024_09_25_MODIFICA_CARICA_LE_PAGINE  = CARICA LE SOTTOFORM SALVATE NELLA TABELLA TMP TB02 _
                                        Recupera i nomi delle sottoform (non dalla SottoFormMaster_GE_Msys) ma dalla tabella tmp che permette la _
                                        visualizzazione delle sottoform: SottoFormMaster_Tb02Qry01_01_Vis_Pagine_Msys_TMP. _

                        TODO: nella SottoFormMaster_Frm00_CARICAMENTO_AUTOMATICO_{SenzaTitolo} correggere l'errore: NEL CAMPO CODICE PROGETTO se il campo è Visibile = True o False cambia sempre il caricamento delle sottoform master

                       
                       '//2024_09_26_MODIFICO_IL_PRINCIPIO_DI_CARICAMENTO=Le sottoform che vengono caricate sono quelle definite dal codice progetto ed Imposto visibile = True. _
                        ATTIVITA CONSEGUENZIALI : _
                            1) imposto a False tutti i campi visible della tabella SottoFormMaster_Tb01_GE_Msys; in questo modo il campo VISIBLE_b ha il solo scopo di _
                                evidenziare quale progetto è stato scelto per il caricamento con il doppio click e non piu come nella MASTER PRINCIPALE prima si _
                                sceglie il progetto con il Flage e poi si carica quello impostato a True. Ribaltando il principio ALLA SCELTA DELL DOPPIO CLICK _
                                si rende piu VELOCE IL CODICE.
                                
                                '//2024_09_26_@01 _
                                    attivo la query di cacellazione del Flag True sulla tabella  todo:continuare ..... fare la query di Upadte flag = FALSE _
                                            ... impostare il Record corrente a True _
                                            ... impostare la requery per l'aggiornamento della Form. ....
                                    
  
                			In questa modifica ribaldo il principio della scelta del progetto da caricare. AL POSTO DI SELEZIONARE PRIMA IL CAMPO VISIBLE = True
                			e poi caricare il progetto con tale flag, cambio principio ed adotto questo criterio:
                			Dall'evento doppio click sul campo COD_PROGETTO_s_TXT 
                					- disabilito con una query tutti i flag visibile = TRUE A FALSE.
                					- imposto a True il campo Visible del record corrente
                					- requery form per l'aggiornamento






==========================================================================================
									LE FAQ
==========================================================================================
FAQ
		TABELLE PER L'ESPORTAZIONE DEGLI OGGETTI
			per esportare gli oggetti di un progetto vengono utilizzate le seguenti tabelle:

				Msys_ESPORTA_DB_EST 			= questa tabella contiene la path ed il database
				Msys_ESPORTA_ELENCO_OGGETTI_DB 	= questa tabella contiene l'elenco degli oggetti come
												query, form, report ecc. da esportare nella path della
												precedente tabella.

		TABELLE PER LA COSTRUZIONE DI UN PROGETTO
			Per la costruzione di un progetto e legare gli oggetti ad un codice vengono utilizzate
			le seguenti tabelle:
				MSys_DF02_TIPO_OGGETTO				= questa tabella definisce gli oggetti di un progetto
													che sono 5, Table, form, Report, Query e Macro

				Msys_DF11_PROGETTI					= la tabella df11, definisce il progetto ed un codice
													chiave e la path di dove si trova il file di gestione
													che di solito è GE_PROGETTO

				quindi vengono utilizzate le tabelle DF02 E DF11 PER LA DEFINIZIONE.


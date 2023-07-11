MSYS_REDME

	PROGETTO_MSYS_REDME
			Note
				il progetto MySys si riferisce al progetto relativo a tutti gli oggetti Access
				utili per gli altri progetti.
				Il progetto Msys si trova in questa path e file:
					c:\CASA\CDM\GE_CDM\MSYS\MSYS_MDB\MSYS_N01_TABELLE.mdb
					@PROGETTO@MSYS


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
			Note
				form a caricamento automatico di 12 pagine



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
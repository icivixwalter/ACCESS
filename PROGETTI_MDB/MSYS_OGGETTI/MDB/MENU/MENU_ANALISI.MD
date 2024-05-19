MENU_ANALISI.MD

	Note
		@analisi@menu_(analisi sui menu)

		Per i menu vi il pogetto MENU_GESTIONE.mdb	  che è duputato
		alla gestione dei menu in collegamento con le altre due 
		voci mdb
			MENU_TB01_VOCI.mdb			= qui vengo salvate le voci dei menu
			MENU_TB02_FUNZIONI.mdb		= qui sono trascritti le funzioni dei
										singoli menu.


			MENU_TB03_OGGETTI.mdb		= quest'ultimo progetto è da utilizzare??


			La path della gestione dei menu:
				@apri@path@menu_(@path@menu si trova in questa directory detta
					anche @menu@msys o @menu@mysys)

				"C:\Program Files\totalcmd\TOTALCMD64.EXE" left: c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MDB\MENU\, right: c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MDB\ 



	VOCI MENU
		Note
			la tabella che contiene le voci menu è la seguente:

		GEST_MENU_TB01_VOCI
			Note
				vengono salvati i menu principale e la tabella si collega con
				la successiva delle funzioni tramite
				DATABASE=C:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MDB\MENU\MENU_TB01_VOCI.mdb;TABLE=GEST_MENU_Tb01_VOCI

				campi master: CodMenu_s	
				campi client: CodMenu_s	

				CODICE MENU VOCI ----> GEST_MENU_TB01
			QUERY
				GEST_MENU_Tb01Qry01_{@===================================@}
				GEST_MENU_Tb01Qry01_00_}-------------------------------VOCI_MENU
				GEST_MENU_Tb01Qry01_01_VOCI
					Note
						@estraggo_(tutti i @menu, schema )
					schema
						SELECT 
							GEST_MENU_Tb01_VOCI.CodMenu_s, GEST_MENU_Tb01_VOCI.Menu_s, GEST_MENU_Tb01_VOCI.NoteMenu_s, GEST_MENU_Tb01_VOCI.COMANDO_s, GEST_MENU_Tb01_VOCI.COD_PROGETTO_s, GEST_MENU_Tb01_VOCI.PROGETTO_s, GEST_MENU_Tb01_VOCI.bCanc, GEST_MENU_Tb01_VOCI.bAgg, GEST_MENU_Tb01_VOCI.bScel, GEST_MENU_Tb01_VOCI.IDKeyArch, GEST_MENU_Tb01_VOCI.IdOrd, GEST_MENU_Tb01_VOCI.Msg01, GEST_MENU_Tb01_VOCI.ID, GEST_MENU_Tb01_VOCI.DATA_INS, GEST_MENU_Tb01_VOCI.DATA_AGG, GEST_MENU_Tb01_VOCI.TIMEOPER

						FROM 
							GEST_MENU_Tb01_VOCI
						WITH 
							OWNERACCESS OPTION;

			FORM
				GEST_MENU_Tb01Frm00_{@======================================@}
				GEST_MENU_Tb01Frm01_}------------------------------------------@
				GEST_MENU_Tb01Frm01_M01_GE_VOCI
				GEST_MENU_Tb01Frm01_S01_VOCI
					form per le voci di menu collegata alla sottofom
					tb01_FUNZIONI con queste chiavi

						campi master: CodMenu_s	
						campi client: CodMenu_s	



		GEST_MENU_TB02_FUNZIONI
			Note
				tutte le funzioni collegate al menu principale

				CODICE VOCI FUNZIONI ----> GEST_MENU_TB02

			QUERY
				GEST_MENU_Tb02Qry01_}----------------------------------------@
				GEST_MENU_TB02Qry01_01_FUNZIONI
					Note
						@funzioni_(estraggo le funzioni collegate ai menu)
					schema

						SELECT 
							GEST_MENU_TB02_FUNZIONI.CodFunzioni_s, GEST_MENU_TB02_FUNZIONI.Funzioni_s, GEST_MENU_TB02_FUNZIONI.ClassFunzioni_s, GEST_MENU_TB02_FUNZIONI.NoteFunzioni_s, GEST_MENU_TB02_FUNZIONI.CodMenu_s, GEST_MENU_TB02_FUNZIONI.bCanc, GEST_MENU_TB02_FUNZIONI.bAgg, GEST_MENU_TB02_FUNZIONI.bScel, GEST_MENU_TB02_FUNZIONI.IDKeyArch, GEST_MENU_TB02_FUNZIONI.IdOrd, GEST_MENU_TB02_FUNZIONI.Msg01, GEST_MENU_TB02_FUNZIONI.DATA_INS, GEST_MENU_TB02_FUNZIONI.DATA_AGG, GEST_MENU_TB02_FUNZIONI.TIMEOPER, GEST_MENU_TB02_FUNZIONI.ID, "->" AS CampiTxt, GEST_MENU_TB02_FUNZIONI.CodFunzioni_s AS CODICE_FUNZIONI, 
							GEST_MENU_TB02_FUNZIONI.Funzioni_s AS FUNZIONI, "chiavi" AS TxtKey, 
							GEST_MENU_TB02_FUNZIONI.CodMenu_s AS KeyCodMenu_s
							FROM 
							GEST_MENU_TB02_FUNZIONI
							ORDER BY GEST_MENU_TB02_FUNZIONI.ID
							;


					
			FORM
				GEST_MENU_Tb02Frm00_{@======================================@}
				GEST_MENU_Tb02Frm01_}------------------------------------------@
				GEST_MENU_Tb02Frm01_M01_GE_FUNZIONI
				GEST_MENU_Tb02Frm01_S01_FUNZIONI



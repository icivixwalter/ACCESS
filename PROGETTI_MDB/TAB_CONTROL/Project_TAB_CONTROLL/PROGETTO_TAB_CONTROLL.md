PROGETTO_TAB_CONTROLL.md

	Nota 	
		progetto del tabl controll e della sua gestioen che
		si trova in questo file e path:
			c:\CASA\LINGUAGGI\ACCESS\ACCESS_PROGETTI_MDB\
			TAB_CONTROLL.mdb

		Il progetto deve controllare l'esistenza di tutti oggetti associati alla form master principale.

	FORM
		GEST_CONDOMINIO
			Nota
				@form@master che contiene 8 tab
				Le sottoform sono:

			Option
				Opzione_02_txt
					Note
						Controllo le proprieta della form corrente limitatamente alla sottoform selezionata, esempio
						se è stata selezionata la pagina 1 vengono iterati tutti i controlli della sottoform 1 e la 
					   l'oggetto inserito.
					   N.B. NON ITERA NELLA PROPRIETA DELL'OGGETTO SOTTOFORM es. SottoForm_01 PERCHE'????

					   TODO: errore non riempe il recordset ??
					   	@opzione@02_(opzione 2 attiva il controllo oggetti della pagina del tab controll attiva)

			Sottoform

				MSys_FrmDF02_S01_TIPO_OGGETTO
				Nota
					Il tipo oggetto che compongono tutti i controlli oggetto.
					La form si basa sulla query di estrazione della @vista@tipo@oggetti
						@form@MSys_DF02_(@form@tipo@oggetto che definisci i tipi degli oggetti )

				MSys_FrmDF19_}-------------------------------------------------@
				MSys_FrmDF20_N01_Controls_GRUPPO_CONTROLLI
				Nota
					Form che evidenzia tutti i controlli salvati nella form master e sottocontrolli, si 
					basa sulla query
						@query@MSys@TB19_(query della form GRUPPO COTROLLI)
				sottoform


				MSys_FrmDF20_}-------------------------------------------------@
				MSys_FrmDF20_N01_Controls_GRUPPO_CONTROLLI

				Nota
					form che raggruppa tutti i controlli della master e controlli
						@form@MSys@DF20_(@form@gruppo@controlli che apparterngono alla form master )



	TABELLE
		MSys_DF02_TIPO_OGGETTO
			Note
				tabelle che definisce il tipo oggetto di tutti i controlli
					codice	
						@tabella@MSys@DF02_(tabella dei @tipo@oggetto definiti, @DF02)




		MSys_DF19_CODICI_CONTROLS

			Note
				tabelle che contiene i codici dei controlli e la loro denominazione salvati sull'oggetto
				form master e quelli incorporati.

						@tabella@MSys@DF19_(tabella degli @oggetti@incorporati e @oggetti@salvatio nel progetto)


	QUERY
		MSsys_DF02_}---------------------------------------------------@
		MSsys_DF02_Qry01_TIPO_OGGETTO
			
			Note
				VISTA tipo oggetti
						@query@MSys@DF02_(query di vista dei @tipo@oggetto definiti @DF02)_@vista@tipo@oggetti

		MSys_QryTB19_N01_Controls
			Note
				Vista di tutti i controlli salvati nella form master
						@query@MSys@TB19_(query di vista dei @CONTROLLI salvati nella form master)

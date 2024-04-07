START "apri il progetto dei comandi PER IL TUTORIAL DI 7z" "c:\\CASA\\LINGUAGGI\ZIP\\ZIP.sublime-project"

:-------@_PROGETTO_BAT_INCREMENTALE_@-------
		
	
	:01--------------------------------TITOLO COSTRUZIONE DI UN BAT INCREMENTALE

			GOTO DoveSiTrova

				Note
					questo esempio di bat incrementale si trova i questa path dello Zzip:	
						c:\CASA\LINGUAGGI\ACCESS\ACCESS_PROGETTI_MDB\MSYS_OGGETTI\SALVATAGGI_MSYS\


				PER COSTRUIRE UN BAT INCREMENTALE DENOMINATO 7ZIP_INCREMENTALE.BAT 
				occorre fare riferimento a questo progetto .BAT che si trova nel TUTORIAL ZIP E 7ZIP.

				Per individuare il progetto del bat incrementale occorre far riferimento al progetto 
				.MD  ZIP che si trova in questa path. Ogni modifica del progetto BAT incrementale 
				nelle sue istruzioni occorre riportarla in questa progettazione:


				TUTORIAL ZIP RAR O 7ZIP 
					@DOVE@SI@TROVA_(@7zip dove si trova il 7z con il tutorial)_@7z
							apro il dos con il power shell

							  @APRI_@COMANDI_@7z_(la path ed il comando per aprire il progetto 7zip)
								START "apri il progetto dei comandi PER IL TUTORIAL DI 7z" "c:\\CASA\\LINGUAGGI\ZIP\\ZIP.sublime-project"


				TUTTI I PROGETTI 
					Apro tutti i progetti con powershell
					START "apri il progetto dei comandi di TUTTI@PROGETTI" "c:\\CASA\\LINGUAGGI\\TROVA_PROGETTI\\TROVA_PROGETTI.sublime-project"


				SUBLIMETEXT 

					Per avere un tutorial veloce in sublimetex utiliza questo link
						https://www.w3schools.io/editor/sublime-text-terminal/
				:DoveSiTrova

	:02--------------------------------PARAMETRI_GENERALI, DISCO, PATH DI DESTINAZIONE, NOME ARCHIVIO.ZIP NOME LISTA FILE

				@REM 					LE SOSTITUZIONI GENERALI
				@REM @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
				@REM
				@REM 				@ESEMPIO@SALVATAGGIO@INCREMENTALE_(bat di esempio di salvataggio dei file incrementali)_@BAT@INCREMENTALE
				@REM
				@REM  		Disco e codice Disco
				@REM .......................................................
				@REM -----> 	c:
				@REM -----> 	_c
				@REM .......................................................


				@REM		directory y dove archiviare i dati = path di destinazione
				@REM   		@path@destinazione_(il parametro della path di destinazione si trova qui)
				@REM .......................................................
					SET PATH_PARTENZA_S=c:\CASA\LINGUAGGI\ACCESS\ACCESS_PROGETTI_MDB\MSYS_OGGETTI\
					SET PATH_DEST_S=c:\CASA\LINGUAGGI\ACCESS\ACCESS_PROGETTI_MDB\MSYS_OGGETTI\SALVATAGGI_MSYS\
				@REM .......................................................


				@REM 		Il nome di tutti i file di archivio dei dati, 7ZIP
				@REM   		@nome@file@salvataggio_(il parametro del nome del file per il salvataggio dei dati si trova qui)
				@REM .......................................................
					SET FILE_S=GE_COND_ARCHIVIO.zip


				@REM 		Il file di archivio dei dati è composito è formato dal nome dell'archivio + lista file.txt
				@REM   		@nome@file@lista_(il parametro del nome della lista file per il salvataggio dei dati si trova qui)
				@REM .......................................................
					SET LISTA_FILE_S=%FILE_S%_ListaFile.txt

					
				@ECHO .              CONTROLLO VARIABILI CREATE
				@ECHO ........................................................
					@ECHO .
					@ECHO 		CONTROLLO PARAMETRI PATH DI DESTINAZIONE FILE DI SALVATAGGIO E LISTA FILE

					@ECHO .
					@ECHO   	@controllo@Parametri_(Viene esegito il controllo dei parametri creati prima dei salvataggi)
					@ECHO.
					@ECHO  "path di PARTENZA 			:" %PATH_PARTENZA_S%
					@ECHO  "path di destinazione 		:" %PATH_DEST_S%
					@ECHO  "file di archiviazione.zip 	:" %FILE_S%
					@ECHO  "lista file processati	 	:" %LISTA_FILE_S%
				

				@REM 					LE SOSTITUZIONI GENERALI *** FINE ***
				@REM @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

				@REM ----> PAUSE
	:03-------------------------------CHIAMO DECODIFICA


					@ECHO  CHIAMO DECODIFICA.BAT
					@ECHO. DEVO POSIZIONARMI SULLA DIRECTORY, CHIAMARE IL BAT E TORNARE INDIETRO ALTRIMENTI NON FUNZIONA IL BAT
					@ECHO. in quanto il bat esegue la today ma non trova il file completo.
					
						CD c:\CASA\LINGUAGGI\ACCESS\ACCESS_PROGETTI_MDB\MSYS_OGGETTI\
						CALL Zip_N70_DECODIFICA.bat"
						CD..

	:04-------------------------------TIPO FILE DA ARCHIVIARE SE NON ESISTE LA LISTA FILE

				@REM @tipi@di@file_(da salvare raggruppati in una variabile per ora SONO BLOCCATI )
				@REM
				@REM imposta i file da salvare su diverse variabili (ATTENZIONE AGLI SPAZI DOPO VARIABILE=*.XLS ... NO CI DEVONO ESSERE
				@REM								ERRATO VARIABILE e spazio tra uguale e variabile es. --> TipoFile_Office_S =*.MDB)
				@SET TipoFile_TEXT_S=*.AUX *GZ *.LOG *.OUT *.TEX *.Tws
				@SET TipoFile_JAVA_S=*.CLASS *.EXE *HTLM *.IML *.JAR *.JAVA *.LOG *.MF *.LNK *.XML
				@SET TipoFile_TXT_S=*.TEX *.TWS *.TXT *BAT
				@SET TipoFile_Office_S=*.MDB *.XLS *.PDF

				:------------------------CONCATENA controlla ED ESEGUI
				SET TIPO_FILE_S=%TipoFile_TEXT_S% %TipoFile_JAVA_S% %TipoFile_TXT_S% %TipoFile_Office_S%

				ECHO.	
				ECHO. "--------------------------------------------------"
				ECHO. "ELENCO FILE CON ESTENSIONE DA SALVARE "
				ECHO. controllo lista file della variabile salvata
				ECHO. %TIPO_FILE_S%
				ECHO DIR  %TIPO_FILE_S%

				ECHO. TIPO DI FILE PER ORA BLOCCATO

				CLS

	:05-------------------------------CICLO FOR CON GESTIONE DATA CORRENTE DA RECUPERARE

					@REM CICLO FOR I E II CASO PER LA GESTIONE DELLA DATA
					@REM ========================================================================================================================
					:----------------------------CICLO FOR I CASO per la gestione della data con le sottostringhe ma aggiunge lo 0 se <10
					@for /f "skip=1" %%x in ('wmic os get localdatetime') do if not defined MyDate set MyDate=%%x
					@echo solo il giorno:
					@echo %MyDate:~6,2%

					@echo I CASO la data con le sottostringhe con separataore l'undescore (_)
					@set today=%MyDate:~0,4%_%MyDate:~4,2%_%MyDate:~6,2%

					@echo.
					@echo today in formato stringa: 
					@echo %today%
					@echo.

					:----------------------------CICLO FOR  II CASO cicolo for per per la data AAA MM GG senza sottostringhe

					@ECHO E' possibile ottenere la data corrente in modo indipendente dalle impostazioni locali utilizzando
					@ECHO SENZA armeggiare con le sottostringhe
					@echo vedi il link: https://qastack.it/programming/10945572/windows-batch-formatted-date-into-variable
					@echo .

					@echo. II CASO la data senza le sottostringhe solo numerico e con separatore il trattino (-)

					@REM for /f %%x in ('wmic path win32_localtime get /format:list ^| findstr "="') do set %%x
					@REM set today=%Year%_%Month%_%Day%

					@echo. IL GIORNO DEFINITO IN VARIABILE
					@ECHO %TODAY%
					echo.

					@REM CICLO FOR I E II CASO PER LA GESTIONE DELLA DATA  *** FINE ***
					@REM ========================================================================================================================

	:06------------------------------ SALVO TUTTI I FILE NELLE SOTTOCARTELLE /S con 7zip uso u -r
			@REM DIR in dos ha diversi parametri per visualizzare le cartelle ed i file che sono:
			@REM 		:  /A:-d mostra solo i file    
			@REM 		/O:n  mostra i file in ordine alfabetico   
			@REM		/W  per vedere l’elenco formato solo dai nomi dei file 
			@REM		BLOCCATO -----> DIR /S /A:-D /O:n /W %TIPO_FILE_S%> %PATH_DEST_S%%today%_ZIP_SALVATAGGI_ARCHIVI_ARCHIVIO_TIPO_FILE_SALVATI.TXT


			@REM -tzip 	=	parametro -t per tipo file processato cioè il file zip.
			@REM -r 	=	iterazione delle sottocartelle
			@REM @file 	=	lista file da processare
			@REM 		 	attenzione nella listafile.txt indicare la path e al posto di *.* occore solo * es. c:\dos\* pero per riportare la path completa
			@REM 		 	della sottodirectory occorre il parametro -r = sottodirectory e nel file .txt indicare la path completa senza * es.
			@REM 		 	c:\CASA\LINGUAGGI\ACCESS\
			@REM path   =	la path di salvataggio, per posizionare l'archivio in un'altra cartella occore scrivere il nome 
			@REM 		 	Dell'archivio preceduto dalla path estesa es.
			@REM 			C:\CASA\SALVATAGGI\archive.zip
			@REM -bsp1	= swithc che permette il reindirizzamento dell'output su file per l'elenco di tutti 
			@REM 		quelli processati e salvati su un file .txt, ma deve essere eseguito su doppio comando.
			@REM NOME DATA CORRENTE : per il nome sulla data corrente occorre utilizzare la variabile tmp
			@REM 			che contiene il nome in AAMMGG_, Attenzione per espandere la variabile occorre utilizzare ! ! altrimenti non funziona;
			@REM 			per ora la espanzione delle variabile con (!!) non è stata utilizzata 
			@REM 			es. %!PATH_DEST_S!%%today%_ZIP_SALVATAGGI_ARCHIVI.RAR %TIPO_FILE_S% %PATH_DEST_S%

			@REM  ATTENZIONE in @Zip_N70_ElencoFile.txt devi inserire solo le directory senza *.* e no le cartelle interne perchè altrimenti
			@REM 				salva solo le sottocartelle e non la cartella madre. Quindi nel file.txt indicare le cartelle ORIGINE in questo
			@REM  				modo: c:\CASA\LINGUAGGI\ACCESS\
			@REM  				modo: c:\CASA\LINGUAGGI\CPP\  ... ecc.


			@ECHO INIZIO SALVATAGGIO 
			@ECHO prima dell'operazione di salvataggio effettu un controllo della  path corrente dove prelevare i file da processare
			@ECHO .
			@ECHO . 
			cd %PATH_PARTENZA_S%

			dir 
			pause



			@REM qui effettuo il salvataggio con il nome DATA CORENTE AAAMMGG
			7z u -tzip -r %PATH_DEST_S%%TODAY%_%FILE_S% @Zip_N70_ElencoFile.txt
			
			@REM qui salvo i file processati nel .txt che è stato impostato nella variabile LISTA_FILE_S
			@REM ma devo prima spostarmi di directory e poi creare la lista con il command l e 
			@REM l'ozione -bsp1

			cd %PATH_DEST_S%
			7z l %TODAY%_%FILE_S% -bsp1 > %TODAY%_%FILE_S%_LISTA.TXT
			
	:07------------------------------ CHIAMO CRIPTA
					@ECHO  QUI CHIAMO CRIPTA.BAT
					@ECHO. SENZA POSIZIONARSI NELLA  DIRECTORY e non so il perchè qui fuzione e sopra no!!!
					

						CALL %PATH_DEST_S%Zip_N70_CRIPTA.bat

	:08------------------------------ SOSPENSIONE ED USCITA DALLA PROCEDURA
				
				@ECHO "--------------------------> USCITA DALLA PROCEDURA E SOSPENSIONE "

				
				@echo.
				@echo."=================================================================="
				@echo."			MSG - EXIT_PROCEDURA "
				@echo.
				@echo."ATTENZIONE '!!!' USCITA DALLA PROCEDURA : - USCITA - "			
				@echo."=================================================================="

				
					@REM 01)
					@REM sospensione per 1 secondo con crononometro 
					@REM ---->	timeout /t 2 /nobreak > NUL
					@REM ---->	
					@TIMEOUT /T 2 /NOBREAK

				EXIT

	GOTO NoteDos

			NOTE DOS DI UTILIZZO DELLA DIR CON I PARAMETRI

				@REM DIR @ORDINAMENTO_ALFABETICO_DECRESCENTE_PER_DATA_DI_CREAZIONE
				@REM ------------------------------------------------------------------------------------------- 
				@REM Sintassi 	DIR  [unità:][percorso][nomefile] [/P] [/W] [/O[[:]ordinamento]] [/S]
				@REM      	[/B] [/L] [/V]
					

				@REM  /A          Visualizza i file con gli attributi specificati.
				@REM  attributo	a:-d  (-d) = non le directory
				@REM		-h = non i file nascosti
				@REM /O		 Elenca file ordinandoli.
				@REM attributo	o:-n (-n) = escluso ordine alfabetico
				@REM /T		Controlla campo data visualizzato o utilizzato per ordinare.
				@REM attributo   t:c (:c)= ordina per creazione

					
					@REM ORDINA tutti i file escluso le directory ed i file nascosti in ordine alfabetico
					@REM decrescente e per data di creazione e salva nel file di testo.
					dir "c:\GESTIONI\GESTIONE_LLPP\02_SCANNER\" /a:-d -h /o:-n /t:c >>DOS_DIR_N01.txt

				@REM DIR @ORDINAMENTO_ALFABETICO_DECRESCENTE_PER_DATA_DI_CREAZIONE *** fine ***
				@REM ----------------------------------------------------






				@REM				PARAMETRI DOSE CON VARI ESEMPI DI COMANDO 
				@REM *************************************************************************************************************************
						@OPZIONI@DOS@DI_(ELENCO DI OPZIONI del comando DIR in dos)


					le opzioni principali sono: 
					/p per vedere l’elenco a schermate 
					/w per vedere l’elenco formato solo dai nomi dei file 
					/a:h mostra i file nascosti 
					/a:s mostra i file di sistema 
					/a:d mostra solo le subdirectory 
					/a:-d mostra solo i file 
					/o:n mostra in ordine alfabetico 
					/o:-n mostra in ordine alfabetico inverso 
					/o:d mostra in ordine di data 
					/o:-d mostra in ordine di data inverso 
					/o:s mostra in ordine di grandezza 
					/o:-s mostra in ordine di grandezza inverso 
					/s mostra tutto il disco 
					>prn stampa l’elenco 
						Esempio: 
						c:\> DIR c:\LETTERE\ROSSI /a:-d/o:n >prn

						stampa l’elenco, in ordine alfabetico, di tutti i file contenuti nella 
						subdirectory \LETTERE\ROSSI, omettendo i nomi delle eventuali subdirectory presenti. 




						@REM DIR @ORDINAMENTO_ALFABETICO_DECRESCENTE_PER_DATA_DI_CREAZIONE
						@REM ------------------------------------------------------------------------------------------- 
						@REM Sintassi 	DIR  [unità:][percorso][nomefile] [/P] [/W] [/O[[:]ordinamento]] [/S]
						@REM      	[/B] [/L] [/V]


						@REM  /A          Visualizza i file con gli attributi specificati.
						@REM  attributo	a:-d  (-d) = non le directory
						@REM			-h = non i file nascosti
						@REM /O		 Elenca file ordinandoli.
						@REM attributo	o:-n (-n) = escluso ordine alfabetico
						@REM
						@REM /T		Controlla campo data visualizzato o utilizzato per ordinare.
						@REM attributo   t:c (:c)= ordina per creazione


							@REM ORDINA tutti i file escluso le directory ed i file nascosti in ordine alfabetico
							@REM decrescente e per data di creazione e salva nel file di testo.
							dir "c:\GESTIONI\GESTIONE_LLPP\02_SCANNER\" /a:-d -h /o:-n /t:c >>DOS_DIR_N01.txt

						@REM DIR @ORDINAMENTO_ALFABETICO_DECRESCENTE_PER_DATA_DI_CREAZIONE *** fine ***
						@REM ----------------------------------------------------





						@REM DOS STRINGHE CONCATENA  @ESEMPIO@DOS@CONCATENA_(concatena le stringhe in dos)
						@REM ----------------------------------------------------

							:-------------------------------imposto le variabili set
							set str1="Hello"
							set str2="World"
							set str3= "1"

							:-------------------------------unisco le stringhe su una variabile 
							set "str3=%str1% %str3% %str2%"
							set "str4=%str1% %str2%"

							@REM qui scrivo str1 + la parola dos  + str2
							set "str1=%str1% DOS %str2%"

							:-------------------------------controllo concatenamento
							echo "str3 concatenata ->" %str3%
							echo "str4 concatenata ->" %str4%
							echo "str1 concatenata ->" %str1%




						@REM DOS STRINGHE CONCATENA	*** fine ***
						@REM ----------------------------------------------------





				:NoteDos

				@REM *************************************************************************************************************************

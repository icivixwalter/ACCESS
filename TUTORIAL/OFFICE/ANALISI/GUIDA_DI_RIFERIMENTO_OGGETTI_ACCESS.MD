GUIDA_DI_RIFERIMENTO_OGGETTI_ACCESS





	Guida di riferimento a DAO
	Novità di ADO 
				@RIFERIMENTO@OGGETTI@ACCESS_(le nuove caratteristiche ADO )

		Le nuove caratteristiche e la documentazione ampliata seguenti sono incluse in ADO versione 2.5. Questo elenco fa riferimento ad ADO, ADO MD e ADOX.
		Nuove caratteristiche
			Oggetti Record e Stream

				Questa versione di ADO include l'oggetto Record, che può rappresentare e gestire elementi quali directory e file in un file system, nonché cartelle e messaggi in un sistema di posta elettronica. Un oggetto Record può inoltre rappresentare una riga di un oggetto Recordset, sebbene gli oggetti Record e Recordset dispongano di metodi e proprietà diversi.

				Il nuovo oggetto Stream consente di leggere, scrivere e gestire il flusso binario di byte o testo che racchiude un file o un flusso di messaggi.

			Utilizzo degli URL

				In questa versione è inoltre incluso l'utilizzo degli URL (Uniform Resource Locator) come alternativa alle stringhe di connessione e al testo del comando, per denominare gli oggetti archivio di dati. È possibile utilizzare gli URL con gli oggetti Connection e Recordset esistenti nonché con i nuovi oggetti Record e Stream.

				In questa versione, ADO supporta i provider OLE DB che riconoscono i propri schemi degli URL. Ad esempio, OLE DB Provider for Internet Publishing, che accede ai file Windows 2000, riconosce lo schema HTTP esistente.

			Campi speciali per provider di origine di documenti

				Una classe speciale di provider, detti provider di origine di documenti, consente di gestire cartelle e documenti. Quando un oggetto Record rappresenta un documento o un oggetto Recordset rappresenta una cartella di documenti, il provider di origine di documenti popola questi oggetti con un set di campi univoco, che descrive le caratteristiche del documento. Questi campi costituiscono un oggetto Record o Recordset di una risorsa.
		
			Nuovi argomenti di riferimento
				In questa versione sono incluse le nuove proprietà seguenti.
		
				Proprietà 			Descrizione 
				Charset 				Indica il set di caratteri nel quale convertire il contenuto di un oggetto Stream di
				 						testo. 
				EOS 					Indica se la posizione corrente si trova alla fine del flusso. 
				LineSeparator 		Indica il carattere binario da utilizzare come separatore di riga negli oggetti Stream di
				 						testo. 
				Mode 					Indica le autorizzazioni disponibili per la modifica dei dati in un oggetto Connection, 
										Record o Stream. 
				ParentURL 			Indica una stringa di URL assoluto che punta al Record padre dell'oggetto Record corrente. 
				Position 			Indica la posizione corrente all'interno di un oggetto Stream. 
				RecordType 			Indica il tipo di un oggetto Record. 
				Size 					Indica le dimensioni del flusso in numero di byte. 
				Source 				Indica l'entità rappresentata dall'oggetto Record. 
				State 				Indica, per tutti gli oggetti applicabili, se lo stato dell'oggetto è aperto o chiuso. 
										Indica, per tutti gli oggetti applicabili che eseguono un metodo asincrono, se l'oggetto corrente è in corso di connessione, esecuzione o recupero.
				 
				Type 					Indica il tipo di dati incluso nell'oggetto Stream (binario o di testo). 


				NUOVI METODI
					In questa versione sono inclusi i nuovi metodi seguenti.


				Metodo Descrizione 

					CopyRecord 			Copia un file o directory, insieme al suo contenuto, in un'altra posizione. 
					CopyTo 				Copia il numero specificato di caratteri o byte (in base all'oggetto Type) dell'oggetto 
											Stream in un altro oggetto Stream. 
					DeleteRecord 		Elimina un file o una directory e tutte le relative sottodirectory. 
					Flush Forza 		il contenuto dell'oggetto Stream rimanente nel buffer ADO nell'oggetto sottostante al 
											quale l'oggetto Stream è associato. 
					GetChildren 		Restituisce un oggetto Recordset le cui righe rappresentano i file e le sottodirectory 
											nella directory rappresentata da questo Record. 
					LoadFromFile 		Carica il contenuto di un file esistente in un oggetto Stream. 

					MoveRecord 			Sposta un file o una directory, insieme al suo contenuto, in un'altra posizione. 
					Open 					Apre un oggetto Record esistente oppure crea un nuovo file o directory. 
					Open 					Apre un oggetto Stream per modificare i flussi di dati di testo o binari. 
					Read 					Legge un numero specificato di byte da un oggetto Stream binario. 
					ReadText 			Legge un numero specificato di caratteri da un oggetto Stream di testo. 
					SaveToFile 			Salva il contenuto binario di un oggetto Stream in un file. 
					SetEOS 				Imposta la posizione alla fine del flusso. 

				*** da finire ***
				Documentazione nuova e ampliata
					Argomenti relativi a esempi di codice

					Gli esempi sono stati ampliati per includere esempi di codice scritti in Microsoft Visual C++® e Microsoft Visual J++®. È possibile copiare e incollare questi esempi di codice nell'Editor in uso.

					Argomenti relativi ai provider

					È stato incluso un nuovo argomento che illustra l'utilizzo di ADO con OLE DB Provider for Internet Publishing.

					Programmazione con ADO

					In questa nuova sezione sono inclusi consigli e suggerimenti per l'utilizzo di ADO con vari linguaggi di programmazione. Essa include gli indici della sintassi esistenti per le estensioni di Visual C++ per ADO e ADO/WFC, nonché nuove informazioni specifiche per gli sviluppatori che utilizzano Microsoft Visual Basic®, Microsoft Visual Basic®, Scripting Edition (VBScript), Microsoft JScript®, Microsoft Visual C++ o Microsoft Visual J++.

				SkipLine 		Ignora un'intera riga durante la lettura di un oggetto Stream 
				Write 			Scrive dati binari in un oggetto Stream. 
				WriteText 		Scrive una stringa di testo specificata in un oggetto Stream. 
		


	Application
		Note
			@OGGETTO@APPLICATION
			
				TUTTA L'APPLICAZIONE
					L'oggetto Application fa riferimento all'applicazione di Microsoft Access attiva.
					L'oggetto Application contiene tutti gli oggetti di Microsoft Access (oggetto di Microsoft Access: Oggetto definito da Access, correlato ad Access e alla relativa interfaccia o alle maschere, ai report e alle pagine di accesso ai dati di un'applicazione e che viene utilizzato per programmare gli elementi dell'interfaccia utilizzati per l'immissione e la visualizzazione dei dati.) e gli insiemi (insieme: Oggetto che contiene un insieme di oggetti correlati. La posizione di qualsiasi oggetto nell'insieme può variare in caso di modifica dell'insieme.).

				UTILIZZI METODI E PROPRIETA PER L'INTERA APPLICAZIONE
					È possibile utilizzare l'oggetto Application per applicare metodi o impostazioni di proprietà all'intera applicazione di Microsoft Access. È possibile, ad esempio, utilizzare il metodo SetOption dell'oggetto Application per impostare le opzioni di database da Visual Basic. Nell'esempio seguente viene illustrato come impostare la casella di controllo Visualizza barra di stato (Display Status Bar) della scheda Database corrente della finestra di dialogo Opzioni di Access.
				
				OPZIONI DEL DABATASE CORRENTE - SETOPTION
					Nell'esempio seguente viene illustrato come impostare la casella di controllo Visualizza barra di stato 
					(Display Status Bar) della scheda Database corrente della finestra di dialogo Opzioni di Access.


						@OPZIONI del database corente si trovano in access 2007 nella menu: windows ----> OpzioniAccess ----> Visualizza ----> 

				Es.1_Barra di stato

					Esempio - @IMPOSTA@CASELLA@DI@CONTROLLO - esempio di applicazione del metodo set all'intera applicazione
					facendo visualizzare la barra di stato, impostando la scheda delle proprieta generale del db corrente:

						Application.SetOption "Show Status Bar", True 		//setta la proprieta visualizza barra di stato del db

			COM
				Microsoft Access è un componente COM che supporta l'automazione (automazione: Metodo per utilizzare gli oggetti di un'applicazione da un'altra applicazione o da un altro strumento di sviluppo. Precedentemente nota come automazione OLE, l'automazione è uno standard del settore e una caratteristica del modello COM (Component Object Model).), in precedenza nota come Automazione OLE. È possibile manipolare oggetti di Microsoft Access da un'altra applicazione in grado, anch'essa, di supportare l'automazione. A tale scopo, utilizzare l'oggetto Application.

				Microsoft Visual Basic, ad esempio, è un componente COM. È possibile aprire un database di Microsoft Access da Visual Basic e utilizzarne gli oggetti. Da Visual Basic creare innanzitutto un riferimento alla libreria di oggetti Microsoft Office Access 2007. Creare quindi una nuova istanza (istanza: Oggetto creato dalla classe che contiene la relativa definizione. Più istanze di una classe maschera condividono lo stesso codice e vengono caricate utilizzando gli stessi controlli con cui è stata strutturata la classe maschera.) della classe (classe: Definizione di un oggetto, inclusi il nome, le proprietà, i metodi e gli eventi disponibili per l'oggetto. Quando si crea un'istanza di una classe, si crea un nuovo oggetto con tutte le caratteristiche definite da quella classe.) Application e fare in modo che una variabile oggetto punti a tale istanza, come nell'esempio seguente:
					Es. Creare una istanza dell'oggetto application:
							
							Dim appAccess As New Access.Application 

				Dalle applicazioni che non supportano la parola chiave New è possibile creare una nuova istanza della classe Application utilizzando la funzione CreateObject:
					Es. creare una istanza senza new
							Dim appAccess As Object
							Set appAccess = CreateObject("Access.Application") 
			MANIPOLAZIONE OGGETTI
				Dopo avere creato una nuova istanza della classe Application, è possibile aprire un database o crearne uno nuovo tramite il metodo OpenCurrentDatabase o il metodo NewCurrentDatabase. È quindi possibile impostare le proprietà dell'oggetto Application e richiamarne i metodi. 
				PROPRIETA COMMANDSBAR ACCEDI ALLA BARRA DEI COMANDI : Quando si restituisce un riferimento all'oggetto CommandBars utilizzando la proprietà CommandBars dell'oggetto Application, è possibile accedere a tutti gli insiemi e gli oggetti della barra dei comandi (barra dei comandi: Le barre degli strumenti, le barre di menu e i menu di scelta rapida (clic destro) sono definiti collettivamente ''barre degli strumenti'' nell'interfaccia utente e ''barre dei comandi'' in VBA. Le barre dei comandi possono essere di tre tipi: barre dei menu, barre degli strumenti e menu popup.) di Microsoft Office System 2007 tramite questo riferimento.

				È inoltre possibile manipolare altri oggetti di Microsoft Access tramite l'oggetto Application. Utilizzando il metodo OpenForm dell'oggetto DoCmd di Microsoft Access, ad esempio, è possibile aprire una maschera di Microsoft Access da Microsoft Excel:
						Dim appAccess As New Access.Application 
						appAccess.DoCmd.OpenForm "Orders"
			FAQ
						sull'intera applicazione con l'oggetto APPLCATION si possono fare le seguenti operazioni :
							- applicare un metodo od una proprieta generale;
							- aggiungere o rimuovere una barra dei comandi;
							- impostare le opzioni del database,
							- componente COM puoi manipolare oggetti da un'altra applicazione;
							- aprire un database esterno 
							- creare un nuovo database
							- manipolare oggetti di accessa da un'altra applicazione come ad es. excel;

		Membri oggetto Application
			Note
				@MEMBRI@OGGETTO.APPLICATION = i membri dell'oggetto application sono:  

			Insieme Forms
			Insieme Modules
			Insieme References
			Insieme Reports
			Oggetto DoCmd
			Oggetto Module
			Oggetto Screen


				Metodi
					Nome
					AccessError 
					AddToFavorites 
					BuildCriteria 
					CloseCurrentDatabase 
					CodeDb 
					ColumnHistory 
					CompactRepair 
					ConvertAccessProject 
					CreateAccessProject 
					CreateAdditionalData 
					CreateControl 
					CreateForm 
					CreateGroupLevel 
					CreateNewWorkgroupFile 
					CreateReport 
					CreateReportControl 
					CurrentDb 
					CurrentUser 
					DAvg 
					DCount 
					DDEExecute 
					DDEInitiate 
					DDEPoke 
					DDERequest 
					DDETerminate 
					DDETerminateAll 
					DefaultWorkspaceClone 
					DeleteControl 
					DeleteReportControl 
					DFirst 
					DLast 
					DLookup 
					DMax 
					DMin 
					DStDev 
					DStDevP 
					DSum 
					DVar 
					DVarP 
					Echo 
					EuroConvert 
					Eval 
					ExportNavigationPane 
					ExportXML 
					FollowHyperlink 
					GetHiddenAttribute 
					GetOption 
					GUIDFromString 
					HtmlEncode 
					hWndAccessApp 
					HyperlinkPart 
					ImportNavigationPane 
					ImportXML 
					LoadCustomUI 
					LoadPicture 
					NewAccessProject 
					NewCurrentDatabase 
					Nz 
					OpenAccessProject 
					OpenCurrentDatabase 
					PlainText 
					Quit 
					RefreshDatabaseWindow 
					RefreshTitleBar 
					Run 
					RunCommand 
					SetDefaultWorkgroupFile 
					SetHiddenAttribute 
					SetOption
						Il metodo SetOption imposta il valore corrente di un'opzione disponibile nella finestra di dialogo Opzioni.
						@opzioni@SetOption@imposta.valore 

						OPZIONE FINESTRA DI DIALOGO = Il metodo SetOption imposta il valore corrente di un'opzione nella finestra di dialogo Opzioni di accesso.

						Sintassi

							espressione.SetOption(NomeOpzione, Impostazione)

							espressione   Variabile che rappresenta un oggetto Application.

						Parametri

						Nome 		Obbligatorio/Facoltativo	Tipo di dati 	Descrizione 
						NomeOpzione 	Obbligatorio 			String 		Nome dell'opzione. Per un elenco delle stringhe dell'argomento
														NomeOpzione, vedere Impostare le opzioni da Visual Basic. 
						Impostazione 	Obbligatorio 			Variant 	Valore corrispondente all'impostazione dell'opzione. Il valore
														dell'argomento Impostazione dipende dalle impostazioni possibili per un'opzione specifica. 

						Valore restituito
							Nessuno

						Note
							Le impostazioni disponibili per le opzioni dipendono dal tipo di opzione che viene impostata. Sono disponibili tre tipi generali di opzioni:

							Opzioni Sì/No che è possibile impostare attivando o disattivando una casella di controllo (casella di controllo: Controllo che indica se un'opzione è selezionata. Quando l'opzione è selezionata, nella casella viene visualizzato un segno di spunta.). 
							Opzioni che è possibile impostare immettendo una stringa (stringa: Insieme di caratteri che può includere numeri e testo.) o un valore numerico. 
							Opzioni predefinite che è possibile selezionare in una casella di riepilogo (casella di riepilogo: Controllo che fornisce un elenco di scelte. Una casella di riepilogo è costituita da un elenco e da un'etichetta facoltativa.), una casella combinata (casella combinata: Controllo utilizzato in una maschera per fornire le funzionalità combinate di una casella di riepilogo e di una casella di testo. È possibile digitare un valore o fare clic sul controllo per visualizzare un elenco e selezionarne una delle voci.) o un gruppo di opzioni (gruppo di opzioni: Cornice che può contenere caselle di controllo, interruttori e pulsanti di opzione in una maschera, report o pagina di accesso ai dati. Si utilizza per presentare alternative tra cui è possibile selezionare una sola opzione.). 

						OPZIONI SELEZIONATA DALL'UTENTE CASELLA DI CONTROLLO = Per le opzioni che vengono impostate dall'utente selezionando o deselezionando una casella di controllo, quando si utilizza il metodo SetOption specificare True o False per l'argomento impostazione, come nell'esempio seguente:
							Application.SetOption "Show Status Bar", True 
						OPZIONI MODELLO DI MASCHERA : Per impostare un tipo di opzione utilizzando il metodo SetOption, specificare la stringa o il valore numerico che si desidera digitare nella finestra di dialogo. Nell'esempio riportato di seguito viene impostato il modello di maschera predefinito su OrderTemplate:
							Application.SetOption "Form Template", "OrderTemplate" 

						OPZIONI CON CASELLE DI RIEPILOGO O CASELLE COMBINATE =  Per le opzioni le cui impostazioni sono selezioni in caselle di riepilogo o caselle combinate, specificare la posizione numerica dell'opzione all'interno dell'elenco come argomento Impostazione per il metodo SetOption. Nell'esempio seguente l'opzione Tipo di campo predefinito (Default Field Type) viene impostata su Contatore:
							Application.SetOption "Default Field Type", 5 		// 5 = NUMERO OPZIONE

						OPZIONE DI UN GRUPPO OPZIONI = Per impostare un'opzione che è membro di un gruppo di opzioni, specificare il numero di indice dell'opzione all'interno del gruppo. Nell'esempio riportato di seguito Impostazioni selezione viene impostato su Consenti selezione parziale:
							Application.SetOption "Selection Behavior", 1      	//1= NUMERO DI INDICE



						 Note 
							Quando si utilizza il metodo SetOption per impostare un'opzione nella in finestra di dialogo v, non occorre specificare la singola scheda nella quale si trova l'opzione. 
							Se il database è in esecuzione su una versione di Microsoft Access in una lingua diversa da quella in cui è stato creato, sarà necessario fornire in inglese gli argomenti per il metodo SetOption. 
							Quando si esce da Microsoft Access, è possibile reimpostare tutte le opzioni secondo le impostazioni originali utilizzando il metodo SetOption su tutte le opzioni modificate. È possibile che si desideri creare variabili pubbliche per memorizzare i valori delle impostazioni originali. È possibile includere del codice per reimpostare le opzioni nella routine evento Close per una maschera, o in una routine di uscita personalizzata che è necessario eseguire per uscire dall'applicazione.











					StringFromGUID 
					SysCmd 
					TransformXML 


				Proprieta
					Nome
					AnswerWizard 
					Application 
					Assistance 
					Assistant 
					AutoCorrect 
					AutomationSecurity 
					BrokenReference 
					Build 
					CodeContextObject 
					CodeData 
					CodeProject 
					COMAddIns 
					CommandBars 
							Classe
								CommandBar
									Membri della commandBar

										AdaptiveMenu
										Application
										BuiltIn
										Context
										Controls
										Creator
										Delete
										Enabled
										FindControl
										Height
										Index
										Left
										Name
										NameLocal
										Parent
										Position
										Protection
										Reset
										RowIndex
										ShowPopup
										Top
										Type
										Visible
										Width
					CurrentData 
					CurrentObjectName 
					CurrentObjectType 
					CurrentProject 
					DBEngine 
					DoCmd 
					FeatureInstall 
					FileDialog 
					Forms 
					IsCompiled 
					LanguageSettings 
					MacroError 
					MenuBar 
					Modules 
					Name 
					NewFileTaskPane 
					Parent 
					Printer 
					Printers 
					ProductCode 
					References 
					Reports 
					Screen 
					ShortcutMenuBar 
					TempVars 
					UserControl 
					VBE 
					Version 
					Visible 




		
		CommandBars
			Note



				METODI	
						Note 
						I metodi della CLASSE CommandBar sono i seguenti: 		@CommandBar@METODI
					
					

					Delete 
						Elimina l'oggetto della barra dei comandi dall'insieme in cui è contenuto.
							Nota
					FindControl
						Ottiene un oggetto CommandBarControl che soddisfa i criteri specificati.

							Nota 
								CommandBar.FindControl, metodo CHE 	Ottiene un oggetto CommandBarControl che soddisfa i criteri specificati.
								L'utilizzo delle barre dei comandi in alcune applicazioni di Microsoft Office è stato sostituito dalla nuova interfaccia utente basata sulla barra multifunzione. Per ulteriori informazioni, cercare "barra multifunzione" nella Guida. 

								Sintassi

								espressione.FindControl(Type, Id, Tag, Visible, Recursive)

								espressione   Variabile che rappresenta un oggetto CommandBar.

								Parametri

								Nome 		Obbligatorio/Facoltativo Tipo di dati Descrizione 
								Type 		Facoltativo Variant Specifica il tipo di controllo. 
								Id 			Facoltativo Variant Identificatore del controllo. 
								Tag 		Facoltativo Variant Valore del tag del controllo. 
								Visible 	Facoltativo Variant Deve avere valore True per includere nella ricerca solo i controlli delle barre dei comandi visibili. Il valore 			predefinito è False. Per barre dei comandi visibili si intendono tutte le barre degli strumenti visibili e tutti i menu aperti al 
											momento dell'esecuzione del metodo FindControl. 

								Recursive Facoltativo Variant Deve avere valore True per includere nella ricerca la barra dei comandi e tutte le relative barre degli strumenti popup secondarie. Questo argomento è valido solo per l'oggetto CommandBar. Il valore predefinito è False. 

								Valore restituito
									
									CommandBarControl

								Note


									RESTITUISCE IL PRIMO CONTROLLO
									Se l'insieme CommandBars contiene due o più controlli che soddisfano i criteri di ricerca, 
									FindControl restituirà il primo controllo trovato. Se non viene trovato alcun controllo che soddisfa i criteri di ricerca FindControl restituirà Nothing.
					ShowPopup
						Visualizza una barra dei comandi sotto forma di menu di scelta rapida in base alle coordinate specificate o alle coordinate correnti del puntatore. 

							Nota

								CommandBar.ShowPopup, metodo 
								Visualizza una barra dei comandi sotto forma di menu di scelta rapida in base alle coordinate specificate o alle coordinate correnti del puntatore.
								Nota 
								Se la proprietà Position della barra dei comandi non è impostata su msoBarPopup, verrà generato un errore. 

								Sintassi

								espressione.ShowPopup(x, y)

								espressione   Variabile che rappresenta un oggetto CommandBar.

								Parametri

								Nome Obbligatorio/Facoltativo Tipo di dati Descrizione 
								x Facoltativo Variant Coordinata X sulla quale si basa la posizione del menu di scelta rapida. Se questo argomento viene omesso, verrà utilizzata la coordinata X corrente del puntatore. 
								y Facoltativo Variant Coordinata Y sulla quale si basa la posizione del menu di scelta rapida. Se questo argomento viene omesso, verrà utilizzata la coordinata Y corrente del puntatore. 

								Note


								Quando i menu sono allineati a sinistra, l'angolo superiore sinistro del menu di scelta rapida visualizzato dal metodo ShowPopup si troverà in corrispondenza di (x, y + 1). Quando i menu sono allineati a destra, l'angolo superiore destro del menu di scelta rapida si troverà in corrispondenza di (x + 1, y + 1). È possibile utilizzare la funzione GetSystemMetrics(SM_MENUDROPALIGNMENT) di Windows per controllare la metrica del sistema per l'allineamento dei menu di scelta rapida.

								Quando la posizione sullo schermo delle coordinate (x, y) comporta la visualizzazione di tutto il menu di scelta rapida o di parte di esso oltre il bordo dello schermo visibile, il menu di scelta rapida viene spostato in modo da rientrare nell'area visualizzabile.


								Esempio

								In questo esempio viene creato un menu di scelta rapida contenente due controlli. Il metodo ShowPopup viene utilizzato per rendere visibile il menu di scelta rapida

								Visual Basic, Application Edition 
									Set myBar = CommandBars _
									    .Add(Name:="Custom", Position:=msoBarPopup, Temporary:=False)
									With myBar
									    .Controls.Add Type:=msoControlButton, Id:=3
									    .Controls.Add Type:=msoControlComboBox
									End With
									myBar.ShowPopup 


																
								Nota 
								L'utilizzo delle barre dei comandi in alcune applicazioni di Microsoft Office è stato sostituito dalla nuova interfaccia utente basata sulla barra multifunzione. Per ulteriori informazioni, cercare "barra multifunzione" nella Guida. 



					Reset 
						Ripristina la configurazione predefinita di una barra dei comandi incorporata.



		OGGETTO_COMMANDBARS
			
			Oggetto CommandBars

				INSIEME OGGETTI COMMANDBARS
				Insieme di oggetti CommandBar che rappresentano le barre dei comandi dell'applicazione contenitore.
				  Nota
		  			Rappresenta una barra dei comandi nell'applicazione contenitore. L'oggetto CommandBar è un membro dell'insieme CommandBars.
	 
					SOSTITUZIONE: L'utilizzo delle barre dei comandi in alcune applicazioni di Microsoft Office è stato sostituito dalla nuova interfaccia utente basata sulla barra multifunzione. Per ulteriori informazioni, cercare "barra multifunzione" nella Guida. 

					BARRA DEI COMANDI
						 La barra dei comandi in Access 2000 è una barra degli strumenti nella parte superiore della finestra di Access che contiene i pulsanti per i comandi di uso comune, come il salvataggio di un database o creazione di una nuova tabella. È possibile accedere ai comandi sulla barra dei comandi anche tramite la barra dei menu o utilizzando le scorciatoie da tastiera.  
						 	@BARRA@DEI@COMANDI@ACCESS.2000
						 	@BARRA@DEGLI@STRUMENTI






			

				Esempio 1

					Utilizzare la proprietà CommandBars per restituire l'insieme CommandBars. Nell'esempio seguente nella finestra Immediata vengono visualizzati il nome e il nome locale di ogni barra dei menu e barra degli strumenti, nonché un valore che indica se la barra dei menu o la barra degli strumenti è visibile.

						VBA_FOR_EACH_IN_COMMANDBARS
							For Each cbar in CommandBars
							    Debug.Print cbar.Name, cbar.NameLocal, cbar.Visible
							Next 

				Esempio 2
					...
				Esempio 3 RICERCA BARRA FORMS

					Utilizzare CommandBars(index), dove index è il nome o il numero di indice di una barra dei comandi, per restituire un oggetto CommandBar. Nell'esempio seguente viene esaminato l'insieme delle barre dei comandi allo scopo di trovare la barra dei comandi denominata "Forms". Se la barra viene trovata, verrà visualizzata e verrà protetto il relativo stato di ancoraggio. In questo esempio la variabile cb rappresenta un oggetto CommandBar.
						
						VBA_RICERCA BARRA FORMS



							Private Sub Form_Load()
							
							    Dim foundFlag  As Boolean
			
							    foundFlag = False
							'//cicla la barre e controlla se esiste Forms
							For Each cb In CommandBars
							    If cb.Name = "Forms" Then
								cb.Protection = msoBarNoChangeDock
								cb.Visible = True
								foundFlag = True
							    End If
							Next cb
							'//se non trova form msg errore
							If Not foundFlag Then
							    MsgBox "The collection does not contain a Forms command bar. - BARRA FORM NON TROVATA"
							End If

			
							End Sub

				Esempio 4 SPECIFICARE BARRA DEI MENU O UNA BARRA DEGLI STRUMENTI ESISTENTE
					Per specificare una barra dei menu o una barra degli strumenti esistente nell'elenco delle barre disponibili all'interno dell'applicazione contenitore, è possibile utilizzare il nome o il numero di indice. Per specificare un menu, un menu di scelta rapida o un sottomenu (tutti rappresentati da oggetti CommandBar), è tuttavia necessario utilizzare il nome. In questo esempio viene aggiunta un nuova voce di menu in fondo al menu Strumenti. Quando si fa clic sulla nuova voce di menu, viene avviata la routine denominata qtrReport.


					Visual Basic, Application Edition 
						
						'//ESEMPIO CHE SI TROVA NELLA FORM ----> Oggetto_Frm01_CommandBar_02_SpecificaMenu+BarraSTrumenti
						
						'//ALLA CHIUSURA DELLA FORM = dovrei cancellare la barra ma NON CI RIESCO
						Private Sub Form_Close()
						
						Dim CMD As CommandBar
						'//cancello se gia esiste in quanto ad ogni apertura della form aggiunge lo stesso menu MAKE REPORT
						'//NON RIESCO A CANCELLARLO ??? con delete, bisogna impostare l'oggetto.
						 ''//CommandBars("Tools").Delete        non funziona
						
						End Sub
						
						
						'//LOAD = CREO all'apertura della form il menu Make Report
						Private Sub Form_Load()
						
						    
						     
						    '//SPECIFICARE UNA BARRA: Per specificare una barra dei menu o una barra degli strumenti esistente nell'elenco _
						      delle barre disponibili all'interno dell'applicazione contenitore, è possibile utilizzare _
						      il nome o il numero di indice. _
						      SPECIFICARE UN MENU : Per specificare un menu, un menu di scelta rapida _
						      o un sottomenu (tutti rappresentati da oggetti CommandBar), è tuttavia necessario _
						      utilizzare il nome. In questo esempio viene aggiunta un nuova voce di menu _
						      in fondo al menu Strumenti. _
						      Quando si fa clic sulla nuova voce di menu, viene avviata la routine denominata qtrReport.
						    
						    '//all'apertura della form CREA LA BARRA TOOLS ed AGGIUNTE IL COMANDO Il menu Make Report
						    Set newItem = CommandBars("Tools").Controls.Add(Type:=msoControlButton)
						    With newItem
							
							.BeginGroup = True
							.Caption = "Make Report"
							.FaceId = 0
							.OnAction = "qtrReport"
						    End With
						    
						    
						
						
						End Sub

				Esempio 3 VERIFICA RESTITUZIONE MENU CORRETTO - INDIVIDUAZIONE CONTROLLO POPUP ***
					Se due o più menu o sottomenu personalizzati hanno lo stesso nome, CommandBars(index) restituirà il primo nome. Per verificare che venga restituito il menu o il sottomenu corretto, individuare il controllo popup che consente di visualizzare tale menu. 
					Applicare quindi la proprietà CommandBar al controllo popup per restituire la barra dei comandi che rappresenta tale menu. 
					In questo esempio si presume che il terzo controllo sulla barra degli strumenti denominata Strumenti aggiuntivi sia un controllo popup. 
					Viene quindi aggiunto il comando Save in fondo a tale menu.

					Visual Basic, Application Edition  *** non funziona ***
						Set viewMenu = CommandBars("Custom Tools").Controls(3)
						viewMenu.Controls.Add ID:=3    'ID of Save command is 3 
				


				
				Proprieta
					
					ActionControl *
						Ottiene l'oggetto CommandBarControl la cui proprietà OnAction è impostata sulla routine in esecuzione. Sola lettura.
					ActiveMenuBar 
						Ottiene un oggetto CommandBar che rappresenta la barra dei menu attiva nell'applicazione contenitore. Sola lettura.
							Nota 
								L'utilizzo delle barre dei comandi in alcune applicazioni di Microsoft Office è stato sostituito dalla nuova interfaccia utente basata sulla barra multifunzione. Per ulteriori informazioni, cercare "barra multifunzione" nella Guida. 
							Sintassi

								espressione.ActiveMenuBar

								espressione   Variabile che rappresenta un oggetto CommandBars.

							Esempio
								In questo esempio un controllo popup temporaneo denominato "Custom" viene posizionato in fondo alla barra dei menu attiva. 
								Viene quindi aggiunto un pulsante denominato "Import" alla barra dei comandi visualizzata dal controllo. 

								Visual Basic, Application Edition *** @NON@FUNZIONA
									Set myMenuBar = CommandBars.ActiveMenuBar
									Set newMenu = myMenuBar.Controls.Add(Type:=msoControlPopup, Temporary:=True)
									newMenu.Caption = "Custom"
									Set ctrl1 = newMenu.CommandBar.Controls _
									    .Add(Type:=msoControlButton, Id:=1)
									With ctrl1
									    .Caption = "Import"
									    .TooltipText = "Import"
									    .Style = msoButtonCaption
									End With 
					Position **
						Ottiene o imposta una costante MsoBarPosition che rappresenta la posizione di una barra dei comandi. Proprietà di lettura/scrittura.
							Nota 
								L'utilizzo delle barre dei comandi in alcune applicazioni di Microsoft Office è stato sostituito dalla nuova interfaccia utente basata sulla barra multifunzione. Per ulteriori informazioni, cercare "barra multifunzione" nella Guida. CommandBar.Position, proprietà 

							Sintassi

							espressione.Position

							espressione   Variabile che rappresenta un oggetto CommandBar.

							Esempio *** NON FUNZIONA ***
							In questo esempio viene esaminato l'insieme delle barre dei comandi, le barre personalizzate vengono ancorate nella parte inferiore della finestra dell'applicazione e le barre predefinite vengono ancorate nella parte superiore. 

							Visual Basic, Application Edition 
							For Each bar In CommandBars
							    If bar.Visible = True Then
								If bar.BuiltIn Then
								    bar.Position = msoBarTop
								 Else
								    bar.Position = msoBarBottom
								End If
							    End If
							Next 

					Item **
							Ottiene un oggetto CommandBar dall'insieme CommandBars. Sola lettura.
					
						Nota 
							L'utilizzo delle barre dei comandi in alcune applicazioni di Microsoft Office è stato sostituito dalla nuova interfaccia utente basata sulla barra multifunzione. Per ulteriori informazioni, cercare "barra multifunzione" nella Guida. 
								@CommandBars.Item, @proprieta@item

						Sintassi

							espressione.Item(Index)

							espressione   Argomento obbligatorio. Variabile che rappresenta un oggetto CommandBars.

						Parametri

							@parametri@item
							Nome Obbligatorio/Facoltativo Tipo di dati Descrizione 
							Index Obbligatorio Variant Specifica il nome o il numero di indice dell'oggetto da restituire. 

						@Esempio@item

								Item è il membro predefinito dell'oggetto o dell'insieme. Ad esempio, le due istruzioni seguenti assegnano entrambe 
								un oggetto CommandBar a cmdBar.
										@oggetto@commandBars@assegnazione@oggetto
										Set cmdBar = CommandBars.Item("Standard")
										Set cmdBar = CommandBars("Standard") 


				Metodi

					Add
						Note
							CommandBars.Add 

							metodo 
							Crea una nuova barra dei comandi e la aggiunge all'insieme delle barre dei comandi.L'utilizzo delle barre dei comandi in alcune applicazioni di Microsoft Office è stato sostituito dalla nuova interfaccia utente basata sulla barra multifunzione. Per ulteriori informazioni, cercare "barra multifunzione" nella Guida.

							Parametri

							Nome Obbligatorio/Facoltativo Tipo di dati Descrizione 
							Name Facoltativo Variant Nome della nuova barra dei comandi. Se questo argomento viene omesso, alla barra dei comandi viene assegnato un nome predefinito, ad esempio Custom 1. 
							Position Facoltativo Variant Specifica la posizione o il tipo della nuova barra dei comandi. Può essere una delle costanti MsoBar. 
							MenuBar Facoltativo Variant Se ha valore True, sostituisce la barra dei menu attiva con la nuova barra dei comandi. Il valore predefinito è False. 
							Temporary Facoltativo Variant Se ha valore True, rende temporanea la nuova barra dei comandi. Le barre dei comandi vengono eliminate alla chiusura dell'applicazione contenitore. Il valore predefinito è False. 



						Sintassi

							espressione.Add(Name, Position, MenuBar, Temporary)

							espressione   Argomento obbligatorio. Variabile che rappresenta un oggetto CommandBars.

							esempio
									Application.CommandBars.Add(Name, Position, MenuBar, Temporary)

			

	PROCEDURE
				 
			01) PROCEDURA DI AGGIUNTA E GESTIONE BARRE MENU E VOCI MENU

				CREO UNA BARRA MENU
					  Nota 
						L'utilizzo delle barre dei comandi in alcune applicazioni di Microsoft Office è stato sostituito dalla nuova interfaccia utente basata sulla barra multifunzione. Per ulteriori informazioni, cercare "barra multifunzione" nella Guida. 
						In alcune applicazioni contenitore non è disponibile un modo per creare nuove barre dei menu ed è quindi necessario crearle utilizzando Visual Basic. Dopo aver creato una barra dei menu in Visual Basic, è possibile personalizzarla utilizzando l'interfaccia dell'applicazione contenitore oppure Visual Basic.
							@CREAZIONE@BARRA.MENU

						Aggiunta di barre dei menu in fase di esecuzione
							Quando viene aggiunta una barra dei menu a un'applicazione in fase di esecuzione, è possibile utilizzare il metodo Add per l'insieme CommandBars e specificare True per l'argomento MenuBar. Nell'esempio seguente viene aggiunta una barra dei menu che può essere spostata. La barra dei menu viene inoltre ancorata lungo il lato destro della finestra dell'applicazione.
							Questo esempio si trova nella form MSys_OGGETTI_Frm13_BARRA_Elencovoci, il codice è -----> @CREO@MENU.BAR

							@ERRORE@NON@VISIBILE


							 '//Procedura: Aggiunta e gestione di barre dei menu e di voci di menu
								'//Aggiunto Dim e delete altrimenti da errore.
								    Dim MenuBar  As CommandBar
		
								    
								    '//cancello se gia esiste INSERITO PERCHE SE DA ERRORE
								    CommandBars("mBar").Delete
								    
								    '//CREA LA MENU BAR E A LA AGGIUNGE A SX ma non la posso modificare??
								    
								     Set MenuBar = CommandBars.Add _
									(Name:="mBar", Position:=msoBarRight, MenuBar:=True)
		
								    '//NON VEDO LA BARRA???'
								    With MenuBar
									.Protection = msoBarNoMove      '//protezione anti spostamento
									'.Visible = True                '//bloccato perche lo nasconde ??
									MenuBar.Position = msoBarBottom
									
									
								    End With
				    
				   



	
	
	OGGETTI_ACCESS    
		
		AccessObject 
			Oggetto AccessObject
			Un oggetto AccessObject fa riferimento a un oggetto (oggetto: Tabella, grafico, elemento grafico, 
			equazione o altri tipi di informazioni. Gli oggetti creati in un'applicazione, ad esempio in un foglio di calcolo, e collegati o incorporati in un'altra applicazione sono oggetti OLE.) di Microsoft Access specifico. 
		Note

			L'oggetto AccessObject include informazioni su un'istanza di un oggetto. 
			Nella tabella seguente sono inclusi i tipi di oggetto descritti da ogni oggetto AccessObject, il nome del relativo insieme e il tipo di informazioni contenute nell'oggetto AccessObject.

				AccessObject 		Insieme 			Contiene informazioni su 
				Diagramma di database 	AllDatabaseDiagrams 		Diagrammi di database salvati 
				Maschera 		AllForms 			Maschere salvate 
				Funzione 		AllFunctions 			Funzioni salvate 
				Macro 			AllMacros 			Macro salvate 
				Modulo 			AllModules 			Moduli salvati 
				Query 			AllQueries 			Query salvate 
				Report 			AllReports 			Report salvati 
				Stored procedure 	AllStoredProcedures 		Stored procedure salvate 
				Tabella 		AllTables 			Tabelle salvate 
				Visualizzazione 	AllViews 			Visualizzazioni salvate 


			OGGETTO ESISTENTE NON SI PUO CREARE O CANCELLARE  = Poiché un oggetto AccessObject corrisponde a un oggetto esistente, 
			non è possibile creare nuovi oggetti AccessObject o eliminarne di esistenti. 
			Per fare riferimento a un oggetto AccessObject in un insieme in base al numero ordinale o in base all'impostazione della proprietà Name, utilizzare una delle seguenti forme sintattiche: @FARE@RIFERIMENTO@ALL'INSIEME

				Insieme di oggetti

				AllForms(0) 
				AllForms("nome") 
				AllForms![nome] 
	
		Membri dell'oggetto AccessObject
			Un oggetto AccessObject fa riferimento a un oggetto di Microsoft Access specifico. 

			Metodi
				GetDependencyInfo 			Restituisce un oggetto DependencyInfo che rappresenta gli
				 					oggetti di database dipendenti dall'oggetto specificato.
					AccessObject.GetDependencyInfo
						 
						METODO - OGGETTI DIPENDENTI: Restituisce un oggetto DependencyInfo che rappresenta gli oggetti di database dipendenti dall'oggetto specificato.

						Sintassi

							espressione.GetDependencyInfo

							espressione   Variabile che rappresenta un oggetto AccessObject.
						Valore restituito
							DependencyInfo

						ERRORI RESTITUITI
							Questo metodo restituisce un errore di run-time se si verifica una delle condizioni seguenti:

							CASELLA DI CONTROLLO = La casella di controllo Tieni traccia delle informazioni di correzione (menu Strumenti, finestra di dialogo Opzioni, scheda Generale) è deselezionata. Per selezionare la casella Tieni traccia delle informazioni di correzione (Track Name AutoCorrect Info) e aggiornare le informazioni sulle dipendenze per tutti gli oggetti nel database, è possibile utilizzare il codice seguente:

								Application.SetOption "Track Name AutoCorrect Info", 1

							ERRORE PER MANCANZA DI AUTORIZZAZIONE = Non si dispone di autorizzazioni sufficienti per verificare le informazioni sulle dipendenze per l'oggetto AccessObject specificato


							PROGETTO ADP CHIAMA QUESTO METODO : Questo metodo viene chiamato da un progetto di Access (adp) (progetto di Microsoft Access: File di Access connesso a un database di Microsoft SQL Server e utilizzato per creare applicazioni client/server. Un file di progetto non contiene dati o oggetti basati su query di definizione dati, come tabelle o viste.).



							VBA NO RICERCA INFORMAZIONI = In Access la ricerca delle informazioni sulle dipendenze non viene eseguita nel codice di Visual Basic, Applications Edition (VBA), nelle macro (macro: Azione o gruppo di azioni che è possibile utilizzare per automatizzare le operazioni.), né nelle pagine di accesso ai dati (pagina di accesso ai dati: Pagina Web pubblicata da Access e connessa a un database. In una pagina di accesso ai dati è possibile visualizzare i dati memorizzati nel database, aggiungere altri dati oppure modificare e manipolare quelli esistenti.Una pagina di accesso ai dati può anche includere dati provenienti da altre origini, ad esempio Excel.).
				IsDependentUpon 		 	Restituisce un valore Boolean che indica se l'oggetto
			 						specificato dipende dall'oggetto di database 
									definito nell'argomento NomeOggetto.
					***
			Proprieta
						*** TUTTI DA FINIRE ***
				CurrentView 				È possibile utilizzare questa proprietà per restituire la
				 					visualizzazione corrente dell'oggetto di 
									Microsoft Access specificato. Costante AcCurrentView di sola lettura.
					***
				DateCreated 				Restituisce un valore Date che indica la data e l'ora 
									dell'ultima modifica apportata alla struttura 
									dell'oggetto specificato. Elemento in sola lettura.
					***
				DateModified 				Restituisce un valore Date che indica la data e l'ora 
									dell'ultima modifica apportata alla struttura 
									dell'oggetto specificato. Elemento di sola lettura.
						***
				FullName 				È possibile utilizzare questa proprietà per impostare o
				 					restituire il percorso completo, incluso il 
									nome del file, di un oggetto specifico. Elemento String di sola lettura.
					***
					***
				IsLoaded				È possibile utilizzare la proprietà IsLoaded per determinare 
									se un oggetto AccessObject è 
									attualmente caricato. Elemento Boolean in sola lettura. ***
				Name 					È possibile utilizzare la proprietà Name per determinare 
									l'espressione stringa che identifica 
									il nome un oggetto. Elemento String di sola lettura.
					***
				Parent 					Restituisce l'oggetto padre per l'oggetto specificato. 
									Elemento di sola lettura.
					**
				Properties 				Restituisce un riferimento all'insieme AccessObjectProperties 
									dell'oggetto AccessObject. 
									Elemento di sola lettura.
					**
				Type 					È possibile utilizzare questa proprietà per restituire il 
									valore di un tipo di oggetto AccessObject. 
									Costante AcObjectType di sola lettura.


		AccessObjectProperties **



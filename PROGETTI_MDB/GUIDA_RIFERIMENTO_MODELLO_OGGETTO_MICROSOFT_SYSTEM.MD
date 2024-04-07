OFFICE VISUAL BASIC 2007

	GUIDA INTRODUTTIVA  **
	NOVITA   **
	CONCETTI	 **
	RICERCA PER CATEGORIE **
	

	GUIDA_RIFERIMENTO_MODELLO_OGGETTO_MICROSOFT_SYSTEM_2007

	Oggetto AnswerWizard **
	Oggetto AnswerWizard **
	Oggetto AnswerWizardFiles **
	Oggetto Assistant **
	Oggetto Balloon **
	Oggetto BalloonCheckbox **
	Oggetto BalloonCheckboxes **
	Oggetto BalloonLabel **
	Oggetto BalloonLabels **
	Oggetto BulletFormat2 **
	Oggetto COMAddIn **
	Oggetto COMAddIns **
	Oggetto CommandBar **
	Oggetto CommandBarButton **
	Oggetto CommandBarComboBox **
	Oggetto CommandBarControl

		CONTROLLO BARRA DEI COMANDI
			UN CONTROLLO DELLA BARRA DEI COMANDI: Rappresenta un controllo barra dei comandi. L'oggetto CommandBarControl è un membro dell'insieme CommandBarControls. Le proprietà e i metodi dell'oggetto CommandBarControl sono tutti condivisi dagli oggetti 
				CommandBarButton, 
				CommandBarComboBox e 
				CommandBarPopup.

		UTILIZZO
			L'utilizzo delle barre dei comandi in alcune applicazioni di Microsoft Office è stato sostituito dalla nuova interfaccia utente basata sulla barra multifunzione. Per ulteriori informazioni, cercare "barra multifunzione" nella Guida.


		NOTE
		 	Quando si scrive codice di Visual Basic che verrà utilizzato per lavorare con i controlli della barra dei comandi personalizzata, è possibile utilizzare gli oggetti CommandBarButton, CommandBarComboBox e CommandBarPopup.
		 	Quando si scrive codice che verrà utilizzato per lavorare con i controlli incorporati nell'applicazione contenitore che non possono essere rappresentati da uno dei tre oggetti, è possibile utilizzare l'oggetto CommandBarControl. 
		 	COMMANDBARCONTROLL : Utilizzare Controls(index), dove index è il numero d'indice di un controllo, per restituire un oggetto CommandBarControl. (La proprietà Type del controllo deve essere impostata su 
		 		msoControlLabel, 
		 		msoControlExpandingGrid, 
		 		msoControlSplitExpandingGrid, 
		 		msoControlGrid o 
		 		msoControlGauge). 
		 		Alle variabili dichiarate come CommandBarControl è possibile assegnare i valori CommandBarButton, CommandBarComboBox e CommandBarPopup.


	 	Esempio

			È inoltre possibile utilizzare il metodo FindControl per restituire un oggetto CommandBarControl. Nell'esempio seguente viene eseguita la ricerca di un controllo di tipo msoControlGauge. Se ne viene trovato uno, verranno visualizzati il relativo numero di indice e il nome della barra dei comandi che lo contiene. In questo esempio la variabile lbl rappresenta un oggetto CommandBarControl.

			Visual Basic, Application Edition 
					
					Set lbl = CommandBars.FindControl(Type:= msoControlGauge)
					If lbl Is Nothing Then
					    MsgBox "A control of type msoControlGauge was not found."
					Else
					    MsgBox "Control " & lbl.Index & " on command bar " _
					        & lbl.Parent.Name & " is type msoControlGauge"
					End If 

	
	Oggetto Application
			Application
				Note
						L'oggetto Application fa riferimento all'applicazione di Microsoft Access attiva.
							@oggetto@Application_(Oggetto che fa rifermento all'applicazione access attiva, metodi e proprieta)
				
				Osservazioni
				
					L'oggetto Application contiene tutti gli oggetti e gli insiemi di Microsoft Access.
				
					È possibile utilizzare l'oggetto Application per applicare metodi o impostazioni di proprietà all'intera applicazione di Microsoft Access. Per impostare da Visual Basic le opzioni del database, è possibile, ad esempio, utilizzare il metodo SetOption dell'oggetto Application. L'esempio riportato di seguito indica come sia possibile impostare la casella di controllo Barra di stato in Mostra della scheda Visualizza della finestra di dialogo Opzioni.
					
					Application.SetOption "Mostra barra di stato", TrueMicrosoft Access è un componente COM che supporta l'automazione, precedentemente nota come Automazione OLE. È possibile manipolare oggetti di Microsoft Access da un'altra applicazione in grado, anch'essa, di supportare l'automazione. A tale scopo, utilizzare l'oggetto Application.
					
					Microsoft Visual Basic è, ad esempio, un componente COM. È possibile aprire un database di Microsoft Access da Visual Basic e lavorare con i relativi oggetti. Da Visual Basic, creare innanzitutto un riferimento alla libreria di oggetti di Microsoft Access 9.0. Creare quindi una nuova istanza della classe Application e puntare ad essa una variabile di oggetto, come nell'esempio seguente:
					
					Dim appAccess As New Access.ApplicationDalle applicazioni che non supportano la parola chiave New è possibile creare una nuova istanza della classe Application utilizzando la funzione CreateObject:
					
					Dim appAccess As Object
					Set appAccess = CreateObject("Access.Application.9")Una volta creata una nuova istanza della classe Application, è possibile aprire un database o crearne uno nuovo tramite i metodi OpenCurrentDatabase o NewCurrentDatabase. È quindi possibile impostare le proprietà dell'oggetto Application e richiamare i suoi metodi. Quando si restituisce un riferimento all'oggetto CommandBars utilizzando la proprietà CommandBars dell'oggetto Application, è possibile accedere a ogni oggetto e insieme della barra dei comandi OFF9TBD utilizzando questo riferimento.
					
					È inoltre possibile manipolare altri oggetti di Microsoft Access tramite l'oggetto Application. Utilizzando il metodo OpenForm dell'oggetto DoCmd di Microsoft Access, ad esempio, da Microsoft Excel è possibile aprire una maschera di Microsoft Access.
					
					appAccess.DoCmd.OpenForm "MascheraOrdine"Per ulteriori informazioni sulla creazione di un riferimento e sul controllo di oggetti tramite l'automazione, consultare la documentazione relativa all'applicazione che al momento risulta componente COM.

				Esempio di oggetto Application

					Nell'esempio riportato di seguito viene indicato come stampare alcune impostazioni correnti di proprietà dell'oggetto Application, impostare un'opzione e quindi uscire dall'applicazione salvando tutti gli oggetti:
					
					Sub InformazioniApplicazione()
				    	' Stampa nome e tipo di oggetto corrente.
				    	Debug.Print Application.CurrentObjectName
				    	Debug.Print Application.CurrentObjectType
				    	' Imposta opzione Oggetti nascosti in Visualizza
				    	' nella scheda Visualizzazione della finestra Opzioni.
				    	Application.SetOption "Mostra oggetti nascosti", True
				    	' Esce da Microsoft Access, salvando tutti gli oggetti.
				    	Application.Quit acSaveYes
					End Sub
					L'esempio successivo mostra come utilizzare Microsoft Access come componente COM. Da Microsoft Excel, Visual Basic o da un'altra applicazione che agisce come componente COM, creare un riferimento a Microsoft Access scegliendo Riferimenti dal menu Strumenti nella finestra Moduli. Selezionare la casella di controllo posta accanto a Microsoft Access 9.0 Object Library. Immettere quindi il codice riportato di seguito in un modulo di Visual Basic all'interno di tale applicazione e chiamare la routine RicercaDati.
					
					Nell'esempio, a una routine che crea una nuova istanza della classe Application vengono passati un nome di database e un nome di report, viene aperto il database e stampato il report indicato.
					
					' Dichiara variabile di oggetto nella sezione Dichiarazioni di un modulo
				    	Dim appAccess As Access.Application
					
					Sub RicercaDati()
				    	Dim strDB As String
				    	Dim strNomeReport As String
					
				    	' Inizializza stringa a percorso database.
				    	strDB = "C:\Program Files\Microsoft Office\Office\Esempi\Northwind.mdb"
				    	' Inizializza stringa a nome report.
				    	strNomeReport = "Catalogo"
				    	PrintAccessReport strDB, strNomeReport
					End Sub
					
					Sub StampaReport(strDB As String, strNomeReport As String)
				    	' Restituisce riferimento a oggetto Application di Microsoft Access.
				    	Set appAccess = New Access.Application
				    	' Apre database in Microsoft Access.
				    	appAccess.OpenCurrentDatabase strDB
				    	' Stampa report.
				    	appAccess.DoCmd.OpenReport strNomeReport
				    	MsgBox "Fare clic su OK al " & strNomeReport & _
				        	" termine della stampa"
				    	appAccess.CloseCurrentDatabase
				    	Set appAccess = Nothing
					End Sub
			Gli oggetti sottostanti dell'Oggetto Application
				Insieme CommandBars
				Insieme DataAccessPages
				Insieme Forms
				Insieme Modules
				Insieme References
				Insieme Reports
				Oggetto AnswerWizard
				Oggetto Assistant
				Oggetto CodeData
				Oggetto CurrentData
				Oggetto CurrentProject
				Oggetto DefaultWebOptions
				Oggetto DoCmd
				Oggetto FileSearch
				Oggetto LanguageSettings
            Oggetto Screen
            Oggetto VBE    

			Elenco_Metodi_e_Proprieta
				Metodi

					Nome Metodo	               	|					@DESCRIZIONE@METODO
					-----------------------------------------------------------------------------------------------------------------------------------
					AccessError 			    |	È possibile utilizzare il metodo AccessError per restituire la stringa descrittiva associata a un 
												|	errore di Microsoft Access o a un errore DAO.
					AddToFavorites 		   		|	Il metodo AddToFavorites aggiunge un indirizzo di collegamento ipertestuale alla cartella Preferiti.
					BuildCriteria 			   	|	Il metodo BuildCriteria restituisce una stringa di criteri analizzata, nel modo in cui viene 
												|	visualizzata nella griglia i 
												|	struttura della query, nella modalità Filtro in base a maschera o nella modalità Filtro su server in 
												|	base a maschera. Si 
												|	supponga, ad esempio, di voler impostare  la proprietà Filter o ServerFilter di una maschera in base 
												|	a criteri variabili stabiliti dall'utente. È possibile utilizzare il metodo uildCriteria per
												|  	costruire l'argomento espressione stringa per la proprietà Filter o ServerFilter. Elemento di tipo 
												|	stringa.
					CloseCurrentDatabase    	|	È possibile utilizzare il metodo CloseCurrentDatabase per chiudere il database corrente, ovvero un 
												|	database di Microsoft 	Access o un progetto di Microsoft Access (adp), da un'altra applicazione in 
												|	cui è stato aperto un database mediante automazione.
					CodeDb 					    |	È possibile utilizzare il metodo CodeDb in un modulo di codice per determinare il nome dell'oggetto 
												|	Database che fa riferimento al database in cui il codice è attualmente in esecuzione. Utilizzare il 
												|	metodo CodeDb per accedere a oggetti DAO che fanno parte Di un database di libreria.
					ColumnHistory 			    |	Gets the history of values that have been stored in a Memo field.
					CompactRepair 			    |   Compatta e ripristina il database o il file di progetto (adp) di Microsoft Access specificato. 
												|	Restituisce un valore Boolean. Restituisce True se il processo è stato eseguito in modo corretto.
					ConvertAccessProject    	|	Converte il file di Microsoft Access specificato da una versione a un'altra.
					CreateAccessProject 	   	|	È possibile utilizzare il metodo CreateAccessProject per creare un nuovo progetto di Microsoft Access 
												|	(adp) su disco.
					CreateAdditionalData    	|	È possibile utilizzare questo metodo per creare un oggetto AdditionalData che può essere utilizzato 
												|	per aggiungere nuove tabelle e query alla tabella padre esportata tramite il metodo ExportXML .
					CreateControl 			   	|	Il metodo CreateControl consente di creare un controllo nella maschera aperta specificata. Si 
												|	supponga ad esempio di creare una procedura guidata personalizzata per consentire agli utenti di 
												|	Creare in modo semplice una determinata maschera. È possibile utilizzare il metodo CreateControl 
												|	nella procedura guidata per aggiungere alla maschera i controlli appropriati.
					CreateForm 				   	|	Il metodo CreateForm crea una maschera e restituisce un oggetto Form.
					CreateGroupLevel 		   	|	È possibile utilizzare il metodo CreateGroupLevel per specificare un'espressione o un campo in base 
												|	al quale raggruppare o ordinare i dati di un report.
					CreateNewWorkgroupFile 		|	Crea un nuovo file di gruppo di lavoro in modo che un utente possa accedere a un database in modo 
												|	protetto.
					CreateReport 				|	Il metodo CreateReport crea un report e restituisce un oggetto Report. Si supponga, ad esempio, di 
												|	creare una procedura guidata personalizzata per creare un rapporto vendite. È possibile utilizzare il 
												|	metodo CreateReport nella procedura guidata per creare un nuovo report basato su un modello di report 
												|	specifico.
					CreateReportControl 		|	Il metodo CreateReportControl crea un controllo in un report aperto specificato. Per ulteriori 
												|	informazioni, vedere il metodo CreateControl .
					CurrentDb 					| 	Il metodo CurrentDb restituisce una variabile oggetto di tipo Database che rappresenta il database 
												|	attualmente aperto nella finestra di Microsoft Access.
					CurrentUser 				|	È possibile utilizzare il metodo CurrentUser per restituire il nome dell'utente corrente del database.
					DAvg 						|	È possibile utilizzare la funzione DAvg per calcolare la media di un set di valori contenuti in un 
												|	set di record specifico, ovvero un dominio.
					DCount 						|	È possibile utilizzare la funzione DCount per stabilire il numero di record contenuti in un set di 
												|	record specifico, ovvero un dominio.
					DDEExecute 					|	È possibile utilizzare l'istruzione DDEExecute per inviare un comando da un'applicazione client a 
												|	un'applicazione server 	attraverso un canale DDE (Dynamic Data Exchange) aperto.
					DDEInitiate 				|	È possibile utilizzare la funzione DDEInitiate per iniziare una conversazione DDE (Dynamic Data 
												|	Exchange) con un'altra 	applicazione. La funzione DDEInitiate apre un canale DDE per trasferire i 
												|	dati da un server DDE a un'applicazione client.
					DDEPoke 					|	È possibile utilizzare l'istruzione DDEPoke per trasferire dati di tipo testo da un'applicazione 
												|	client a un'applicazione server attraverso un canale DDE (Dynamic Data Exchange) aperto.
					DDERequest 					|	È possibile utilizzare la funzione DDERequest attraverso un canale DDE (Dynamic Data Exchange) aperto 
												|	per richiedere un'informazione a un'applicazione server DDE.
					DDETerminate 				|	È possibile utilizzare l'istruzione DDETerminate per chiudere un canale DDE (Dynamic Data Exchange) 
												|	specifico.
					DDETerminateAll 			|	È possibile utilizzare l'istruzione DDETerminateAll per chiudere tutti i canali DDE (Dynamic Data 
												|	Exchange) aperti.
					DefaultWorkspaceClone 		|	È possibile utilizzare il metodo DefaultWorkspaceClone per creare un nuovo oggetto Workspace senza 
												|	richiedere all'utente di riconnettersi. Se, ad esempio, è necessario eseguire simultaneamente due set 
												|	di transazioni in aree di lavoro separate, è possibile utilizzare il metodo DefaultWorkspaceClone per 
												|	creare un secondo oggetto Workspace con lo stesso nome dell'utente e la stessa password in modo che 
												|	non sia necessario richiedere di nuovo tali informazioni.
					DeleteControl 				|	Il metodo DeleteControl elimina un controllo specificato da una maschera.
					DeleteReportControl 		|	Il metodo DeleteReportControl elimina un controllo specificato da un report.
					DFirst 						|	È possibile utilizzare la funzione DFirst per restituire un record casuale da un determinato campo di 
												|	una tabella o query quando è necessario disporre solo di un valore qualsiasi di tale campo.
					DLast 						|	È possibile utilizzare la funzione DLast per restituire un record casuale da un campo specifico di 
												|	una tabella o di una query quando è necessario recuperare solo un valore qualsiasi di tale campo.
					DLookup 					|	È possibile utilizzare la funzione DLookup per ottenere il valore di un determinato campo da un set 
												|	di record specificato, ovvero un dominio.
					DMax 						|	È possibile utilizzare la funzione DMax per determinare il valore massimo di un set di record 
												|	specifico, ovvero un dominio.
					DMin 						|	È possibile utilizzare la funzione DMin per determinare il valore minimo in un set di record 
												|	specifico, ovvero un dominio.
					DStDev 						|	Calcola la deviazione standard su un campione di popolazione in un set di record specifico, ovvero un 
												|	dominio.
					DStDevP 					|	Calcola la deviazione standard di una popolazione in un set di record specifico, ovvero un dominio.
					DSum 						|	Calcola la varianza su un campione in un set di record specifico, ovvero un dominio.
					DVar 						|	Calcola la varianza su un campione in un set di record specifico, ovvero un dominio.
					DVarP 						|	Calcola la varianza di una popolazione in un set di record specifico, ovvero un dominio.
					Echo 						|	Il metodo Echo specifica se Microsoft Access debba aggiornare la visualizzazione del contenuto dello 
												|	schermo.
					EuroConvert 				|	È possibile utilizzare la funzione EuroConvert per convertire in euro un valore espresso in una 
												|	valuta facente parte dell'euro e viceversa. È anche possibile eseguire la conversione da una valuta 
												|	facente parte dell'euro a un'altra valuta utilizzando l'euro come valore intermedio, effettuando 
												|	pertanto una triangolazione. La funzione EuroConvert utilizza i tassi di conversione fissi stabiliti 
												|	dall'Unione Europea.
					Eval 						|	È possibile utilizzare la funzione Eval per restituire un'espressione risultante in una stringa di 
												|	testo o in un valore numerico.
					ExportNavigationPane 		|	Saves the current configuration of the Navigation Pane to an XML file.
					ExportXML 					|	Il metodo ExportXML consente agli sviluppatori di esportare schemi e dati XML e informazioni sulle 
												|	presentazioni da Microsoft SQL Server 2000 Desktop Engine (MSDE 2000), Microsoft SQL Server 6.5 o 
												|	versione successiva o dal modulo di gestione di database Microsoft Access.
					FollowHyperlink 			|	Il metodo FollowHyperlink apre il documento o la pagina Web specificati da un indirizzo di 
												|	collegamento ipertestuale.
					GetHiddenAttribute 			|	Il metodo GetHiddenAttribute restituisce il valore di un attributo nascosto di un oggetto di 
												|	Microsoft Access nella finestra di dialogo Proprietà dell'oggetto, visualizzata se si seleziona 
												|	l'oggetto nella finestra del database e si sceglie Proprietà dal menu Visualizza.
					GetOption 					|	Il metodo GetOption restituisce il valore corrente di un'opzione della finestra di dialogo Opzioni, 
												|	disponibile scegliendo Opzioni dal menu Strumenti. Elemento Variant.
					GUIDFromString 				|	La funzione GUIDFromString converte una stringa in un GUID, cioè una matrice di tipo Byte.
					HtmlEncode 					|	Converts a string to an HTML-encoded string.
					hWndAccessApp 				|	È possibile utilizzare il metodo hWndAccessApp per determinare l'handle assegnato da Microsoft 
												|	Windows alla finestra principale di Microsoft Access.
					HyperlinkPart 				|	Il metodo HyperlinkPart restituisce informazioni sui dati memorizzati come tipo di dati Collegamento 
												|	ipertestuale.
					ImportNavigationPane 		|	Loads a saved Navigation Pane configuration from disk.
					ImportXML 					|	Il metodo ImportXML consente agli sviluppatori di importare informazioni su schemi e/o dati XML in 
												|	Microsoft SQL Server 2000 Desktop Engine (MSDE 2000), Microsoft SQL Server 7.0 o versione successiva 
												|	o nel modulo di gestione di database Microsoft Access.
					LoadCustomUI 				|	Loads XML markup that represents a customized Ribbon.
					LoadPicture 				|	Il metodo LoadPicture carica un'immagine in un controllo ActiveX.
					NewAccessProject 			|	È possibile utilizzare il metodo NewAccessProject per creare ed aprire un nuovo progetto di Microsoft 
												|	Access (adp) come progetto corrente nella finestra di Microsoft Access.
					NewCurrentDatabase 			|	È possibile utilizzare il metodo NewCurrentDatabase per creare un nuovo database di Microsoft Access (
												|	mdb) nella finestra di Microsoft Access.
					Nz 							|	È possibile utilizzare la funzione Nz per restituire zero, una stringa di lunghezza zero (" ") o un 
												|	altro valore specificato quando un valore di tipo Variant è Null. Ad esempio, è possibile utilizzare 
												|	questa funzione per convertire un valore Null in un altro valore ed evitare che si propaghi in 
												|	un'espressione.
					OpenAccessProject 			|	È possibile utilizzare il metodo OpenAccessProject per aprire un progetto di Microsoft Access (adp) 
												|	esistente come progetto corrente nella finestra di Microsoft Access.
					OpenCurrentDatabase 		|	È possibile utilizzare il metodo OpenCurrentDatabase per aprire un database di Microsoft Access 
												|	esistente come database corrente.
					PlainText 					|	Strips the rich text formatting from a string and returns an unformatted text string.
					Quit 						|	Il metodo Quit consente di uscire da Microsoft Access. È possibile selezionare una delle opzioni 
												|	disponibili per salvare un oggetto di database prima di uscire.
					RefreshDatabaseWindow 		|	Il metodo RefreshDatabaseWindow aggiorna la finestra del database dopo la creazione, l'eliminazione o 
												|	la ridenominazione di un oggetto di database.
					RefreshTitleBar 			|	Il metodo RefreshTitleBar aggiorna la barra del titolo di Microsoft Access in seguito 
												|	all'impostazione in Visual Basic della proprietà AppTitle o AppIcon.
					Run 						|	È possibile utilizzare il metodo Run per eseguire una routine Function o Sub di Microsoft Access o 
												|	definita dall'utente specifica. Elemento Variant.
					RunCommand 					|	Il metodo RunCommand esegue un comando incorporato.
					SetDefaultWorkgroupFile 	|	Imposta il file specificato come file del gruppo di lavoro predefinito.
					SetHiddenAttribute 			|	Il metodo SetHiddenAttribute imposta l'attributo nascosto di un oggetto di Access.
					SetOption 					|	Il metodo SetOption imposta il valore corrente di un'opzione disponibile nella finestra di dialogo 
												|	Opzioni.
					StringFromGUID 				|	La funzione StringFromGUID converte un GUID, ovvero una matrice di tipo Byte, in una stringa.
					SysCmd						| 	È possibile utilizzare il metodo SysCmd per visualizzare nella barra di stato un indicatore di 
												|	avanzamento o una stringa di testo facoltativa, restituire informazioni su Microsoft Access e i file 
												|	associati oppure restituire informazioni sullo stato di un oggetto di database specificato, per 
												|	indicare se l'oggetto è aperto, se si tratta di un nuovo oggetto oppure di un oggetto modificato ma 
												|	non salvato.
					TransformXML 				|	Applica un foglio di stile XSL (Extensible Stylesheet Language) a un file di dati XML e scrive il 
												|	codice XML risultante in un file di dati XML.

				Proprieta
					Nome proprieta	            |					@DESCRIZIONE@PROPRIETA
					-----------------------------------------------------------------------------------------------------------------------------------
					AnswerWizard 				|	Restituisce un riferimento all'oggetto AnswerWizard corrente e alle relative proprietà. Oggetto 
												|	AnswerWizard in sola lettura.
					Application 				|	È possibile utilizzare la proprietà Application per accedere all'oggetto Application attivo di 
												|	Microsoft Access e alle relative proprietà. Oggetto Application in sola lettura.
					Assistance 					|	Returns an IAssistance object that represents the Microsoft Office Help Viewer. Elemento in sola 
												|	lettura.
					Assistant 					|	ClippyIsDeadÈ possibile utilizzare la proprietà Assistant per restituire un riferimento all'oggetto 
												|	Assistant. Oggetto Assistant in sola lettura.
					AutoCorrect 				|	Restituisce un oggetto AutoCorrect che rappresenta le impostazioni di correzione automatica per 
												|	1st_Access12. Elemento in sola lettura.
					AutomationSecurity 			|	Restituisce o imposta una costante MsoAutomationSecurity che rappresenta la modalità di protezione 
												|	utilizzata automaticamente in Microsoft Access durante l'apertura di file a livello di 
												|	programmazione. Elemento in lettura e scrittura.
					BrokenReference 			|	Restituisce un valore Boolean che indica se il database corrente include riferimenti interrotti a 
												|	database o librerie dei tipi. True se sono presenti riferimenti interrotti. Elemento in sola lettura.
					Build 						|	Restituisce un valore Long che rappresenta il numero di build della copia installata di 1st_Access12. 
												|	Elemento in sola lettura.
					CodeContextObject 			|	È possibile utilizzare la proprietà CodeContextObject per determinare l'oggetto in cui è in 
												|	esecuzione una macro o codice di Visual Basic. Elemento Object in sola lettura.
					CodeData 					|	È possibile utilizzare la proprietà CodeData per accedere all'oggetto CodeData e agli insiemi a esso 
												|	correlati. Oggetto CodeData in sola lettura.
					CodeProject 				|	È possibile utilizzare la proprietà CodeProject per accedere all'oggetto CodeProject e agli insiemi, 
												|	alle proprietà e ai metodi a esso correlati. Oggetto CodeProject in sola lettura.
					COMAddIns 					|	È possibile utilizzare la proprietà COMAddIns per restituire un riferimento all'oggetto insieme 
												|	COMAddIns corrente e alle relative proprietà. Oggetto COMAddIns in sola lettura.
					CommandBars 				|	È possibile utilizzare la proprietà CommandBars per restituire un riferimento all'oggetto insieme 
												|	CommandBars. Oggetto CommandBars in sola lettura.
					CurrentData 				|	È possibile utilizzare la proprietà CurrentData per accedere all'oggetto CurrentData e agli insiemi a 
												|	esso correlati. Oggetto CurrentData in sola lettura.
					CurrentObjectName 			|	È possibile utilizzare la proprietà CurrentObjectName con l'oggetto Application per determinare il 
												|	nome dell'oggetto di database attivo. L'oggetto di database attivo è l'oggetto che ha lo stato attivo 
												|	o in cui è in esecuzione il codice. Elemento String in sola lettura.
					CurrentObjectType 			|	È possibile utilizzare la proprietà CurrentObjectType con l'oggetto Application per determinare il 
												|	tipo di oggetto di database attivo (tabella, query, maschera, report, macro, modulo, pagina di 
												|	accesso ai dati, visualizzazione del server, diagramma di database o stored procedure). L'oggetto di 
												|	database attivo è l'oggetto che ha lo stato attivo o in cui è in esecuzione il codice. Elemento 
												|	AcObjectType in sola lettura.
					CurrentProject 				|	È possibile utilizzare la proprietà CurrentProject per accedere all'oggetto CurrentProject e agli 
												|	insiemi, alle proprietà e ai metodi a esso correlati. Oggetto CurrentProject in sola lettura.
					DBEngine 					|	È possibile utilizzare la proprietà DBEngine in Visual Basic per accedere all'oggetto DBEngine 
												|	corrente e alle relative proprietà. Elemento DBEngine in sola lettura.
					DoCmd 						|	È possibile utilizzare la proprietà DoCmd per accedere all'oggetto DoCmd in sola lettura e ai 
												|	relativi metodi. Elemento DoCmd in sola lettura.
					FeatureInstall 				|	È possibile utilizzare la proprietà FeatureInstall per specificare o determinare il modo in cui in 
												|	1st_Access12 vengono gestiti i metodi e le proprietà che richiedono caratteristiche non ancora 
												|	installate. Elemento MsoFeatureInstall in lettura e scrittura.
					FileDialog 					|	Restituisce un oggetto FileDialog che rappresenta una singola istanza di una finestra di dialogo per 
												|	la gestione di file. Elemento in sola lettura.
					Forms 						|	È possibile utilizzare la proprietà Forms per restituire un riferimento in sola lettura all'insieme 
												|	Forms e alle relative proprietà.
					IsCompiled 					|	La proprietà IsCompiled restituisce un valore Boolean che indica se il progetto di Visual Basic è 
												|	stato compilato. Elemento Boolean in sola lettura.
					LanguageSettings 			|	È possibile utilizzare la proprietà LanguageSettings per restituire un riferimento in sola lettura 
												|	all'oggetto LanguageSettings corrente e alle relative proprietà.
					MacroError 					|	Returns a MacroError object that contains information about the latest error to occur in a macro. 
												|	Elemento in sola lettura.
					MenuBar 					|	Specifica un menu personalizzato da visualizzare in un database di Microsoft Access. Stringa di 
												|	lettura/scrittura.
					Modules 					|	È possibile utilizzare la proprietà Modules per accedere all'insieme Modules e alle relative 
												|	proprietà. Oggetto Modules in sola lettura.
					Name 						|	Elemento in sola lettura
					NewFileTaskPane 			|	Restituisce un oggetto NewFile che rappresenta un documento elencato nel riquadro attività Nuovo 
												|	file. Oggetto NewFile in sola lettura.
					Parent 						|	Elemento in sola lettura
					Printer 					|	È possibile utilizzare questa proprietà per restituire o impostare un oggetto Printer che rappresenta 
												|	la stampante predefinita del sistema. Elemento di lettura e scrittura.
					Printers 					|	Restituisce l'insieme Printers che rappresenta tutte le stampanti disponibili nel sistema. Insieme 
												|	Printers in lettura e scrittura.
					ProductCode 				|	È possibile utilizzare la proprietà ProductCode per determinare l'identificatore univoco globale (
												|	GUID) di 1st_Access12. Elemento String in sola lettura.
					References 					|	È possibile utilizzare la proprietà References per accedere all'insieme References e ai relativi 
												|	eventi, metodi e proprietà. Insieme References in sola lettura.
					Reports 					|	È possibile utilizzare la proprietà Reports per accedere all'insieme in sola lettura Reports e alle 
												|	relative proprietà.
					Screen 						|	È possibile utilizzare la proprietà Screen per restituire un riferimento all'oggetto Screen e alle 
												|	relative proprietà. Elemento in sola lettura.
					ShortcutMenuBar 			|	È possibile utilizzare la proprietà ShortcutMenuBar per specificare il menu di scelta rapida che 
												|	verrà visualizzato facendo clic con il pulsante destro del mouse sull'oggetto specificato. Elemento 
												|	String in lettura e scrittura.
					TempVars 					|	Returns the collection of TempVar objects. Read-only TempVars.
					UserControl 				|	È possibile utilizzare la proprietà UserControl per determinare se l'applicazione 1st_Access12 
												|	corrente è stata avviata dall'utente o da un'altra applicazione tramite automazione, precedentemente 
												|	denominata automazione OLE. Elemento Boolean in lettura e scrittura.
					VBE 						|	È possibile utilizzare la proprietà VBE per restituire un riferimento all'oggetto VBE corrente e alle 
												|	relative proprietà. La proprietà VBE dell'oggetto Application rappresenta l'editor di Microsoft 
												|	Visual Basic, Applications Edition. Oggetto VBE in sola lettura.
					Version 					|	Restituisce un valore String che indica il numero di versione della copia installata di 1st_Access12. 
												|	Elemento in sola lettura.
					Visible 					|	Elemento in lettura e scrittura

















































































































































 							






































































********************

				Prorpieta


	Oggetto CommandBarControls

			CommandBars 
				
				
			  	Nota 
					Insieme di oggetti @CommandBar che rappresentano le barre dei comandi dell'applicazione contenitore
						@oggetto@Command@Bar_(Oggetto CommandBars che rappresentano le barre dei comandi)_@Barre@dei@comandi


					L'utilizzo delle barre dei comandi in alcune applicazioni di Microsoft Office è stato sostituito dalla nuova interfaccia utente basata sulla barra multifunzione. Per ulteriori informazioni, cercare "barra multifunzione" nella Guida.


				Esempio
					Utilizzare la proprietà CommandBars per restituire l'insieme CommandBars. Nell'esempio seguente nella finestra Immediata vengono visualizzati il nome e il nome locale di ogni 
								barra dei menu e 
								barra degli strumenti, 
					nonché un valore che indica se la barra dei menu o la barra degli strumenti è visibile.

						ES.VISUALIZZA BARRA DEI MENU E STRUMENTI
							For Each cbar in CommandBars
							    Debug.Print cbar.Name, cbar.NameLocal, cbar.Visible
							Next 


						ES. AD AGGIUNGI UNA NUOVA BARRA DEI COMANDI ALL'INSIEME - CUSTOM1 -
							Utilizzare il metodo Add per aggiungere una nuova barra dei comandi all'insieme. Nell'esempio seguente viene creata una barra degli strumenti personalizzata denominata "Custom1" che viene visualizzata come una barra degli strumenti mobile.

							Set cbar1 = CommandBars.Add(Name:="Custom1", Position:=msoBarFloating)
							cbar1.Visible = True


						ES. NUMERO INDICE BARRA DEI COMANDI
							Utilizzare enumName, dove index è il nome o il numero di indice di una barra dei comandi, per restituire un singolo oggetto CommandBar. Nell'esempio seguente la barra degli strumenti denominata "Custom1" viene ancorata nella parte inferiore della finestra dell'applicazione.

								Visual Basic, Application Edition 
								CommandBars("Custom1").Position = msoBarBottom 

				  Nota 
						Per specificare una 
							barra dei menu o 
							una barra degli strumenti 
						esistente nell'elenco delle barre disponibili all'interno dell'applicazione contenitore, è possibile utilizzare il nome o il numero di indice. 
						Per specificare un menu, 
							un menu di scelta rapida o 
							un sottomenu 
						(tutti rappresentati da oggetti CommandBar), è tuttavia necessario utilizzare il nome. 

						Se due o più menu o sottomenu personalizzati hanno lo stesso nome, enumName restituisce il primo nome. Per assicurarsi che venga restituito il menu o il sottomenu corretto, individuare il controllo popup che consente di visualizzare tale menu. 

						Applicare quindi la proprietà CommandBar al controllo popup per restituire la barra dei comandi che rappresenta tale menu. 





					faq
						@barra@dei@comandi
						@barre@oggetti@applicazione@contenitore
			
			Membri dell'oggetto CommandBars 
						Insieme di oggetti CommandBar che rappresentano le barre dei comandi dell'applicazione contenitore.


							'//per applicazione con New - istanza di un oggetto application
							Dim appAccess As New Access.Application 
							
							//utilizzare nelle applicazioni senza new - istanza di un oggetto application
							Dim appAccess As Object
							Set appAccess = CreateObject("Access.Application")


				Metodi
					FindControls
					....

				Proprietà
					ActionControl



TROVARE_TUTTI_I_PROGETTI
      
        //@DOVE@SI@TROVANO.I@PROGETTI   
        //@TUTTI.I.PROGETTI_(tutti i progetti Dos, React, SublimeTex, Java, C++, Powershell ecc.)
        //-----------------------------------------------------------------------------------------------------------//
        
            Se vuoi trovare tutti progetti devi aprire la path in cui c'e il riepilo di tutti i progetti
            quali il DOS, PAWERSHELL, REACT ecc..

            Per aprire questa cartella fare riferimento a queste chiavi di ricerca:
        
            Questo è il progetto di TUTTI@PROGETTI 
            TROVA_PROGETTI.MD
            @TUTTI@I@PROGETTI_(questo è il file contiene i riferimenti a tutti progetti)
            
            Apro il progetto con TUTTI i progetti con il DOS
            Invoke-Item "c:\\CASA\\LINGUAGGI\\TROVA_PROGETTI\\TROVA_PROGETTI.sublime-project"
            
            Apro tutti i progetti con powershell
            START "apri il progetto dei comandi di TUTTI@PROGETTI" "c:\\CASA\\LINGUAGGI\\TROVA_PROGETTI\\TROVA_PROGETTI.sublime-project"
            
        //-----------------------------------------------------------------------------------------------------------//


TROVARE_TUTTI_I_PROGETTI
      
        //@DOVE@SI@TROVANO.I@PROGETTI   
        //@TUTTI.I.PROGETTI_(tutti i progetti Dos, React, SublimeTex, Java, C++, Powershell ecc.)
        //-----------------------------------------------------------------------------------------------------------//
        
            Se vuoi trovare tutti progetti devi aprire la path in cui c'e il riepilo di tutti i progetti
            quali il DOS, PAWERSHELL, REACT ecc..

            Per aprire questa cartella fare riferimento a queste chiavi di ricerca:
        
            Questo è il progetto di TUTTI@PROGETTI 
            TROVA_PROGETTI.MD
            @TUTTI@I@PROGETTI_(questo è il file contiene i riferimenti a tutti progetti)
            
            Apro il progetto con TUTTI i progetti con il DOS
            Invoke-Item "c:\\CASA\\LINGUAGGI\\TROVA_PROGETTI\\TROVA_PROGETTI.sublime-project"
            
            Apro tutti i progetti con powershell
            START "apri il progetto dei comandi di TUTTI@PROGETTI" "c:\\CASA\\LINGUAGGI\\TROVA_PROGETTI\\TROVA_PROGETTI.sublime-project"
            
        //-----------------------------------------------------------------------------------------------------------//

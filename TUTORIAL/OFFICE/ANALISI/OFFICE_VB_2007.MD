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

	Oggetto CommandBarControls

			CommandBars 
				
				
			  	Nota 
					Insieme di oggetti CommandBar che rappresentano le barre dei comandi dell'applicazione contenitore


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


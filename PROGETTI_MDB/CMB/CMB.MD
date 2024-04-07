CMB.MD
	
	Note
		@casella@combinata_(proprieta e metodi della casella combinata)

		La casella combinata ha diverse proprieta e diversi metodi raggruppati con l'iniziale
		della lettera.



	PROPRIETA
		
		H
            Hyperlink, proprietà
		I
			ItemData, proprietà
				@item_(proprieta della casella combinata)

				La proprietà ItemData restituisce i dati nella colonna associata (colonna associata: Colonna in una casella di riepilogo, combinata o di riepilogo a discesa associata al campo specificato dalla proprietà Origine controllo del controllo.) per la riga specificata di una casella combinata (casella combinata: Controllo utilizzato in una maschera per fornire le funzionalità combinate di una casella di riepilogo e di una casella di testo. È possibile digitare un valore o fare clic sul controllo per visualizzare un elenco e selezionarne una delle voci.). Elemento Variant di sola lettura. 
				
				Sintassi
				espressione.ItemData(Index)

				espressione   Variabile che rappresenta un oggetto ComboBox.


				Parametri

					Nome 	Obbligatorio/Facoltativo 	Tipo di dati 	Descrizione 
					Index 	Obbligatorio 				Long 			Riga della casella combinata o della casella di
					 													riepilogo che contiene i dati che si desidera vengano 
					 													restituiti. Le righe in tali caselle sono indicizzate a 
					 													partire da zero. Per la restituzione di un elemento 
					 													della sesta riga di una casella combinata ad esempio 
					 													specificare 5 nell'argomento Index. 


				Note
					La proprietà ItemData consente di scorrere l'elenco delle voci di una casella combinata o di una casella di riepilogo. Per scorrere ad esempio tutte le voci di una casella di riepilogo per cercare una voce particolare, utilizzare la proprietà ListCount per stabilire il numero di righe della casella di riepilogo, quindi utilizzare la proprietà ItemData per restituire i dati per la colonna associata a ogni riga.

					È inoltre possibile utilizzare la proprietà ItemData per restituire i dati solo dalle righe selezionate di una casella di riepilogo. È possibile scorrere l'insieme ItemsSelected per verificare quali righe della casella sono state selezionate e utilizzare la proprietà ItemData per restituire i dati provenienti da tali righe. Impostare la proprietà MultiSelect della casella di riepilogo su Semplice o Estesa, a seconda che si desideri consentire all'utente di selezionare rispettivamente una o più righe alla volta.

					È possibile utilizzare la proprietà Column per restituire i dati provenienti da una riga e una colonna specifiche, anche se la colonna specificata non è una colonna associata.

				Faq

					RECUPERO INDICE DELLA CASELLA COMBINATA
						Nota
							come @recuperare l'@indice della casella combinata

						Esempio
							'proprieta ListIndex  @indice@della@casella@combinata_(come recupare l'indice della selezione sulla casella combinata)
			                'Indica su quale elemento è stata effettuata la selezione : es.
			                'Elementi da 0 a 4  listIndex= 3 il quarto elemento.
			                vV1 = Me.Cmb_01_TXT.ListIndex

            IMEHold, proprietà
            IMEMode, proprietà
            IMESentenceMode, proprietà
            InheritValueList, proprietà
            InputMask, proprietà
            InSelection, proprietà
            IsHyperlink, proprietà
            IsVisible, proprietà





Metodi
		Nome			Descrizione
	A
		AddItem 		È possibile utilizzare questo metodo per aggiungere un nuovo elemento all'elenco dei valori visualizzato dal controllo casella 
						combinata specificato.
	D
		Dropdown 		È possibile utilizzare il metodo Dropdown per forzare la discesa dell'elenco nella casella combinata specificata.
	M
		Move 			Sposta l'oggetto indicato in corrispondenza delle coordinate specificate dai valori degli argomenti.
	R
		RemoveItem 		È possibile utilizzare questo metodo per rimuovere un elemento dall'elenco dei valori visualizzato dal controllo casella 
						combinata specificato.

		Requery 		Il metodo Requery aggiorna i dati sottostanti un determinato controllo contenuto nella maschera attiva rieseguendo una query 
						sull'origine dei dati per il controllo.
	S
		SetFocus 		Con il metodo SetFocus, la maschera specificata, il controllo specificato sulla maschera attiva o il campo specificato sul 
						foglio dati attivo acquisiscono lo stato attivo.
		SizeToFit 		È possibile utilizzare il metodo SizeToFit per ridimensionare un controllo in modo che si adatti al testo o all'immagine che 
						contiene.
	U
		Undo 			È possibile utilizzare il metodo Undo per reimpostare un controllo o una maschera in seguito alla modifica del relativo valore.


	           










	METODI


Proprietà
Lettera		Nome				Descrizione
	A
			AddColon 			Specifica se il testo contenuto nelle etichette dei nuovi controlli deve essere seguito dai due punti. Elemento 
								Boolean in lettura 	e scrittura.
			AfterUpdate 		Restituisce o imposta la macro, la routine evento o la funzione definita dall'utente che viene eseguita quando si 
								verifica  l'evento AfterUpdate. Elemento String in lettura e scrittura.
	B
	C     
			Column 					È possibile utilizzare la proprietà Column per fare riferimento a una determinata colonna o a una combinazione di 
									colonna e riga contenute in una casella combinata o casella di riepilogo a più colonne. Elemento Variant in sola 
									lettura.


									È possibile utilizzare la proprietà Column per fare riferimento a una determinata colonna (colonna: Posizione in una tabella di database che memorizza un particolare tipo di dati. È anche la rappresentazione visiva di un campo in un foglio dati e, in un database di Access, la griglia di struttura delle query o dei filtri.) o a una combinazione di colonna e riga contenute in una casella combinata (casella combinata: Controllo utilizzato in una maschera per fornire le funzionalità combinate di una casella di riepilogo e di una casella di testo. È possibile digitare un valore o fare clic sul controllo per visualizzare un elenco e selezionarne una delle voci.) o casella di riepilogo (casella di riepilogo: Controllo che fornisce un elenco di scelte. Una casella di riepilogo è costituita da un elenco e da un'etichetta facoltativa.) a più colonne. Elemento Variant in sola lettura. 

								Sintassi

									espressione.Column(Index, Row)

									espressione   Variabile che rappresenta un oggetto ComboBox.


									Parametri

									Nome Obbligatorio/Facoltativo 	Tipo di dati 	Descrizione 
									Index Obbligatorio 			   	Long 			Numero intero Long compreso tra 0 e l'impostazione della proprietà 
																					ColumnCount meno 1. 
									Row   Facoltativo 				Variant 		Numero intero compreso tra 0 e l'impostazione della proprietà 
																					ListCount meno 1. 


								Note


									Utilizzare 0 per fare riferimento alla prima colonna, 1 per fare riferimento alla seconda e così via. Utilizzare 0 per fare riferimento alla prima riga, 1 per fare riferimento alla seconda e così via. In una casella di riepilogo che include, ad esempio, una colonna con gli ID dei clienti e una colonna con i nomi dei clienti, è possibile fare riferimento al nome del cliente presente nella seconda colonna e nella quinta riga come indicato di seguito:

										Forms!Contacts!Customers.Column(1, 4) 


									È possibile utilizzare la proprietà Column per assegnare il contenuto di una casella combinata o di una casella di riepilogo a un altro controllo, quale una casella di testo (casella di testo: Controllo, detto anche campo di modifica, utilizzato in una maschera, un report o una pagina di accesso ai dati per visualizzare testo o accettare l'immissione di dati. Può avere un'etichetta associata.). Per impostare ad esempio la proprietà ControlSource di una casella di testo sul valore presente nella seconda colonna di una casella di riepilogo, è possibile utilizzare l'espressione seguente:

										=Forms!Customers!CompanyName.Column(1) 


									Se non si effettua alcuna selezione quando si fa riferimento a una colonna in una casella combinata o in una casella di riepilogo, l'impostazione della proprietà Column sarà Null. È possibile utilizzare la funzione IsNull per determinare se è stata effettuata una selezione, come nell'esempio seguente:

										If IsNull(Forms!Customers!Country)
										    Then MsgBox "No selection."
										End If 

							  Nota 
									Per determinare il numero di colonne di una casella combinata o di una casella di riepilogo, è possibile controllare l'impostazione della proprietà ColumnCount. 

							Esempio


									Negli esempi riportati di seguito vengono utilizzate le proprietà Column e ColumnCount per stampare i valori selezionati in una casella di riepilogo.


							Visual Basic, Application Edition 

									Public Sub Read_ListBox()

									    Dim intNumColumns As Integer
									    Dim intI As Integer
									    Dim frmCust As Form

									    Set frmCust = Forms!frmCustomers
									    If frmCust!lstCustomerNames.ItemsSelected.Count > 0 Then

									        ' Any selection?
									        intNumColumns = frmCust!lstCustomerNames.ColumnCount
									        Debug.Print "The list box contains "; intNumColumns; _
									            IIf(intNumColumns = 1, " column", " columns"); _
									             " of data."

									        Debug.Print "The current selection contains:"
									        For intI = 0 To intNumColumns - 1
									            ' Print column data.
									            Debug.Print frmCust!lstCustomerNames.Column(intI)
									        Next intI
									    Else
									        Debug.Print "You haven't selected an entry in the " _
									            & "list box."
									    End If

									    Set frmCust = Nothing

									End Sub 


							FAQ
								come individuare la colonna della combo box

								ESEMPIO di recupero valore della combo box
									   'proprieta ListIndex  @indice@della@casella@combinata_(come recupepare l'indice della selezione sulla casella combinata)
					                      'Indica su quale elemento è stata effettuata la selezione : es.
					                      'Elementi da 0 a 4  listIndex= 3 il quarto elemento.
					                      vV1 = Me.CMB_03_TXT.ListIndex					'//questo recupera il record della combo box ossi l'indice in cui si trova il record selezionato

					    				  '//nella combo box 3 ci sono 2 colonne con indice 0,1'	
					                      Me.Cmb_04_TXT = Me.CMB_03_TXT.Column(1)       '//imposto la combo con il valore della combo 3 colonna 2
					    












		
	I
			ItemData 			La proprietà ItemData restituisce i dati nella colonna associata per la riga specificata di una casella combinata. 
								Elemento Variant di sola lettura.

							@item_(proprieta della casella combinata)

								La proprietà ItemData restituisce i dati nella colonna associata (colonna associata: Colonna in una casella di riepilogo, combinata o di riepilogo a discesa associata al campo specificato dalla proprietà Origine controllo del controllo.) per la riga specificata di una casella combinata (casella combinata: Controllo utilizzato in una maschera per fornire le funzionalità combinate di una casella di riepilogo e di una casella di testo. È possibile digitare un valore o fare clic sul controllo per visualizzare un elenco e selezionarne una delle voci.). Elemento Variant di sola lettura. 
								
								Sintassi
								espressione.ItemData(Index)

								espressione   Variabile che rappresenta un oggetto ComboBox.


								Parametri

									Nome 	Obbligatorio/Facoltativo 	Tipo di dati 	Descrizione 
									Index 	Obbligatorio 				Long 			Riga della casella combinata o della casella di
									 													riepilogo che contiene i dati che si desidera vengano 
									 													restituiti. Le righe in tali caselle sono indicizzate a 
									 													partire da zero. Per la restituzione di un elemento 
									 													della sesta riga di una casella combinata ad esempio 
									 													specificare 5 nell'argomento Index. 


								Note
									La proprietà ItemData consente di scorrere l'elenco delle voci di una casella combinata o di una casella di riepilogo. Per scorrere ad esempio tutte le voci di una casella di riepilogo per cercare una voce particolare, utilizzare la proprietà ListCount per stabilire il numero di righe della casella di riepilogo, quindi utilizzare la proprietà ItemData per restituire i dati per la colonna associata a ogni riga.

									È inoltre possibile utilizzare la proprietà ItemData per restituire i dati solo dalle righe selezionate di una casella di riepilogo. È possibile scorrere l'insieme ItemsSelected per verificare quali righe della casella sono state selezionate e utilizzare la proprietà ItemData per restituire i dati provenienti da tali righe. Impostare la proprietà MultiSelect della casella di riepilogo su Semplice o Estesa, a seconda che si desideri consentire all'utente di selezionare rispettivamente una o più righe alla volta.

									È possibile utilizzare la proprietà Column per restituire i dati provenienti da una riga e una colonna specifiche, anche se la colonna specificata non è una colonna associata.

								Faq

									RECUPERO INDICE DELLA CASELLA COMBINATA
										Nota
											come @recuperare l'@indice della casella combinata

										Esempio
											'proprieta ListIndex  @indice@della@casella@combinata_(come recupare l'indice della selezione sulla casella combinata)
							                'Indica su quale elemento è stata effettuata la selezione : es.
							                'Elementi da 0 a 4  listIndex= 3 il quarto elemento.
							                vV1 = Me.Cmb_01_TXT.ListIndex

	DA_FINIRE

		AllowAutoCorrect 		È possibile utilizzare la proprietà AllowAutoCorrect per indicare se i dati immessi dall'utente verranno corretti 
								automaticamente dal controllo specificato. Elemento Boolean di lettura e scrittura.

		AllowValueListEdits 	Recupera o imposta se il comando Modifica voci di elenco è disponibile quando si fa clic con il pulsante destro del 
								mouse su una casella combinata. Valore Boolean di lettura/scrittura.

		Application 			È possibile utilizzare la proprietà Application per accedere all'oggetto Application attivo di Microsoft Access e alle 
								relative proprietà. Oggetto Application in sola lettura.

		AutoExpand 				È possibile utilizzare la proprietà AutoExpand per specificare se in 1st_Access12 la parte casella di testo di una 
								casella combinata viene completata automaticamente con un valore del relativo elenco, corrispondente ai caratteri 
								digitati nella casella combinata. In questo modo, è possibile immettere rapidamente un valore esistente in una casella 
								combinata senza visualizzare la parte casella di riepilogo. Elemento Boolean in lettura e scrittura.

		AutoLabel 				Specifica se le etichette devono essere automaticamente create e associate ai nuovi controlli. Elemento Boolean in 
								lettura e scrittura.

		BackColor 				Ottiene o imposta il colore interno dell'oggetto specificato. Elemento Long in lettura e scrittura.

		BackStyle 				È possibile utilizzare la proprietà BackStyle per specificare se un controllo dovrà essere trasparente. Elemento Byte 
								di lettura e scrittura.
		BeforeUpdate 	Restituisce o imposta la macro, la routine evento o la funzione definita dall'utente che viene eseguita quando si verifica l'evento BeforeUpdate. Elemento String in lettura e scrittura.
		BorderColor 	È possibile utilizzare la proprietà BorderColor per specificare il colore del bordo di un controllo. Elemento Long in lettura e scrittura.
		BorderStyle 	Specifica la modalità di visualizzazione del bordo di un controllo. Elemento Byte in lettura e scrittura.
		BorderWidth 	È possibile utilizzare la proprietà BorderWidth per specificare lo spessore del bordo di un controllo. Elemento Byte in lettura e scrittura.
		BottomMargin 	Insieme alle proprietà LeftMargin, RightMargin e TopMargin, specifica la posizione delle informazioni visualizzate in un controllo casella di testo Valore Integer di lettura/scrittura.
		BottomPadding 	Recupera o imposta la quantità di spazio, espressa in pollici, tra la casella combinata e la corrispondente griglia inferiore. Valore Integer di lettura/scrittura.
		BoundColumn 	Quando si effettua una selezione da una casella combinata, la proprietà BoundColumn indica a Microsoft Access quali tra i valori della colonna utilizzare come valore del controllo. Se il controllo è associato a un campo, nel campo indicato nella proprietà ControlSource verrà memorizzato il valore contenuto nella colonna specificata dalla proprietà BoundColumn. Elemento Long di lettura e scrittura.
		CanGrow 	Recupera o imposta se il controllo specificato viene automaticamente adattato in senso verticale per stampare o visualizzare in anteprima di stampa tutti i dati inclusi nel controllo. Valore Boolean di lettura/scrittura.
		CanShrink 	Recupera o imposta se il controllo specificato viene automaticamente adattato in senso verticale in modo da consentire la stampa o la visualizzazione in anteprima di stampa di tutti i dati inclusi nel controllo. Valore Boolean di lettura/scrittura.
		
		

		ColumnCount 	È possibile utilizzare la proprietà ColumnCount per specificare il numero di colonne visualizzate in una casella di riepilogo o nella parte casella di riepilogo di una casella combinata oppure inviate a oggetti OLE in un controllo grafico o cornice oggetto non associato. Elemento Integer di lettura e scrittura.
		ColumnHeads 	È possibile utilizzare la proprietà ColumnHeads per visualizzare una singola riga di intestazioni di colonna per caselle di riepilogo, caselle combinate e oggetti OLE che supportano intestazioni di colonna. È inoltre possibile utilizzare questa proprietà per creare un'etichetta per ogni elemento di un controllo grafico. L'impostazione della proprietà RowSourceType dell'oggetto determina l'intestazione di colonna visualizzata per la prima riga. Valore di tipo Boolean di lettura/scrittura.
		ColumnHidden 	È possibile utilizzare la proprietà ColumnHidden per mostrare o nascondere una determinata colonna in visualizzazione Foglio dati. Elemento Boolean in lettura e scrittura.
		ColumnOrder 	È possibile utilizzare la proprietà ColumnOrder per specificare l'ordine delle colonne in visualizzazione Foglio dati. Elemento Integer in lettura e scrittura.
		ColumnWidth 	È possibile utilizzare la proprietà ColumnWidth per specificare la larghezza di una colonna in visualizzazione Foglio dati. Elemento Integer in lettura e scrittura.
		ColumnWidths 	È possibile utilizzare la proprietà ColumnWidths per specificare la larghezza di ogni colonna in una casella combinata a più colonne. Elemento String di lettura e scrittura.
		Controls 	Elemento in sola lettura
		ControlSource 	È possibile utilizzare la proprietà ControlSource per specificare quali dati vengono visualizzati in un controllo. I dati che possono essere visualizzati e modificati sono quelli associati a un campo di una tabella, di una query o di un'istruzione SQL. È inoltre possibile visualizzare il risultato di un'espressione. Elemento String in lettura e scrittura.
		ControlTipText 	È possibile utilizzare la proprietà ControlTipText per specificare il testo da visualizzare in una descrizione comando quando si sposta il puntatore del mouse su un controllo. Elemento String in lettura e scrittura.
		ControlType 	È possibile utilizzare la proprietà ControlType in Visual Basic per determinare il tipo di controllo di una maschera o report. Valore Byte di lettura/scrittura.
		DecimalPlaces 	È possibile utilizzare la proprietà DecimalPlaces per specificare il numero di posizioni decimali utilizzate da Microsoft Access per la visualizzazione di numeri. Elemento Byte di lettura e scrittura.
		DefaultValue 	Specifica un valore che viene immesso automaticamente in un campo alla creazione di un nuovo record. In una tabella Indirizzi è ad esempio possibile impostare il valore predefinito del campo Città su Milano. Quando gli utenti aggiungono un record alla tabella, possono accettare questo valore o immettere il nome di una città diversa. Elemento String in lettura e scrittura.
		DisplayAsHyperlink 	Gets or sets an AcDisplayAsHyperlink constant that specifies whether to display the contents of the specified combo box as a hyperlink. Valore di lettura/scrittura.
		DisplayWhen 	È possibile utilizzare la proprietà DisplayWhen per specificare quali controlli di una maschera si desidera visualizzare o stampare. Elemento Byte di lettura e scrittura.
		Enabled 	La proprietà Enabled imposta o restituisce lo stato del formato condizionale nell'oggetto FormatCondition. Elemento Boolean in lettura e scrittura.
		EventProcPrefix 	Elemento in lettura e scrittura
		FontBold 	È possibile utilizzare la proprietà FontBold per specificare se un carattere deve essere visualizzato in grassetto nelle situazioni seguenti: Quando vengono visualizzati o stampati controlli di maschere e report. Quando viene utilizzato il metodo Print in un report. Elemento Boolean in lettura e scrittura.
		FontItalic 	È possibile utilizzare la proprietà FontItalic per specificare se il testo deve essere formattato in corsivo nelle situazioni seguenti: Quando vengono visualizzati o stampati controlli di maschere e report. Quando viene utilizzato il metodo Print in un report. Elemento Boolean in lettura e scrittura.
		FontName 	È possibile utilizzare la proprietà FontName per specificare il tipo di carattere del testo nelle situazioni seguenti: Quando vengono visualizzati o stampati controlli di maschere e report. Quando viene utilizzato il metodo Print in un report. Elemento String in lettura e scrittura.
		FontSize 	È possibile utilizzare la proprietà FontSize per specificare la dimensione del testo in punti nelle situazioni seguenti: Quando vengono visualizzati o stampati controlli di maschere e report. Quando viene utilizzato il metodo Print in un report. Elemento Integer in lettura e scrittura.
		FontUnderline 	È possibile utilizzare la proprietà FontUnderline per specificare se il testo deve essere sottolineato nelle situazioni seguenti: Quando vengono visualizzati o stampati controlli di maschere e report. Quando viene utilizzato il metodo Print in un report. Elemento Boolean in lettura e scrittura.
		FontWeight 	È possibile utilizzare la proprietà DatasheetFontWeight per specificare lo spessore della riga del carattere utilizzato per visualizzare e stampare i caratteri dei nomi di campo e dei dati in visualizzazione Foglio dati. Elemento Integer in lettura e scrittura.
		ForeColor 	È possibile utilizzare la proprietà ForeColor per specificare il colore per il testo in un controllo. Elemento Long in lettura e scrittura.
		Format 	È possibile utilizzare la proprietà Format per personalizzare il modo in cui vengono visualizzati e stampati numeri, date, ore e testo. Elemento String di lettura e scrittura.
		FormatConditions 	È possibile utilizzare la proprietà FormatConditions per restituire un riferimento di sola lettura all'insieme FormatConditions e alle proprietà a esso correlate.
		GridlineColor 	Recupera o imposta il colore della griglia della casella combinata specificata. Valore Long di lettura/scrittura.
		GridlineStyleBottom 	Recupera o imposta lo stile della griglia inferiore della casella combinata specificata. Valore Byte di lettura/scrittura.
		GridlineStyleLeft 	Recupera o imposta lo spessore della griglia inferiore della casella combinata specificata. Valore Byte di lettura/scrittura.
		GridlineStyleRight 	Recupera o imposta lo stile della griglia destra della casella combinata specificata. Valore Byte di lettura/scrittura.
		GridlineStyleTop 	Recupera o imposta lo stile della griglia superiore della casella combinata specificata. Valore Byte di lettura/scrittura.
		GridlineWidthBottom 	Recupera o imposta lo spessore della griglia inferiore della casella combinata specificata. Valore Byte di lettura/scrittura.
		GridlineWidthLeft 	Recupera o imposta lo spessore della griglia sinistra della casella combinata specificata. Valore Byte di lettura/scrittura.
		GridlineWidthRight 	Recupera o imposta lo spessore della griglia destra della casella combinata specificata. Valore Byte di lettura/scrittura.
		GridlineWidthTop 	Recupera o imposta lo spessore della griglia superiore della casella combinata specificata. Valore Byte di lettura/scrittura.
		Height 	Elemento in lettura e scrittura
		HelpContextId 	Elemento in lettura e scrittura
		HideDuplicates 	È possibile utilizzare la proprietà HideDuplicates per nascondere un controllo di un report quando il valore è uguale a quello contenuto nel record precedente. Elemento Boolean in lettura e scrittura.
		HorizontalAnchor 	Recupera o imposta una costante AcHorizontalAnchor che indica come la casella combinata verrà ancorata in senso orizzontale all'interno del relativo layout. Valore di lettura/scrittura.
		Hyperlink 	È possibile utilizzare la proprietà Hyperlink per restituire un riferimento a un oggetto Hyperlink. È inoltre possibile utilizzare la proprietà Hyperlink per accedere alle proprietà e ai metodi di un collegamento ipertestuale di un controllo. Valore di sola lettura.
		IMEHold 	Informazioni specifiche della lingua È possibile utilizzare la proprietà IMEHold/Hold KanjiConversionMode per indicare se la modalità di conversione Kanji viene mantenuta quando il controllo perde lo stato attivo. Elemento Boolean di lettura e scrittura.
		IMEMode 	Informazioni specifiche della linguaÈ possibile utilizzare la proprietà IMEMode per impostare la modalità di conversione Kanji di un controllo quando su di esso viene spostato lo stato attivo. Elemento di lettura e scrittura.
		IMESentenceMode 	Informazioni specifiche della lingua È possibile utilizzare la proprietà IMESentenceMode per specificare o determinare la modalità frase IME dei campi di una tabella o dei controlli di una maschera da utilizzare quando si sposta lo stato attivo. Elemento di lettura e scrittura.
		InheritValueList 	Recupera o imposta se l'elenco di valori di una casella combinata viene ereditato dal relativo campo. Valore Boolean di lettura/scrittura.
		InputMask 	È possibile utilizzare la proprietà InputMask per semplificare l'immissione di dati e definire i valori che gli utenti possono immettere in un controllo casella combinata. Elemento String di lettura e scrittura.
		InSelection 	Elemento in lettura e scrittura
		IsHyperlink 	È possibile utilizzare la proprietà IsHyperlink per specificare o determinare se i dati contenuti in una casella combinata siano un collegamento ipertestuale. Elemento Boolean di lettura e scrittura.
		IsVisible 	Elemento in lettura e scrittura





		ItemsSelected 	Elemento di sola lettura.
		KeyboardLanguage 	Informazioni specifiche della lingua È possibile utilizzare la proprietà KeyboardLanguage per specificare o determinare la lingua della tastiera quando si accede a un controllo. Elemento Byte di lettura e scrittura.
		LabelAlign 	La proprietà specifica l'allineamento del testo contenuto nelle etichette associate ai nuovi controlli. Elemento Byte in lettura e scrittura.
		LabelX 	Informazioni specifiche della linguaLa proprietà LabelX, insieme alla proprietà LabelY, specifica la posizione dell'etichetta per un nuovo controllo. Elemento Integer in lettura e scrittura.
		LabelY 	Informazioni specifiche della linguaLa proprietà LabelY, insieme alla proprietà LabelX, specifica la posizione dell'etichetta per un nuovo controllo. Elemento Integer in lettura e scrittura.
		Layout 	Restituisce il tipo di layout della casella combinata specificata. Costante AcLayoutType di sola lettura.
		LayoutID 	Restituisce l'identificatore univoco del layout contenente la casella combinata specificata. Valore Long di sola lettura.
		Left 	Elemento in lettura e scrittura
		LeftMargin 	Insieme alle proprietà TopMargin, RightMargin e BottomMargin, specifica la posizione delle informazioni visualizzate in un controllo casella di testo Valore Integer di lettura/scrittura.
		LeftPadding 	Recupera o imposta la quantità di spazio, espressa in pollici, tra la casella combinata e la corrispondente griglia sinistra. Valore Integer di lettura/scrittura.
		LimitToList 	È possibile utilizzare la proprietà LimitToList per limitare ai soli elementi elencati i valori inclusi in una casella combinata. Elemento Boolean in lettura e scrittura.
		ListCount 	È possibile utilizzare la proprietà ListCount per determinare il numero di righe della parte casella di riepilogo di una casella combinata. Elemento Long di lettura e scrittura.
		ListIndex 	È possibile utilizzare la proprietà ListIndex per determinare quale sia l'elemento selezionato in una casella combinata. Elemento Long di lettura e scrittura.
		ListItemsEditForm 	Recupera o imposta il nome della maschera visualizzata quando si fa clic su Modifica voci di elenco. Stringa di lettura/scrittura.
		ListRows 	È possibile utilizzare la proprietà ListRows per impostare il numero massimo di righe da visualizzare nella parte casella di riepilogo di una casella combinata. Elemento Integer in lettura e scrittura.
		ListWidth 	È possibile utilizzare la proprietà ListWidth per impostare la larghezza della parte casella di riepilogo di una casella combinata. Elemento String in lettura e scrittura.
		Locked 	La proprietà Locked specifica se è possibile modificare i dati di un controllo in visualizzazione Maschera. Elemento Boolean in lettura e scrittura.
		Name 	Elemento in lettura e scrittura
		NumeralShapes 	Informazioni specifiche della linguaÈ possibile utilizzare la proprietà NumeralShapes per specificare o determinare le forme dei numerali da visualizzare e stampare in una casella combinata, un'etichetta, una casella di riepilogo o una casella di testo. Elemento Byte in lettura e scrittura.
		OldBorderStyle 	È possibile utilizzare questa proprietà per impostare o restituire il valore non modificato della proprietà BorderStyle di una maschera o un controllo. Questa proprietà è utile nel caso in cui sia necessario ripristinare uno stile di bordo non modificato o preferito. Elemento Byte in lettura e scrittura.
		OldValue 	È possibile utilizzare la proprietà OldValue per determinare il valore non modificato di un controllo associato. Elemento Variant in sola lettura.
		OnChange 	Imposta o restituisce il valore della casella Su modifica nella finestra di dialogo Proprietà relativa a uno degli oggetti dell'elenco Si applica a. Elemento String in lettura e scrittura.
		OnClick 	Restituisce o imposta il valore della casella Su clic della finestra Proprietà. Elemento String in lettura e scrittura.
		OnDblClick 	Restituisce o imposta il valore della casella Su doppio clic della finestra Proprietà. Elemento String in lettura e scrittura.
		OnDirty 	È possibile utilizzare questa proprietà per impostare o restituire il valore della casella In attesa della finestra Proprietà di una maschera o di un report. Elemento String di lettura e scrittura.
		OnEnter 	Restituisce o imposta il valore della casella Su INVIO della finestra Proprietà dell'oggetto specificato. Elemento String in lettura e scrittura.
		OnExit 	Restituisce o imposta il valore della casella Su uscita della finestra Proprietà dell'oggetto specificato. Elemento String in lettura e scrittura.
		OnGotFocus 	Restituisce o imposta il valore della casella Su attivato della finestra Proprietà dell'oggetto specificato. Elemento String in lettura e scrittura.
		OnKeyDown 	Restituisce o imposta il valore della casella Su tasto giù della finestra Proprietà. Elemento String in lettura e scrittura.
		OnKeyPress 	Restituisce o imposta il valore della casella Su pressione della finestra Proprietà. Elemento String in lettura e scrittura.
		OnKeyUp 	Restituisce o imposta il valore della casella Su tasto su della finestra Proprietà. Elemento String in lettura e scrittura.
		OnLostFocus 	Restituisce o imposta il valore della casella Su disattivato della finestra Proprietà dell'oggetto specificato. Elemento String in lettura e scrittura.
		OnMouseDown 	Restituisce o imposta il valore della casella Su pulsante mouse giù della finestra Proprietà. Elemento String in lettura e scrittura.
		OnMouseMove 	Restituisce o imposta il valore della casella Su mouse spostato della finestra Proprietà. Elemento String in lettura e scrittura.
		OnMouseUp 	Restituisce o imposta il valore della casella Su pulsante mouse su della finestra Proprietà. Elemento String in lettura e scrittura.
		OnNotInList 	Restituisce o imposta il valore della casella Su non in elenco della finestra Proprietà di una casella combinata. Elemento String in lettura e scrittura.
		OnUndo 	È possibile utilizzare questa proprietà per restituire o impostare un valore String che indica la macro, la routine evento o la funzione definita dall'utente che viene eseguita quando si verifica l'evento Undo. Elemento di lettura e scrittura.
		Parent 	Elemento in sola lettura
		Properties 	È possibile utilizzare questa proprietà per restituire un riferimento all'oggetto insieme Properties di un controllo. Elemento di sola lettura.
		ReadingOrder 	È possibile utilizzare la proprietà ReadingOrder per specificare o determinare l'ordine di lettura delle parole di un testo. Elemento Byte in lettura e scrittura.
		Recordset 	È possibile utilizzare questa proprietà per restituire o impostare l'oggetto ADO Recordset o l'oggetto DAO Recordset che rappresenta l'origine record dell'oggetto specificato. Elemento Object di lettura e scrittura.
		RightMargin 	Along with the TopMargin, Left Margin, and BottomMargin properties, specifies the location of information displayed within a combo box control. Valore Integer di lettura/scrittura.
		RightPadding 	Recupera o imposta la quantità di spazio, espressa in pollici, tra la casella combinata e la corrispondente griglia destra. Valore Integer di lettura/scrittura.
		RowSource 	È possibile utilizzare la proprietà RowSource, insieme alla proprietà RowSourceType, per indicare a Microsoft Access come fornire dati all'oggetto specificato. Elemento String di lettura e scrittura.
		RowSourceType 	È possibile utilizzare la proprietà RowSourceType, insieme alla proprietà RowSource, per indicare a Microsoft Access come fornire dati all'oggetto specificato. Elemento String di lettura e scrittura.
		ScrollBarAlign 	È possibile utilizzare la proprietà ScrollBarAlign per specificare o determinare l'allineamento di una barra di scorrimento verticale. Elemento Byte di lettura e scrittura.
		Section 	È possibile identificare questi controlli in base alla sezione di una maschera o di un report in cui è incluso il controllo. Elemento Integer di lettura e scrittura.
		Selected 	È possibile utilizzare la proprietà Selected in Visual Basic per determinare se un elemento incluso in una casella combinata è selezionato. Valore Long di lettura/scrittura.
		SelLength 	La proprietà SelLength specifica o determina il numero di caratteri selezionati nella parte casella di testo di una casella combinata. Elemento Integer di lettura e scrittura.
		SelStart 	La proprietà SelStart specifica o determina il punto di inizio del testo selezionato o la posizione del punto di inserimento se non è stato selezionato del testo. Elemento Integer in lettura e scrittura.
		SelText 	La proprietà SelText restituisce una stringa contenente il testo selezionato. Elemento String di lettura e scrittura.
		SeparatorCharacters 	Recupera o imposta il separatore visualizzato tra valori se la casella combinata è associata a un campo multivalore. Valore AcSeparatorCharacters di lettura/scrittura.
		ShortcutMenuBar 	È possibile utilizzare la proprietà ShortcutMenuBar per specificare il menu di scelta rapida che verrà visualizzato facendo clic con il pulsante destro del mouse sull'oggetto specificato. Elemento String in lettura e scrittura.
		ShowOnlyRowSourceValues 	Recupera o imposta se la casella combinata può visualizzare valori non specificati tramite la proprietà RowSource. Valore Boolean di lettura/scrittura.
		SmartTags 	Utilizzare questa proprietà per restituire un insieme SmartTags che rappresenta l'insieme di smart tag aggiunti a un controllo.
		SpecialEffect 	È possibile utilizzare la proprietà SpecialEffect per specificare se all'oggetto specificato si applica una formattazione speciale. Elemento Byte in lettura e scrittura.
		StatusBarText 	È possibile utilizzare la proprietà StatusBarText per specificare il testo visualizzato sulla barra di stato quando viene selezionato un controllo. Stringa di lettura/scrittura.
		TabIndex 	È possibile utilizzare la proprietà TabIndex per specificare la posizione di un controllo nell'ordine di tabulazione di una maschera o di un report. Valore di tipo Integer di lettura/scrittura.
		TabStop 	È possibile utilizzare la proprietà TabStop per specificare se sia consentito l'uso del tasto di tabulazione per spostare lo stato attivo su un controllo. Valore di tipo Boolean di lettura/scrittura.
		Tag 	Consente di memorizzare informazioni aggiuntive su una maschera, un report, una sezione o un controllo richieste da un'applicazione di Microsoft Access. Stringa di lettura/scrittura.
		Text 	È possibile utilizzare la proprietà Text per impostare o restituire il testo contenuto nella parte casella di testo di una casella combinata. Elemento String di lettura e scrittura.
		TextAlign 	La proprietà TextAlign specifica l'allineamento del testo nei nuovi controlli. Elemento Byte di lettura e scrittura.
		Top 	Elemento in lettura e scrittura
		TopMargin 	Insieme alle proprietà LeftMargin, RightMargin e BottomMargin, specifica la posizione delle informazioni visualizzate in un controllo casella di testo Valore Integer di lettura/scrittura.
		TopPadding 	Recupera o imposta la quantità di spazio, espressa in pollici, tra la casella combinata e la corrispondente griglia superiore. Valore Integer di lettura/scrittura.
		ValidationRule 	È possibile utilizzare la proprietà ValidationRule per specificare i requisiti relativi a dati immessi in un record, campo o controllo. Quando si immettono dati che violano l'impostazione della proprietà ValidationRule, è possibile specificare il messaggio da visualizzare all'utente impostando la proprietà ValidationText. Elemento String di lettura e scrittura.
		ValidationText 	È possibile utilizzare la proprietà ValidationText per specificare un messaggio da visualizzare quando vengono immessi dati che violano un'impostazione di ValidationRule per un record, campo o controllo. Elemento String in lettura e scrittura.
		Value 	È possibile utilizzare questa proprietà per determinare o specificare quale valore o opzione è stata selezionata nella casella combinata. Elemento Variant di lettura e scrittura.
		VerticalAnchor 	Recupera o imposta una costante AcVerticalAnchor che indica come la casella combinata specificata verrà ancorata in senso verticale all'interno del relativo layout. Valore di lettura/scrittura.
		Visible 	Elemento in lettura e scrittura
		Width 	Elemento in lettura e scrittura

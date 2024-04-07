GUIDA_RIFERIMENTI_MODELLI_DAO

		Guida di riferimento a DAO
				@MODELLI@DAO@riferimenti_(ci anche le novita @ado)

		Novità di ADO 

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


		Guida riferimento Dao
			...
				TODO: finire la guida DAo
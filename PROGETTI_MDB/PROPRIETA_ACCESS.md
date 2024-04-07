PROPRIETA_ACCESS.md

	F
		FORMAT
			Note
				Esempio di proprieta format in access che permette la formattazione dei dati.
					@access@format_(@proprieta per la @formattazione@dei@dati in una casella di testo)





			TextBox.Format, proprietà
				
				È possibile utilizzare la proprietà Format per personalizzare il modo in cui vengono visualizzati e stampati numeri, date, ore e testo. Elemento String di lettura e scrittura.

				Sintassi

					espressione.Format

					espressione   Variabile che rappresenta un oggetto TextBox.

				Note


					È possibile utilizzare uno dei formati predefiniti o creare un formato personalizzato utilizzando i simboli di formattazione.

					La proprietà Format utilizza diverse impostazioni per diversi tipi di dati (tipo di dati di un campo: Caratteristica di un campo che determina il tipo di dati che è possibile memorizzare. In un campo con tipo di dati Testo, ad esempio, è possibile memorizzare dati costituiti da testo o numeri, ma in un campo Numerico è possibile memorizzare solo dati numerici.). Per ulteriori informazioni relative alle impostazioni per tipi di dati specifici, selezionare uno dei seguenti argomenti:

				Tipo di dati Data/ora
				Tipi di dati Numerico e Valuta
				Tipi di dati Testo e Memo
				Tipo di dati Sì/No

				In Visual Basic immettere un'espressione stringa (espressione stringa: Qualsiasi espressione che restituisca una sequenza di caratteri adiacenti. Gli elementi dell'espressione possono essere funzioni che restituiscono una stringa o una stringa di tipo Variant (VarType 8), una variabile letterale, una costante, una variabile o un valore Variant.) corrispondente a uno dei formati predefiniti oppure immettere un formato personalizzato.

				La proprietà Format influenza solo il modo in cui i dati vengono visualizzati, non la memorizzazione dei dati.

				In Microsoft Access sono disponibili formati predefiniti per i tipi di dati Data/ora, Numerico e Valuta, Testo e Memo e Sì/No. I formati predefiniti variano in base al paese o area specificato nella finestra di dialogo Opzioni internazionali del Pannello di controllo di Windows. In Microsoft Access vengono visualizzati i formati appropriati a seconda del paese o area selezionato. Se, ad esempio, nella scheda Generale si seleziona Inglese (Stati Uniti), nel formato Valuta il numero 1234,56 viene visualizzato come $1,234.56. Se viene indicato Inglese (Gran Bretagna), il numero viene visualizzato come £1,234.56.

				Se la proprietà Format di un campo viene impostata Visualizzazione struttura della tabella, tale formato viene utilizzato per visualizzare i dati contenuti nei fogli dati. Viene inoltre applicata la proprietà Format del campo ai nuovi controlli contenuti nelle maschere e nei report.

				Nei formati personalizzati è possibile utilizzare i simboli riportati di seguito per qualsiasi tipo di dati.

				Simbolo Significato 
				(Spazio) Visualizza gli spazi come caratteri letterali. 
				"ABC" Visualizza ciò che è racchiuso tra virgolette come caratteri letterali. 
				! Forza l'allineamento a sinistra anziché a destra. 
				* Riempie lo spazio disponibile con il carattere successivo. 
				\ Visualizza il carattere successivo come carattere letterale. Per ottenere lo stesso risultato è anche possibile racchiudere i caratteri letterali tra virgolette. 
				[colore] Visualizza i dati formattati nel colore indicato tra parentesi. I colori disponibili sono: Nero, Blu, Verde, Azzurro, Rosso, Magenta, Giallo, Bianco. 

				Non è possibile utilizzare insieme simboli di formattazione personalizzati per tipi di dati Numerico e Valuta con simboli di formattazione Data/ora, Sì/No o Testo e Memo.

				Nel caso in cui sia stata definita una maschera di input (maschera di input: Formato composto da caratteri letterali visualizzati (come parentesi, punti e trattini) e caratteri maschera che specificano la posizione di immissione dei dati, i tipi di dati consentiti e il numero di caratteri ammesso.) e impostata la proprietà Format per gli stessi dati, al momento della visualizzazione dei dati, la proprietà Format ha la precedenza. Nel caso in cui, ad esempio, sia stata creata una maschera di input Password in visualizzazione Struttura della tabella e che nella tabella oppure in un controllo della maschera sia stata impostata anche la proprietà Format per lo stesso campo, la maschera di input Password viene ignorata e i dati vengono visualizzati in base alla proprietà Format.
				Esempio

				Nei tre esempi riportati di seguito viene impostata la proprietà Format utilizzando un formato predefinito.

				Visual Basic, Application Edition 
				Me!Date.Format = "Medium Date"

				Me!Time.Format = "Long Time"

				Me!Registered.Format = "Yes/No" 

				Nell'esempio successivo viene indicato come impostare la proprietà Format utilizzando un formato personalizzato. Viene visualizzata la data nel formato: gen 2006.

				Visual Basic, Application Edition 
				Forms!Employees!HireDate.Format = "mmm yyyy" 

				Nell'esempio riportato di seguito viene indicata una funzione di Visual Basic che formatta dati numerici nel formato Valuta e formatta dati di testo in lettere maiuscole. La funzione viene chiamata dall'evento OnLostFocus di un controllo non associato denominato TaxRefund.

				Visual Basic, Application Edition 
					Function FormatValue() As Integer
					    Dim varEnteredValue As Variant

					    varEnteredValue = Forms!Survey!TaxRefund.Value
					    If IsNumeric(varEnteredValue) = True Then
					        Forms!Survey!TaxRefund.Format = "Currency"
					    Else
					        Forms!Survey!TaxRefund.Format = ">"
					    End If
					End Function 


			Esempio di funzione Format
						@FORMATTA@TESTO_(il testo di una casella di testo formattato in numeri con la funzione @format)
				
				In questo esempio vengono illustrati vari utilizzi della funzione Format per formattare i valori utilizzando sia formati predefiniti che formati definiti dall'utente. Per quanto riguarda il separatore di data (/), il separatore di ora (:) e i valori letterali AM/PM, il formato dell'output effettivamente visualizzato dipenderà dalle impostazioni internazionali del sistema in cui si esegue il codice. Quando date e orari vengono visualizzati in ambiente di sviluppo, verranno utilizzati il formato breve di data e il formato breve di ora delle impostazioni internazionali del codice. Quando date e orari vengono visualizzati dal codice in esecuzione, verranno utilizzati i formati brevi di data e ora delle impostazioni internazionali del sistema. Nell'esempio seguente, le impostazioni internazionali sono Italia/Italiano.

				MyTime e MyDate verranno visualizzate nell'ambiente di sviluppo utilizzando le impostazioni correnti di sistema per il formato breve di data e il formato breve di ora.

				Dim MyTime, MyDate, MyStr
				MyTime = #17:04:23#
				MyDate = #Gennaio 27, 1993#

				' Restituisce l'ora corrente di sistema nel formato
				' esteso impostato nel sistema.
				MyStr = Format(Time, "Long Time")

				' Restituisce la data corrente di sistema nel formato
				' esteso impostato nel sistema.
				MyStr = Format(Date, "Long Date")      '// @formatta@date_(IN ACCESS viene formatta un un data long)

				MyStr = Format(MyTime, "h:m:s")    ' Restituisce "17.04.23".
				MyStr = Format(MyTime, "hh:mm:ss AMPM")    ' Restituisce "05.04.23".
				MyStr = Format(MyDate, "dddd, mmm d yyyy")    ' Restituisce "mercoledì, gen 27 1993".
				' Se il formato non viene indicato, viene restituita una stringa.
				MyStr = Format(23)    ' Restituisce "23".

				' Formati definiti dall'utente.
				MyStr = Format(5459.4, "##,##0.00")    ' Restituisce "5.459,40".
				MyStr = Format(334.9, "###0.00")    ' Restituisce "334,90".
				MyStr = Format(5, "0.00%")    ' Restituisce "500,00%".
				MyStr = Format("SALVE", "<")    ' Restituisce "salve".
				MyStr = Format("Ecco", ">")    ' Restituisce "ECCO".

					@ESEMPIO@DI@FORMATTAZIONE_(di un testo in un @numero @double, @formatta da @un @milione di euro)
					Me.Parent!TXT_FILTRO_01.Value = Format(CDbl(CVar(Me.IMPORTO_PROGETTO_lng_TXT)), "#,###,###.00") '// @FORMATTA UN 1.000.000,00 @IN@EURO





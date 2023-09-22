ACCESS_PROPRIETA_ESEMPIO_FORMATTA_DATE.MD
	Note
		Come formattare le date in access.
		@farmattare@le@date_(le formattazioni della data in access)_@formatta@date
	
	Proprietà Format - Tipo di dati Data/ora

		È possibile impostare la proprietà Format su formati di data e ora predefiniti o utilizzare formati personalizzati per i tipi di dati Data/Ora.

		Impostazione Formati predefiniti
		Nella tabella seguente vengono illustrate le impostazioni predefinite della proprietà Format per il tipo di dati Data/ora.

		Impostazione Descrizione 
		Data generica Impostazione predefinita. Se il valore indica solo la data, non viene visualizzata l'ora; se il valore indica 
						solo l'ora, non viene visualizzata la data. Questa impostazione è la combinazione delle impostazioni Data 
						breve e Ora estesa. 
						Esempi: 04/03/96 17.34.00 oppure 04/03/96 17.34.00.
		 
		Data estesa 	Equivalente all'impostazione Data estesa nelle opzioni internazionali di Windows. 
						Esempio: sabato 3 aprile 1996.
		 
		Data breve 		Esempio: 3 apr 96. 
		Data in cifre 	Equivalente all'impostazione Data breve nelle opzioni internazionali di Windows. 
						Esempio: 04/03/96.

						Avvertenza  L'impostazione Data in cifre suppone che le date tra il 01/01/00 e 31/12/29 facciano riferimento al ventunesimo secolo, vale a dire agli anni compresi tra il 2000 e il 2029. Le date comprese tra il 01/10/30 e 31/12/99 fanno, invece, riferimento al ventesimo secolo, vale a dire agli anni compresi tra il 1930 e il 1999.
		 
		Ora estesa 		Equivalente all'impostazione delle opzioni internazionali di Windows nella scheda Ora.
						Esempio: 17.34.23.
		 
		Ora breve 12h 	Esempio: 17.34.00. 
		Ora breve 24h 	Esempio: 17.34. 



		Formati personalizzati
				È possibile creare formati di data e ora personalizzati utilizzando i simboli seguenti:

				Simbolo 		Descrizione 
				. (punto) 		Separatore (separatore: Carattere che separa unità di testo o numeri.) di ora. I separatori vengono 
								impostati nelle opzioni internazionali di Windows. 
				/ 				Separatore di data. 
				c 				Identico al formato predefinito per la data generica. 
				g 				Giorno del mese espresso con una o due cifre numeriche, in base alle specifiche esigenze (da 1 a 
								31). 
				gg 				Giorno del mese espresso con due cifre numeriche (da 01 a 31). 
				ggg 			Prime tre lettere del giorno della settimana (da lun a dom). 
				gggg 			Nome per esteso del giorno della settimana (da lunedì a domenica). 
				ggggg 			Identico al formato predefinito per la data breve. 
				gggggg 			Identico al formato predefinito per la data estesa. 
				s 				Giorno della settimana (da 1 a 7). 
				ss 				Settimana dell'anno (da 1 a 53). 
				m 				Mese dell'anno espresso con una o due cifre numeriche, in base alle specifiche esigenze (da 1 a 
								12). 
				mm 				Mese dell'anno espresso con due cifre numeriche (da 01 a 12). 
				mmm 			Prime tre lettere del mese (da gen a dic). 
				mmmm 			Nome per esteso del mese (da gennaio a dicembre). 
				q 				Data visualizzata come trimestre dell'anno (da 1 a 4). 
				a 				Numero del giorno dell'anno (da 1 a 366). 
				aa 				Ultime due cifre dell'anno (da 01 a 99). 
				aaaa 			Anno per esteso (da 0100 a 9999). 
				h 				Ora espressa con due cifre numeriche, in base alle specifiche esigenze (da 0 a 23). 
				hh 				Ora espressa con due cifre numeriche (da 00 a 23). 
				n 				Minuti espressi con una o due cifre numeriche, in base alle esigenze (da 0 a 59). 
				nn 				Minuti espressi con due cifre numeriche (da 00 a 59). 
				s 				Secondi espressi con una o due cifre numeriche, in base alle specifiche esigenze (da 0 a 59). 
				ss 				Secondi espressi con due cifre numeriche (da 00 a 59). 
				ttttt 			Identico al formato predefinito dell'ora estesa. 
				AM/PM 			Indicazione oraria espressa nel formato 12 ore seguita dall'abbreviazione "AM" o "PM" appropriata 
								in caratteri maiuscoli. 
				am/pm 			Indicazione oraria espressa nel formato 12 ore seguita dall'abbreviazione "am" o "pm" appropriata 
								in caratteri minuscoli. 
				A/P 			Indicazione oraria espressa nel formato 12 ore seguita dall'abbreviazione "A" o "P" appropriata in 
								caratteri maiuscoli. 
				a/p 			Indicazione oraria espressa nel formato 12 ore seguita dall'abbreviazione "a" o "p" appropriata in 
								caratteri minuscoli. 
				AM PM 			Ora espressa nel formato 12 ore seguita dall'abbreviazione appropriata per le ore antimeridiane o 
								pomeridiane definita nelle opzioni internazionali di Windows. 



				I formati personalizzati vengono visualizzati in base alle impostazioni specificate nelle opzioni internazionali di Windows. I formati personalizzati che non corrispondono alle impostazioni specificate nelle opzioni internazionali di Windows vengono ignorati.

				  Nota 
				Se si desidera aggiungere una virgola o altri separatori al formato personalizzato, racchiudere ciascun separatore tra virgolette, ad esempio: g mmm ","aaaa. 

				Esempio
				Negli esempi seguenti vengono illustrati i formati di data/ora personalizzati.

				Impostazione Visualizzazione 
				ggg g mmm aaa lun 2 giu 1997 
				gg mmmm aaaa 02 giugno 1996 
				Settimana numero ss Settimana numero 22 
				Oggi è gggg Oggi è martedì 



				È possibile utilizzare un formato personalizzato per visualizzare "d.C." dopo o "a.C." prima dell'anno a seconda che venga inserito un numero positivo o negativo. Per esaminare un esempio di questo tipo di formato personalizzato, creare un nuovo campo di tabella, impostare il tipo di dati su Numerico e inserire un formato di questo tipo:

				"d.C. " #;# " a.C."

				I valori positivi vengono visualizzati con il numero relativo all'anno preceduto da "d.C.". I valori negativi vengono con visualizzati con il numero dell'anno seguito da "a.C."

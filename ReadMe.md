ReadMe.md
	c:\CASA\LINGUAGGI\ACCESS\



PROFILO GIT SE DA ERRORE
	
		P
			PROFILO GIT
				per i profilo git questa è l'utenza 
				email: icivixwalter@gmail.com
				passw: walter_01_

				per configurare il profilo per tutto utilizzo globale con questi comandi
				 	git config --global user.email "you@example.com"
  					git config --global user.name "Your Name"

				se vuoi impostar il profilo solo per la  repository
				corrente utilizza:
				git config user.email "you@example.com"
				git config user.name "Your Name"

					esempio
					git config user.email "icivixwalter@gmail.com"
					git config user.name "walter"




CREARE I RAMI
	Gestione delle filiali
		Un singolo repository Git può mantenere più rami di sviluppo. Per creare 
		un nuovo ramo denominato , utilizzareexperimental
			git branch experimental
			git branch ramo2			= 2 esempio

		per controllare quanti rami esistono eseguire
			git branch
		Verrà visualizzato un elenco di tutti i rami esistenti:

		Io ho il main ma posso creare un ramo secondario di appoggio
			git branch appoggio

		ho creato il brach e mi trovo ancora su main con l'asterisco
		per passare al nuovo branch digitare
			git switch appoggio


		Ora modifica un file, esegui il commit del file modificare e tornare al ramo
			git commit -a
			git branch main

		A questo punto i due rami si sono divisi, con cambiamenti diversi fatto in ciascuno. Per unire le modifiche apportate in ,  
			git merge appoggio

		Se le modifiche non sono in conflitto, il gioco è fatto. In caso di conflitti, I marcatori verranno lasciati nei file problematici che mostrano il conflitto;

			git diff


		per veder tutte le differenze tra i git digitare
			gitk
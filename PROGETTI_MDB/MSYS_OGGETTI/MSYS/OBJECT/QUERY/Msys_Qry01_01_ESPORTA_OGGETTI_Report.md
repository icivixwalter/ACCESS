NOTA_QUERY:
	Query di esportazione oggetti REPORT

	@query@ESPORTA@OGGETTI@report_(la query estrae GLI OGGETTI REPORT da @esportare)

SELECT

	Msys_Qry01_01_ESPORTA_OGGETTI_Report.md



	Msys_Qry01_01_ESPORTA_OGGETTI_Report_TMP.md
		Note
			estrae i report con codice PROGETTO da Tmp



				SCHEMA_SELECT_IN_TMP
		
					SELECT 
						Msys_ESPORTA_DB_EST.PATH_s, 
						Msys_ESPORTA_ELENCO_OGGETTI_DB.NRO_OGGETTO_i, 
						Msys_ESPORTA_ELENCO_OGGETTI_DB.TIPOGGETTO_s, 
						Msys_ESPORTA_ELENCO_OGGETTI_DB.NOMEOGGETTO_s, 
						Msys_ESPORTA_DB_EST.Scel_b

						FROM 
							Msys_ESPORTA_DB_EST_TMP 
							INNER JOIN 
								(Msys_ESPORTA_DB_EST INNER JOIN Msys_ESPORTA_ELENCO_OGGETTI_DB 
								ON 
								(Msys_ESPORTA_DB_EST.COD_PROGETTO_s = Msys_ESPORTA_ELENCO_OGGETTI_DB.COD_PROGETTO_s) 
								AND 
								(Msys_ESPORTA_DB_EST.PATH_s = Msys_ESPORTA_ELENCO_OGGETTI_DB.PATH_s)) 
								ON 
								(Msys_ESPORTA_DB_EST.PATH_s = Msys_ESPORTA_DB_EST_TMP.PATH_s) 
								AND 
								(Msys_ESPORTA_DB_EST_TMP.COD_PROGETTO_s = Msys_ESPORTA_DB_EST.COD_PROGETTO_s)

						WHERE 
								(((Msys_ESPORTA_ELENCO_OGGETTI_DB.TIPOGGETTO_s)="REPORTS") 
							AND 
								((Msys_ESPORTA_ELENCO_OGGETTI_DB.NOMEOGGETTO_s)>""))

						WITH 
							OWNERACCESS OPTION;
MSys_TABELLE_Qry01_01_SELECT_TUTTE_TMP.md

	Note
		la nuova query preleva dalle tabella ma solo con il codice della tabella progetto tmp


	SCHEMA

		SELECT 
				MSys_TABELLE.*, 
				MSys_TABELLE.COD_PROGETTO_s AS ORD_COD_PROGETTO_s, 
				MSys_TABELLE.NRO_OGGETTO_i AS ORD

			FROM 
				MSys_TABELLE INNER JOIN 
				PROGETTI_Msys_TB01_PROJECT_TMP ON 
				MSys_TABELLE.COD_PROGETTO_s = PROGETTI_Msys_TB01_PROJECT_TMP.COD_PROGETTO_s
			ORDER BY 
				MSys_TABELLE.COD_PROGETTO_s, MSys_TABELLE.NRO_OGGETTO_i;

ARRAY.MD

	Note
		array o matrici esempio di gestioni.



	SCHEMA

				@array@matrice_(esempio di gestione di un array di stringhe)				

			'//la matrice'	
			'//--------------------------------------------------------------------//
				Dim formsToImport As Variant                    '//calcolo form da importare


		   		' Elenco dei nomi delle form da importare
                formsToImport = Array("MSys_DF02_TIPO_OGGETTO", _
                                      "MSys_DF10_DATABASE", _
                                      "Msys_DF11_PROGETTI", _
                                      "Msys_DF12_GE_ObjProgetti")
          		


          		'ITERO NELL'ARRAY O MATRICE =  Importa ogni oggetto table
                For i = LBound(formsToImport) To UBound(formsToImport)
                    
                    tableName = formsToImport(i)
                    
                    '//@controllo@preliminare_(@esistenza della tabella e poi cancello)
                    '//ciclo nelle tabelle se esiste cancello altrimenti no
                    For Each tblCiclo In db.TableDefs
                        
                        Debug.Print tblCiclo.Name
                        If tblCiclo.Name = tableName Then
                            DoCmd.DeleteObject acTable, tableName
                        End If
                    
                    Next tblCiclo
                	
                	'//Esempio'    
                    DoCmd.TransferDatabase acImport, "Microsoft Access", sourceDBPath_s, acTable, tableName, tableName
                
                Next i
                
          '//--------------------------------------------------------------------//
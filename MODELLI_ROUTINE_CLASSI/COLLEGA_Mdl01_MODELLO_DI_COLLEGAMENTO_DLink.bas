Attribute VB_Name = "COLLEGA_Mdl01_LLPP_ATTI_Tb01_Gestione"
Option Compare Database


'//########################################################################################################################################################//
'// COLLEGA LA TABELLA DI SEGUITO INDICATA.
'// TABELLA DI ARCHIVIO   LLPP_ATTI_Tb01_Gestione e la
'// TABELLA TMP           LLPP_ATTI_Tb01_Gestione_tmp
'// NOTE:  per il collegamento occorre definire i seguenti parametri che _
    permetteranno il collegamento della tabella predefinita da un MDB esterno. _
    E' necessario utilizzare il comando DoCmd.TransferDatabase acLink con i seguenti dati di cui si riporta _
    l'esempio completo: _
    MODELLO DI COLLEGAMENTO ESPLICITO _
    DoCmd.TransferDatabase acLink, "Microsoft Access", "c:\GESTIONI\GESTIONE_LLPP\25_GESTIONE_LLPP\LLPP_ARCHVI_MDB\LLPP_GESTIONE\LLPP_ATTI_Tb01_GESTIONE.mdb", acTable, "LLPP_ATTI_Tb01_Gestione", "LLPP_ATTI_Tb01_Gestione", False _

'// MODELLO DI COLLEGAMENTO CON PARAMETRI: _
    Con i Parametri sono i seguenti _
                                     acLink                      = Tabella collegata _
                                     "Microsoft Access"          = Tipo database _
                                     PathDatabaseOrigine_s      = Database completo con path _
                                     acTable                    = tipo oggetto in questo caso una tabella _
                                     TabellaOrigine_s           = tabella originale da importare _
                                     TabellaDestinazione_s      = tabella di destinazione che può avere lo stesso nome di quella originale _
                                     False                      = ????
'//CODICE ----------->  COLLEGA_TABELLA_ATTI_DI_GESTIONE.01
'//########################################################################################################################################################//
Function COLLEGA_Mcr_LLPP_ATTI_Tb01_Gestione()
On Error GoTo COLLEGA_Mcr_LLPP_ATTI_Tb01_Gestione_Err

Dim PathOrigine_s As String
Dim TabellaOrigine_s As String
Dim DatabaseOrigine_s As String
Dim PathDatabaseOrigine_s  As String


 

'//IMPOSTAZIONE DELLE VARIABILI _
solo la path di origine e il dababase mdb che vengono unita nell'ultima variabile
PathOrigine_s = "c:\GESTIONI\GESTIONE_LLPP\25_GESTIONE_LLPP\LLPP_ARCHIVI_MDB\"





            'MODELLO DI COLLEGAMENTO ESPLICITO
            'TABELLA ------------------------> LLPP_ATTI_Tb01_Gestione
            'path della tabella  ------------> :c:\GESTIONI\GESTIONE_LLPP\25_GESTIONE_LLPP\LLPP_ARCHVI_MDB\LLPP_GESTIONE\
            'IL database --------------------> LLPP_ATTI_Tb01_GESTIONE.mdb
            
            'DoCmd.TransferDatabase acLink, "Microsoft Access", "c:\GESTIONI\GESTIONE_LLPP\25_GESTIONE_LLPP\LLPP_ARCHVI_MDB\LLPP_GESTIONE\LLPP_ATTI_Tb01_GESTIONE.mdb", acTable, "LLPP_ATTI_Tb01_Gestione", "LLPP_ATTI_Tb01_Gestione", False
            
            
        '//I° COLLEGAMENTO TABELLA
        '//======================================================================================================================//
        '//NOTE: per il collegamento della tabella occorre inserire i parametri:
        '//Parametri da inserire : acLink = Tabella collegata _
                                   "Microsoft Access"       = Tipo database _
                                    PathDatabaseOrigine_s   = Database completo con path _
                                    acTable                 = tipo oggetto in questo caso una tabella _
                                    TabellaOrigine_s        = tabella originale da importare _
                                    TabellaDestinazione_s   = tabella di destinazione che può avere lo stesso nome di quella originale _
                                    False                   = ????
                
                '//IMPOSTAZIONE DELLE VARIABILI _
                reimposto sia la tabelladi origine che di destinazione. Quest'ultima può essere anche ridenominata.
                TabellaOrigine_s = "LLPP_ATTI_Tb01_Gestione"
                TabellaDestinazione_s = "LLPP_ATTI_Tb01_Gestione"
                DatabaseOrigine_s = "LLPP_ATTI_Tb01_GESTIONE.mdb"
                PathDatabaseOrigine_s = PathOrigine_s & DatabaseOrigine_s

                
                
                '//CONTROLLO ERRORE DI TABELLA ON ERRORE RESUME NEXT
                '//In caso di errore non viene eseguita l'istruzione elimina oggetto ma quella successiva nel caso in cui la tabella collegata _
                era stata già cancella e quindi non esiste come oggetto collegato nel db corrente. Se la tabella collegata è ESISTENTE anche se _
                non funzionante viene eseguito il comando di eliminazione.
                '//CONTROLLO ERRORE SALTA L'OPERAZIONE SUCCESSIVA
                On Error Resume Next

                
                '//ELIMINO L'OGGETTO SOLO SE ESISTE _
                  prima elimina la tabella collegata gia esistente _
                  e dopo attivo il comando CANCELLAZIONE DELLA TABELLA, IN CASO DI ERRORE QUESTO COMANDO _
                  NON VIENE ESEGUITO
                DoCmd.DeleteObject acTable, TabellaOrigine_s
    
                
                '//Comando di collegamento
                DoCmd.TransferDatabase acLink, "Microsoft Access", _
                                              PathDatabaseOrigine_s, _
                                              acTable, TabellaOrigine_s, _
                                              TabellaDestinazione_s, _
                                              False
                                              
                                              
                        
                        '//COLLEGO I SEPARATORI I°
                        '//---------------------------------------------------------------------------------------------------//
                        
                            '//IMPOSTAZIONE DELLE VARIABILI _
                            reimposto sia la tabelladi origine che di destinazione. Quest'ultima può essere anche ridenominata.
                            TabellaOrigine_s = "LLPP_ATTI_Tb01_{@=============================================@}"
                            TabellaDestinazione_s = "LLPP_ATTI_Tb01_{@=============================================@}"
                            
                            '//ELIMINO L'OGGETTO SOLO SE ESISTE _
                            prima elimina la tabella collegata gia esistente _
                            e dopo attivo il comando CANCELLAZIONE DELLA TABELLA, IN CASO DI ERRORE QUESTO COMANDO _
                            NON VIENE ESEGUITO
                            DoCmd.DeleteObject acTable, TabellaOrigine_s
                            
                            
                            '//Comando di collegamento
                            DoCmd.TransferDatabase acLink, "Microsoft Access", _
                                            PathDatabaseOrigine_s, _
                                            acTable, TabellaOrigine_s, _
                                            TabellaDestinazione_s, _
                                            False
                        '//---------------------------------------------------------------------------------------------------//
                        
                        
                        
                        '//COLLEGO I SEPARATORI II°
                        '//---------------------------------------------------------------------------------------------------//
                        
                            '//IMPOSTAZIONE DELLE VARIABILI _
                            reimposto sia la tabelladi origine che di destinazione. Quest'ultima può essere anche ridenominata.
                            TabellaOrigine_s = "LLPP_ATTI_Tb01_}-----------------------------------------------@"
                            TabellaDestinazione_s = "LLPP_ATTI_Tb01_}-----------------------------------------------@"
                            
                            '//ELIMINO L'OGGETTO SOLO SE ESISTE _
                            prima elimina la tabella collegata gia esistente _
                            e dopo attivo il comando CANCELLAZIONE DELLA TABELLA, IN CASO DI ERRORE QUESTO COMANDO _
                            NON VIENE ESEGUITO
                            DoCmd.DeleteObject acTable, TabellaOrigine_s
                            
                            
                            '//Comando di collegamento
                            DoCmd.TransferDatabase acLink, "Microsoft Access", _
                                            PathDatabaseOrigine_s, _
                                            acTable, TabellaOrigine_s, _
                                            TabellaDestinazione_s, _
                                            False
                        '//---------------------------------------------------------------------------------------------------//
                        
                        
                        
                                       
        '//*** fine ***
        '//I° COLLEGAMENTO TABELLA
        '//======================================================================================================================//
                                      

            
        '// II° COLLEGAMENTO CON CANCELLAZIONE DELLA PRECEDENTE SE ESISTE
        '//======================================================================================================================//
        '//NOTE: per il collegamento della tabella occorre inserire i parametri:
        '//Parametri da inserire : acLink = Tabella collegata _
                                    "Microsoft Access"       = Tipo database _
                                     PathDatabaseOrigine_s   = Database completo con path _
                                     acTable                 = tipo oggetto in questo caso una tabella _
                                     TabellaOrigine_s        = tabella originale da importare _
                                     TabellaDestinazione_s   = tabella di destinazione che può avere lo stesso nome di quella originale _
                                     False                   = ????
                    
                    '//IMPOSTAZIONE DELLE VARIABILI _
                    reimposto sia la tabelladi origine che di destinazione. Quest'ultima può essere anche ridenominata.
                    TabellaOrigine_s = "LLPP_ATTI_Tb01_Gestione_TMP"
                    TabellaDestinazione_s = "LLPP_ATTI_Tb01_Gestione_TMP"
                    DatabaseOrigine_s = "LLPP_ATTI_Tb01_GESTIONE_TMP.mdb"
                    PathDatabaseOrigine_s = PathOrigine_s & DatabaseOrigine_s

                    
                    
                    '//CONTROLLO ERRORE DI TABELLA ON ERRORE RESUME NEXT
                    '//In caso di errore non viene eseguita l'istruzione elimina oggetto ma quella successiva nel caso in cui la tabella collegata _
                    era stata già cancella e quindi non esiste come oggetto collegato nel db corrente. Se la tabella collegata è ESISTENTE anche se _
                    non funzionante viene eseguito il comando di eliminazione.
                    '//CONTROLLO ERRORE SALTA L'OPERAZIONE SUCCESSIVA
                    On Error Resume Next
    
                    
                    '//ELIMINO L'OGGETTO SOLO SE ESISTE _
                      prima elimina la tabella collegata gia esistente _
                      e dopo attivo il comando CANCELLAZIONE DELLA TABELLA, IN CASO DI ERRORE QUESTO COMANDO _
                      NON VIENE ESEGUITO
                    DoCmd.DeleteObject acTable, TabellaOrigine_s
        
                    
                    '//Comando di collegamento
                    DoCmd.TransferDatabase acLink, "Microsoft Access", _
                                                  PathDatabaseOrigine_s, _
                                                  acTable, TabellaOrigine_s, _
                                                  TabellaDestinazione_s, _
                                                  False
                                           
            '//-------------------------------------------------------------------------------------------------------//
            
            
                
                        '//COLLEGO I SEPARATORI I°
                        '//---------------------------------------------------------------------------------------------------//
                        
                            '//IMPOSTAZIONE DELLE VARIABILI _
                            reimposto sia la tabelladi origine che di destinazione. Quest'ultima può essere anche ridenominata.
                            TabellaOrigine_s = "LLPP_ATTI_Tb01_Gestione_TM_}----------------------------------@"
                            TabellaDestinazione_s = "LLPP_ATTI_Tb01_Gestione_TM_}----------------------------------@"
                            
                            '//ELIMINO L'OGGETTO SOLO SE ESISTE _
                            prima elimina la tabella collegata gia esistente _
                            e dopo attivo il comando CANCELLAZIONE DELLA TABELLA, IN CASO DI ERRORE QUESTO COMANDO _
                            NON VIENE ESEGUITO
                            DoCmd.DeleteObject acTable, TabellaOrigine_s
                            
                            
                            '//Comando di collegamento
                            DoCmd.TransferDatabase acLink, "Microsoft Access", _
                                            PathDatabaseOrigine_s, _
                                            acTable, TabellaOrigine_s, _
                                            TabellaDestinazione_s, _
                                            False
                        '//---------------------------------------------------------------------------------------------------//
                      
                    
            
        
        '//*** FINE ***
        '// II° COLLEGAMENTO CON CANCELLAZIONE DELLA PRECEDENTE SE ESISTE
        '//======================================================================================================================//
    

COLLEGA_Mcr_LLPP_ATTI_Tb01_Gestione_Exit:
    Exit Function

COLLEGA_Mcr_LLPP_ATTI_Tb01_Gestione_Err:
    MsgBox Error$
    Resume COLLEGA_Mcr_LLPP_ATTI_Tb01_Gestione_Exit

End Function



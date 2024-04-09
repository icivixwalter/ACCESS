Attribute VB_Name = "MODELLI_DEFINIZIONE_Mdl16_DAO_RECORDSET"
Option Compare Database

'//AGGIORNA_NUMERO_PROGRESSIVO_E_NOME_FILE
'//DAO_RECORDSET_pSub
'//========================================================================================================================================//
'//DENOMINAZIONE--->:Itero in un Rs _
                     DAO
                                            
'//Tipo------------>:routine pubblica.
'//Attività-------->:controllo i campi _
                     del Recorset
'//NOTE------------>:Aggiorna il campo del _
                     recordset
'//Parametro------->:par_sSql = tipo stringa ed equivale alla sql o _
                     alla query da utilizzare.
'//Restituisce----->:Null _
'//Codice---------->:DAO_RECORDSET_pSub.01 _
'/
Private Sub DAO_RECORDSET_pSub(par_sSql As String)

'//AGGIORNO I CAMPI NRO + FILE + DATA ED ORA
'//--------------------------------------------------------------------------------------------------
'//Codice---------->:DAO_RECORDSET_pSub.01.COSTRUISCI
        
        
        '//IMPOSTO LA STRINGA SSQL di apertura del Rs
        sSql = ""
        sSql = par_sSql

'//ITERO NELLA TABELLA
'//.....................................................................................................
'//NOTE------------>:Tramite una Select vengono individuati i valori da restiuire.
'//Codice---------->:DAO_RECORDSET_pSub.01.itero

     '//Apro il Database
     Set DaoDB = DBEngine.Workspaces(0).Databases(0)
     '//Apro un Recordset dal parametro ssql
     Set DAORs = DaoDB.OpenRecordset(sSql)
        
    If DAORs.EOF = False And DAORs.BOF = False Then
        '//Posizione Primo record
        DAORs.MoveFirst

       
            '//ITERAZIONE_RECORSET
            '//.....................................................................................................//
            '//Codice---------->: MODELLO_SUB_N01_IterazioneRecord_pSub.01.01
            '//Note------------>: Tramite una Select vengono individuati i valori da restiuire.
        
                            
                            While Not DAORs.EOF
                              '//Blocco iterazione
                                 DoEvents
                                         
                                    '//POSIZIONE_ATTIVITA
                                    ProceduraAttivaEseguita_s = ProceduraAttivaEseguita_s & "_01"
                                         
                                    '//IF DI CONTROLLO CAMPI
                                    '//_______________________________________________________________________
                                    '//NOTE :
                                    
                                        If DAORs.Fields("Campo_01") = par_AnnoImp_i _
                                        And DAORs.Fields("Campo_02") = par_CodiceTributo_s Then
                                    
                                        
                                                    '//Salvo nella Variabile
                                                    CampoCercato_s = DAORs.Fields("Campo_01")
                                    
                                        End If
                                    '//_______________________________________________________________________
                                    
                                    '//Record Successivo
                                    DAORs.MoveNext
                
                            Wend
                  '//** FINE **
                  '//ITERAZIONE_RECORSET
                  '//.....................................................................................................//
                    
                    '//Uscita Rs e chiusura oggetti
                    DAORs.Close
                    Set DAORs = Nothing
        
        End If  '//If DAORs.EOF = False And DAORs.BOF = False Then
        

    '//*** fine ***
    '//ITERO NELLA TABELLA
    '//.....................................................................................................


       
'//--------------------------------------------------------------------------------------------------
'//GESTIONE ERRORI E USCITA ROUTINE
'//NOTA:

Exit_DAO_RECORDSET_pSub:
    Exit Sub

Err_DAO_RECORDSET_pSub:
    
    MsgBox Err.Description & " - Errore Messaggio -> : " & ProceduraMessaggioErrore_s & Chr$(13) & _
           Err.Number
    Debug.Print ProceduraMessaggioErrore_s
    Debug.Print ProceduraAttivaEseguita_s
    Stop
    
    Resume Exit_DAO_RECORDSET_pSub
        
       
End Sub

'//DAO_RECORDSET_pSub
'//========================================================================================================================================//



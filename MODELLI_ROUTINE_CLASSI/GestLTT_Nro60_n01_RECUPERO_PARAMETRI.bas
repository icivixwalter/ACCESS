Attribute VB_Name = "GestLTT_Nro60_n01_RECUPERO_PARAMETRI"
Option Compare Database

'DIM
Dim DaoRs  As DAO.Recordset
Dim dtxDATA_INIZ As Date
Dim dtxDATA_FIN As Date
Dim sxFORM_CHIAMANTE As String
Dim sxSottoFORM_CHIAMANTE As String
'



'RECUPERA LE DATE
Public Function pfc_RECUPERA_I_PARAMETRI(par_ixPage As Integer)

On Error GoTo pfc_RECUPERA_I_PARAMETRI_Err

                
    '??? INSERIRE IF
    'di inviduazione date di estrazione ambo normali oppure
    'ambi non estratti
                
                'RECUPERO DATA INIZIALE E FINALE
                '==================================================================================================
                'Recupero le date di estrazione salvata
         
    
         
                    'Recupero la data di estrazione salvata
                    Set DaoRs = CurrentDb.OpenRecordset("UTIL_Tb10_PARAM_LOTTO")
                    While Not DaoRs.EOF
                    DaoRs.MoveFirst
                        
                        'SE LA PAGINA = 0 ALLORA AMBO
                        'recupero dal campo DATA_INIZ al campo DATA_FIN + campo ESTRAZIONE
                        If par_ixPage = 0 Then
                        
                            dtxESTRAZIONE = DaoRs.Fields("ESTRAZIONE").Value
                            'reset i controlli activex
                            'recupero la data inizio e la data fine
                            dtxDATA_INIZ = DaoRs.Fields("DATA_INIZ").Value
                            dtxDATA_FIN = DaoRs.Fields("DATA_FIN").Value
                        
                        
                        'SE LA PAGINA = 2 ALLORA AMBO NON ESTRATTI
                        'recupero dal campo AMB_NON_ESTRATTI_DATA_INIZ  a data AMB_NON_ESTRATTI_DATA_FIN + campo ESTRAZIONE
                        ElseIf par_ixPage = 1 Then
                        
                            dtxESTRAZIONE = DaoRs.Fields("ESTRAZIONE").Value
                            'reset i controlli activex
                            'recupero la data inizio e la data fine
                            dtxDATA_INIZ = DaoRs.Fields("AMB_NON_ESTRATTI_DATA_INIZ").Value
                            dtxDATA_FIN = DaoRs.Fields("AMB_NON_ESTRATTI_DATA_FIN").Value

                        
                        'SE LA PAGINA = 3 ALLORA TERNO
                        'recupero dal campo DATA_INIZ al campo DATA_FIN + campo ESTRAZIONE
                        ElseIf par_ixPage = 2 Then
                        
                            dtxESTRAZIONE = DaoRs.Fields("ESTRAZIONE").Value
                            'reset i controlli activex
                            'recupero la data inizio e la data fine
                            dtxDATA_INIZ = DaoRs.Fields("DATA_INIZ").Value
                            dtxDATA_FIN = DaoRs.Fields("DATA_FIN").Value
                        
                        
                        
                        'SE LA PAGINA = 4 ALLORA TERNO NON ESTRATTI
                        'recupero dal campo TERN_NON_ESTRATTI_DATA_INIZ al campo TERN_NON_ESTRATTI_DATA_INIZ + campo ESTRAZIONE
                        ElseIf par_ixPage = 3 Then
                        
                            dtxESTRAZIONE = DaoRs.Fields("ESTRAZIONE").Value
                            'reset i controlli activex
                            'recupero la data inizio e la data fine
                            dtxDATA_INIZ = DaoRs.Fields("TERN_NON_ESTRATTI_DATA_INIZ").Value
                            dtxDATA_FIN = DaoRs.Fields("TERN_NON_ESTRATTI_DATA_INIZ").Value
                        
                        
                        'Pagine non controllata parametri base
                        Else
                            dtxESTRAZIONE = DaoRs.Fields("ESTRAZIONE").Value
                            'reset i controlli activex
                            'recupero la data inizio e la data fine
                            dtxDATA_INIZ = DaoRs.Fields("DATA_INIZ").Value
                            dtxDATA_FIN = DaoRs.Fields("DATA_FIN").Value
                        
                        End If
                        
                            'la form principale
                            sxFORM_CHIAMANTE = DaoRs.Fields("FORM_CHIAMANTE").Value
                            
                            'La sottoform
                            If IsNull(DaoRs.Fields("SottoFORM_CHIAMANTE").Value) = False Then
                                sxSottoFORM_CHIAMANTE = DaoRs.Fields("SottoFORM_CHIAMANTE").Value
                            Else
                                sxSottoFORM_CHIAMANTE = ""
                            End If
                    
                    DaoRs.MoveLast
                    DaoRs.MoveNext
                    Wend
                    DaoRs.Close
                    Set DaoRs = Nothing
                    
                                                        
                            'stampa di controllo
                            Debug.Print "pfc_RECUPERA_I_PARAMETRI -  data iniziale: " & dtxDATA_INIZ
                            Debug.Print "pfc_RECUPERA_I_PARAMETRI -  data finale  : " & dtxDATA_FIN
                            
                       
                '==================================================================================================
                
'EXIT E GESTIONE ERRORI
'-----------------------------------------------------------------------------------------------
        
        

pfc_RECUPERA_I_PARAMETRI_Exit:
    Exit Function

pfc_RECUPERA_I_PARAMETRI_Err:
    MsgBox Error$
    Resume pfc_RECUPERA_I_PARAMETRI_Exit


End Function



'LA DATA INIZIALE
Public Function pfc_RESTITUISCI_DATA_INIZ() As Date

On Error GoTo pfc_RESTITUISCI_DATA_INIZ_Err

                    
                'RESTITUISCE LA DATA INIZIALE
                '==================================================================================================
                        
                       'Controlli id per il blocco della procedura
                       If dtxDATA_INIZ > 0 Then
                                                        
                            'DATA INIZIALE
                            pfc_RESTITUISCI_DATA_INIZ = dtxDATA_INIZ
                        
                        Else
                            'Se nulla restituisce la data corrente
                            pfc_RESTITUISCI_DATA_INIZ = Date
                            
                       End If
                       
                '==================================================================================================
                
'EXIT E GESTIONE ERRORI
'-----------------------------------------------------------------------------------------------
        
        

pfc_RESTITUISCI_DATA_INIZ_Exit:
    Exit Function

pfc_RESTITUISCI_DATA_INIZ_Err:
    MsgBox Error$
    Resume pfc_RESTITUISCI_DATA_INIZ_Exit


End Function




'LA DATA FINALE
Public Function pfc_RESTITUISCI_DATA_FIN() As Date

On Error GoTo pfc_RESTITUISCI_DATA_FIN_Err

                    
                'RESTITUISCE LA DATA FINALE
                '==================================================================================================
                        
                       'Controlli id per il blocco della procedura
                       If dtxDATA_FIN > 0 Then
                                                        
                            'DATA INIZIALE
                            pfc_RESTITUISCI_DATA_FIN = dtxDATA_FIN
                        
                        Else
                            'Se nulla restituisce la data corrente
                            pfc_RESTITUISCI_DATA_FIN = Date
                            
                       End If
                       
                '==================================================================================================
                
'EXIT E GESTIONE ERRORI
'-----------------------------------------------------------------------------------------------
        
        

pfc_RESTITUISCI_DATA_FIN_Exit:
    Exit Function

pfc_RESTITUISCI_DATA_FIN_Err:
    MsgBox Error$
    Resume pfc_RESTITUISCI_DATA_FIN_Exit


End Function





'LA FORM CHIAMANTE
Public Function pfc_RESTITUISCI_FORM_CHIAMANTE() As String

On Error GoTo pfc_RESTITUISCI_FORM_CHIAMANTE_Err

                    
                'RESTITUISCE LA form chiamante
                '==================================================================================================
                        
                       'Controlli id per il blocco della procedura
                       If sxFORM_CHIAMANTE > "" Then
                                                        
                            'DATA INIZIALE
                            pfc_RESTITUISCI_FORM_CHIAMANTE = sxFORM_CHIAMANTE
                        
                        Else
                            'Se nulla restituisce la data corrente
                            pfc_RESTITUISCI_FORM_CHIAMANTE = ""
                            
                       End If
                       
                '==================================================================================================
                
'EXIT E GESTIONE ERRORI
'-----------------------------------------------------------------------------------------------
        
        

pfc_RESTITUISCI_FORM_CHIAMANTE_Exit:
    Exit Function

pfc_RESTITUISCI_FORM_CHIAMANTE_Err:
    MsgBox Error$
    Resume pfc_RESTITUISCI_FORM_CHIAMANTE_Exit


End Function



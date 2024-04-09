Attribute VB_Name = "UTILITA_Mdl38_FOR_EACH_RESETTA_ControloSource+Label"

Option Compare Database
Option Explicit
 

'//PULISCO I CAMPI TEXT E LABEL DELLA FORM CORRENTE -
'//NOTE: eseguo il ciclo nell'insieme forms e nell'insieme dei controlli della form filtrata con il nome e condizionata _
dalla if di filtro. Cicolo all'interno delle proprieta delle form e resetto LE CASELLE DI TESTO E LE LABEL.
'//CODICE----------->LoopFormProps_PULISCI_OGGETTI_TEXT_LABEL_PSub
'//###########################################################################################################################//
Public Sub LoopFormProps_PULISCI_OGGETTI_TEXT_LABEL_PSub()
    On Error GoTo err_Handler
    Dim frmLoop As Object
    Dim ctrl As Control
    Dim propLoop As Property
    Dim strFormName As String
    
    
    '//CHIAMO LA ROUTINE DI PULITURA DEI CAMPI TEXT E LABEL
    '//---------------------------------------------------------------------------------//
    '//CODICE----------->LoopFormProps_PULISCI_OGGETTI_TEXT_LABEL_PSub.puliscoTutto
    
        '//ATTIVO LA ROUTINE Call LoopFormProps_PULISCI_OGGETTI_TEXT_LABEL_PSub
    
    '//---------------------------------------------------------------------------------//
    
    For Each frmLoop In CurrentProject.AllForms
        '//TROVA QUELLA CORRENTE
        If frmLoop.NAME = "MOD30_Frm01_S01_ELENCO_SCH01_65txt_campi_1key" Then
      
        '//LA FORM CORRENTE
        '//=====================================================================================================//
                strFormName = frmLoop.NAME
                
                '//APRO LA FORM IN MODALITA DESIGN
                DoCmd.OpenForm strFormName, acDesign, , , , acHidden
                
                '//Qualifico l'oggetto form pe le proprieta
                With Forms(strFormName)
                    
                    For Each ctrl In .Controls
                        
                        DoEvents
                         
                        For Each propLoop In ctrl.Properties
                        
                             If ctrl.NAME = "TXT_01" Then
                                'Stop
                             End If
                                
                            
                        '//PULISCO I CAMPI TEXT E LABEL DELLA FORM CORRENTE - CICOLO FOR EACH -
                        '//==================================================================================================//
                        '//NOTE : pulisco le caselle di testo impostando la proprieta recordsource a null mentre le label _
                                corrispondenti vengono impostate al nome oggetto label corrente.
                        '//CODICE----------->LoopFormProps_PULISCI_OGGETTI_TEXT_LABEL_PSub.puliscoTutto
                        
                        
                            
                            '//PULISCO LA PROPRIETA RECORD SOURCE DELLE CASELLE DI TESTO
                            '//--------------------------------------------------------------------------------------------//
                            '//NELLA CASELLE TEXT riconosciute con il valore ControlSource _
                            viene azzerata la proprieta control source con valore ""
                                If propLoop.NAME = "ControlSource" Then
                                    Debug.Print propLoop.Value
                                    propLoop.Value = ""
                                End If
                                
                            '//--------------------------------------------------------------------------------------------//
                            
                                                    
                            
                            '//PULISCO LA PROPRIETA CAPTION DELLE LABEL
                            '//--------------------------------------------------------------------------------------------//
                            '//se è una label riconosciuta con il nome della proprieta _
                            Caption, viene sostituita il valore del contenuto della _
                            etichetta con il nome della label
                                 If propLoop.NAME = "Caption" Then
                                            Debug.Print propLoop.Value
                                            '//inserisco il nome label?
                                            propLoop.Value = ctrl.NAME
                                                    
                                                    
                                End If
                                
                            '//--------------------------------------------------------------------------------------------//
                        
                        '//*** FINE ***
                        '//PULISCO I CAMPI TEXT E LABEL DELLA FORM CORRENTE - CICOLO FOR EACH -
                        '//==================================================================================================//
                             
                            '//CONTROLLO oggetto
                            Debug.Print "Form [" & .NAME & "] - Control[" & ctrl.NAME & "](" & propLoop.NAME & "):" & propLoop.Value
                            
                        Next '//For Each propLoop In ctrl.Properties
                    
                    
                    Next    '//For Each ctrl In .Controls
                
                            '//SALVATAGGIO FORM APERTA IN DESIGN
                            'Richiama la funzione che chiude la maschera aperta
                            DoCmd.SetWarnings False
                            DoCmd.Close acForm, strFormName, acSaveYes
                            DoCmd.SetWarnings True
                            DoCmd.Close acForm, strFormName

                End With '//With Forms(strFormName)
                
                               
        '//=====================================================================================================//
        
        End If
    Next
    
    
    
 
    Exit Sub
err_Handler:
    Select Case Err.Number
        Case 2186
            '//SE LA PROPRIETA NON ESISTE ATTIVO L'ERRORE E RIENTRO NEL CICLO FOR EACH
            'Property not available in design view  SE NON CI SONO LE  PROPRIETA RIPROVA.
            Resume Next
        Case Else
            MsgBox "Err: " & Err.Number & " - " & Err.Description
    End Select
End Sub
 

Sub prova()


Dim obj As Object
Dim frm As Form
    Dim ctrl As Control

    For Each obj In CurrentProject.AllForms
        If InStr(obj.NAME, "MODELLO_Frm01_M01_GE_FILTRO_M15_H15_L35_NoButton") Then Debug.Print "Forms|" & obj.NAME
        'Set frm = Forms(obj.Name) '- this had problems as well, so I commented out
        For Each ctrl In Forms(obj.NAME).Controls
            If InStr(ctrl.NAME, "LBL_FILTRO_01") Then Debug.Print "Form|" & obj.NAME & "|" & ctrl.NAME
        Next ctrl
    Next obj
    
End Sub

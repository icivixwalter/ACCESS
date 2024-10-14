Attribute VB_Name = "MODELLO_SUB"
'//MODELLO
'//-----------------------------------------------------------------------//
Private Sub MODELLO()

    On Error GoTo Err_MODELLO

    
            'Reset Variabili Oggetti Form
            m_sxTIPOGGETTO = ""
            m_sxPROPRIETA = ""
            m_sxMETODO = ""
            m_sxEVENTO = ""
            
            
'USCITA ED ERRORI
'..............................................................
Exit_MODELLO:
    Exit Sub

Err_MODELLO:
    MsgBox Err.Description
    Resume Exit_MODELLO

                                                      
End Sub

'//-----------------------------------------------------------------------//



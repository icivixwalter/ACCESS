Attribute VB_Name = "CARDIF_MDL01_Archivio"
'//************************************************************************************************
'//                 mdl UTILITA DELLE TABELLE GESTIONE CONDOMINIO



'//************************************************************************************************

'//###########################################################################################
'//Codice :OPTION.01
'//Nolte  : Le opzioni di scrittura

    '//OPZION
    '//........................................................
    Option Compare Text                     '//Le Opzioni di comparazione testo
    Option Explicit                         '//Le Opzioni esplicite per le variabili

    '//*** Fine ***
    '//OPZION
    '//........................................................
            

'//Variabili di database
'//>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>:>>>>>>>>>>>>>>>>>>>>>>>>>>>>



'//VARIABILI_DATABASE
'//###########################################################################################
'//Codice :VARIABILI_DEGLI_INDIRIZZI.01.A
'//
  '//LE_VARIABILI_GENERALI
  '//==============================================================================
  '//Codice   :VariabiliGenerali.01
  '//Note     :Le variabili Generali per la gestione dei Database Dao e Ado.

    '//DAO
    '//........................................................
    '//Codice :DimDbDao.01
    '//Note   :Le variabili del database Dao.
        
    Dim DaoDB As DAO.Database                   '//Database Dao
    Dim DaoWks As DAO.Workspace
    Dim DaoRs As DAO.Recordset
    
    
    '//*** Fine ***
    '//DAO
    '//........................................................
    
    
    '//ADO
    '//........................................................
    '//Codice :DimdDbAdo.01
    '//Note   :Le variabili del database Ado.
    
        Dim ADODB As Database                       '//Database Ado
        Dim AdodaoRs As Recordset
    
        
    '//*** Fine ***
    '//ADO
    '//........................................................
    
    '//LE_VARIABILI_DATABASE_ESTERNO
    '//........................................................
    '//Codice :DimDatabaseEsterno.01
    '//Note   :Le variabili per la ricerca e l'//apertura
    '//       del database esterno.
    
    
        Dim sPathDbEsterno As String                    '//Path del database esterno
        Dim sOption_SEZ As Integer                          '//Numero Sezione di Progetto Scelta
        Dim sName_Tab As String                             '//Nome Tabella da aprire/creare
        Dim Apridbs As Database                     '//Apri il database
        Dim intCicloDb As Integer                           '//Intero per ciclo lettura oggetti del database
        Dim appAccess As Access.Application             '//Applicazione access
        Dim strDB As String                     '//Stringa per il db
        Dim strReportName As String                 '//Stringa per il report.


    '//*** Fine ***
    '//DATABASE_ESTERNO
    '//........................................................
    


    '//LE_VARIABILI_COMUNI
    '//........................................................
    '//Codice :DimVariabiliComuni.01
    '//Note   :Le variabili comuni per la gestione del database
    '//       con quelle che rappresentano i tipi visual basic.

        '//Variabili generali
        Dim Str1 As String
        Dim Int1 As Integer
        Dim Lng1 As Long
        Dim Dbl1 As Double
        Dim Bln1 As Boolean
        Dim Vv1 As Variant
    
        '//Le variabili di Connessione Al db.
        Dim sSql As String                                              '//Stringa sql di estrazione


        '//Contatori                          '//Contatore Integer
        Dim iCount As Integer
        Dim dbl_count As Double                     '//Contatore Double


    '//*** Fine ***
    '//LE_VARIABILI_COMUNI
    '//........................................................
    
    '//LE_VARIABILI_PER_LA_RICERCA_STRINGA
    '//........................................................
    '//Codice :DimRicercaStringa.01
    '//Note   :Le variabili per la ricerca della stringa e
    '//       la sua gestione.
    
    
        Dim SearchString  As String                 '//Stringa da ricerca
        Dim SearchChar As String                    '//Ricerca il carattere
        Dim MyPos As Integer                        '//La posizione.
        Dim MyLen As Integer                                        '//La lughezza della stringa
        Dim sStringaIniz As String                                  '//Stringa fino all'//apostrofo
        Dim MyLenIniz As Integer                                    '//La lughezza della stringa Iniziale
        Dim sStringaFin As String                                   '//Stringa finale senza apostrofo
        Dim MyLenFin As Integer                                     '//La lughezza della stringa Iniziale
        Dim MyLenDiff As Integer                                    '//La lughezza rimanente tra (Stringa Iniziale - Stringa Finale = Diff)
        Dim sStringaRicostr As String                               '//Stringa ricostruita
    
    '//***Fine***
    '//LE_VARIABILI_PER_LA_RICERCA_STRINGA
    '//........................................................
        
    
    '//LE_VARIABILI_PROCEDURE_ERRORE
    '//........................................................
    '//Codice :DimProcedureErrore.01
    '//Note   :Le variabili per la gestione degli erori e dei
    '//       messaggi della procedura.
    
        Dim sxProceduraMessaggioErrore As String            '//Messaggio dei errore della procedura
        Dim sxProceduraAttivaEseguita  As String            '//Ultima procedura eseguita nell'//errore
    
    
    '//***Fine***
    '//LE_VARIABILI_PROCEDURE_ERRORE
    '//........................................................
    
  '//*** Fine ***
  '//LE_VARIABILI_GENERALI
  '//==============================================================================

'//*** Fine ***
'//VARIABILI_DATABASE
'//###########################################################################################


'//==============================================================================================================================
'//(pFunct02)
'//NOME         :(AGGIORNO CAMPO MSG01 DELLA TABELLA CONDOMINIO PARAMETRI GENERALI)
'//TIPO         :Public Function
'//Codice       :Aggiorna_DENOMINAZIONE.01
'//Parametri    :(par_sxMsg01)
'//Attività     :Aggiorna il campo della tabella.
'//Codice       :Aggiorna_DENOMINAZIONE.01
'//==============================================================================================================================
Public Function Aggiorna_DENOMINAZIONE() As String

On Error GoTo Err_Aggiorna_DENOMINAZIONE

'//Reset
'//.............................................................................
    '//Messaggio dei errore della procedura
    sxProceduraMessaggioErrore = ""
    '//Ultima procedura eseguita nell'errore
    sxProceduraAttivaEseguita = ""
'//.............................................................................
   
   
'//Attività     :Aggiorna il campo della tabella.
'//.............................................................................
    'Aggiorna_DENOMINAZIONE(Str1)
'//.............................................................................
        
    '//A)(SALVO IL MESSAGGIO NEL CAMPO MSG_01)
    '//______________________________________________________________________________
    '//Nota       :
    '//00)(Attività eseguite)
    '//01)(APRO IL RS E SALVO NEL CAMPO)
                
        '//Imposto le variabili
        sxProceduraMessaggioErrore = "Errore nella messaggio"
        '//Ultima procedura eseguita nell'errore
        sxProceduraAttivaEseguita = "Aggiorna_DENOMINAZIONE"
                    
        '//1)(APRO IL RS E SALVO NEL CAMPO)
        '//.............................................................................
        '//Note         :Apro il rs nella tabella <<Tb02_CdmParamGen_Tmp>> ed aggiorno
        '//             il campo messaggio.
        '//Codice       :Aggiorna_DENOMINAZIONE.02
            
            Set DaoRs = CurrentDb.OpenRecordset("CARDIF_N02-01_ESTRAZIONE_BASE_SOLO_DENOMINAZIONE")
            
            '//01.2)(RECUPERO CODICE)
            '//----------------------------------------------------------
            '//Note         :
            '//Codice       :Aggiorna_DENOMINAZIONE.02.save
                
             
                

Dim A_xs         As String
Dim B_xs         As String
Dim C_xs         As String
Dim D_xs         As String
Dim E_xs         As String
Dim F_xs         As String
Dim G_xs         As String
Dim H_xs         As String
Dim I_xs         As String
Dim J_xs         As String
Dim K_xs         As String
Dim L_xs         As String
Dim M_xs         As String
Dim ID_A_Lng
Dim ID_B_Lng
Dim ID_C_Lng
Dim ID_D_Lng
Dim ID_E_Lng
Dim ID_F_Lng
Dim ID_G_Lng
Dim ID_H_Lng
Dim ID_I_Lng
Dim ID_J_Lng
Dim ID_K_Lng
Dim ID_L_Lng
Dim ID_M_Lng

'//NORMALIZZA IMPORTI
Dim IMP_A_Lng
Dim IMP_B_Lng
Dim IMP_C_Lng
Dim IMP_D_Lng
Dim IMP_E_Lng
Dim IMP_F_Lng
Dim IMP_G_Lng
Dim IMP_H_Lng
Dim IMP_I_Lng
Dim IMP_J_Lng
Dim IMP_K_Lng
Dim IMP_L_Lng
Dim IMP_M_Lng
      



Dim DENOMINAZIONE_s     As String
Dim TFR_NRO_MECC_s          As String



'//reset campi importo
DoCmd.OpenQuery ("Resetta_N01-01_Campi_IMPORTO")


                '//Se il rs popolato aggiorno il campo
                If DaoRs.EOF = False And DaoRs.BOF = False Then
                    DaoRs.MoveFirst
                           
                           Do
                            
                            '//reset ID RECORD
                            ID_A_Lng = 0
                            ID_B_Lng = 0
                            ID_C_Lng = 0
                            ID_D_Lng = 0
                            ID_E_Lng = 0
                            ID_F_Lng = 0
                            ID_G_Lng = 0
                            ID_H_Lng = 0
                            ID_I_Lng = 0
                            ID_J_Lng = 0
                            ID_K_Lng = 0
                            ID_L_Lng = 0
                            ID_M_Lng = 0


                            '//RESET TESTO RECORD
                            A_xs = ""
                            B_xs = ""
                            C_xs = ""
                            D_xs = ""
                            E_xs = ""
                            F_xs = ""
                            G_xs = ""
                            H_xs = ""
                            I_xs = ""
                            J_xs = ""
                            K_xs = ""
                            L_xs = ""
                            M_xs = ""
                            
                            '//NORMALIZZA IMPORTI RECORD
                            IMP_A_Lng = 0
                            IMP_B_Lng = 0
                            IMP_C_Lng = 0
                            IMP_D_Lng = 0
                            IMP_E_Lng = 0
                            IMP_F_Lng = 0
                            IMP_G_Lng = 0
                            IMP_H_Lng = 0
                            IMP_I_Lng = 0
                            IMP_J_Lng = 0
                            IMP_K_Lng = 0
                            IMP_L_Lng = 0
                            IMP_M_Lng = 0

                            
                            '//A= DENOMINAZIONE
                            '//---------------------------------------------------------
                                
                                A_xs = Left(DaoRs.Fields("TFR_COD"), 2)
                                
                                '//Importo campo A
                                
                                    IMP_A_Lng = CVar(Right(DaoRs.Fields("TFR_COD"), 14))
                                    If IMP_A_Lng > 0 Then
                                        IMP_A_Lng = CVar(Right(DaoRs.Fields("TFR_COD"), 14))
                                    Else
                                        IMP_A_Lng = 0
                                    End If
                                    
                              
                               
                            
                                If Left(DaoRs.Fields("TFR_COD"), 2) = A_xs Then
                                    ID_A_Lng = DaoRs.Fields("ID")
                                    DENOMINAZIONE_s = DaoRs.Fields("DENOMINAZIONE")
                                    
                                    '//controllo apostrofo
                                    Str1 = vControlloApostrofo_pFunct(DENOMINAZIONE_s)
                                                                       
                                    DENOMINAZIONE_s = Str1
                                                                       
                                    '//Aggiorna DENOMINAZIONE LETTERA A
                                    sSql = ""
                                    sSql = sSql & "UPDATE CARDIF_TB01_Archivio SET CARDIF_TB01_Archivio.COD_DENOMINAZIONE = '" & DENOMINAZIONE_s & "'"
                                    sSql = sSql & "WHERE (((CARDIF_TB01_Archivio.ID)=" & ID_A_Lng & "));"
                                    
                                    
                                    '//controllo ed esecuzione
                                    Debug.Print sSql
                                    CurrentDb.Execute sSql
                                    
                                    
                                End If
                                '//aumento il record di 1
                                DaoRs.MoveNext
                            '//---------------------------------------------------------
                            
                            
                            '//B IL TFR
                            '//---------------------------------------------------------
                                
                                B_xs = Left(DaoRs.Fields("TFR_COD"), 2)
                            
                                If Left(DaoRs.Fields("TFR_COD"), 2) = B_xs Then
                                    ID_B_Lng = DaoRs.Fields("ID")
                                
                                     '//Aggiorna DENOMINAZIONE LETTERA B
                                    sSql = ""
                                    sSql = sSql & "UPDATE CARDIF_TB01_Archivio SET CARDIF_TB01_Archivio.COD_DENOMINAZIONE = '" & DENOMINAZIONE_s & "'"
                                    sSql = sSql & "WHERE (((CARDIF_TB01_Archivio.ID)=" & ID_B_Lng & "));"
                                    
                                
                                    '//controllo ed esecuzione
                                    Debug.Print sSql
                                    CurrentDb.Execute sSql
                                    
                                
                                End If
                                
                                            '//B IL NRO MECCANOGRAFICO DEL CAMPO DENOMINAZIONE
                                            '//..............................................................
                                            '//Il controllo viene eseguito sullo stesso record della _
                                               procedura precedente relativa alla denominazione.
                                            
                                                
                                                '//CONTROLLO se ho individuato la lettera B nel campo _
                                                   TFR_COD
                                                B_xs = Left(DaoRs.Fields("TFR_COD"), 2)
                                                
                                                '//controllo il campo precedente DENOMINAZIONE al record _
                                                   corrispondente lettera "B"
                                                If Left(DaoRs.Fields("TFR_COD"), 2) = B_xs Then
                                                    '//Salvo nella variabile solo i 6 caratteri dalla posizione due (Mid) _
                                                       senza spazi finali o iniziali. Il campo è significativo per 5 caratteri, _
                                                       ne prelevo 6 nel caso in cui la matricola sia più lunga. La funzione Trim _
                                                       in ogni caso elimina gli spazi.
                                                    TFR_NRO_MECC_s = Trim(Mid(DaoRs.Fields("DENOMINAZIONE"), 2, 7))
                                                    
                                                    Debug.Print Trim(Mid(DaoRs.Fields("DENOMINAZIONE"), 1, 25))
                                                    
                                                    ID_B_Lng = DaoRs.Fields("ID")
                                                
                                                     '//Aggiorna DENOMINAZIONE LETTERA B
                                                    sSql = ""
                                                    sSql = sSql & "UPDATE CARDIF_TB01_Archivio SET CARDIF_TB01_Archivio.TFR_NRO_MECC = '" & TFR_NRO_MECC_s & "'"
                                                    sSql = sSql & "WHERE (((CARDIF_TB01_Archivio.ID)=" & ID_B_Lng & "));"
                                                    
                                                    '//controllo ed esecuzione
                                                    Debug.Print sSql
                                                    CurrentDb.Execute sSql
                                                                                                   
                                                End If
                                            '//B IL NRO MECCANOGRAFICO DEL CAMPO DENOMINAZIONE
                                            '//..............................................................

                                                
                                '//aumento il record di 1
                                DaoRs.MoveNext
                            
                            
                            
                            '//C
                            '//---------------------------------------------------------
                                
                                C_xs = Left(DaoRs.Fields("TFR_COD"), 2)
                            
                                If Left(DaoRs.Fields("TFR_COD"), 2) = C_xs Then
                                    ID_C_Lng = DaoRs.Fields("ID")
                                
                                     '//Aggiorna DENOMINAZIONE LETTERA C
                                    sSql = ""
                                    sSql = sSql & "UPDATE CARDIF_TB01_Archivio SET CARDIF_TB01_Archivio.COD_DENOMINAZIONE = '" & DENOMINAZIONE_s & "'"
                                    sSql = sSql & "WHERE (((CARDIF_TB01_Archivio.ID)=" & ID_C_Lng & "));"
                                    
                                
                                    '//controllo ed esecuzione
                                    Debug.Print sSql
                                    CurrentDb.Execute sSql
                                    
                                
                                End If
                                '//aumento il record di 1
                                DaoRs.MoveNext
                            '//---------------------------------------------------------
                            
                            
                            
                             
                            '//D
                            '//---------------------------------------------------------
                                
                                D_xs = Left(DaoRs.Fields("TFR_COD"), 2)
                            
                                If Left(DaoRs.Fields("TFR_COD"), 2) = D_xs Then
                                    ID_D_Lng = DaoRs.Fields("ID")
                                
                                     '//Aggiorna DENOMINAZIONE LETTERA D
                                    sSql = ""
                                    sSql = sSql & "UPDATE CARDIF_TB01_Archivio SET CARDIF_TB01_Archivio.COD_DENOMINAZIONE = '" & DENOMINAZIONE_s & "'"
                                    sSql = sSql & "WHERE (((CARDIF_TB01_Archivio.ID)=" & ID_D_Lng & "));"
                                    
                                
                                    '//controllo ed esecuzione
                                    Debug.Print sSql
                                    CurrentDb.Execute sSql
                                    
                                
                                End If
                                '//aumento il record di 1
                                DaoRs.MoveNext
                            '//---------------------------------------------------------
                            
                            
                            '//E
                            '//---------------------------------------------------------
                                
                                E_xs = Left(DaoRs.Fields("TFR_COD"), 2)
                            
                                If Left(DaoRs.Fields("TFR_COD"), 2) = E_xs Then
                                    ID_E_Lng = DaoRs.Fields("ID")
                                
                                     '//Aggiorna DENOMINAZIONE LETTERA E
                                    sSql = ""
                                    sSql = sSql & "UPDATE CARDIF_TB01_Archivio SET CARDIF_TB01_Archivio.COD_DENOMINAZIONE = '" & DENOMINAZIONE_s & "'"
                                    sSql = sSql & "WHERE (((CARDIF_TB01_Archivio.ID)=" & ID_E_Lng & "));"
                                    
                                
                                    '//controllo ed esecuzione
                                    Debug.Print sSql
                                    CurrentDb.Execute sSql
                                    
                                
                                End If
                                '//aumento il record di 1
                                DaoRs.MoveNext
                            '//---------------------------------------------------------
                            
                                                        
                            '//F
                            '//---------------------------------------------------------
                                
                                F_xs = Left(DaoRs.Fields("TFR_COD"), 2)
                            
                                If Left(DaoRs.Fields("TFR_COD"), 2) = F_xs Then
                                    ID_F_Lng = DaoRs.Fields("ID")
                                
                                     '//Aggiorna DENOMINAZIONE LETTERA F
                                    sSql = ""
                                    sSql = sSql & "UPDATE CARDIF_TB01_Archivio SET CARDIF_TB01_Archivio.COD_DENOMINAZIONE = '" & DENOMINAZIONE_s & "'"
                                    sSql = sSql & "WHERE (((CARDIF_TB01_Archivio.ID)=" & ID_F_Lng & "));"
                                    
                                    '//controllo ed esecuzione
                                    Debug.Print sSql
                                    CurrentDb.Execute sSql
                                    
                                    
                                
                                End If
                                '//aumento il record di 1
                                DaoRs.MoveNext
                            '//---------------------------------------------------------
                            
                            
                            
                            '//G
                            '//---------------------------------------------------------
                                
                                G_xs = Left(DaoRs.Fields("TFR_COD"), 2)
                            
                                If Left(DaoRs.Fields("TFR_COD"), 2) = G_xs Then
                                    ID_G_Lng = DaoRs.Fields("ID")
                                
                                     '//Aggiorna DENOMINAZIONE LETTERA G
                                    sSql = ""
                                    sSql = sSql & "UPDATE CARDIF_TB01_Archivio SET CARDIF_TB01_Archivio.COD_DENOMINAZIONE = '" & DENOMINAZIONE_s & "'"
                                    sSql = sSql & "WHERE (((CARDIF_TB01_Archivio.ID)=" & ID_G_Lng & "));"
                                    
                                    '//controllo ed esecuzione
                                    Debug.Print sSql
                                    CurrentDb.Execute sSql
                                    
                                    
                                
                                End If
                                '//aumento il record di 1
                                DaoRs.MoveNext
                            '//---------------------------------------------------------
                            
                            
                            '//H
                            '//---------------------------------------------------------
                                
                                H_xs = Left(DaoRs.Fields("TFR_COD"), 2)
                            
                                If Left(DaoRs.Fields("TFR_COD"), 2) = H_xs Then
                                    ID_H_Lng = DaoRs.Fields("ID")
                                
                                     '//Aggiorna DENOMINAZIONE LETTERA H
                                    sSql = ""
                                    sSql = sSql & "UPDATE CARDIF_TB01_Archivio SET CARDIF_TB01_Archivio.COD_DENOMINAZIONE = '" & DENOMINAZIONE_s & "'"
                                    sSql = sSql & "WHERE (((CARDIF_TB01_Archivio.ID)=" & ID_H_Lng & "));"
                                    
                                
                                    '//controllo ed esecuzione
                                    Debug.Print sSql
                                    CurrentDb.Execute sSql
                                    
                                
                                End If
                                '//aumento il record di 1
                                DaoRs.MoveNext
                            '//---------------------------------------------------------
                            
                            
                              
                            '//I
                            '//---------------------------------------------------------
                                
                                I_xs = Left(DaoRs.Fields("TFR_COD"), 2)
                            
                                If Left(DaoRs.Fields("TFR_COD"), 2) = I_xs Then
                                    ID_I_Lng = DaoRs.Fields("ID")
                                
                                     '//Aggiorna DENOMINAZIONE LETTERA I
                                    sSql = ""
                                    sSql = sSql & "UPDATE CARDIF_TB01_Archivio SET CARDIF_TB01_Archivio.COD_DENOMINAZIONE = '" & DENOMINAZIONE_s & "'"
                                    sSql = sSql & "WHERE (((CARDIF_TB01_Archivio.ID)=" & ID_I_Lng & "));"
                                    
                                    '//controllo ed esecuzione
                                    Debug.Print sSql
                                    CurrentDb.Execute sSql
                                    
                                    
                                
                                End If
                                '//aumento il record di 1
                                DaoRs.MoveNext
                            '//---------------------------------------------------------
                            
                            
                            '//J
                            '//---------------------------------------------------------
                                
                                J_xs = Left(DaoRs.Fields("TFR_COD"), 2)
                            
                                If Left(DaoRs.Fields("TFR_COD"), 2) = J_xs Then
                                    ID_J_Lng = DaoRs.Fields("ID")
                                
                                     '//Aggiorna DENOMINAZIONE LETTERA J
                                    sSql = ""
                                    sSql = sSql & "UPDATE CARDIF_TB01_Archivio SET CARDIF_TB01_Archivio.COD_DENOMINAZIONE = '" & DENOMINAZIONE_s & "'"
                                    sSql = sSql & "WHERE (((CARDIF_TB01_Archivio.ID)=" & ID_J_Lng & "));"
                                    
                                    '//controllo ed esecuzione
                                    Debug.Print sSql
                                    CurrentDb.Execute sSql
                                    
                                    
                                
                                End If
                                '//aumento il record di 1
                                DaoRs.MoveNext
                            '//---------------------------------------------------------
                            
                            
                            
                                
                            '//K
                            '//---------------------------------------------------------
                                
                                K_xs = Left(DaoRs.Fields("TFR_COD"), 2)
                            
                                If Left(DaoRs.Fields("TFR_COD"), 2) = K_xs Then
                                    ID_K_Lng = DaoRs.Fields("ID")
                                
                                     '//Aggiorna DENOMINAZIONE LETTERA K
                                    sSql = ""
                                    sSql = sSql & "UPDATE CARDIF_TB01_Archivio SET CARDIF_TB01_Archivio.COD_DENOMINAZIONE = '" & DENOMINAZIONE_s & "'"
                                    sSql = sSql & "WHERE (((CARDIF_TB01_Archivio.ID)=" & ID_K_Lng & "));"
                                    
                                    '//controllo ed esecuzione
                                    Debug.Print sSql
                                    CurrentDb.Execute sSql
                                    
                                    
                                
                                End If
                                '//aumento il record di 1
                                DaoRs.MoveNext
                            '//---------------------------------------------------------
                            
                            '//L
                            '//---------------------------------------------------------
                                
                                L_xs = Left(DaoRs.Fields("TFR_COD"), 2)
                            
                                If Left(DaoRs.Fields("TFR_COD"), 2) = L_xs Then
                                    ID_L_Lng = DaoRs.Fields("ID")
                                
                                     '//Aggiorna DENOMINAZIONE LETTERA K
                                    sSql = ""
                                    sSql = sSql & "UPDATE CARDIF_TB01_Archivio SET CARDIF_TB01_Archivio.COD_DENOMINAZIONE = '" & DENOMINAZIONE_s & "'"
                                    sSql = sSql & "WHERE (((CARDIF_TB01_Archivio.ID)=" & ID_L_Lng & "));"
                                    
                                    '//controllo ed esecuzione
                                    Debug.Print sSql
                                    CurrentDb.Execute sSql
                                    
                                    
                                
                                End If
                                '//aumento il record di 1
                                DaoRs.MoveNext
                            '//---------------------------------------------------------
                            
                            
                            
                            '//M
                            '//---------------------------------------------------------
                                
                                M_xs = Left(DaoRs.Fields("TFR_COD"), 2)
                            
                                If Left(DaoRs.Fields("TFR_COD"), 2) = M_xs Then
                                    ID_M_Lng = DaoRs.Fields("ID")
                                
                                     '//Aggiorna DENOMINAZIONE LETTERA K
                                    sSql = ""
                                    sSql = sSql & "UPDATE CARDIF_TB01_Archivio SET CARDIF_TB01_Archivio.COD_DENOMINAZIONE = '" & DENOMINAZIONE_s & "'"
                                    sSql = sSql & "WHERE (((CARDIF_TB01_Archivio.ID)=" & ID_M_Lng & "));"
                                    
                                    '//controllo ed esecuzione
                                    Debug.Print sSql
                                    CurrentDb.Execute sSql
                                    
                                End If
                                '//aumento il record di 1
                                DaoRs.MoveNext
                            '//---------------------------------------------------------
                                                       
                                                    
                            
                          
                            
                           
                            Loop While Not DaoRs.EOF
                            

                            
                End If
                
            '//----------------------------------------------------------
            
            '//chiudo e resetto il rs
            DaoRs.Close
            Set DaoRs = Nothing
        '//*** FINE ***
        '//01)(APRO IL RS E SALVO NEL CAMPO)
        '//.............................................................................
    
    
                'AGGIORNO TUTTI I CAMPI IMPORTO
                DoCmd.OpenQuery ("CARDIF_N10-01_AggiornaCampo->TFR_COD_IMP")
                DoCmd.OpenQuery ("CARDIF_N10-02_AggiornaCampo->TFR_RIV_IMP")
                DoCmd.OpenQuery ("CARDIF_N10-03_AggiornaCampo->TFR_ANNO_IMP")
                DoCmd.OpenQuery ("CARDIF_N10-04_AggiornaCampo->TFR_ANTIC_IMP")
                DoCmd.OpenQuery ("CARDIF_N10-05_AggiornaCampo->TFR_ESERC_IMP")
                DoCmd.OpenQuery ("CARDIF_N10-06_AggiornaCampo->TFR_QUOTA_LIQ_IMP")
                DoCmd.OpenQuery ("CARDIF_N10-07_AggiornaCampo->TFR_QUOTA_FIN_IMP")
                DoCmd.OpenQuery ("CARDIF_N10-08_AggiornaCampo->ID_LETT")
                DoCmd.OpenQuery ("CARDIF_N10-10_AggiornaCampo->TFR_NRO_MECC") 'aggiorno i campo NRO MECC (MATRICOLA)
                '
                
                MsgBox "Fine Aggiornamento"
                
                
                
    
    
    '//*** FINE ***
    '//A)(SALVO IL MESSAGGIO NEL CAMPO MSG_01)
    '//______________________________________________________________________________


'//USCITA  E GESTIONE ERRORI
'//..............................................................................................................

Exit_Aggiorna_DENOMINAZIONE:
    Exit Function

Err_Aggiorna_DENOMINAZIONE:
    MsgBox Err.Description & " ---> " & sxProceduraMessaggioErrore & " ---> " & sxProceduraAttivaEseguita
    Vv1 = Err.Description & " ---> " & sxProceduraMessaggioErrore & " ---> " & sxProceduraAttivaEseguita
    Debug.Print Vv1
    Resume Exit_Aggiorna_DENOMINAZIONE

End Function

'//*** FINE ***
'//(pFunct02)
'//==============================================================================================================================


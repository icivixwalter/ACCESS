Attribute VB_Name = "UTIL_MDL40_N06_CONTROLLO_OGGETTI_DB_TABLES_CONTROLLO_ATTRIBUTI"
Option Compare Database


'//https://docs.microsoft.com/it-it/office/troubleshoot/access/tabledef-attributes-usage
'Come utilizzare la proprietà Attributes per gli oggetti TableDef in Access

'Attributi TableDef _
La Attributes proprietà di un TableDef oggetto specifica le caratteristiche della tabella _
rappresentata dall' TableDef oggetto. La Attributes proprietà viene archiviata come un singolo intero lungo _
e corrisponde alla somma delle costanti lunghe seguenti:


'Costante                   Descrizione
'dbAttachExclusive          Per i database che utilizzano il modulo di _
                            gestione di database Microsoft Jet, indica che la tabella è una tabella collegata aperta per l'utilizzo esclusivo.
                
'dbAttachSavePWD            Per i database che utilizzano il modulo di _
                            gestione di database Jet, indica che l'ID utente e la password per la tabella collegata devono essere salvati con le informazioni sulla connessione.
'dbSystemObject             Indica che la tabella è una tabella di sistema.

'dbHiddenObject             Indica che la tabella è una tabella nascosta (per uso temporaneo).
'dbAttachedTable            Indica che la tabella è una tabella collegata da un database _
                            ODBC (non-Open Database Connectivity), ad esempio Microsoft Access o Paradox.
                
'dbAttachedODBC             Indica che la tabella è una tabella collegata da un database _
                            ODBC, ad esempio Microsoft SQL Server o ORACLE Server.
                
'Per un TableDef oggetto, l'utilizzo della Attributes proprietà dipende dallo stato di TableDef _
, come illustrato nella tabella seguente:

'TableDef                           Usage
'Oggetto non accodato all'insieme   Lettura/scrittura
'Tabella di base                    Sola lettura
'Tabella collegata                  Sola lettura


'Quando si controlla l'impostazione di questa proprietà, è possibile utilizzare l'operatore AND per eseguire _
il test di un attributo specifico. Ad esempio, per determinare se un oggetto Table è una tabella di sistema, _
eseguire un confronto logico tra la TableDef proprietà Attributes e la dbSystemObject costante. _
ESEMPIO  Attrib = (T.Attributes And dbSystemObject)



'//LE COSTANTI ACCESS
'https://docs.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/tabledefattributeenum-enumeration-dao

'TableDefAttributeEnum enumeration(DAO)
'Used with the Attributes property to determine attributes of a TableDef object.

'Name                   Value       Description

'dbAttachedODBC         536870912       Linked ODBC database table.

'dbAttachedTable        1073741824      Linked non-ODBC database table.

'dbAttachExclusive      65536           Opens a linked Microsoft Access database engine table for exclusive use.

'dbAttachSavePWD        131072          Saves user ID and password for linked remote table.

'dbHiddenObject                 1       Hidden table (for temporary use).

'dbSystemObject      -2147483646        System table.








'//****************************************************************************************************************************
'//                    CONTROLLO ATTRIBUTI DEGLI OGGETTI TABELLA
'//
'//****************************************************************************************************************************
'//NOTE     : Come utilizzare la proprietà Attributes per gli oggetti TableDef in Access _
                https://docs.microsoft.com/it-it/office/troubleshoot/access/tabledef-attributes-usage

'//         La seguente funzione di esempio definita dall'utente esegue un ciclo tra _
            tutte le tabelle di un database e visualizza una finestra di messaggio che elenca ogni nome _
            di tabella e indica se la tabella è o meno una tabella di sistema _
            È possibile utilizzare la Attributes proprietà di un TableDef oggetto per determinare proprietà _
            specifiche della tabella. Ad esempio, è possibile utilizzare la Attributes proprietà per individuare _
            se una tabella è una tabella di sistema o una tabella collegata (allegata).


'//CODICE   : ShowTableAttribs.attributiTabelle

Function ShowTableAttribs()
   
   
On Error GoTo ShowTableAttribs_Err
   
   Dim DB As DAO.Database
   Dim T As DAO.TableDef
   Dim TType As String
   Dim TName As String
   Dim Attrib As String
   Dim I As Integer

Set DB = CurrentDb()
    
    Debug.Print "                  INIZIO CONTROLLO ATTRIBUTI DELLA TABELLA"
    Debug.Print "=================================================================================="
    Debug.Print
    
    '//CICLO OGGETTI TABLE
    For I = 0 To DB.TableDefs.Count - 1
          
          '// DO EVENTS EVENTI WINDOS DA BLOCCARE
          DoEvents
          
          Set T = DB.TableDefs(I)
          TName = T.Name
          Attrib = (T.Attributes And dbSystemObject)
          
          '//MESSAGGIO PER OGNI TABELLA _
             ---- SOSPESO ---
          MsgBox TName & IIf(Attrib, ": System Table - TABELLA DI SISTEMA", ": Not System NON E' UNA TABELLA DI SISTEMA" & "Table")
          
            '//---------------------------------------------------------//
                
                Debug.Print "TABELLA            -------------------------> " & T.Name
                Debug.Print "controllo attributi della tabella ----------> " & T.Attributes
                
                '//CONTROLLO
                If T.Attributes = 0 Then
                 Debug.Print "Tabella normale  con attributo            =  " & T.Attributes
                 
                 ElseIf T.Attributes = 2 Then
                    Debug.Print "Tabella DI SISTEMA  con attributo      =  " & T.Attributes
                 
                End If
                
                Debug.Print "valore di dbSystemObject          ----------> " & dbSystemObject
                
                
                
                Debug.Print "valore di                  Attrib ----------> " & Attrib
                
            '//---------------------------------------------------------//
          
    Next I

Debug.Print "                   *** FINE CONTROLLO ****"
Debug.Print "=================================================================================="

'EXIT E GESTIONE ERRORI
'-----------------------------------------------------------------------------------------------

ShowTableAttribs_Exit:
    Exit Function

ShowTableAttribs_Err:
    MsgBox Error$
    Resume ShowTableAttribs_Exit

End Function

'//*** FINE ***
'//****************************************************************************************************************************
'//                    ATTRIBUTI DEGLI OGGETTI TABELLA
'//
'//****************************************************************************************************************************




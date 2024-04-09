Attribute VB_Name = "UTILITA_Mdl30_N01_Crea_File_schema_ini"

'CREAZIONE DI UN FILE SCHEMA
'************************************************************************************
    'In questo articolo viene illustrato come scrivere una routine che crea un file
    'schema.ini basato su una tabella esistente nel database.
    'Si suppone che l'utente conosca Visual Basic Applications Edition e sappia
    'creare applicazioni Microsoft Access avvalendosi degli strumenti di programmazione
    'forniti con Microsoft Access. Per ulteriori informazioni su Visual Basic,
    'Applications Edition, fare riferimento alla versione del manuale
    '"Building Applications con Microsoft Access".
    
    'In Microsoft Access 7.0 e Microsoft Access 97, è possibile collegare o aprire file di testo
    'delimitati e a lunghezza fissa. Verrà è in grado di leggere un file di testo
    'direttamente oppure è possibile utilizzare un file di informazioni
    'denominato schema.ini per determinare le caratteristiche del file di testo,
    'ad esempio i nomi di colonna, lunghezze di campo e i tipi di dati.
    'Un file schema.ini è richiesto quando si collega o aprire file di testo a lunghezza fissa;
    'è facoltativo per i file di testo delimitato. Il file schema.ini deve risiedere nella stessa
    'cartella come vengono descritti i file di testo.

 '//ESEMPIO E SPECIFICHE : Prepariamo il file di testo da  salvare su disco C; chiamato _
    Citta.txt e costituito da questa struttura : _
    TORINO;98762 _
    MILANO;123456 _
    ROMA;876 _
    e salvato nel file CITTA.TXT. Poi facciamo un collegamento di prova con il db _
    access ed il file dovrebbe essere riconosciuto come file da due campi separato _
    da virgole. L'importazione o il collegamento riesce perfettamente se utilizziamo _
    come separatore di testo il punto e virgola (;) in quanto si lavora con la _
    macchina windows italiana che riconosce come separatore il ;. Mentre se il sistema _
    americano, occorre utilizzare la virgola (,) per la macchina Windows versione USA. _
    Infatti l'estensione .csv nasce dalle iniziale dell'espressione "comma separated values" _
    ma non si tratta di comma ma di separator, e quindi i csv risentono delle impostazioni _
    internazioni in cui i file .csv viene descritto come "delimitato dal separatore di elenco" _
    e cioè ; in italia e , negli USA.
    
'//COLLEGAMENTO DI FILE DI TESTO MEDIANTE PROGRAMMA = il collegamento di un file di testo con _
   può essere effettuato manualmente oppure con la proprieta Connect dell'Oggetto DAO TableDef, _
   esattamente come quando si collegano tabelle di altri sistemi di database. Nell'esempio di seguito _
   basta modificare i riferimento dei file interessati ed utilizzare la routine con una tabella _
   dBASE III. Vedi Routine COLLEGA_FILE_TESTO_DELIMITATO ()
   



'//PARAMETRI DI COLLEGAMENTO
'//formato File =  Format = Delimited (*) -> i campi sono delimitati da _
                                     asterischi. Invece dell'asterisco si può usare _
                                     qualunque altro carattere, eccetto le virgolette _
                                     doppie. _
                    Format = CsvDelimited -> i campi sono delimitati dal carattere separatore. _
                    Format = TabDelimited -> i campi sono delimitati dal carattere tabulazione.
                            
'************************************************************************************


Option Compare Database



Private Sub chiama_funzione()
Dim bIncFldNames As Boolean
Dim sPath As String
Dim sSectionName As String
Dim sTblQryName As String
    
    '//parametro campi inclusi Si/no = ColNameHeader = False/True
    bIncFldNames = False
    '//nome file .ini che sara crato
    sSectionName = "SezioneProva"
    sPath = "c:\CASA\LINGUAGGI\ACCESS\ACCESS_FILE_INI\"
    
    '//la tabella/query dove saranno presi i dati ed esaminati per creare il file .ini
    sTblQryName = "Fatture_Tb01_Emesse"

    Call CreateSchemaFile(bIncFldNames, sPath, sSectionName, sTblQryName)

End Sub



  Public Function CreateSchemaFile(bIncFldNames As Boolean, _
                                       sPath As String, _
                                       sSectionName As String, _
                                       sTblQryName As String) As Boolean
         
 On Local Error GoTo CreateSchemaFile_Err
         
         Dim Msg As String ' For error handling.
         Dim ws As Workspace, db As Database
         Dim tblDef As TableDef, fldDef As Field
         Dim i As Integer, Handle As Integer
         Dim fldName As String, fldDataInfo As String
         ' -----------------------------------------------
         ' Set DAO objects.
         ' -----------------------------------------------
         Set db = CurrentDb()
         ' -----------------------------------------------
         ' Open schema file for append.
         ' -----------------------------------------------
         Handle = FreeFile
         Open sPath & "schema.ini" For Output Access Write As #Handle
         ' -----------------------------------------------
         ' Write schema header.
         ' -----------------------------------------------
         Print #Handle, "[" & sSectionName & "]"
         '//Porprieta ColNameHeader = campi prima riga True/False
         Print #Handle, "ColNameHeader = " & _
                         IIf(bIncFldNames, "True", "False")
         
         Print #Handle, "CharacterSet = ANSI"
         
         '//fomato campi
         Print #Handle, "Format = TabDelimited"
         ' -----------------------------------------------
         ' Get data concerning schema file.
         ' -----------------------------------------------
         Set tblDef = db.TableDefs(sTblQryName)
         With tblDef
            For i = 0 To .Fields.Count - 1
               Set fldDef = .Fields(i)
               With fldDef
                  fldName = .Name
                  Select Case .Type
                     Case dbBoolean
                        fldDataInfo = "Bit"
                     Case dbByte
                        fldDataInfo = "Byte"
                     Case dbInteger
                        fldDataInfo = "Short"
                     Case dbLong
                        fldDataInfo = "Integer"
                     Case dbCurrency
                        fldDataInfo = "Currency"
                     Case dbSingle
                        fldDataInfo = "Single"
                     Case dbDouble
                        fldDataInfo = "Double"
                     Case dbDate
                        fldDataInfo = "Date"
                     Case dbText
                        fldDataInfo = "Char Width " & Format$(.Size)
                                            
                     Case dbLongBinary
                        fldDataInfo = "OLE"
                     Case dbMemo
                        fldDataInfo = "LongChar"
                     Case dbGUID
                        fldDataInfo = "Char Width 16"
                  End Select
                  Print #Handle, "Col" & Format$(i + 1) _
                                  & "=" & fldName & Space$(1) _
                                  & fldDataInfo
               End With
            Next i
         End With
         MsgBox sPath & "SCHEMA.INI has been created."
         CreateSchemaFile = True
CreateSchemaFile_End:
         Close Handle
         Exit Function
CreateSchemaFile_Err:
         Msg = "Error #: " & Format$(Err.Number) & vbCrLf
         Msg = Msg & Err.Description
         MsgBox Msg
         Resume CreateSchemaFile_End
      End Function
                



'//ROUTINE DI COLLEGAMENTO CON DAO.CONNECT
'//file da collegare prodotti.txt si deve trovare _
   nella directory  c:\CASA\LINGUAGGI\ACCESS\ACCESS_FILE_INI\
Private Sub COLLEGA_FILE_TESTO_DELIMITATO()


    Dim Dbs As DAO.Database
    Dim Tbf As DAO.TableDef
    
    Set Dbs = CurrentDb
    Set Tbf = Dbs.CreateTableDef("CITTA_COLL")
    Tbf.Connect = "TEXT;DATABASE=c:\CASA\LINGUAGGI\ACCESS\ACCESS_FILE_INI\"
    Tbf.SourceTableName = "CITTA.txt"
    
    
    Dbs.TableDefs.Append Tbf
    
End Sub


'//Eliminazione tabella collegata
Private Sub EliminaTabella_COLLEGATA()

Dim Dbs As DAO.Database
Set Dbs = CurrentDb
Dbs.TableDefs.Delete ("CITTA_COLL")
End Sub







'// *** fine ***
'//ESEMPIO DI COLLEGAMENTO AD UN DATABASE ACCESS DIVERSO DA QUELLO CORRENTE
'//*************************************************************************************//




'// INIZIO
'//ESEMPIO DI COLLEGAMENTO AD UN DATABASE ACCESS CORRENTE CON ADOX
'//*************************************************************************************//

Private Sub CollegaTABELLA_DB_con_ADOX()

Dim cat As ADOX.Catalog         '//devi istanziare Microsoft ADO ext 2.8 for DDL and Security
Dim tbl As ADOX.Table

    '//istanza  un oggetto Catalog
    Set cat = New ADOX.Catalog
    cat.ActiveConnection = CurrentProject.Connection
    
    '//istanza un oggetto Table
    Set tbl = New ADOX.Table
    tbl.Name = "TB01_MODULI_COLL"
    Set tbl.ParentCatalog = cat
    
    '//imposta le proprieta Jet OLDEB del nuovo oggetto tAble
    tbl.Properties("Jet OLEDB:Create Link") = True
    tbl.Properties("Jet OLEDB:Link Datasource") = "c:\CASA\LINGUAGGI\ACCESS\ACCESS_FILE_INI\CONSULENTE_FITOSANITARI.mdb"
    tbl.Properties("Jet OLEDB:Link Provider String") = "MS Access; PWD=;"
    
    tbl.Properties("Jet OLEDB:Remote Table Name") = "TB01_MODULI"
    
    '//Accoda il nuovo oggetto tabella all'insieme delle tabelle
    cat.Tables.Append tbl
    
    



End Sub


'// *** fine ***
'//ESEMPIO DI COLLEGAMENTO AD UN DATABASE ACCESS CORRENTE CON ADOX
'//*************************************************************************************//




'//ESEMPIO DI COLLEGAMENTO AD UN DATABASE ACCESS DIVERSO DA QUELLO CORRENTE ??? DA FINIRE PERCH IL DB III sostituire con file access???s
'//*************************************************************************************//
'//ESEMPIO SU WEB
'//https://books.google.it/books?id=Z6HKDXI6ipMC&pg=PA562&lpg=PA562&dq=access+schema.ini+file&source=bl&ots=ROAftcwWdq&sig=dVZPpvJOuedmeXD_Ucb0KT5X03E&hl=it&sa=X&ved=0ahUKEwiW8Nzvh_7bAhWBxRQKHccgB3kQ6AEIVzAG#v=onepage&q=access%20schema.ini%20file&f=false
'//Per collegare una tabella ad un database Access diverso da quello corrente, occorre _
   usare il metodo OpenDatabase per variabile oggetto dbs al posto della funzione CurrentDb _
   che permette il collegamento al db corrente.
   
'//SUB DI COLLEGGA TABELLA AL DB NON APERTO

Public Sub CollegaTabella_DB_NonAperto()
On Error GoTo Err_CollegaTabella_DB_NonAperto

Dim Dbs As DAO.Database
Dim Tbf As DAO.TableDef

    '//Apre un database diverso da quello corrente, _
       che non viene visualizzato e vi collega una nuova tabella.
       
       Set Dbs = OpenDatabase("c:\CASA\LINGUAGGI\ACCESS\ACCESS_FILE_INI\Db.mdb")
       Set Tbf = Dbs.CreateTableDef("OperazioniBancarie")
       
       '//imposta la proprieta Connect e SourceTableNaper la tabella _
          si tratta di una dabella dBase III
          
          '//modificare per Access o Excel
          Tbf.Connect = "Dbase III;DATASE = C:\Esempi"
          Tbf.SourceTableName = "Bancacom"
          
          '//Accoda l'oggetto TableDef all'insieme Tabledefs del database.
          Dbs.TableDefs.Append Tbf
          
Exit_CollegaTabella_DB_NonAperto:
          On Error Resume Next
          Dbs.Close
          Set Dbs = noting
          Exit Sub
          
Err_CollegaTabella_DB_NonAperto:
          MsgBox "Errore : " & Err.Number & " - " & Err.Description
          Resume Exit_CollegaTabella_DB_NonAperto
          

'//errore
On Error GoTo Err_CollegaTabella_DB_NonAperto
End Sub





'//*************************************************************************************//


'//ACCESSO DIRETTO A UN FOGLIO DI LAVORO EXCEL
'//*************************************************************************************//


Public Sub ApreCartellaExcel()

'//Questa routine apre un recordse su glio di lavoro SocietaAlfa _
   in una cartemma Microsoft Excel 8.0 chiamata contoEconomico.xls _
   che si trova nella directory c:\CASA\LINGUAGGI\ACCESS\ACCESS_FILE_INI\. Successivamente conta il _
   numero dei record che si trovano nel recordset.
   
   
Dim Dbs As DAO.Database
Dim RS As DAO.Recordset

    '//Apre la cartella Excel come se fosse un database
    Set Dbs = OpenDatabase("c:casa\ContoEconomico.xls", _
              False, False, "Excel 8.0;HDR=No;")
              
    '//Crea un oggetto Reocrset per il foglio di lavoro SocietaAlfa
    Set rst = Dbs.OpenRecordset("SocietaAlfa$")
    
    '//Si porta sull'ultimo record del recordset e visualizza la _
       proprieta RecordCount.
       
       

End Sub


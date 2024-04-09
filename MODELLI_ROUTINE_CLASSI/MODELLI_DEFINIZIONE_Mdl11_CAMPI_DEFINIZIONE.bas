Attribute VB_Name = "MODELLI_DEFINIZIONE_Mdl11_CAMPI_DEFINIZIONE"
Option Compare Text
Option Explicit

'//                                         NOTE                                                                   //
'//****************************************************************************************************************//
'...........................................:......................................................................
'1.1.1              SEZIONE 01  - SEZIONE DI APPATENENZA
'
'CAMPO_3            :nome tabella                                                                   - NomeCampo
'CAMPO_4            :nome_tabella                                                                   - Nome campo
'
'_________________________________________________________________________________________________________________
'TIPO               :CAMPO CONTATORE
'_________________________________________________________________________________________________________________
'CAMPO_N            :DEFINIZIONECAMPO                                                               - <Nome campo>
'Proprieta_1        :Propriet� FieldSize (Dimensione campo) n.b. solo
'                                                        per i campi Memo e Long Binary             - <tipo>
'Propriet�_2        :NewValues (Nuovi valori)                                                       - <valori>
'Propriet�_3        :Format (Formato)                                                               - <formato>
'Propriet�_3        :Propriet� Caption (Etichetta)                                                  - <etichetta>
'Propriet�_4        :Propriet� Indexed (Indicizzato)                                                - <tipo indice>
'Generale_1         :Nessuna





'...........................................:......................................................................
'1.1.2              SEZIONE 00  - SEZIONE DI APPATENENZA
'
'CAMPO_3            :nome tabella                                                                - NomeCampo
'CAMPO_4            :nome_tabella                                                                - Nome campo
'
'_________________________________________________________________________________________________________________
'TIPO               :CAMPO TESTO
'_________________________________________________________________________________________________________________
'CAMPO_N            :ID                                                                         - <Nome campo>

'..................................... ELENCO PROPRIETA  .........................................................
'Proprieta_1        :Propriet� FieldSize (Dimensione campo) n.b. solo per i campi
'                                                                   Memo e Long Binary          - <tipo>
'Propriet�_2        :Propriet� Format (Formato)                                                 - <formato>
'Propriet�_3        :Propriet� InputMask (Maschera di input)                                    - <maschera di imput>
'Propriet�_4        :Propriet� Caption (Etichetta)                                              - <etichetta>
'Propriet�_5        :Propriet� DefaultValue (Valore predefinito)                                - <tipo valore>
'Propriet�_6        :Propriet� ValidationRule, ValidationText (Valido se, Messaggio errore)     - <valore predefinito>
'Propriet�_7        :Propriet� DefaultValue (Valore predefinito)                                - <valido se>
'Propriet�_8        :Propriet� DefaultValue (Valore predefinito)                                - <messaggio di errore>
'Propriet�_9        :Propriet� Required (Richiesto)                                             - <si/no>
'Propriet�_10       :Propriet� AllowZeroLength (Consenti lunghezza zero)                        - <si/no>
'Propriet�_11       :Propriet� Indexed (Indicizzato)                                            - <si/no>
'Propriet�_12       :Propriet� Compressione unicode (NON FUNZIONA)                              - <si/no>
'..................................... ELENCO RICERCA .............................................................
'Ricerca_1          :Propriet� DisplayControl (Visualizza controllo)                            - <tipo controllo>
'Ricerca_2          :Propriet� RowSourceType, RowSource (Origine riga, Tipo origine riga)       - <tipo origine>
'Ricerca_3          :Propriet� RowSourceType, RowSource (Origine riga, Tipo origine riga)       - <origine riga>
'Ricerca_4          :Propriet� BoundColumn (Colonna associata)                                  - <colonna>
'Ricerca_5          :Propriet� ColumnCount (Numero colonne)                                     - <nro col>
'Ricerca_6          :Propriet� ColumnHeads (Intestazioni colonne)                               - <intest col.>
'Ricerca_7          :Propriet� ColumnWidths (Larghezza colonne)                                 - <largh.>
'Ricerca_8          :Propriet� ListRows (Righe in elenco)                                       - <nro>
'Ricerca_9          :Propriet� ListWidth (Larghezza elenco)                                     - <nro>
'Ricerca_10         :Propriet� LimitToList (Solo in elenco)                                     - <si/no>
'..................................... PROPRIETA TABELLA ..........................................................
'Proprieta_1        :Nessuna                                                                    - <tipo controllo>

'//                                         NOTE    *** FINE ***                                                                   //
'//****************************************************************************************************************//











                                                '____________________________________________________________________________________________________________________________________________
                                                '
                                                '1.1.1              MODELLI : - TIPI DI CAMPO -
                                                '
                                                'CAMPO_1            :LONG                       :CONTATORE                  :PRIMARYKEY
                                                'CAMPO_1            :INTEGER                    :CONTATORE                  :PRIMARYKEY
                                                '
                                                'CAMPO_2            :SEZ1_GE_02_ID_ESERCIZIO                                :LONG - CONTATORE - ANNO DI ESERCIZIO - PRIMARYKEY  :dbText,1
                                                'CAMPO_3            :SEZ1_GE_02_ANNO_ESERCIZIO                              :INTEGER - ANNO DI ESERCIZIO - SecondaryKey -       :dbInteger,1
                                                'CAMPO_4            :SEZ1_GE_02_COD_COND                                    :TXT 10 - CODICE CONDOMINIO -
                                                'CAMPO_5            :SEZ1_GE_02_DATA_INIZIO_GESTIONE                        :DATA - INIZIO GESTIONE - DATA ESTESA
                                                'CAMPO_7            :SEZ1_GE_02_TOTALE_PREVENTIVO                           :LONG - TOTALE PREVENTIVO
                                                'CAMPO_8            :SEZ1_GE_02_TOTALE_CONSUNTIVO                           :LONG - TOTALE CONSUNTIVO
                                                '...........................................:.................................................................................................





                                                '____________________________________________________________________________________________________________________________________________
                                                'CAMPO_1            :CAMPO_LONG                                             :LONG -                               - PRIMARYKEY  :dbLONG
                                                'Propriet�_9        :dbAutoIncrField                                                                                            :<Si/no>
                                                'Propriet�_9        :Propriet� Required (Richiesto)                                                                             :<Si/no>
                                                'Propriet�_10       :Propriet� AllowZeroLength (Consenti lunghezza zero)                                                        :<Si/no>
                                                    
                                                    
                                                    'Creo il LONG AUTOINCREMENTANTE
                                                   'Set FieldNuovo = TableDefNuovo.CreateField("NOME_NN_01_ID_NNNNNNNNN", dbLong)
                                                    
                                                   'FieldNuovo.Attributes = dbAutoIncrField             '... imposta a contatore
                                                    'accodo il campo all'inSieme Fields
                                                   'TableDefNuovo.Fields.Append FieldNuovo
                                                    
                                                   
                                                    
                                                    
                                                    
                                                    '..........................................................
                                                    '   DEFINISCO I CAMPI DELL'INSIEME INDEX - CHIAVE PRIMARIA
                                                        
                                                       'Set idxNuovo_Contatore_Univoco = TableDefNuovo.CreateIndex("NOME_CHIAVE_PRIARIA_1")     '.... imposta Nome index
                                                                                                          
                                                            '   DEFINISCO L'INSiEME INDEX CON UN CAMPO A PRIMARY KEY
                                                            ' Quando imposto la proprieta del campo a chiave primaria
                                                            ' automaticamente le proprieta Index sono impostate nel seguente modo:
                                                            ' Il nome dell'indice viene attribuito con il metodo CreateIndex("Nome Indice"),
                                                            ' le PROPRIETA DELL'INDICE della CHIAVE PRIMARIA sono impostate automaticamente a
                                                            ' 1) Primario = Si (Chiave primaria)
                                                            ' 2) Indicizzato (duplicati non ammesSi)= True,
                                                            ' 3) Ignora Null = no
                                                            ' 4) la proprieta del campo Field - Indicizzato = (Duplicati non ammesSi)
                                                        
                                                        '   campo SEZ1_GE_02_ID_ESERCIZIO = PRIMARY KEY - indicizzata + ignor null
                                                       'With TableDefNuovo
                                                            'idxNuovo_Contatore_Univoco.Fields.Append .CreateField("NOME_NN_01_ID_NNNNNNNNN")
                                                            'idxNuovo_Contatore_Univoco.Primary = True              '.... PROPRIETA PRIMARY = chiave univoca
                                                            ''idxNuovo_Contatore_Univoco.Unique = True              'in automatico imposta  PROPRIETA Indicizzato(Duplicati non ammesSi)= Si + Index = univoco
                                                            ''idxNuovo_Contatore_Univoco.IgnoreNulls = True         'e PROPRIETA Index = Ingnora Null
                                                            '.Indexes.Append idxNuovo_Contatore_Univoco

                                                '____________________________________________________________________________________________________________________________________________
                                                'CAMPO_2            :CAMPO_INTEGER                                          :INTEGER -                            - PRIMARYKEY  :dbInteger
                                                'Propriet�_9        :dbAutoIncrField                                                                                            :<Si/no>
                                                'Propriet�_9        :Propriet� Required (Richiesto)                                                                             :<Si/no>
                                                'Propriet�_10       :Propriet� AllowZeroLength (Consenti lunghezza zero)                                                        :<Si/no>
                                                    
                                                    
                                                    'Creo il INTEGER AUTOINCREMENTANTE
                                                   'Set FieldNuovo = TableDefNuovo.CreateField("NOME_NN_01_ID_NNNNNNNNN", dbLong)
                                                    
                                                   'FieldNuovo.Attributes = dbAutoIncrField             '... imposta a contatore
                                                    'accodo il campo all'inSieme Fields
                                                   'TableDefNuovo.Fields.Append FieldNuovo
                                                    
                                                   
                                                    
                                                    
                                                    
                                                    '..........................................................
                                                    '   DEFINISCO I CAMPI DELL'INSIEME INDEX - CHIAVE PRIMARIA
                                                        
                                                       'Set idxNuovo_Contatore_Univoco = TableDefNuovo.CreateIndex("NOME_CHIAVE_PRIARIA_1")     '.... imposta Nome index
                                                                                                          
                                                            '   DEFINISCO L'INSiEME INDEX CON UN CAMPO A PRIMARY KEY
                                                            ' Quando imposto la proprieta del campo a chiave primaria
                                                            ' automaticamente le proprieta Index sono impostate nel seguente modo:
                                                            ' Il nome dell'indice viene attribuito con il metodo CreateIndex("Nome Indice"),
                                                            ' le PROPRIETA DELL'INDICE della CHIAVE PRIMARIA sono impostate automaticamente a
                                                            ' 1) Primario = Si (Chiave primaria)
                                                            ' 2) Indicizzato (duplicati non ammesSi)= True,
                                                            ' 3) Ignora Null = no
                                                            ' 4) la proprieta del campo Field - Indicizzato = (Duplicati non ammesSi)
                                                        
                                                        '   campo SEZ1_GE_02_ID_ESERCIZIO = PRIMARY KEY - indicizzata + ignor null
                                                       'With TableDefNuovo
                                                            'idxNuovo_Contatore_Univoco.Fields.Append .CreateField("NOME_NN_01_ID_NNNNNNNNN")
                                                            'idxNuovo_Contatore_Univoco.Primary = True              '.... PROPRIETA PRIMARY = chiave univoca
                                                            ''idxNuovo_Contatore_Univoco.Unique = True              'in automatico imposta  PROPRIETA Indicizzato(Duplicati non ammesSi)= Si + Index = univoco
                                                            ''idxNuovo_Contatore_Univoco.IgnoreNulls = True         'e PROPRIETA Index = Ingnora Null
                                                            '.Indexes.Append idxNuovo_Contatore_Univoco
                                                
                                                
                                                
                                                '____________________________________________________________________________________________________________________________________________
                                                'CAMPO_2            :CAMPO_INTEGER  INDICIZZATO SECONDARYKEY                  :INTEGER -                            - PRIMARYKEY  :dbInteger
                                                'Propriet�_11       :Propriet� Indexed (Indicizzato)                                                                            :<Si/no>
                                                        
                                                    
                                                     'Creo il campo_3 testo con dbText e definisco la lunghezza
                                                   'Set FieldNuovo = TableDefNuovo.CreateField("NOME CAMPO INDICIZZATO   ", dbInteger)
                                                   'FieldNuovo.Required = True                              '... imposta Rchiesto = Si
                                                    ' accodo il campo all'inSieme Fields
                                                   'TableDefNuovo.Fields.Append FieldNuovo
     
                                                        '..........................................................
                                                        '   DEFINISCO I CAMPI DELL'INSiEME INDEX - CHIAVE SECONDARIA

                                                       'Set idxNuovo_Contatore_Univoco = TableDefNuovo.CreateIndex("Chiave_Secondaria_NOME_AMPO_INDICIZZATO")     '.... imposta Nome index

                                                        '   campo SEZ1_ANNO_ESERCIZIO = SecondaryKey - indicizzata + ignor null
                                                             
                                                            'idxNuovo_Contatore_Univoco.Fields.Append .CreateField("NOME CAMPO INDICIZZATO  ")
                                                            'idxNuovo_Contatore_Univoco.Unique = True              'in automatico imposta  PROPRIETA Indicizzato(Duplicati non ammesSi)= Si + Index = univoco
                                                            'idxNuovo_Contatore_Univoco.IgnoreNulls = True         'e PROPRIETA Index = Ingnora Null
                                                            '.Indexes.Append idxNuovo_Contatore_Univoco
                                                       'End With



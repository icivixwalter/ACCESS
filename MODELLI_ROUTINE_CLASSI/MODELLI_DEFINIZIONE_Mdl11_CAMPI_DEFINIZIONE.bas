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
'Proprieta_1        :ProprietÓ FieldSize (Dimensione campo) n.b. solo
'                                                        per i campi Memo e Long Binary             - <tipo>
'ProprietÓ_2        :NewValues (Nuovi valori)                                                       - <valori>
'ProprietÓ_3        :Format (Formato)                                                               - <formato>
'ProprietÓ_3        :ProprietÓ Caption (Etichetta)                                                  - <etichetta>
'ProprietÓ_4        :ProprietÓ Indexed (Indicizzato)                                                - <tipo indice>
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
'Proprieta_1        :ProprietÓ FieldSize (Dimensione campo) n.b. solo per i campi
'                                                                   Memo e Long Binary          - <tipo>
'ProprietÓ_2        :ProprietÓ Format (Formato)                                                 - <formato>
'ProprietÓ_3        :ProprietÓ InputMask (Maschera di input)                                    - <maschera di imput>
'ProprietÓ_4        :ProprietÓ Caption (Etichetta)                                              - <etichetta>
'ProprietÓ_5        :ProprietÓ DefaultValue (Valore predefinito)                                - <tipo valore>
'ProprietÓ_6        :ProprietÓ ValidationRule, ValidationText (Valido se, Messaggio errore)     - <valore predefinito>
'ProprietÓ_7        :ProprietÓ DefaultValue (Valore predefinito)                                - <valido se>
'ProprietÓ_8        :ProprietÓ DefaultValue (Valore predefinito)                                - <messaggio di errore>
'ProprietÓ_9        :ProprietÓ Required (Richiesto)                                             - <si/no>
'ProprietÓ_10       :ProprietÓ AllowZeroLength (Consenti lunghezza zero)                        - <si/no>
'ProprietÓ_11       :ProprietÓ Indexed (Indicizzato)                                            - <si/no>
'ProprietÓ_12       :ProprietÓ Compressione unicode (NON FUNZIONA)                              - <si/no>
'..................................... ELENCO RICERCA .............................................................
'Ricerca_1          :ProprietÓ DisplayControl (Visualizza controllo)                            - <tipo controllo>
'Ricerca_2          :ProprietÓ RowSourceType, RowSource (Origine riga, Tipo origine riga)       - <tipo origine>
'Ricerca_3          :ProprietÓ RowSourceType, RowSource (Origine riga, Tipo origine riga)       - <origine riga>
'Ricerca_4          :ProprietÓ BoundColumn (Colonna associata)                                  - <colonna>
'Ricerca_5          :ProprietÓ ColumnCount (Numero colonne)                                     - <nro col>
'Ricerca_6          :ProprietÓ ColumnHeads (Intestazioni colonne)                               - <intest col.>
'Ricerca_7          :ProprietÓ ColumnWidths (Larghezza colonne)                                 - <largh.>
'Ricerca_8          :ProprietÓ ListRows (Righe in elenco)                                       - <nro>
'Ricerca_9          :ProprietÓ ListWidth (Larghezza elenco)                                     - <nro>
'Ricerca_10         :ProprietÓ LimitToList (Solo in elenco)                                     - <si/no>
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
                                                'ProprietÓ_9        :dbAutoIncrField                                                                                            :<Si/no>
                                                'ProprietÓ_9        :ProprietÓ Required (Richiesto)                                                                             :<Si/no>
                                                'ProprietÓ_10       :ProprietÓ AllowZeroLength (Consenti lunghezza zero)                                                        :<Si/no>
                                                    
                                                    
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
                                                'ProprietÓ_9        :dbAutoIncrField                                                                                            :<Si/no>
                                                'ProprietÓ_9        :ProprietÓ Required (Richiesto)                                                                             :<Si/no>
                                                'ProprietÓ_10       :ProprietÓ AllowZeroLength (Consenti lunghezza zero)                                                        :<Si/no>
                                                    
                                                    
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
                                                'ProprietÓ_11       :ProprietÓ Indexed (Indicizzato)                                                                            :<Si/no>
                                                        
                                                    
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



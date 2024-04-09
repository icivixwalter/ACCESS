Attribute VB_Name = "MODELLI_DEFINIZIONE_Mdl05_MODELLI_SQL"
Option Compare Text
Option Explicit



Private Sub sQL()
Dim sSql As String


'SELECT FROM
'----------------------------------------------------------------------
'       modello di QUERY SEMPLICE CON SELECT


                '...... a 1 campo

sSql = ""
sSql = sSql & "SELECT "
sSql = sSql & "[+++Tabella++++].[--campo---] AS [***AS*alias***], "
sSql = sSql & "[+++Tabella++++].[--campo---] AS [***AS*alias***] "
'       ..... FROM ....
sSql = sSql & " FROM [+++Tabella++++];"

                '......a 4 CAMPI

sSql = ""
sSql = sSql & "SELECT "
sSql = sSql & "[+++Tabella++++].[--campo---] AS [***AS*alias***], "
sSql = sSql & "[+++Tabella++++].[--campo---] AS [***AS*alias***], "
sSql = sSql & "[+++Tabella++++].[--campo---] AS [***AS*alias***], "
sSql = sSql & "[+++Tabella++++].[--campo---] AS [***AS*alias***] "
'       ..... FROM ....
sSql = sSql & " FROM [+++Tabella++++];"

                

                '......a TUTTI I  CAMPI + 1 ESPLICITO con 2 clausole WHERE con preposizione AND
                '      1° parametro stringa e il 2° parametro Long


sSql = ""
sSql = sSql & "SELECT SEZ2_T50_TRACCIATO_RECORD.*, "
sSql = sSql & "[SEZ2_T50_TRACCIATO_RECORD].[COD_COND] "
sSql = sSql & "FROM SEZ2_T50_TRACCIATO_RECORD "
'sSql = sSql & "WHERE ((([SEZ2_T50_TRACCIATO_RECORD].[COD_COND])='" & Me.cmb_Cod_Cond & "') "
'sSql = sSql & "AND (([SEZ2_T50_TRACCIATO_RECORD].[ANNO_INIZIO_ESERC])=" & Me.cmb_ANNO_INIZIO_ESERC & "));"



End Sub

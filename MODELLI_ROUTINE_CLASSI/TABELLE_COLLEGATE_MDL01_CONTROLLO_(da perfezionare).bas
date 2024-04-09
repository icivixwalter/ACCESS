Attribute VB_Name = "TABELLE_COLLEGATE_Mdl01_CONTROLLO_(da perfezionare)"
Option Compare Database


Sub CONTROLLO_N01_TABELLE_COLLEGATE()

Dim Tdf As DAO.TableDef
Dim Db As DAO.Database
Set Db = CodeDb

'//SE LA TABELLA NON ESISTE ERRORE = 3265
Set Tdf = Db.TableDefs("GE_CASA_DF01_DEFINIZ_CODICI")

'//SE LA TABELLA NON E' COLLEGATA ERRRE = 3044
Tdf.Connect = ";DATABASE=C:\Dati\Db2.Mdb"
Tdf.RefreshLink


Set Tdf = Nothing
Set Db = Nothing

End Sub

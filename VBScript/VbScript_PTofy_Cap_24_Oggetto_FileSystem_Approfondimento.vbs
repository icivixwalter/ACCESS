Option Explicit
On Error Resume Next
Dim fso,drv,inkb,inmb,ingb
'Variabili (anche per la conversione)

Set fso = CreateObject("Scripting.FileSystemObject")
'Creazione oggetto

drv = InputBox("Inserisci il drive di cui vuoi sapere la dimensione")
'Richiesta del drive

inkb = CCur(fso.GetDrive(drv).TotalSize) / 1024
'Da byte a kilobyte (byte / 1024)

inmb = CCur(inkb) / 1024
'Da kilobyte a megabyte (kb / 1024)

ingb = CCur(inmb) / 1024
'Da megabyte a gigabyte (mb / 1024)

wscript.echo("Dimensione del disco rigido " & drv & " GigaByte (GB) e TeraByte (TB) :")
'"Introduce"

wscript.echo("")
'Visualizza una riga vuota


wscript.echo(Round(CDbl(ingb)) & " GB")
'Visualizza in forma arrotondata Round(<valore>)
'la dimensione in GigaByte

wscript.echo(Round(CDbl(ingb)) / 1024 & " TB")
'Visualizza in forma arrotondata Round(<valore>)
'la dimensione in TeraByte

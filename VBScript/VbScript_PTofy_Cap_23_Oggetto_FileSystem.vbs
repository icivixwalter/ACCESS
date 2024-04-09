Option Explicit
Dim fso,drv
Set fso = CreateObject("Scripting.FileSystemObject")
drv = "C:" 'Usiamo il drive C
wscript.echo(fso.GetDrive(drv).TotalSize)& " FUNZIONE-> fso.GetDrive(drv).TotalSize) =  ci restituisce in bytes la dimensione totale del drive"

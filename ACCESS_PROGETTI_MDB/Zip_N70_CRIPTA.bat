@ECHO OFF

@REM		directory y dove archiviare i dati = path di destinazione
@REM .......................................................
SET PATH_DEST_S=c:\CASA\LINGUAGGI\ACCESS\ACCESS_PROGETTI_MDB\ZIP_SALVATAGGI\

@REM//CRIPTA i file .rar .zip in GE614

@RENAME %PATH_DEST_S%*ZIP_SALVATAGGI*.RAR *ZIP_SALVATAGGI_*.GE614 

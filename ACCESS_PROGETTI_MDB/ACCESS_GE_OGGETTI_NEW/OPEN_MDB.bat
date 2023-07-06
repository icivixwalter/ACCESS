@ECHO OFF

@REM		directory y dove archiviare i dati = path di destinazione
@REM .......................................................
SET PATH_DEST_S=c:\CASA\LINGUAGGI\ACCESS\ACCESS_PROGETTI_MDB\ACCESS_GE_OGGETTI_NEW\GE_OGGETTI_NEW_MDB\
SET FILE_OPEN_s=GE_OGGETTI_Access_2000_NEW.mdb
start "apri file" %PATH_DEST_S%%FILE_OPEN_s% &^exit
ECHO OFF
@REM  ***************************************************************

@REM     file batch : Salva-C rar
@REM     directory = c:\CASA\LINGUAGGI\ACCESS\ACCESS_PROGETTI_MDB\MSYS_OGGETTI\SALVATAGGI_MSYS\ ; \DOC  e \STAMPE
@REM     file WinRAR

@REM  ***************************************************************

SET PATH_DEST_S=c:\CASA\LINGUAGGI\ACCESS\ACCESS_PROGETTI_MDB\MSYS_OGGETTI\SALVATAGGI_MSYS\




@REM CICLO FOR I E II CASO PER LA GESTIONE DELLA DATA
@REM ========================================================================================================================
:----------------------------CICLO FOR I CASO per la gestione della data con le sottostringhe ma aggiunge lo 0 se <10
for /f "skip=1" %%x in ('wmic os get localdatetime') do if not defined MyDate set MyDate=%%x
echo solo il giorno:
echo %MyDate:~6,2%

echo I CASO la data con le sottostringhe con separataore l'undescore (_)
set today=%MyDate:~0,4%_%MyDate:~4,2%_%MyDate:~6,2%

echo.
echo today in formato stringa: 
echo %today%
echo.

:----------------------------CICLO FOR  II CASO cicolo for per per la data AAA MM GG senza sottostringhe

ECHO E' possibile ottenere la data corrente in modo indipendente dalle impostazioni locali utilizzando
ECHO SENZA armeggiare con le sottostringhe
echo vedi il link: https://qastack.it/programming/10945572/windows-batch-formatted-date-into-variable
echo .

echo. II CASO la data senza le sottostringhe solo numerico e con separatore il trattino (-)

@REM for /f %%x in ('wmic path win32_localtime get /format:list ^| findstr "="') do set %%x
@REM set today=%Year%_%Month%_%Day%

echo.
ECHO %TODAY%
echo.

@REM CICLO FOR I E II CASO PER LA GESTIONE DELLA DATA  *** FINE ***
@REM ========================================================================================================================







@REM//CRIPTA i file .ZIP .zip in GE614

@RENAME "%!TODAY!%*.GE614" "%TODAY%*.ZIP"


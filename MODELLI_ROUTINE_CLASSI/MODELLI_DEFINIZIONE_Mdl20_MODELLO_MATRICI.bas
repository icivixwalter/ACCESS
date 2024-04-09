Attribute VB_Name = "MODELLI_DEFINIZIONE_Mdl20_MODELLO_MATRICI"
Option Compare Database

'//PASSAGGIO_MATRICE_DI_ARGOMENTI
'//Per passare una matrice di argomenti a una routine, è possibile utilizzare una matrice di parametri. _
Per la definizione della routine non è necessario conoscere il numero di elementi della matrice. Le matrici di parametri possono _
essere identificate con la parola chiave ParamArray. La matrice deve essere dichiarata come di tipo Variant _
e deve essere l'ultimo argomento nella definizione della routine. Nell 'esempio seguente è _
illustrata la definizione di una routine mediante una matrice di parametri. _
Negli esempi seguenti vengono illustrate diverse modalità di richiamo di questa routine. _
AnyNumberArgs "Stefania", 10, 26, 32, 15, 22, 24, 16 _
AnyNumberArgs "Paola", "Alto", "Basso", "Medio", "Alto" _


'//Chiamo matrice con Stringhe + numeri
Private Sub ChimaMatrice_TIPO_01()
    AnyNumberArgs "Stefania", 10, 26, 32, 15, 22, 24, 16
End Sub


'//Chiamo matrice con Stringhe
Private Sub ChimaMatrice_TIPO_02()
    AnyNumberArgs "Paola", "Alto", "Basso", "Medio", "Alto"

End Sub

'LLPP_IMPEGNI_Tb01_ATTI_DI_IMPEGNO
'LLPP_IMPEGNI_Tb02_ELENCO_DI_SPESA


'//Chiamo matrice con Stringhe che rappresentano delle tabelle _
Impegni, Presenze
Private Sub ChimaMatrice_TIPO_03()
    AnyNumberArgs "LLPP_IMPEGNI_Tb01_ATTI_DI_IMPEGNO", "LLPP_IMPEGNI_Tb02_ELENCO_DI_SPESA", _
    "PRES3000_Tb01_Calendario", "PRES3000_Tb02_ElencoGiornaliere", "PRES3000_Tb01_Calendario", "Indirizzi_Tb01_INTESTATARI", _
    "Indirizzi_Tb02_ELENCO"

End Sub

Sub AnyNumberArgs(strName As String, ParamArray intScores() As Variant)
    Dim intI As Integer
    
    Debug.Print
    Debug.Print strName; " Punteggi"
    ' Utilizza la funzione UBound per definire il
    ' limite superiore della matrice.
    For intI = 0 To UBound(intScores())
        Debug.Print "          "; intScores(intI)
    Next intI

End Sub



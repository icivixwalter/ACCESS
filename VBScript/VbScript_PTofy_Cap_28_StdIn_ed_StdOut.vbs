'//STDIN (input)
Dim StdIn
Set StdIn = WScript.StdIn

'//STDOUT (output)
Dim StdOut
Set StdOut = WScript.StdOut 

Function ReadLine()
    Dim StdIn
    Set StdIn = WScript.StdIn 'Impostazione
    ReadLine = StdIn.ReadLine 'Ritorno l'input
End Function

Sub WriteLine(text)
    Dim StdOut
    Set StdOut = WScript.StdOut 'Impostazione
    StdOut.WriteLine(text) 'Scrivo l'output
End Sub


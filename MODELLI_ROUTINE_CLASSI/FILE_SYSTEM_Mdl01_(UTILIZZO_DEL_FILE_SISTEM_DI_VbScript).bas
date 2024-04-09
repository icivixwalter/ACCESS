Attribute VB_Name = "FILE_UTILIZZO_DEL_FILE_SISTEM_DI_VbScript"
'//NOTE per far riferimento al File <<New FileSystemObject>> occorre selezionare _
menu "Project" ~> "References...",seleziona dalla ListBox la voce "Microsoft Scripting Runtime",


Option Compare Database
Option Explicit




      Private Sub attiva_Loop_Fie()
        
        'Declare variables.
        Dim fso As New FileSystemObject
        Dim ts As TextStream
        'Open file.
        Set ts = fso.OpenTextFile(Environ("windir") & "\system.ini")
        'Loop while not at the end of the file.
        Do While Not ts.AtEndOfStream
          Debug.Print ts.ReadLine
        Loop
        'Close the file.
        ts.Close
      End Sub

      '//RESTITUISCE IL NUMERO DI FILE DELLA DIRECTORY
      Private Sub attiva_GETFILE()
         Dim fso As New FileSystemObject
         Dim f As File
         'Get a reference to the File object.
         Set f = fso.GetFile(Environ("windir") & "\system.ini")
         MsgBox f.Size 'displays size of file
      
      End Sub

      '//STAMPA LE DIRECTORY
      Private Sub CONTROLLA_PATH()
         Dim fso As New FileSystemObject
         Dim f As Folder, sf As Folder, path As String
         'Initialize path.
         path = Environ("windir")
         'Get a reference to the Folder object.
         Set f = fso.GetFolder(path)
         'Iterate through subfolders.
         For Each sf In f.SubFolders
           Debug.Print sf.Name
         Next
      End Sub
        
        
      '//EVIDENZIA LA LETTERA DEL DRIVE CORRENTE
      Private Sub MIO_DRIVE()
         Dim fso As New FileSystemObject
         Dim mydrive As Drive
         Dim path As String
         'Initialize path.
         path = "C:\"
         'Get object.
         Set mydrive = fso.GetDrive(path)
         'Check for success.
         MsgBox mydrive.DriveLetter 'displays "C"
      End Sub



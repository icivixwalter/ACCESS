Attribute VB_Name = "UTIL_MDL42_02_CONTROLLO_OGGETTI_TABLE_(Controllo_ATTRIBUTI)"
Option Compare Database

Private Sub AttributesX()
 
 Dim dbsNorthwind As Database
 Dim fldLoop As Field
 Dim relLoop As Relation
 Dim tdfloop As TableDef
 
 '//DATABASE ESTERNO
 'Set dbsNorthwind = OpenDatabase("Northwind.mdb")
 Set dbsNorthwind = CurrentDb
 
 With dbsNorthwind
 
 ' Display the attributes of a TableDef object's
 ' fields.
 Debug.Print "============================================================================"
 Debug.Print "I campi della tabella CONTROLLATA : "
 Debug.Print
 Debug.Print "Attributes of fields in " & _
    .TableDefs(0).Name & " table:"
    
    For Each fldLoop In .TableDefs(0).Fields
            Debug.Print " " & fldLoop.Name & " = " & _
            fldLoop.Attributes
    Next fldLoop
    
    ' Display the attributes of the Northwind database's
    ' relations.
    Debug.Print "Attributes of relations in " & _
            .Name & ":"
    For Each relLoop In .Relations
        Debug.Print " " & relLoop.Name & " = " & _
        relLoop.Attributes
    Next relLoop
    
    ' Display the attributes of the Northwind database's
    ' tables.
    Debug.Print "...................................................."
    Debug.Print " LE RELAZIONI DELLE TABELLE"
    Debug.Print
    Debug.Print "Attributes of tables in " & .Name & ":"
    For Each tdfloop In .TableDefs
        Debug.Print " " & tdfloop.Name & " = " & _
        tdfloop.Attributes
    Next tdfloop
    
 .Close
 End With
 
End Sub
 



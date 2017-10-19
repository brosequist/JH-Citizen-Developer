Sub HelloWorld()

  Dim dbs As Database
  Dim tbldef As TableDef
  
  Set dbs = CurrentDb
  
  For Each tbldef In dbs.TableDefs
    If tbldef.Name = "World" Then
      dbs.Execute "DROP TABLE " & tbldef.Name, dbFailOnError
    End If
  Next tbldef
  
  dbs.Execute "CREATE TABLE World " _
    & "(Message VARCHAR);", dbFailOnError
  
  dbs.Execute "INSERT INTO World " _
    & "(Message) VALUES " _
    & "('Hello!');"
  
  dbs.Close
  
  Set tbldef = Nothing
  Set dbs = Nothing
  
End Sub

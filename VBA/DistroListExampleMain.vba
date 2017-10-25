Sub TestArrayCollector()
  
  Dim Result() As String
  
  'MsgBox ActiveCell.Address
  
	Result = StoreVerticalListToArray(Range(ActiveCell.Address))
	
	'Exit the subroutine if nothing is passed to the Result dimension
	If Len(Join(Result)) < 1 Then
		MsgBox ("No Values Found")
		Exit Sub
	End If
	
	'Loop to print array results for testing if necessary
	'If Len(Join(Result)) > 0 Then
	'    For N = LBound(Result) To UBound(Result)
	'        MsgBox Result(N)
	'    Next
	'End If
	
	Dim MySubject As String
	MySubject = "Please see the attached workbook"
	
	Call SendToDistributionList(Result, MySubject)

End Sub

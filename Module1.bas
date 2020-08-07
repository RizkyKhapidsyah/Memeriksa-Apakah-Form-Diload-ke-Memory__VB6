Attribute VB_Name = "Module1"
Public Function FormLoadedByName(FormName As String) As Boolean
Dim i As Integer, fnamelc As String
  fnamelc = LCase$(FormName)
  FormLoadedByName = False
  For i = 0 To Forms.Count - 1
    If LCase$(Forms(i).Name) = fnamelc Then
       FormLoadedByName = True
       Exit Function
    End If
  Next
End Function



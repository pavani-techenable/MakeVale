Public Class clsattchment
    Public Property data As List(Of dataAtt)
    Public Property included As List(Of includedc)
End Class

Public Class dataAtt
    Public Property id As Integer
    Public Property type As String
    Public Property attributes
End Class
Public Class includedc
    Public Property id As Integer
    Public Property attributes

End Class

Public Class attrib
    Public Property file_id As Integer
    Public Property file_name As String
    Public Property url As String
    Public Property text As String
End Class
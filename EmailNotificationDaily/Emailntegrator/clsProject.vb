
Public Class clsProject
    Public Property data As List(Of datap)


    'Public teams As String
    'Public attributes As List(Of Attribute)



End Class

Public Class clsProjectSNG
    Public Property data As datap


    'Public teams As String
    'Public attributes As List(Of Attribute)



End Class



Public Class clsProject_N
    Public Property data As List(Of dataN)


    'Public teams As String
    'Public attributes As List(Of Attribute)



End Class

Public Class dataN
    Public Property id As Integer
    Public Property type As String
    ' Public Property attributes
End Class
Public Class datap
    Public Property id As Integer
    Public Property type As String
    Public Property archived As String
    Public Property attributes
End Class
Public Class attrip

    'INSERT INTO projects (t_id ,s_id ,name ,visibility ,start_date ,archived ,created_at ,updated_at) VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}') 

    Public Property name
    Public Property description
    Public Property visibility As String
    Public start_date As DateTime
    Public Property archived

    Public created_at As DateTime
    Public updated_at As DateTime

    'INSERT INTO dbo.tasks (t_id ,p_id ,e_id ,s_id ,name ,started_on ,due_date ,description ,state ,archived ,status_id ,status_name ,prev_status_id ,prev_status_name ,next_status_id ,next_status_name ,x ,y ,created_at ,updated_at) VALUES ('{0}','{1}' ,'{2}' ,'{3}','{4}' ,'{5}','{6}' ,'{7}' ,'{8}' ,'{9}' ,'{10}' ,'{11}' ,'{12}' ,'{13}' ,'{14}' ,'{15}' ,'{16}' ,'{17}' ,'{18}' ,'{19}')

    Public Property started_on
    Public Property due_date

    Public Property state

    Public Property status_id
    Public Property status_name
    Public Property prev_status_id
    Public Property prev_status_name
    Public Property next_status_id
    Public Property next_status_name
    Public Property x
    Public Property y

End Class
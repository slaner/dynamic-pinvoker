''' <summary>
''' 매개변수의 형식이 잘못되었을 때 발생하는 예외입니다.
''' </summary>
Public Class InvalidParameterTypeException
    Inherits Exception

    Public Sub New()
        MyBase.New()
    End Sub
    Public Sub New(ByVal Message As String)
        MyBase.New(Message)
    End Sub
    Public Sub New(ByVal Message As String, ByVal innerException As Exception)
        MyBase.New(Message, innerException)
    End Sub

End Class
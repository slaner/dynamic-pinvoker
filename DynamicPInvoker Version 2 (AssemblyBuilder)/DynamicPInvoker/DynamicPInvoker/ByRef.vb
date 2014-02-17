''' <summary>
''' 참조에 의해 값을 전달하는 방식을 구현하는 클래스입니다.
''' </summary>
Public Class [ByRef]

    Private g_Value As Object
    Private g_ValueType As Type

    ''' <summary>
    ''' 참조에 의해 값을 전달하는 ByRef 클래스의 새 개체를 초기화합니다.
    ''' </summary>
    ''' <param name="Value">참조로 넘길 개체를 입력합니다.</param>
    Public Sub New(ByVal Value As Object)

        g_Value = Value
        g_ValueType = Value.GetType().MakeByRefType

    End Sub

    ''' <summary>
    ''' 개체의 형식을 가져옵니다.
    ''' </summary>
    Public ReadOnly Property Type As Type
        Get
            Return g_ValueType
        End Get
    End Property

    ''' <summary>
    ''' 개체의 값을 가져옵니다.
    ''' </summary>
    Public ReadOnly Property Value As Object
        Get
            Return g_Value
        End Get
    End Property

End Class
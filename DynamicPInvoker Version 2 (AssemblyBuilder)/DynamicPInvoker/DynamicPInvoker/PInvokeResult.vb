''' <summary>
''' 플랫폼 호출의 결과값과 매개변수를 저장하는 클래스입니다.
''' </summary>
Public Class PInvokeResult

    Private g_Result As Object
    Private g_Parameters() As Object

    ''' <summary>
    ''' 결과 값 및 매개변수 목록을 지정하여, PInvokeResult 클래스의 새 개체를 초기화합니다.
    ''' </summary>
    ''' <param name="Result">결과 값을 입력합니다.</param>
    ''' <param name="Parameters">매개변수 목록을 입력합니다.</param>
    Friend Sub New(ByVal Result As Object, ByVal Parameters() As Object)
        g_Result = Result
        g_Parameters = Parameters
    End Sub

    ''' <summary>
    ''' 플랫폼 호출의 결과 값을 가져옵니다.
    ''' </summary>
    Public ReadOnly Property Result As Object
        Get
            Return g_Result
        End Get
    End Property

    ''' <summary>
    ''' 매개변수 목록을 가져옵니다.
    ''' </summary>
    Public ReadOnly Property Parameters As Object()
        Get
            Return g_Parameters
        End Get
    End Property

End Class
Imports System.Runtime.InteropServices

Partial Class DynamicPInvoker

    ''' <summary>
    ''' 호출 규약을 가져오거나 설정합니다.
    ''' </summary>
    Public Property CallingConvention As CallingConvention
        Get
            Return g_CallingConvention
        End Get
        Set(ByVal value As CallingConvention)
            g_CallingConvention = value
        End Set
    End Property

    ''' <summary>
    ''' 문자셋을 가져오거나 설정합니다.
    ''' </summary>
    Public Property CharSet As CharSet
        Get
            Return g_CharSet
        End Get
        Set(ByVal value As CharSet)
            g_CharSet = value
        End Set
    End Property

    ''' <summary>
    ''' 라이브러리 이름을 가져오거나 설정합니다.
    ''' </summary>
    Public Property LibraryName As String
        Get
            Return g_LibraryName
        End Get
        Set(ByVal value As String)
            g_LibraryName = value
        End Set
    End Property

    ''' <summary>
    ''' 메서드 이름을 가져오거나 설정합니다.
    ''' </summary>
    Public Property MethodName As String
        Get
            Return g_MethodName
        End Get
        Set(ByVal value As String)
            g_MethodName = value
        End Set
    End Property

    ''' <summary>
    ''' API의 반환 자료형을 가져오거나 설정합니다.
    ''' </summary>
    Public Property ReturnType As Type
        Get
            Return g_ReturnType
        End Get
        Set(ByVal value As Type)
            g_ReturnType = value
        End Set
    End Property

End Class
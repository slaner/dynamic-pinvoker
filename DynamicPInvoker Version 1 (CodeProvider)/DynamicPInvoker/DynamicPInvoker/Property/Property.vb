Imports System.Runtime.InteropServices

Partial Class DynamicPInvoker

    ''' <summary>
    ''' 포함할 네임스페이스의 목록을 가져옵니다.
    ''' </summary>
    Public Shared ReadOnly Property ImportedNamespaces As List(Of String)
        Get
            Return g_ImportedNamespaces
        End Get
    End Property

    ''' <summary>
    ''' 참조할 어셈블리의 목록을 가져옵니다.
    ''' </summary>
    Public Shared ReadOnly Property ReferencedAssemblies As List(Of String)
        Get
            Return g_ReferencedAssemblies
        End Get
    End Property

    ''' <summary>
    ''' 사용자 정의 자료형 선언 문자열을 가져오거나 설정합니다.
    ''' </summary>
    Public Shared Property UserDefinedTypeDefinition As String
        Get
            Return g_UserDefinedTypeDefinition
        End Get
        Set(ByVal value As String)
            g_UserDefinedTypeDefinition = value
        End Set
    End Property

    ''' <summary>
    ''' 매개변수의 목록을 가져옵니다.
    ''' </summary>
    Public ReadOnly Property ParamList As List(Of ParamInfo)
        Get
            Return g_ParamType
        End Get
    End Property

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
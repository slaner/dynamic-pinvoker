Imports System.Runtime.InteropServices

Partial Class DynamicPInvoker

    ''' <summary>
    ''' 라이브러리 이름과 메서드 이름을 사용하여 DynamicPInvoker 클래스의 새 개체를 초기화합니다.
    ''' </summary>
    ''' <param name="LibraryName">라이브러리 이름을 입력합니다.</param>
    ''' <param name="MethodName">메서드 이름을 입력합니다.</param>
    Public Sub New(ByVal LibraryName As String, ByVal MethodName As String)

        g_LibraryName = LibraryName
        g_MethodName = MethodName

    End Sub

    ''' <summary>
    ''' 라이브러리 이름, 메서드 이름과 반환 형식을 사용하여 DynamicPInvoker 클래스의 새 개체를 초기화합니다.
    ''' </summary>
    ''' <param name="LibraryName">라이브러리 이름을 입력합니다.</param>
    ''' <param name="MethodName">메서드 이름을 입력합니다.</param>
    ''' <param name="ReturnType">반환 형식을 입력합니다.</param>
    Public Sub New(ByVal LibraryName As String, ByVal MethodName As String, ByVal ReturnType As Type)

        g_LibraryName = LibraryName
        g_MethodName = MethodName
        g_ReturnType = ReturnType

    End Sub

    ''' <summary>
    ''' 라이브러리 이름, 메서드 이름, 반환 형식, 호출 규약 및 문자셋을 이용하여 DynamicPInvoker 클래스의 새 개체를 초기화합니다.
    ''' </summary>
    ''' <param name="LibraryName">라이브러리 이름을 입력합니다.</param>
    ''' <param name="MethodName">메서드 이름을 입력합니다.</param>
    ''' <param name="ReturnType">반환 형식을 입력합니다.</param>
    ''' <param name="CallingConvention">호출 규약을 입력합니다.</param>
    ''' <param name="CharSet">문자셋을 입력합니다.</param>
    Public Sub New(ByVal LibraryName As String, ByVal MethodName As String, ByVal ReturnType As Type, ByVal CallingConvention As CallingConvention, ByVal CharSet As CharSet)

        g_LibraryName = LibraryName
        g_MethodName = MethodName
        g_ReturnType = ReturnType
        g_CallingConvention = CallingConvention
        g_CharSet = CharSet

    End Sub

End Class
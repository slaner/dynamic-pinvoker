﻿Imports System.Runtime.InteropServices

''' <summary>
''' 동적 플랫폼 호출를 호출하는 작업을 구현하는 클래스입니다.
''' </summary>
Public Class DynamicPInvoker

    Private g_LibraryName As String
    Private g_MethodName As String
    Private g_CallingConvention As CallingConvention = System.Runtime.InteropServices.CallingConvention.StdCall
    Private g_CharSet As CharSet = System.Runtime.InteropServices.CharSet.Auto
    Private g_ReturnType As Type = Nothing
    Private g_ParamType As New List(Of ParamInfo)
    Private Shared g_UserDefinedTypeDefinition As String
    Private Shared g_ReferencedAssemblies As New List(Of String)
    Private Shared g_ImportedNamespaces As New List(Of String)

End Class
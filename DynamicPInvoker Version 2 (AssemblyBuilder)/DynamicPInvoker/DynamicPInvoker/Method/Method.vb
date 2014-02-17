Imports System.Reflection
Imports System.Reflection.Emit
Imports System.Runtime.InteropServices

Partial Class DynamicPInvoker

    ''' <summary>
    '''  지정된 라이브러리 및 메서드 이름, 문자셋과 호출 규약과 매개변수를 사용하여 동적으로 플랫폼 호출 메서드를 호출하고, 결과와 매개변수 목록을 포함하는 PInvokeResult 클래스의 개체를 반환합니다.
    ''' </summary>
    ''' <param name="parameters">API 호출에 사용할 매개변수를 입력합니다.</param>
    Public Function Invoke(ByVal ParamArray parameters() As Object) As PInvokeResult

        Dim paramTypes() As Type = Nothing,
            paramValues() As Object = Nothing

        ' 매개변수의 길이가 0이 아닐 경우에만 처리한다.
        If parameters.Length <> 0 Then

            ReDim paramTypes(parameters.Length - 1)
            ReDim paramValues(parameters.Length - 1)

            For i As Int32 = 0 To parameters.Length - 1

                ' 참조에 의해 값을 전달하는 경우:
                If parameters(i).GetType() Is GetType([ByRef]) Then
                    paramTypes(i) = parameters(i).Type
                    paramValues(i) = parameters(i).Value
                Else
                    paramTypes(i) = parameters(i).GetType
                    paramValues(i) = parameters(i)
                End If

            Next

        End If

        ' 동적 플랫폼 호출에 필요한 어셈블리, 모듈, 형식과 메서드를 생성합니다.
        Dim asmBuilder As AssemblyBuilder = AppDomain.CurrentDomain.DefineDynamicAssembly(New AssemblyName("TeamDEV.DynamicPInvoker"), AssemblyBuilderAccess.Run)
        Dim modBuilder As ModuleBuilder = asmBuilder.DefineDynamicModule("DynamicPInvokerModule")
        Dim typeBuilder As TypeBuilder = modBuilder.DefineType("DynamicPInvokerMethod", TypeAttributes.Public)
        Dim methodBuilder As MethodBuilder = typeBuilder.DefinePInvokeMethod(g_MethodName, g_LibraryName, MethodAttributes.Public Or MethodAttributes.PinvokeImpl Or MethodAttributes.Static, CallingConventions.Standard, g_ReturnType, paramTypes, g_CallingConvention, g_CharSet)

        ' 메서드의 구현 방식을 설정합니다.
        methodBuilder.SetImplementationFlags(methodBuilder.GetMethodImplementationFlags() Or MethodImplAttributes.PreserveSig)

        ' 형식을 생성하고, 생성된 형식에서 메서드를 검색합니다.
        Dim methodType As Type = typeBuilder.CreateType()
        Dim method As MethodInfo = methodType.GetMethod(g_MethodName)

        ' 결과 값을 저장하기 위한 변수를 선언하고,
        ' 메서드를 호출합니다.
        Dim RETVAL As Object
        RETVAL = method.Invoke(Nothing, paramValues)

        ' 결과 값과 매개변수 목록을 포함하는 PInvokeResult 개체를 반환합니다.
        Return New PInvokeResult(RETVAL, paramValues)

    End Function



    ''' <summary>
    ''' 지정된 라이브러리 이름, 메서드 이름, 반환 형식, 기본 호출 규약 및 문자셋과 매개변수를 사용하여 동적으로 플랫폼 호출 메서드를 호출하고, 결과와 매개변수 목록을 포함하는 PInvokeResult 클래스의 개체를 반환합니다.
    ''' </summary>
    ''' <param name="LibraryName">라이브러리 이름을 입력합니다.</param>
    ''' <param name="MethodName">메서드 이름을 입력합니다.</param>
    ''' <param name="ReturnType">반환 형식을 입력합니다.</param>
    ''' <param name="parameters">API 호출에 사용할 매개변수를 입력합니다.</param>
    Public Shared Function Invoke(ByVal LibraryName As String, ByVal MethodName As String, ByVal ReturnType As Type, ByVal ParamArray parameters() As Object) As PInvokeResult

        Dim paramTypes() As Type = Nothing,
            paramValues() As Object = Nothing

        ' 매개변수의 길이가 0이 아닐 경우에만 처리한다.
        If parameters.Length <> 0 Then

            ReDim paramTypes(parameters.Length - 1)
            ReDim paramValues(parameters.Length - 1)

            For i As Int32 = 0 To parameters.Length - 1

                ' 참조에 의해 값을 전달하는 경우:
                If parameters(i).GetType() Is GetType([ByRef]) Then
                    paramTypes(i) = parameters(i).Type
                    paramValues(i) = parameters(i).Value
                Else
                    paramTypes(i) = parameters(i).GetType
                    paramValues(i) = parameters(i)
                End If

            Next

        End If

        ' 동적 플랫폼 호출에 필요한 어셈블리, 모듈, 형식과 메서드를 생성합니다.
        Dim asmBuilder As AssemblyBuilder = AppDomain.CurrentDomain.DefineDynamicAssembly(New AssemblyName("TeamDEV.DynamicPInvoker"), AssemblyBuilderAccess.Run)
        Dim modBuilder As ModuleBuilder = asmBuilder.DefineDynamicModule("DynamicPInvokerModule")
        Dim typeBuilder As TypeBuilder = modBuilder.DefineType("DynamicPInvokerMethod", TypeAttributes.Public)
        Dim methodBuilder As MethodBuilder = typeBuilder.DefinePInvokeMethod(MethodName, LibraryName, MethodAttributes.Public Or MethodAttributes.PinvokeImpl Or MethodAttributes.Static, CallingConventions.Standard, ReturnType, paramTypes, Runtime.InteropServices.CallingConvention.StdCall, Runtime.InteropServices.CharSet.Auto)

        ' 메서드의 구현 방식을 설정합니다.
        methodBuilder.SetImplementationFlags(methodBuilder.GetMethodImplementationFlags() Or MethodImplAttributes.PreserveSig)

        ' 형식을 생성하고, 생성된 형식에서 메서드를 검색합니다.
        Dim methodType As Type = typeBuilder.CreateType()
        Dim method As MethodInfo = methodType.GetMethod(MethodName)

        ' 결과 값을 저장하기 위한 변수를 선언하고,
        ' 메서드를 호출합니다.
        Dim RETVAL As Object
        RETVAL = method.Invoke(Nothing, paramValues)

        ' 결과 값과 매개변수 목록을 포함하는 PInvokeResult 개체를 반환합니다.
        Return New PInvokeResult(RETVAL, paramValues)

    End Function

    ''' <summary>
    ''' 지정된 라이브러리 이름, 메서드 이름, 반환 형식, 호출 규약 및 문자셋과 매개변수를 사용하여 동적으로 플랫폼 호출 메서드를 호출하고, 결과와 매개변수 목록을 포함하는 PInvokeResult 클래스의 개체를 반환합니다.
    ''' </summary>
    ''' <param name="LibraryName">라이브러리 이름을 입력합니다.</param>
    ''' <param name="MethodName">메서드 이름을 입력합니다.</param>
    ''' <param name="ReturnType">반환 형식을 입력합니다.</param>
    ''' <param name="CallingConvention">호출 규약을 입력합니다.</param>
    ''' <param name="CharSet">문자셋을 입력합니다.</param>
    ''' <param name="parameters">API 호출에 사용할 매개변수를 입력합니다.</param>
    Public Shared Function Invoke(ByVal LibraryName As String, ByVal MethodName As String, ByVal ReturnType As Type, ByVal CallingConvention As CallingConvention, ByVal CharSet As CharSet, ByVal ParamArray parameters() As Object) As PInvokeResult

        Dim paramTypes() As Type = Nothing,
            paramValues() As Object = Nothing

        ' 매개변수의 길이가 0이 아닐 경우에만 처리한다.
        If parameters.Length <> 0 Then

            ReDim paramTypes(parameters.Length - 1)
            ReDim paramValues(parameters.Length - 1)

            For i As Int32 = 0 To parameters.Length - 1

                ' 참조에 의해 값을 전달하는 경우:
                If parameters(i).GetType() Is GetType([ByRef]) Then
                    paramTypes(i) = parameters(i).Type
                    paramValues(i) = parameters(i).Value
                Else
                    paramTypes(i) = parameters(i).GetType
                    paramValues(i) = parameters(i)
                End If

            Next

        End If

        ' 동적 플랫폼 호출에 필요한 어셈블리, 모듈, 형식과 메서드를 생성합니다.
        Dim asmBuilder As AssemblyBuilder = AppDomain.CurrentDomain.DefineDynamicAssembly(New AssemblyName("TeamDEV.DynamicPInvoker"), AssemblyBuilderAccess.Run)
        Dim modBuilder As ModuleBuilder = asmBuilder.DefineDynamicModule("DynamicPInvokerModule")
        Dim typeBuilder As TypeBuilder = modBuilder.DefineType("DynamicPInvokerMethod", TypeAttributes.Public)
        Dim methodBuilder As MethodBuilder = typeBuilder.DefinePInvokeMethod(MethodName, LibraryName, MethodAttributes.Public Or MethodAttributes.PinvokeImpl Or MethodAttributes.Static, CallingConventions.Standard, ReturnType, paramTypes, CallingConvention, CharSet)

        ' 메서드의 구현 방식을 설정합니다.
        methodBuilder.SetImplementationFlags(methodBuilder.GetMethodImplementationFlags() Or MethodImplAttributes.PreserveSig)

        ' 형식을 생성하고, 생성된 형식에서 메서드를 검색합니다.
        Dim methodType As Type = typeBuilder.CreateType()
        Dim method As MethodInfo = methodType.GetMethod(MethodName)

        ' 결과 값을 저장하기 위한 변수를 선언하고,
        ' 메서드를 호출합니다.
        Dim RETVAL As Object
        RETVAL = method.Invoke(Nothing, paramValues)

        ' 결과 값과 매개변수 목록을 포함하는 PInvokeResult 개체를 반환합니다.
        Return New PInvokeResult(RETVAL, paramValues)

    End Function

End Class
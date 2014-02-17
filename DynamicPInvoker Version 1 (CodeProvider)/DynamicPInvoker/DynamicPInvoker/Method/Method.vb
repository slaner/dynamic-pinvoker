Imports System.Runtime.InteropServices
Imports System.CodeDom.Compiler
Imports System.Reflection

Partial Class DynamicPInvoker

    ''' <summary>
    ''' Api를 호출합니다.
    ''' </summary>
    ''' <param name="parameters">Api 호출에 사용될 매개변수를 입력합니다.</param>
    Public Function Invoke(ByVal ParamArray parameters() As Object) As PInvokeResult

        ' Api를 포함하는 어셈블리 생성
        Dim dacr As DynamicApiHelper.DynamicApiCompileResults = DynamicApiHelper.CompileDynamicApiAssembly(Me)

        ' 오류가 발생한 경우, 예외를 발생시킨다.
        For Each e As CompilerError In dacr.CompilerResults.Errors
            Debug.Print(e.ErrorNumber & vbTab & " | " & e.ErrorText)
            Throw New InvalidOperationException("내부 오류가 발생했습니다.")
        Next

        ' Api 메서드의 정보를 가져온다.
        Dim ApiMethod As MethodInfo = dacr.CompilerResults.CompiledAssembly.GetType(dacr.ModuleName).GetMethod("DynamicApiMethod")

        ' 메서드 정보를 가져오지 못한 경우, 예외를 발생시킨다.
        If IsNothing(ApiMethod) Then
            Throw New InvalidOperationException("내부 오류가 발생했습니다.")
        End If

        ' 매개변수의 형식을 어셈블리에서 외부로 노출하는 형식으로 변환한 개체 배열을 만들고,
        ' Api를 호출하고 결과 값을 저장합니다.
        Dim convertedParameters As Object() = DynamicApiHelper.ConvertParameters(dacr, parameters),
            ar As Object = ApiMethod.Invoke(Nothing, convertedParameters)

        ' Api 호출 결과와 매개변수의 배열을 가지는 ApiResult 클래스를 반환합니다.
        Return New PInvokeResult(ar, convertedParameters)

    End Function

End Class
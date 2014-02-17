Imports System.Reflection
Imports System.CodeDom.Compiler

Friend MustInherit Class DynamicApiHelper

    Public Shared DefaultTypes() As String = {"Int16", "Int32", "Int64", "IntPtr", "UInt16", "UInt32", "UInt64", "UIntPtr", "String", "Char", "Boolean", "Object", "Date", "Single", "Double", "SByte", "Byte", "Decimal", "Char"}

    ''' <summary>
    ''' 동적 Api를 외부에 노출하는 어셈블리를 컴파일한 결과와 모듈의 이름을 저장하는 구조체입니다.
    ''' </summary>
    Public Structure DynamicApiCompileResults

        ''' <summary>
        ''' 어셈블리 내에 Api를 외부에 노출하는 모듈의 이름을 저장합니다.
        ''' </summary>
        Public ModuleName As String

        ''' <summary>
        ''' 컴파일 결과를 저장합니다.
        ''' </summary>
        Public CompilerResults As CompilerResults

    End Structure



    ''' <summary>
    ''' Api를 호출하는 어셈블리에서 정의된 형식을 Api가 정의된 어셈블리에서 노출하는 형식으로 변환한 값을 반환합니다.
    ''' </summary>
    ''' <param name="DACR">CompileDynamicApiAssembly 메서드를 호출하여 반환된 DynamicApiCompileResults 개체를 입력합니다.</param>
    ''' <param name="parameters">변환할 매개변수를 입력합니다.</param>
    Public Shared Function ConvertParameters(ByVal DACR As DynamicApiCompileResults, ByVal ParamArray parameters() As Object) As Object()

        ' 매개변수가 없을 경우, 빈 값을 반환한다.
        If IsNothing(parameters) OrElse parameters.Length = 0 Then
            Return Nothing
        End If

        Dim convertedParams(parameters.Length - 1) As Object

        For i As Int32 = 0 To parameters.Length - 1
            Dim p As Object = parameters(i),
                paramType As Type = parameters(i).GetType()

            ' 매개변수의 형식이 기본 형식일 경우엔 값을 그대로 적용하고
            ' 다음 순환을 시작한다.
            If IsDefaultType(paramType) Then

                convertedParams(i) = p
                Continue For

            End If

            ' 기본 형식이 아닐 경우 (= 사용자 정의 형식)

            ' 매개변수의 형식을 어셈블리에서 찾는다.
            Dim exportedType As Type = DACR.CompilerResults.CompiledAssembly.GetType(DACR.ModuleName & "+" & paramType.Name, False, True)

            ' 어셈블리에서 매개변수의 형식을 외부로 노출하는 경우:
            If Not IsNothing(exportedType) Then

                ' 매개변수 형식의 필드 정보와 어셈블리에서 노출하는 형식의 필드 정보를 가져온다.
                Dim inField As FieldInfo() = paramType.GetFields(),
                    outField As FieldInfo() = exportedType.GetFields()

                ' 필드 목록의 갯수가 다를 경우, 예외를 발생시킨다.
                If inField.Length <> outField.Length Then
                    Throw New InvalidParameterTypeException("매개변수의 형식과 어셈블리에서 노출하는 형식이 서로 다릅니다.")
                End If

                ' 노출하는 형식의 개체를 생성합니다.
                Dim typeInstance As ValueType = Activator.CreateInstance(exportedType)

                For fieldCounter As Int32 = 0 To inField.Length - 1

                    ' 필드의 이름이 다를 경우, 예외를 발생시킨다.
                    If inField(fieldCounter).Name <> outField(fieldCounter).Name Then
                        Throw New InvalidParameterTypeException("필드의 이름이 일치하지 않습니다.")
                    End If

                    outField(fieldCounter).SetValue(typeInstance, DynamicApiHelper.GetFieldValue(p, inField(fieldCounter).Name))

                Next

                convertedParams(i) = typeInstance

            End If

        Next

        Return convertedParams

    End Function

    ''' <summary>
    ''' 지정한 형식이 기본 형식인지 확인합니다.
    ''' </summary>
    ''' <param name="t">확인할 형식을 입력합니다.</param>
    Public Shared Function IsDefaultType(ByVal t As Type) As Boolean

        For Each tn In DefaultTypes
            If t.FullName = String.Format("System.{0}", tn) Then
                Return True
            End If
        Next
        Return False

    End Function

    ''' <summary>
    ''' 필드의 값을 가져옵니다.
    ''' </summary>
    ''' <param name="Source">값을 가져올 필드가 정의되어 있는 개체를 입력합니다.</param>
    ''' <param name="FieldName">필드의 이름을 입력합니다.</param>
    ''' <param name="AccessModifier">필드의 접근 수식자를 입력합니다.</param>
    Public Shared Function GetFieldValue(ByVal Source As Object, ByVal FieldName As String, Optional ByVal AccessModifier As BindingFlags = BindingFlags.Instance Or BindingFlags.Public) As Object

        Return Source.GetType().GetField(FieldName, AccessModifier).GetValue(Source)

    End Function

    ''' <summary>
    ''' 지정한 매개변수와 반환 형식을 가진 Api를 외부에 노출하는 어셈블리를 만들고, DynamicApiCompileResults 구조체를 반환합니다.
    ''' </summary>
    ''' <param name="Api">동적 Api 호출을 구현하는 DynamicApi 클래스의 유효한 개체를 입력합니다.</param>
    Public Shared Function CompileDynamicApiAssembly(ByVal Api As DynamicPInvoker) As DynamicApiCompileResults

        ' 위에서부터 차례대로
        '   VB 코드 컴파일러 개체
        '   컴파일 옵션을 저장하는 개체
        '   컴파일 결과를 저장할 변수
        '   임포트 하는 네임스페이스 저장할 문자열
        '   모듈 이름을 저장할 문자열
        '   특성을 저장할 문자열
        '   반환 형식을 저장할 문자열
        '   매개변수를 저장할 문자열
        '   메서드 형식을 저장할 문자열
        '   최종 컴파일할 코드를 저장할 문자열
        '   최종 컴파일할 메서드 코드를 저장할 문자열
        Dim vbc As CodeDomProvider = CodeDomProvider.CreateProvider("VB"),
            cp As New CompilerParameters(),
            cr As CompilerResults = Nothing,
            strNamespaceImports As String = Nothing,
            ModuleName As String = Nothing,
            strAttr As String = String.Format("<System.Runtime.InteropServices.DllImport({0}{1}{0}, EntryPoint:={0}{2}{0}, CallingConvention:={3}, CharSet:={4})>", Chr(34), Api.LibraryName, Api.MethodName, Val(Api.CallingConvention), Val(Api.CharSet) + 1),
            strReturnType As String = Nothing,
            strParameters As String = Nothing,
            strMethodType As String = Nothing,
            strCompleteCode As String = Nothing,
            strCompleteMethod As String = Nothing

        ' 컴파일 옵션:
        '   결과 파일 생성하지 않음
        '   메모리에 결과를 생성
        '   디버그 정보를 포함하지 않음
        '   참조할 어셈블리를 추가한다
        cp.GenerateExecutable = False
        cp.GenerateInMemory = True
        cp.IncludeDebugInformation = False
        cp.ReferencedAssemblies.AddRange(DynamicPInvoker.ReferencedAssemblies.ToArray)

        ' 포함할 네임스페이스 목록을 문자열로 만든다.
        For Each n As String In DynamicPInvoker.ImportedNamespaces
            strNamespaceImports &= "Imports " & n & vbCrLf
        Next

        ' 모듈 이름을 저장한다.
        ModuleName = String.Format("__DynAPI_APIInvokerModule{0:X8}", Environment.TickCount)

        ' Api의 반환 형식이 Void 혹은 Nothing인 경우, 반환 자료형을 설정하지 않는다.
        If Api.ReturnType Is Nothing OrElse Api.ReturnType Is GetType(Void) Then

            strMethodType = "Sub"

        Else

            ' 반환 형식이 명시되어있고, Void형이 아닌 경우
            ' 반환 자료형을 설정한다.
            strMethodType = "Function"
            strReturnType = "As " & Api.ReturnType.Name

        End If

        If Api.ParamList.Count <= 0 Then

            ' 매개변수가 없을 경우 괄호만 사용한다.
            strParameters = "()"

        Else

            strParameters = "("
            For i As Int32 = 0 To Api.ParamList.Count - 1
                Dim p As ParamInfo = Api.ParamList(i)

                Dim bPassTypeSet As Boolean = False,
                    bOptionalParameter As Boolean = False

                ' 매개변수 전달 방식을 검사한다.
                If p.ParamFlag And (ParamInfo.ParamFlags.ByVal And ParamInfo.ParamFlags.ByRef) Then
                    Throw New InvalidOperationException("매개변수의 옵션에는 ByVal과 ByRef을 같이 사용할 수 없습니다.")
                End If

                ' 매개변수의 자료 전달 방식을 적용한다.
                If p.ParamFlag And ParamInfo.ParamFlags.Optional Then
                    strParameters &= "Optional "
                    bOptionalParameter = True
                End If
                If p.ParamFlag And ParamInfo.ParamFlags.ByRef Then
                    strParameters &= "ByRef "
                    bPassTypeSet = True
                End If
                If p.ParamFlag And ParamInfo.ParamFlags.ByVal Then
                    strParameters &= "ByVal "
                    bPassTypeSet = True
                End If

                ' 매개변수의 자료 전달 방식이 설정되지 않았을 경우, ByVal 방식을 기본으로 적용한다.
                If Not bPassTypeSet Then
                    strParameters &= "ByVal "
                End If

                ' 생략 가능한 매개변수인 경우, 기본값을 지정해준다.
                If bOptionalParameter Then
                    strParameters &= String.Format("param{0} As {1} = {2}", i, p.ParamType.Name, p.DefaultValue)
                Else
                    strParameters &= String.Format("param{0} As {1}", i, p.ParamType.Name)
                End If

                ' 마지막 매개변수가 아닌 경우엔 쉼표를 찍는다.
                If i < Api.ParamList.Count - 1 Then
                    strParameters &= ", "
                End If

            Next
            strParameters &= ")"

        End If

        '[Attribute]
        'Public [Function/Sub] [Name] [As Type]
        'End [Function/Sub]

        ' 메서드 형식에 맞게 문자열을 처리한다.
        strCompleteMethod = String.Format("{0}" & vbCrLf & _
                                            "Public {1} DynamicApiMethod{2} {3}" & vbCrLf & _
                                            "End {1}",
                                            strAttr, strMethodType, strParameters, strReturnType)

        ' 컴파일할 코드를 만든다.
        strCompleteCode = String.Format("' DynAPI API Invoker Module - Auto generated at {0}" & vbCrLf & _
                        "{1}" & vbCrLf & _
                        "Module {2}" & vbCrLf & _
                        "{3}" & vbCrLf & _
                        "{4}" & vbCrLf & _
                        "End Module",
                        Now, strNamespaceImports, ModuleName, DynamicPInvoker.UserDefinedTypeDefinition, strCompleteMethod)

        cr = vbc.CompileAssemblyFromSource(cp, strCompleteCode)
        vbc.Dispose()
        Return New DynamicApiCompileResults() With {.CompilerResults = cr, .ModuleName = ModuleName}

    End Function

End Class
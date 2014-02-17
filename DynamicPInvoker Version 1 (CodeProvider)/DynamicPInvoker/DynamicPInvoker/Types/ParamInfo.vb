''' <summary>
''' 매개변수의 정보를 저장합니다.
''' </summary>
Public Structure ParamInfo

    ''' <summary>
    ''' 매개변수의 형식을 결정하는 값들을 열거합니다.
    ''' </summary>
    Public Enum ParamFlags

        ''' <summary>
        ''' 값에 의한 전달을 사용합니다.
        ''' </summary>
        [ByVal] = 1

        ''' <summary>
        ''' 참조에 의한 전달을 사용합니다.
        ''' </summary>
        [ByRef] = 2

        ''' <summary>
        ''' 생략 가능한 매개변수를 나타냅니다.
        ''' </summary>
        [Optional] = 4

    End Enum

    ''' <summary>
    ''' 매개변수의 자료형을 저장합니다.
    ''' </summary>
    Public ParamType As Type

    ''' <summary>
    ''' 매개변수의 형식을 저장합니다.
    ''' </summary>
    Public ParamFlag As ParamFlags

    ''' <summary>
    ''' 생략 가능한 매개변수의 기본 값을 저장합니다.
    ''' </summary>
    Public DefaultValue As Object

End Structure
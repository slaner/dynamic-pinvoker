Imports TeamDEV
Imports System.Runtime.InteropServices

Module LibraryTarget

    Structure ProcessEntry32

        Public dwSize As Int32
        Public cntUsage As Int32
        Public th32ProcessID As Int32
        Public th32DefaultHeapID As IntPtr
        Public th32ModuleID As Int32
        Public cntThreads As Int32
        Public th32ParentProcessID As Int32
        Public pcPriClassBase As Int32
        Public dwFlags As Int32
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=260)> _
        Public szExeFile As String

    End Structure

    Sub Main()

        ' Define an API
        Dim GetLastError As New DynamicPInvoker("kernel32", "GetLastError", GetType(Int32)),
            CreateToolhelp32Snapshot As New DynamicPInvoker("kernel32", "CreateToolhelp32Snapshot", GetType(IntPtr)),
            Process32First As New DynamicPInvoker("kernel32", "Process32First", GetType(Int32)),
            Process32Next As New DynamicPInvoker("kernel32", "Process32Next", GetType(Int32)),
            CloseHandle As New DynamicPInvoker("kernel32", "CloseHandle", GetType(Void)),
            PE32 As New ProcessEntry32 With {.dwSize = Marshal.SizeOf(GetType(ProcessEntry32))},
            ProcessEntry32Ref As New ParamInfo With {.ParamType = GetType(ProcessEntry32), .ParamFlag = ParamInfo.ParamFlags.ByRef},
            IntPtrParam As New ParamInfo With {.ParamType = GetType(IntPtr), .ParamFlag = ParamInfo.ParamFlags.ByVal},
            Int32Param As New ParamInfo With {.ParamType = GetType(Int32), .ParamFlag = ParamInfo.ParamFlags.ByVal}

        DynamicPInvoker.ImportedNamespaces.Add("System")
        DynamicPInvoker.ImportedNamespaces.Add("System.Runtime.InteropServices")
        DynamicPInvoker.UserDefinedTypeDefinition = "    Structure ProcessEntry32" & vbCrLf & _
                                                    "        Public dwSize As Int32" & vbCrLf & _
                                                    "        Public cntUsage As Int32" & vbCrLf & _
                                                    "        Public th32ProcessID As Int32" & vbCrLf & _
                                                    "        Public th32DefaultHeapID As IntPtr" & vbCrLf & _
                                                    "        Public th32ModuleID As Int32" & vbCrLf & _
                                                    "        Public cntThreads As Int32" & vbCrLf & _
                                                    "        Public th32ParentProcessID As Int32" & vbCrLf & _
                                                    "        Public pcPriClassBase As Int32" & vbCrLf & _
                                                    "        Public dwFlags As Int32" & vbCrLf & _
                                                    "        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=260)> _" & vbCrLf & _
                                                    "        Public szExeFile As String" & vbCrLf & _
                                                    "    End Structure"

        CreateToolhelp32Snapshot.ParamList.Add(Int32Param)
        CreateToolhelp32Snapshot.ParamList.Add(Int32Param)

        Process32First.ParamList.Add(IntPtrParam)
        Process32First.ParamList.Add(ProcessEntry32Ref)

        Process32Next.ParamList.Add(IntPtrParam)
        Process32Next.ParamList.Add(ProcessEntry32Ref)

        CloseHandle.ParamList.Add(IntPtrParam)

        Dim Snapshot As IntPtr = CreateToolhelp32Snapshot.Invoke(2, 0).Result

        ' If unable to create a snapshot; display text with error code.
        If Snapshot = IntPtr.Zero Then

            Console.WriteLine("CreateToolhelp32Snapshot failed! (ErrorCode: {0})", GetLastError.Invoke().Result)

        Else

            ' Invoke Process32First and Process32Next to enumerate all running processes.
            Dim res As PInvokeResult = Process32First.Invoke(Snapshot, PE32)

            If res.Result = 0 Then

                Console.WriteLine("Process32First failed! (ErrorCode: {0})", GetLastError.Invoke().Result)
                CloseHandle.Invoke(Snapshot)
                Return

            End If

            ' If Process32First succeed, display PID and Image name to console screen.
            Do While res.Result

                Console.WriteLine("ID: {0} ({1})", res.Parameters(1).th32ProcessID, res.Parameters(1).szExeFile)

                ' Call Process32Next to get next process's information.
                res = Process32Next.Invoke(Snapshot, PE32)

            Loop

            ' Close a snapshot handle
            CloseHandle.Invoke(Snapshot)

        End If

    End Sub

End Module
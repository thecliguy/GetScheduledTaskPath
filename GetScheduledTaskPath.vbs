Function GetScheduledTaskPath
    '***************************************************************************
    ' DETAILS
    '   Copyright (C) 2020
    '   Adam Russell <adam[at]thecliguy[dot]co[dot]uk> 
    '   https://www.thecliguy.co.uk
    '   
    '   Licensed under the MIT License.
    '
    ' PURPOSE
    '   Intended for use in scripts that are run as a scheduled task, this
    '   function will return the scheduled task's path.
    '
    ' HOW IT WORKS
    '   When a scheduled task invokes a VBScript interpreter (cscript or 
    '   wscript), the task's EnginePID is populated with the interpreter's 
    '   process ID (PID). This function compares the EnginePID of each running 
    '   task against the script interpreter's PID and returns the task path of 
    '   any matches.
    '
    ' CAVEATS
    '   This function will only work in scripts where the interpreter is invoked
    '   directly by a scheduled task (scheduled task -> interpreter). It will 
    '   not work in scripts where the interpreter is invoked indirectly, such as 
    '   a task that invokes a batch file which in turn invokes a VBScript 
    '   interpreter (scheduled task -> batch file -> interpreter).
    '
    '   This script has been tested on Windows 7 and is known **not** to work.
    '   I took the decision not to complicate the code to accommodate Windows 7 
    '   since the OS went end-of-life on 14th January 2020. I would therefore
    '   assume that this script won't work on any version of Windows prior to
    '   Windows 7.
    '
    ' DEPENDENCIES
    '   This function uses Windows PowerShell (powershell.exe) to obtain the 
    '   script interpreter's PID.
    '
    ' RETURN VALUE
    '   No result: An empty string.
    '
    '   One result: A string, the value of which is a scheduled task path.
    '
    '   Multiple results: A string, the value of which is a comma separated list
    '   of scheduled task paths.
    '
    '   NB: Whilst it's technically possible for multiple scheduled tasks to run 
    '       under the same EnginePID, the only examples I've personally observed 
    '       are those triggered 'At log on' with an action that launches a COM 
    '       handler.
    '***************************************************************************
    ' DEVELOPMENT LOG
    '
    ' 1.0.0, 2020-09-26, Adam Russell
    '   * First release
    '
    '***************************************************************************
    
    Dim objShell, objExec, strPsExe, strPsArg, objScheduleService
    Dim taskCollection, registeredTask, i, arrTaskPaths(), strStdOut, strStdErr
    Dim fDebug, strProcResult, intExitCode, intCustomErrNum, intPID, objFSO
        
    Const cEXCLUDE_HIDDEN_TASKS = 0
    Const cINCLUDE_HIDDEN_TASKS = 1
    
    fDebug = False
    
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    intCustomErrNum = VbObjectError + 1
        
    strPsExe = "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"
    strPsArg = "-command (gwmi win32_process | Where-Object processid -eq  $PID).ParentProcessID"
    
    If Not objFSO.FileExists(strPsExe) Then
        Err.Raise intCustomErrNum, WScript.ScriptName, "PowerShell executable not found: '" & strPsExe & "'."
    End If
    
    Set objShell = WScript.CreateObject("WScript.Shell")
    Set objExec = objShell.Exec(strPsExe & " " & strPsArg)
    
    strStdOut = objExec.StdOut.ReadAll()
    strStdErr = objExec.StdErr.ReadAll()
    intExitCode = objExec.ExitCode
    
    strProcResult =              "--------------------------------------------------------------------------------" & VbCrLf
    strProcResult = strProcResult & "STDOUT:" & VbCrLf
    strProcResult = strProcResult & strStdOut & VbCrLf
    strProcResult = strProcResult & "--------------------------------------------------------------------------------" & VbCrLf
    strProcResult = strProcResult & "STDERR:" & VbCrLf
    strProcResult = strProcResult & strStdErr & VbCrLf
    strProcResult = strProcResult & "--------------------------------------------------------------------------------" & VbCrLf
    strProcResult = strProcResult & "EXIT CODE: " & intExitCode & VbCrLf
    strProcResult = strProcResult & "--------------------------------------------------------------------------------" & VbCrLf
        
    If fDebug Then
        wscript.echo strProcResult
    End If
        
    If intExitCode <> 0 Or strStdErr <> "" Then
        Err.Raise intCustomErrNum, WScript.ScriptName, VbCrLf & strProcResult
    End If
    
    intPID = Cint(strStdOut)
    
    Set objScheduleService = CreateObject("Schedule.Service")
    objScheduleService.Connect()
    
    ' TaskService.GetRunningTasks method
    ' https://docs.microsoft.com/en-us/windows/win32/taskschd/taskservice-getrunningtasks
    Set taskCollection = objScheduleService.GetRunningTasks(cINCLUDE_HIDDEN_TASKS)
            
    If taskCollection.Count <> 0 Then
        i = 0
        
        If fDebug Then
            WScript.Echo "SCHEDULED TASK(S):"
        End If
        
        For Each registeredTask In taskCollection
            If fDebug Then
                WScript.Echo "Task Name: " & registeredTask.Name
                WScript.Echo "  * Path: "          & registeredTask.Path
                WScript.Echo "  * InstanceGuid: "  & registeredTask.InstanceGuid
                WScript.Echo "  * CurrentAction: " & registeredTask.CurrentAction
                WScript.Echo "  * EnginePID: "     & CInt(registeredTask.EnginePID) 
            End If
                        
            If CInt(registeredTask.EnginePID) = intPID Then
                Redim Preserve arrTaskPaths(i)
                arrTaskPaths(i) = registeredTask.Path
                i = i + 1
            End If
        Next
        
        If fDebug Then
            wscript.echo "--------------------------------------------------------------------------------"
        End If      
    End If
    
    If i = 0 Then
        GetScheduledTaskPath = ""
    Else
        GetScheduledTaskPath = Join(arrTaskPaths, ", ")
    End If
End Function

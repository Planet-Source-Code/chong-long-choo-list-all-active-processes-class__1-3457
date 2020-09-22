<div align="center">

## List All Active Processes \(Class\)


</div>

### Description

List All Active Processes (Class)
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Chong Long Choo](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/chong-long-choo.md)
**Level**          |Unknown
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Windows System Services](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-system-services__1-35.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/chong-long-choo-list-all-active-processes-class__1-3457/archive/master.zip)

### API Declarations

```
Option Explicit
' Name:   List All Active Processes
' Author:  Chong Long Choo
' Email: chonglongchoo@hotmail.com
' Date:   09 September 1999
'<--------------------------Disclaimer------------------------------->
'
'This sample is free. You can use the sample in any form. Use this
'sample at your own risk! I have no warranty for this sample.
'
'<--------------------------Disclaimer------------------------------->
'---------------------------------------------------------------------------------
'How to use
'---------------------------------------------------------------------------------
'  Dim i As Integer
'  Dim objItem As ListItem
'  Dim NumOfProcess As Long
'  Dim objActiveProcess As SQLSysInfo.clsActiveProcess
'  Set objActiveProcess = New SQLSysInfo.clsActiveProcess
'  NumOfProcess = objActiveProcess.GetActiveProcess
'  For i = 1 To NumOfProcess
'    Set objItem = ListView2.ListItems.Add(, , objActiveProcess.szExeFile(i))
'    With objItem
'      .SubItems(1) = objActiveProcess.th32ProcessID(i)
'      .SubItems(2) = objActiveProcess.th32DefaultHeapID(i)
'      .SubItems(3) = objActiveProcess.thModuleID(i)
'      .SubItems(4) = objActiveProcess.cntThreads(i)
'      .SubItems(5) = objActiveProcess.th32ParentProcessID(i)
'    End With
'  Next
'  Set objActiveProcess = Nothing
Private Const TH32CS_SNAPPROCESS As Long = 2&
Private Const MAX_PATH As Integer = 260
Private Type PROCESSENTRY32
  dwSize As Long
  cntUsage As Long
  th32ProcessID As Long
  th32DefaultHeapID As Long
  th32ModuleID As Long
  cntThreads As Long
  th32ParentProcessID As Long
  pcPriClassBase As Long
  dwFlags As Long
  szExeFile As String * MAX_PATH
End Type
Private Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, ByVal lProcessID As Long) As Long
Private Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Sub CloseHandle Lib "kernel32" (ByVal hPass As Long)
Dim ListOfActiveProcess() As PROCESSENTRY32
```


### Source Code

```
Public Function szExeFile(ByVal Index As Long) As String
  szExeFile = ListOfActiveProcess(Index).szExeFile
End Function
Public Function dwFlags(ByVal Index As Long) As Long
  dwFlags = ListOfActiveProcess(Index).dwFlags
End Function
Public Function pcPriClassBase(ByVal Index As Long) As Long
  pcPriClassBase = ListOfActiveProcess(Index).pcPriClassBase
End Function
Public Function th32ParentProcessID(ByVal Index As Long) As Long
  th32ParentProcessID = ListOfActiveProcess(Index).th32ParentProcessID
End Function
Public Function cntThreads(ByVal Index As Long) As Long
  cntThreads = ListOfActiveProcess(Index).cntThreads
End Function
Public Function thModuleID(ByVal Index As Long) As Long
  thModuleID = ListOfActiveProcess(Index).th32ModuleID
End Function
Public Function th32DefaultHeapID(ByVal Index As Long) As Long
  th32DefaultHeapID = ListOfActiveProcess(Index).th32DefaultHeapID
End Function
Public Function th32ProcessID(ByVal Index As Long) As Long
  th32ProcessID = ListOfActiveProcess(Index).th32ProcessID
End Function
Public Function cntUsage(ByVal Index As Long) As Long
  cntUsage = ListOfActiveProcess(Index).cntUsage
End Function
Public Function dwSize(ByVal Index As Long) As Long
  dwSize = ListOfActiveProcess(Index).dwSize
End Function
Public Function GetActiveProcess() As Long
  Dim hToolhelpSnapshot As Long
  Dim tProcess As PROCESSENTRY32
  Dim r As Long, i As Integer
  hToolhelpSnapshot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
  If hToolhelpSnapshot = 0 Then
    GetActiveProcess = 0
    Exit Function
  End If
  With tProcess
    .dwSize = Len(tProcess)
    r = ProcessFirst(hToolhelpSnapshot, tProcess)
    ReDim Preserve ListOfActiveProcess(20)
    Do While r
      i = i + 1
      If i Mod 20 = 0 Then ReDim Preserve ListOfActiveProcess(i + 20)
      ListOfActiveProcess(i) = tProcess
      r = ProcessNext(hToolhelpSnapshot, tProcess)
    Loop
  End With
  GetActiveProcess = i
  Call CloseHandle(hToolhelpSnapshot)
End Function
```


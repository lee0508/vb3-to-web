Attribute VB_Name = "ParaInfo"
Option Explicit

Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" _
                 (ByVal lpBuffer As String, nSize As Long) As Long
Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" _
                 (ByVal lpBuffer As String, nSize As Long) As Long
                 
'============================================================
' Registry Key
'============================================================
Const RegProName = "ParaProName"
Const RegSection = "ParaSection"

Public Type ParaInfoM
    ParaA          As String
    ParaB          As String
    ParaC          As String
    ParaD          As String
    ParaE          As String
    ParaM          As String  '"".단독실행, "0".검색할 컬럼번호
    ParaS          As String  '검색할 번호내용
    ParaX          As String  'O.실행시작, X.실행종료
End Type

Public Function ParaInfoM_Read() As ParaInfoM
Dim u As ParaInfoM
    u.ParaA = GetSetting(RegProName, RegSection, "ParaA")
    u.ParaB = GetSetting(RegProName, RegSection, "ParaB")
    u.ParaC = GetSetting(RegProName, RegSection, "ParaC")
    u.ParaD = GetSetting(RegProName, RegSection, "ParaD")
    u.ParaE = GetSetting(RegProName, RegSection, "ParaE")
    u.ParaM = GetSetting(RegProName, RegSection, "ParaM")
    u.ParaS = GetSetting(RegProName, RegSection, "ParaS")
    u.ParaX = GetSetting(RegProName, RegSection, "ParaX")
    ParaInfoM_Read = u
End Function

Public Sub ParaInfoM_Save(ByRef RegI As ParaInfoM)
    SaveSetting RegProName, RegSection, "ParaA", RegI.ParaA
    SaveSetting RegProName, RegSection, "ParaB", RegI.ParaB
    SaveSetting RegProName, RegSection, "ParaC", RegI.ParaC
    SaveSetting RegProName, RegSection, "ParaD", RegI.ParaD
    SaveSetting RegProName, RegSection, "ParaE", RegI.ParaE
    SaveSetting RegProName, RegSection, "ParaM", RegI.ParaM
    SaveSetting RegProName, RegSection, "ParaS", RegI.ParaS
    SaveSetting RegProName, RegSection, "ParaX", RegI.ParaX
End Sub

Public Sub ParaInfoM_Delete()
    On Error Resume Next
    DeleteSetting RegProName, RegSection
    Exit Sub
End Sub


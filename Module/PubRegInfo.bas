Attribute VB_Name = "PubRegInfo"
Option Explicit

'+--------------+
'| Registry Key
'+--------------+
Const RegProName = "ymh"
Const RegSection = "ymhUserinfoU"

'+----------+
'| UserinfoU
'+----------+
Public Type UserinfoU
    UserComputerName   As String   '1. WorkStation Name
    UserClientName     As String   '2. Client Wondows Login Name
    UserServerDate     As String   '3. 연결시의 서버일자
    UserServerTime     As String   '4. 연결시의 서버시간
    UserClientDate     As String   '5. 프로그램 실행일자(작업일자)
    UserClientTime     As String   '6. 프로그램 실행시간
    '+------------------------+
    UserBranchCode     As String   '7. 사업장코드
    UserBranchName     As String   '8. 사업장명
    UserCode           As String   '9. 사용자코드
    UserName           As String   '10.사용자성명
    UserLoginPasswd    As String   '11.사용자비밀번호
    UserSanctionPasswd As String   '12.사용자결재비밀번호
    UserAuthority      As String   '13.사용자권한
    '+------------------------+
    UserMJGbn          As String   '14.미지급금발생구분
    UserMSGbn          As String   '15.미지급금발생구분
End Type

Public Function UserinfoU_Read() As UserinfoU
Dim u As UserinfoU
    u.UserComputerName = GetSetting(RegProName, RegSection, "UserComputerName")
    u.UserClientName = GetSetting(RegProName, RegSection, "UserClientName")
    u.UserServerDate = GetSetting(RegProName, RegSection, "UserServerDate")
    u.UserServerTime = GetSetting(RegProName, RegSection, "UserServerTime")
    u.UserClientDate = GetSetting(RegProName, RegSection, "UserClientDate")
    u.UserClientTime = GetSetting(RegProName, RegSection, "UserClientTime")
    '+-------------------------+
    u.UserBranchCode = GetSetting(RegProName, RegSection, "UserBranchCode")
    u.UserBranchName = GetSetting(RegProName, RegSection, "UserBranchName")
    u.UserCode = GetSetting(RegProName, RegSection, "UserCode")
    u.UserName = GetSetting(RegProName, RegSection, "UserName")
    u.UserLoginPasswd = GetSetting(RegProName, RegSection, "UserLoginPasswd")
    u.UserSanctionPasswd = GetSetting(RegProName, RegSection, "UserSanctionPasswd")
    u.UserAuthority = GetSetting(RegProName, RegSection, "UserAuthority")
    '+-------------------------+
    u.UserMJGbn = GetSetting(RegProName, RegSection, "UserMJGbn")
    u.UserMSGbn = GetSetting(RegProName, RegSection, "UserMSGbn")
    UserinfoU_Read = u
End Function

Public Sub UserinfoU_Save(ByRef RegUserinfoU As UserinfoU)
    SaveSetting RegProName, RegSection, "UserComputerName", RegUserinfoU.UserComputerName
    SaveSetting RegProName, RegSection, "UserClientName", RegUserinfoU.UserClientName
    SaveSetting RegProName, RegSection, "UserServerDate", RegUserinfoU.UserServerDate
    SaveSetting RegProName, RegSection, "UserServerTime", RegUserinfoU.UserServerTime
    SaveSetting RegProName, RegSection, "UserClientDate", RegUserinfoU.UserClientDate
    SaveSetting RegProName, RegSection, "UserClientTime", RegUserinfoU.UserClientTime
    '+-------------------------+
    SaveSetting RegProName, RegSection, "UserBranchCode", RegUserinfoU.UserBranchCode
    SaveSetting RegProName, RegSection, "UserBranchName", RegUserinfoU.UserBranchName
    SaveSetting RegProName, RegSection, "UserCode", RegUserinfoU.UserCode
    SaveSetting RegProName, RegSection, "UserName", RegUserinfoU.UserName
    SaveSetting RegProName, RegSection, "UserLoginPasswd", RegUserinfoU.UserLoginPasswd
    SaveSetting RegProName, RegSection, "UserSanctionPasswd", RegUserinfoU.UserSanctionPasswd
    SaveSetting RegProName, RegSection, "UserAuthority", RegUserinfoU.UserAuthority
    '+-------------------------+
    SaveSetting RegProName, RegSection, "UserMJGbn", RegUserinfoU.UserMJGbn
    SaveSetting RegProName, RegSection, "UserMSGbn", RegUserinfoU.UserMSGbn
End Sub

Public Sub UserinfoU_Delete()
    On Error Resume Next
    DeleteSetting RegProName, RegSection
    Exit Sub
End Sub

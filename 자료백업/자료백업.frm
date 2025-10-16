VERSION 5.00
Begin VB.Form frm자료백업 
   BorderStyle     =   1  '단일 고정
   Caption         =   "자료 백업"
   ClientHeight    =   2040
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3255
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "자료백업.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   3255
   StartUpPosition =   1  '소유자 가운데
   Begin VB.Frame Frame1 
      Caption         =   "[ 백업명 선택 ]"
      Height          =   1335
      Left            =   120
      TabIndex        =   2
      ToolTipText     =   "백업될 파일의 이름 형태를 지정"
      Top             =   100
      Width           =   3015
      Begin VB.OptionButton optName 
         Caption         =   "yyyy/mm/dd hh:mm:ss (초)"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   5
         Top             =   960
         Width           =   2535
      End
      Begin VB.OptionButton optName 
         Caption         =   "yyyy/mm/dd hh       (시)"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   4
         Top             =   600
         Width           =   2535
      End
      Begin VB.OptionButton optName 
         Caption         =   "yyyy/mm/dd          (일)"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Value           =   -1  'True
         Width           =   2535
      End
   End
   Begin VB.CommandButton cmdExec 
      Caption         =   "실행(&E)"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   1560
      Width           =   1455
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "종료(&X)"
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   1560
      Width           =   1455
   End
End
Attribute VB_Name = "frm자료백업"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'+-------------------------------------------------------------------------------------------------------+
' 프로그램 제 목 : 자료백업
' 사용된 Control :
' 참조된 Table   :
' 업  무  설  명 :
'+-------------------------------------------------------------------------------------------------------+
Option Explicit
Private P_blnActived       As Boolean
Private P_adoRec           As New ADODB.Recordset
Private P_intButton        As Integer
Private Const PC_intRowCnt As Integer = 22  '그리드 한 페이지 당 행수(FixedRows 포함)

'+--------------------------------+
'/// LOAD FORM ( 한번만 실행 ) ///
'+--------------------------------+
Private Sub Form_Load()
    P_blnActived = False
End Sub

'+-------------------------------------------+
'/// ACTIVATE FORM 활성화 ( 한번만 실행 ) ///
'+-------------------------------------------+
Private Sub Form_Activate()
Dim strSQL      As String
Dim strDateTime As String '서버일시
    If P_blnActived = False Then
       Screen.MousePointer = vbHourglass
       On Error GoTo ERROR_FORM_ACTIVATE_CONNECTION_SERVER
       P_blnActived = True
       Screen.MousePointer = vbDefault
    End If
    Exit Sub
ERROR_FORM_ACTIVATE_CONNECTION_SERVER:
    MsgBox Err.Number & Err.Description & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "(서버와의 연결 실패)"
    Unload Me
    Exit Sub
End Sub

'+-----------+
'/// 실행 ///
'+-----------+
Private Sub cmdExec_Click()
Dim strSQL            As String
Dim lngR              As Long
Dim lngC              As Long
Dim blnOK             As Boolean
Dim intRetVal         As Integer
Dim strNextYM         As String
Dim strServerDateTime As String
Dim strExecDateTime   As String
Dim strToBackUpPath   As String
Dim strToBackUpFile   As String
Dim intParR_Status    As Integer
Dim strParR_MSG       As String
    P_adoRec.CursorLocation = adUseClient
    'PATH
    strSQL = "SELECT ISNULL(T1.백업폴더, '') AS 백업폴더 FROM 사업장 T1 WHERE T1.사업장코드 = '" & PB_regUserinfoU.UserBranchCode & "' "
    On Error GoTo ERROR_FORM_ACTIVATE_CONNECTION_SERVER
    P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
    If P_adoRec.RecordCount = 1 Then
       strToBackUpPath = P_adoRec("백업폴더")
    End If
    P_adoRec.Close
    If Len(strToBackUpPath) = 0 Then
       MsgBox "사업장정보의 백업폴더를 먼저확인하세요!", vbCritical + vbOKOnly, "백업폴더 오류"
       Exit Sub
    End If
    '서버일시
    strSQL = "SELECT CONVERT(VARCHAR(19),GETDATE(), 120) AS 서버일시 "
    On Error GoTo ERROR_FORM_ACTIVATE_CONNECTION_SERVER
    P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
    strServerDateTime = P_adoRec("서버일시")
    P_adoRec.Close
    For lngR = optName.LBound To optName.UBound
        If optName(lngR).Value = True Then
           If lngR = 0 Then '일
              strServerDateTime = Format(DTOS(Mid(strServerDateTime, 1, 10)), "0000/00/00")
              strExecDateTime = DTOS(Mid(strServerDateTime, 1, 10))
           ElseIf _
              lngR = 1 Then '시
              strServerDateTime = Format(DTOS(Mid(strServerDateTime, 1, 10)), "0000/00/00") + " " + Mid(strServerDateTime, 12, 2)
              strExecDateTime = DTOS(Mid(strServerDateTime, 1, 10)) + Mid(strServerDateTime, 12, 2)
           Else             '초
              strServerDateTime = Format(DTOS(Mid(strServerDateTime, 1, 10)), "0000/00/00") + " " + Right(strServerDateTime, 8)
              strExecDateTime = DTOS(Mid(strServerDateTime, 1, 10)) + Format(Right(strServerDateTime, 8), "hhmmss")
           End If
        End If
    Next lngR
    intRetVal = MsgBox("[" + strServerDateTime + "] 자료를 백업하시겠습니까 ?", vbQuestion + vbYesNo + vbDefaultButton2, "자료백업")
    If intRetVal = vbNo Then
       Exit Sub
    End If
    cmdExec.Enabled = False
    Screen.MousePointer = vbHourglass
    PB_adoCnnSQL.BeginTrans
    '자료백업
    strToBackUpFile = strExecDateTime + ".bak"
    strSQL = "spBackUpYmhDB '" & strToBackUpPath & "', '" & strToBackUpFile & "', '', 0, '' "
    On Error GoTo ERROR_STORED_PROCEDURE
    P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
    If P_adoRec.RecordCount = 0 Then
    Else
       intParR_Status = P_adoRec(0)
       strParR_MSG = P_adoRec(1)
    End If
    P_adoRec.Close
    '결과 보기
    MsgBox strParR_MSG, IIf(intParR_Status = 0, vbCritical, vbInformation), "자료 백업 결과"
    PB_adoCnnSQL.CommitTrans
    cmdExec.Enabled = True
    Screen.MousePointer = vbDefault
    Exit Sub
ERROR_FORM_ACTIVATE_CONNECTION_SERVER:
    MsgBox Err.Number & Err.Description & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "(서버와의 연결 실패)"
    Unload Me
    Exit Sub
ERROR_STORED_PROCEDURE:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Number & Err.Description & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "자료 백업 실패"
    Unload Me
    Exit Sub
End Sub

'+-----------+
'/// 종료 ///
'+-----------+
Private Sub cmdExit_Click()
    Unload Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Screen.MousePointer = vbDefault
    If P_adoRec.State <> adStateClosed Then
       P_adoRec.Close
    End If
    Set P_adoRec = Nothing
    Set frm자료백업 = Nothing
End Sub

'+--------------------+
'/// Sub Procedure ///
'+--------------------+

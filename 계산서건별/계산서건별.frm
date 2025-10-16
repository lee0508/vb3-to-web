VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm계산서건별 
   BorderStyle     =   1  '단일 고정
   Caption         =   "세금계산서(건별)처리"
   ClientHeight    =   5790
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12255
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "계산서건별.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5792.867
   ScaleMode       =   0  '사용자
   ScaleWidth      =   12255
   StartUpPosition =   2  '화면 가운데
   Begin VB.Frame Frame1 
      Appearance      =   0  '평면
      ForeColor       =   &H80000008&
      Height          =   675
      Left            =   60
      TabIndex        =   20
      Top             =   0
      Width           =   12075
      Begin VB.ComboBox cboPrinter 
         Height          =   300
         Left            =   8640
         Style           =   2  '드롭다운 목록
         TabIndex        =   44
         Top             =   240
         Width           =   3315
      End
      Begin VB.ListBox lstPort 
         Height          =   240
         Left            =   3960
         TabIndex        =   30
         Top             =   240
         Visible         =   0   'False
         Width           =   495
      End
      Begin Crystal.CrystalReport CrystalReport1 
         Left            =   7920
         Top             =   195
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowBorderStyle=   0
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         PrintFileLinesPerPage=   60
         WindowShowGroupTree=   -1  'True
      End
      Begin VB.Label Label2 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00008000&
         BorderStyle     =   1  '단일 고정
         Caption         =   "세금계산서(건별)처리"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   18
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   420
         Left            =   3765
         TabIndex        =   21
         Top             =   180
         Width           =   4650
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  '평면
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   5115
      Left            =   60
      TabIndex        =   19
      Top             =   630
      Width           =   12075
      Begin VB.TextBox Text1 
         Appearance      =   0  '평면
         Height          =   285
         Index           =   10
         Left            =   1680
         TabIndex        =   16
         Top             =   4000
         Width           =   5175
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  '평면
         Enabled         =   0   'False
         Height          =   285
         IMEMode         =   10  '한글 
         Index           =   9
         Left            =   6840
         TabIndex        =   15
         Top             =   3600
         Width           =   2415
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  '오른쪽 맞춤
         Appearance      =   0  '평면
         Height          =   285
         IMEMode         =   8  '영문
         Index           =   4
         Left            =   4920
         TabIndex        =   8
         Top             =   2100
         Width           =   1815
      End
      Begin VB.CommandButton cmdExit 
         Height          =   390
         Left            =   4080
         Picture         =   "계산서건별.frx":030A
         Style           =   1  '그래픽
         TabIndex        =   43
         Top             =   4560
         Width           =   1095
      End
      Begin VB.CommandButton cmdSave 
         Height          =   390
         Left            =   2880
         Picture         =   "계산서건별.frx":0C58
         Style           =   1  '그래픽
         TabIndex        =   17
         Top             =   4560
         Width           =   1095
      End
      Begin VB.CommandButton cmdClear 
         Height          =   375
         Left            =   1680
         Picture         =   "계산서건별.frx":14DF
         Style           =   1  '그래픽
         TabIndex        =   18
         Top             =   4560
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         Height          =   285
         IMEMode         =   8  '영문
         Index           =   6
         Left            =   5760
         MaxLength       =   4
         TabIndex        =   10
         Top             =   2450
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         Height          =   285
         IMEMode         =   8  '영문
         Index           =   7
         Left            =   1680
         MaxLength       =   1
         TabIndex        =   11
         Top             =   2800
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         Height          =   285
         IMEMode         =   8  '영문
         Index           =   8
         Left            =   1680
         MaxLength       =   1
         TabIndex        =   12
         Top             =   3200
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  '평면
         Height          =   285
         Index           =   5
         Left            =   1680
         TabIndex        =   9
         Top             =   2450
         Width           =   3495
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  '오른쪽 맞춤
         Appearance      =   0  '평면
         Height          =   285
         IMEMode         =   8  '영문
         Index           =   3
         Left            =   1680
         TabIndex        =   7
         Top             =   2100
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         Height          =   285
         IMEMode         =   8  '영문
         Index           =   0
         Left            =   1680
         MaxLength       =   1
         TabIndex        =   0
         Top             =   300
         Width           =   735
      End
      Begin VB.CheckBox chkPrint 
         Caption         =   "저장후 인쇄"
         Height          =   375
         Left            =   240
         TabIndex        =   28
         Top             =   4575
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  '평면
         Enabled         =   0   'False
         Height          =   285
         IMEMode         =   8  '영문
         Index           =   2
         Left            =   3720
         TabIndex        =   5
         Top             =   1440
         Width           =   3015
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         Height          =   285
         IMEMode         =   8  '영문
         Index           =   1
         Left            =   2400
         MaxLength       =   8
         TabIndex        =   4
         Top             =   1440
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker dtpF_Date 
         Height          =   270
         Left            =   2160
         TabIndex        =   2
         Top             =   1080
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   476
         _Version        =   393216
         Format          =   56623105
         UpDown          =   -1  'True
         CurrentDate     =   37763
      End
      Begin MSComCtl2.DTPicker dtpT_Date 
         Height          =   270
         Left            =   4320
         TabIndex        =   3
         Top             =   1080
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   476
         _Version        =   393216
         Format          =   56623105
         UpDown          =   -1  'True
         CurrentDate     =   37763
      End
      Begin MSComCtl2.DTPicker dtpP_Date 
         Height          =   270
         Left            =   1680
         TabIndex        =   1
         Top             =   720
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   476
         _Version        =   393216
         Format          =   56623105
         UpDown          =   -1  'True
         CurrentDate     =   37763
      End
      Begin MSComCtl2.DTPicker dtpS_Date 
         Height          =   270
         Left            =   1680
         TabIndex        =   13
         Top             =   3600
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   476
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   56623105
         UpDown          =   -1  'True
         CurrentDate     =   37763
      End
      Begin MSComCtl2.DTPicker dtpM_Date 
         Height          =   270
         Left            =   4320
         TabIndex        =   14
         Top             =   3600
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   476
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   56623105
         UpDown          =   -1  'True
         CurrentDate     =   37763
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "적요"
         Height          =   240
         Index           =   14
         Left            =   360
         TabIndex        =   49
         Top             =   4050
         Width           =   1095
      End
      Begin VB.Label lblMDate 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "만기일자"
         Height          =   240
         Left            =   3000
         TabIndex        =   48
         Top             =   3645
         Width           =   1095
      End
      Begin VB.Label lblBillNo 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "어음번호"
         Height          =   240
         Left            =   5640
         TabIndex        =   47
         Top             =   3645
         Width           =   1095
      End
      Begin VB.Label lblSDate 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "수금일자"
         Height          =   240
         Left            =   360
         TabIndex        =   46
         Top             =   3650
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "부가세액"
         Height          =   240
         Index           =   7
         Left            =   3600
         TabIndex        =   45
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Line Line1 
         X1              =   240
         X2              =   11760
         Y1              =   4440
         Y2              =   4440
      End
      Begin VB.Label Label1 
         Caption         =   "(0.현금, 1.수표, 2.어음, 3.외상미수금)"
         Height          =   240
         Index           =   13
         Left            =   2640
         TabIndex        =   42
         Top             =   2880
         Width           =   3615
      End
      Begin VB.Label Label1 
         Alignment       =   2  '가운데 맞춤
         Caption         =   "종"
         Height          =   240
         Index           =   12
         Left            =   6480
         TabIndex        =   41
         Top             =   2520
         Width           =   375
      End
      Begin VB.Label Label1 
         Alignment       =   2  '가운데 맞춤
         Caption         =   "외"
         Height          =   240
         Index           =   8
         Left            =   5280
         TabIndex        =   40
         Top             =   2520
         Width           =   375
      End
      Begin VB.Label lblJanMny 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "0"
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   1680
         TabIndex        =   6
         Top             =   1800
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "(0.영수함, 1.청구함)"
         Height          =   240
         Index           =   21
         Left            =   2640
         TabIndex        =   39
         Top             =   3240
         Width           =   2055
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "위금액을"
         Height          =   240
         Index           =   1
         Left            =   360
         TabIndex        =   38
         Top             =   3240
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "구분"
         Height          =   240
         Index           =   9
         Left            =   360
         TabIndex        =   37
         Top             =   2880
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "품명및규격"
         Height          =   240
         Index           =   20
         Left            =   360
         TabIndex        =   36
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "계산서금액"
         Height          =   240
         Index           =   19
         Left            =   360
         TabIndex        =   35
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "현장부잔액"
         Height          =   240
         Index           =   18
         Left            =   360
         TabIndex        =   34
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "매출처코드"
         Height          =   240
         Index           =   0
         Left            =   360
         TabIndex        =   33
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "발행일자"
         Height          =   240
         Index           =   17
         Left            =   360
         TabIndex        =   32
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "(0.거래, 1.임의)"
         Height          =   360
         Index           =   11
         Left            =   2280
         TabIndex        =   31
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "발행구분"
         Height          =   240
         Index           =   10
         Left            =   360
         TabIndex        =   29
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "까지"
         Height          =   240
         Index           =   6
         Left            =   5760
         TabIndex        =   27
         Top             =   1125
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "부터"
         Height          =   240
         Index           =   5
         Left            =   3600
         TabIndex        =   26
         Top             =   1125
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "]"
         Height          =   240
         Index           =   4
         Left            =   6675
         TabIndex        =   25
         Top             =   1125
         Width           =   135
      End
      Begin VB.Label Label1 
         Caption         =   "["
         Height          =   240
         Index           =   3
         Left            =   1680
         TabIndex        =   24
         Top             =   1125
         Width           =   135
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "발행기간"
         Height          =   240
         Index           =   2
         Left            =   360
         TabIndex        =   23
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "[Home]"
         Height          =   240
         Index           =   15
         Left            =   1680
         TabIndex        =   22
         Top             =   1485
         Width           =   615
      End
   End
End
Attribute VB_Name = "frm계산서건별"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'+-------------------------------------------------------------------------------------------------------+
' 프로그램 제 목 : 계산서건별
' 사용된 Control : VideoSoft VSFlexGrid 7.0(OLEDB) = vsflex7.ocx
' 참조된 Table   : 사업장, 매출처, 자재입출내역, 세금계산서
' 업  무  설  명 : 입출고구분(8.미수금발생금액에만 포함)
'+-------------------------------------------------------------------------------------------------------+
Option Explicit
Private P_blnActived         As Boolean
Private P_adoRec             As New ADODB.Recordset
Private P_adoRecW            As New ADODB.Recordset
Private P_intButton          As Integer
Private P_strFindString2     As String
Private Const PC_intRowCnt1  As Integer = 25   '그리드1의 한 페이지 당 행수(FixedRows 포함)

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
Dim strSQL             As String
Dim inti               As Integer

Dim p                  As Printer
Dim strDefaultPrinter  As String
Dim aryPrinter()       As String
Dim strBuffer          As String

    frmMain.SBar.Panels(4).Text = "현금은 집계에서 제외, 매출세금계산서장부에도 저장, 3.외상미수금이 아닌경우 미수금 수금내역에 저장됩니다. "
    If P_blnActived = False Then
       Screen.MousePointer = vbHourglass
       On Error GoTo ERROR_FORM_ACTIVATE_CONNECTION_SERVER
       
       'Printer Setting and Seaching(API)
       strBuffer = Space(1024)
       inti = GetProfileString("windows", "Device", "", strBuffer, Len(strBuffer))
       aryPrinter = Split(strBuffer, ",")
       strDefaultPrinter = Trim(Trim(aryPrinter(0)))
       For Each p In Printers
           cboPrinter.AddItem Trim(p.DeviceName)
           lstPort.AddItem p.Port
       Next
       For inti = 0 To cboPrinter.ListCount - 1
           cboPrinter.ListIndex = inti
           If UCase(Trim(cboPrinter.Text)) = UCase(Trim(strDefaultPrinter)) Then
              Exit For
           End If
       Next inti
       '---
       Select Case Val(PB_regUserinfoU.UserAuthority)
              'Case Is <= 10 '조회
              '     cmdPrint.Enabled = False: cmdClear.Enabled = False: cmdFind.Enabled = True
              '     cmdSave.Enabled = False: cmdDelete.Enabled = False: cmdExit.Enabled = True
              'Case Is <= 20 '인쇄, 조회
              '     cmdPrint.Enabled = True: cmdClear.Enabled = False: cmdFind.Enabled = True
              '     cmdSave.Enabled = False: cmdDelete.Enabled = False: cmdExit.Enabled = True
              'Case Is <= 40 '추가, 저장, 인쇄, 조회
              '     cmdPrint.Enabled = True: cmdClear.Enabled = False: cmdFind.Enabled = True
              '     cmdSave.Enabled = True: cmdDelete.Enabled = False: cmdExit.Enabled = True
              'Case Is <= 50 '삭제, 추가, 저장, 조회, 인쇄
              '     cmdPrint.Enabled = True: cmdClear.Enabled = False: cmdFind.Enabled = True
              '     cmdSave.Enabled = True: cmdDelete.Enabled = True: cmdExit.Enabled = True
              'Case Is <= 99 '삭제, 추가, 저장, 조회, 인쇄
              '     cmdPrint.Enabled = True: cmdClear.Enabled = False: cmdFind.Enabled = True
              '     cmdSave.Enabled = True: cmdDelete.Enabled = True: cmdExit.Enabled = True
              'Case Else
              '     cmdPrint.Enabled = False: cmdClear.Enabled = False: cmdFind.Enabled = False
              '     cmdSave.Enabled = False: cmdDelete.Enabled = False: cmdExit.Enabled = True
       End Select
       SubOther_FILL
       P_blnActived = True
       Screen.MousePointer = vbDefault
    End If
    Exit Sub
ERROR_FORM_ACTIVATE_CONNECTION_SERVER:
    MsgBox Err.Number & Err.Description & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "계산서건별(서버와의 연결 실패)"
    Unload Me
    Exit Sub
End Sub

'+--------------------+
'--- Select Printer ---
'+--------------------+
Private Sub cboPrinter_Click()
    lstPort.ListIndex = cboPrinter.ListIndex
End Sub

'+-------------------+
'/// Text1(index) ///
'+-------------------+
Private Sub Text1_GotFocus(Index As Integer)
    With Text1(Index)
         .SelStart = 0
         .SelLength = Len(Trim(.Text))
    End With
End Sub
Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim strSQL     As String
Dim strExeFile As String
Dim varRetVal  As Variant
    If (Index = 1 And (KeyCode = vbKeyHome Or KeyCode = vbKeyReturn)) Then  '매출처검색
       PB_strSupplierCode = UPPER(Trim(Text1(Index).Text))
       PB_strSupplierName = "" 'Trim(Text1(Index + 1).Text)
       frm매출처검색.Show vbModal
       If (Len(PB_strSupplierCode) + Len(PB_strSupplierName)) = 0 Then '검색에서 취소(ESC)
       Else
          Text1(Index).Text = PB_strSupplierCode
          Text1(Index + 1).Text = PB_strSupplierName
       End If
       If PB_strSupplierCode <> "" Then
          SendKeys "{TAB}"
       End If
       PB_strSupplierCode = "": PB_strSupplierName = ""
    Else
       If KeyCode = vbKeyReturn Then
          Select Case Index
          End Select
          SendKeys "{tab}"
       End If
    End If
    Exit Sub
ERROR_TABLE_SELECT:
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "매출처 읽기 실패"
    Unload Me
    Exit Sub
End Sub

'+-------------+
'/// 매출처 ///
'+-------------+
Private Sub Text1_LostFocus(Index As Integer)
Dim strSQL   As String
Dim lngR     As Long
Dim intIndex As Integer
    intIndex = Index
    With Text1(Index)
         .Text = Trim(.Text)
         Select Case Index
                Case 0 '발행구분
                     If Trim(.Text) = "0" Then '거래
                        If Vals(Text1(3).Text) = 0 Then
                           Text1(3).Enabled = False: Text1(4).Enabled = False
                        Else
                           Text1(3).Enabled = True: Text1(4).Enabled = True
                        End If
                     ElseIf _
                        Trim(.Text) = "1" Then '임의
                        Text1(3).Enabled = True: Text1(4).Enabled = True
                     Else
                        Text1(3).Enabled = False: Text1(4).Enabled = False
                     End If
                Case 1 '매출처코드
                     .Text = UPPER(Trim(.Text))
                     If Len(Trim(.Text)) < 1 Then
                        .Text = ""
                        Text1(2).Text = ""
                        Exit Sub
                     End If
                     'P_adoRec.CursorLocation = adUseClient
                     'strSQL = "SELECT * FROM 매출처 " _
                             & "WHERE 사업장코드 = '" & PB_regUserinfoU.UserBranchCode & "' " _
                               & "AND 매출처코드 = '" & Trim(.Text) & "' "
                     'On Error GoTo ERROR_TABLE_SELECT
                     'P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
                     'If P_adoRec.RecordCount = 0 Then
                     '   P_adoRec.Close
                     '   .Text = ""
                     '   .SetFocus
                     '   Exit Sub
                     'End If
                     'Text1(2).Text = P_adoRec("매출처명")
                     'P_adoRec.Close
                     SubCompute_JanMny PB_regUserinfoU.UserBranchCode, PB_regUserinfoU.UserClientDate, Text1(1).Text '미수금잔액
                     If Trim(Text1(0).Text) = "0" Then '거래(세금계산서금액)
                        SubCompute_TaxBillMny PB_regUserinfoU.UserBranchCode, DTOS(dtpF_Date.Value), _
                                              DTOS(dtpT_Date.Value), Trim(Text1(1).Text)
                        If Vals(Text1(3).Text) = 0 Then
                           Text1(3).Enabled = False: Text1(4).Enabled = False
                        Else
                           Text1(3).Enabled = True: Text1(4).Enabled = True
                        End If
                     End If
                Case 3 '세금계산서금액(공급가액)
                     .Text = Format(Vals(.Text), "#,0.00")
                     Text1(4).Text = Format(Fix(Vals(.Text) * (PB_curVatRate)), "#,0.00")
                Case 4 '세금계산서금액(부가세)
                     .Text = Format(Vals(.Text), "#,0.00")
                Case 7 '구분(7), 영수구분(8)
                     dtpS_Date.Enabled = False: dtpM_Date.Enabled = False: Text1(9).Enabled = False
                     If Text1(7).Text = "0" Or Text1(7).Text = "1" Then   '현금 또는 수표
                        Text1(8).Text = "0"
                        dtpS_Date.Enabled = True: dtpM_Date.Enabled = False: Text1(9).Enabled = False
                     ElseIf _
                        Text1(7).Text = "2" Then                           '어음
                        Text1(8).Text = "0"
                        dtpS_Date.Enabled = True: dtpM_Date.Enabled = True: Text1(9).Enabled = True
                     Else
                        Text1(7).Text = "3"
                        Text1(8).Text = "1"
                     End If
         End Select
    End With
    Exit Sub
ERROR_TABLE_SELECT:
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "TABLE 읽기 실패"
    Unload Me
    Exit Sub
End Sub

'+---------------+
'/// 발행일자 ///
'+---------------+
Private Sub dtpP_Date_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub
Private Sub dtpP_Date_LostFocus()
    dtpS_Date.Value = dtpP_Date.Value
End Sub
'+---------------+
'/// 발행기간 ///
'+---------------+
Private Sub dtpF_Date_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub
Private Sub dtpT_Date_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
       SendKeys "{tab}"
    End If
End Sub
'+---------------+
'/// 수금일자 ///
'+---------------+
Private Sub dtpS_Date_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub
'+---------------+
'/// 만기일자 ///
'+---------------+
Private Sub dtpM_Date_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub

'+-----------+
'/// 추가 ///
'+-----------+
Private Sub cmdClear_Click()
    SubClearText
    Text1(Text1.LBound).SetFocus
End Sub

'+-----------+
'/// 저장 ///
'+-----------+
Private Sub cmdSave_Click()
Dim p              As Printer
Dim blnSaveOK      As Boolean
Dim strSQL         As String
Dim lngR           As Long
Dim lngC           As Long
Dim blnOK          As Boolean
Dim intRetVal      As Integer
Dim lngLogCnt      As Long
Dim strMakeYear    As String
Dim lngLogCnt1     As Long
Dim lngLogCnt2     As Long
Dim strJukyo       As String  '적요(자재입출내역. 미수금내역)
Dim intSactionWay  As Integer '결제방법(미수금내역)
Dim curInputMoney  As Long
Dim curOutputMoney As Long
Dim strServerTime  As String
Dim strTime        As String
Dim intKijang      As Integer '1.기장
    
    For Each p In Printers
        If Trim(p.DeviceName) = Trim(cboPrinter.Text) And lstPort.List(cboPrinter.ListIndex) = p.Port Then
           Set Printer = p
           Exit For
        End If
    Next
    
    If dtpF_Date.Value > dtpT_Date.Value Then
       dtpF_Date.SetFocus
       Exit Sub
    End If
    '입력내역 검사
    FncCheckTextBox lngC, blnOK
    If blnOK = False Then
       If Text1(lngC).Enabled = False Then
          Text1(lngC).Enabled = True
       End If
       Text1(lngC).SetFocus
       Exit Sub
    End If
    intRetVal = MsgBox("세금계산서를 저장하시겠습니까 ?", vbQuestion + vbYesNo + vbDefaultButton1, "자료 저장")
    If intRetVal = vbNo Then
       Exit Sub
    End If
    
    '서버시간 구하기
    cmdSave.Enabled = False
    Screen.MousePointer = vbHourglass
    P_adoRec.CursorLocation = adUseClient
    strSQL = "SELECT RIGHT(CONVERT(VARCHAR(23),GETDATE(), 121),12) AS 서버시간 "
    On Error GoTo ERROR_FORM_ACTIVATE_CONNECTION_SERVER
    P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
    strServerTime = Mid(P_adoRec("서버시간"), 1, 2) + Mid(P_adoRec("서버시간"), 4, 2) + Mid(P_adoRec("서버시간"), 7, 2) _
                  + Mid(P_adoRec("서버시간"), 10)
    P_adoRec.Close
    strTime = strServerTime
    '책번호, 일련번호 구하기
    PB_adoCnnSQL.BeginTrans
    strMakeYear = Mid(DTOS(dtpP_Date.Value), 1, 4)
    strSQL = "spLogCounter '세금계산서', '" & PB_regUserinfoU.UserBranchCode + strMakeYear & "', " _
                         & "0, 0, '" & PB_regUserinfoU.UserServerDate & "', '" & PB_regUserinfoU.UserCode & "' "
    On Error GoTo ERROR_STORED_PROCEDURE
    P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
    lngLogCnt1 = P_adoRec(0)
    lngLogCnt2 = P_adoRec(1)
    P_adoRec.Close
    '세금계산서 추가
    strSQL = "INSERT INTO 세금계산서(사업장코드, 작성년도, 책번호, 일련번호, " _
                                  & "매출처코드, 작성일자, 품목및규격, 수량, " _
                                  & "공급가액, 세액, 금액구분, 영청구분, " _
                                  & "발행여부, 작성구분, 미수구분, 적요, 사용구분, " _
                                  & "수정일자, 사용자코드) VALUES(" _
    & "'" & PB_regUserinfoU.UserBranchCode & "', '" & strMakeYear & "', " & lngLogCnt1 & ", " & lngLogCnt2 & "," _
    & "'" & Trim(Text1(1).Text) & "', '" & DTOS(dtpP_Date.Value) & "','" & Trim(Text1(5).Text) & "'," & Vals(Text1(6).Text) & ", " _
    & "" & Vals(Text1(3).Text) & ", " & Vals(Text1(4).Text) & ", " & Vals(Text1(7).Text) & "," & Vals(Text1(8).Text) & ", " _
    & "0, " & Vals(Text1(0).Text) & ", 1, '" & Text1(10).Text & "', 0, " _
    & "'" & PB_regUserinfoU.UserServerDate & "','" & PB_regUserinfoU.UserCode & "') "
    On Error GoTo ERROR_TABLE_INSERT
    PB_adoCnnSQL.Execute strSQL
    If Trim(Text1(0).Text = "0") Then '거래분 세금계산서이면 (자재입출내역 변경)
       strSQL = "UPDATE 자재입출내역 SET " _
                     & "계산서발행여부 = 1, " _
                     & "작성년도 = '" & strMakeYear & "', " _
                     & "책번호 = " & lngLogCnt1 & ", " _
                     & "일련번호 = " & lngLogCnt2 & " " _
               & "WHERE 사업장코드 = '" & PB_regUserinfoU.UserBranchCode & "' AND 매출처코드 = '" & Trim(Text1(1).Text) & "' " _
                 & "AND 입출고일자 BETWEEN '" & DTOS(dtpF_Date.Value) & "' AND '" & DTOS(dtpT_Date.Value) & "' " _
                 & "AND 입출고구분 = 2 AND 사용구분 = 0 AND 현금구분 = 0 AND 계산서발행여부 = 0 "
       On Error GoTo ERROR_TABLE_UPDATE
       PB_adoCnnSQL.Execute strSQL
    Else                                 '임의분 세금계산서
       intRetVal = MsgBox("미수금에 기장하겠습니까 ?", vbCritical + vbYesNo + vbDefaultButton1, "미수금 기장")
       If intRetVal = vbYes Then
          intKijang = 1
       End If
       If Vals(Trim(Text1(6).Text)) = 0 Then '수량
          strJukyo = Trim(Text1(5).Text)
       Else
          strJukyo = Trim(Text1(5).Text) & " 외 " + CStr(Vals(Trim(Text1(6).Text))) + "종"
       End If
       '거래번호 구하기
       'P_adoRec.CursorLocation = adUseClient
       'strSQL = "spLogCounter '자재입출내역', '" & PB_regUserinfoU.UserBranchCode + DTOS(dtpP_Date.Value) + "2" & "', " _
       '                        & "0, 0, '" & PB_regUserinfoU.UserServerDate & "', '" & PB_regUserinfoU.UserCode & "' "
       'On Error GoTo ERROR_STORED_PROCEDURE
       'P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
       'lngLogCnt = P_adoRec(0)
       'P_adoRec.Close
       If intKijang = 1 Then
          'strSQL = "INSERT INTO 자재입출내역(사업장코드, 분류코드, 세부코드, 입출고구분," _
                                          & "입출고일자, 입출고시간, 입고수량, 입고단가," _
                                          & "입고부가, 출고수량, 출고단가, 출고부가," _
                                          & "매입처코드, 매출처코드, 원래입출고일자, 직송구분," _
                                          & "발견일자, 발견번호, 거래일자, 거래번호," _
                                          & "계산서발행여부, 현금구분, 감가구분, 적요, 작성년도, 책번호," _
                                          & "일련번호, 사용구분, 수정일자, 사용자코드, 재고이동사업장코드) VALUES(" _
                 & "'" & PB_regUserinfoU.UserBranchCode & "', '', '', 8," _
                 & "'" & DTOS(dtpP_Date.Value) & "','" & strServerTime & "', 0, 0," _
                 & "0, " & IIf(Vals(Text1(3).Text) < 0, -1, 1) & ", " & Abs(Vals(Text1(3).Text)) & ", " & Abs(Vals(Text1(4).Text)) & "," _
                 & "'', '" & Trim(Text1(1).Text) & "', '" & DTOS(dtpP_Date.Value) & "', 0," _
                 & "'', 0, '" & DTOS(dtpP_Date.Value) & "', " & lngLogCnt & "," _
                 & "1, " & IIf(Vals(Trim(Text1(7).Text)) = 0 Or Vals(Trim(Text1(7).Text)) = 1, 1, 0) & ", 0, '" & strJukyo & "', " _
                 & "'" & strMakeYear & "', " & lngLogCnt1 & ", " _
                 & "" & lngLogCnt2 & ", 0,'" & PB_regUserinfoU.UserServerDate & "','" & PB_regUserinfoU.UserCode & "', '')"
          'On Error GoTo ERROR_TABLE_INSERT
          'PB_adoCnnSQL.Execute strSQL
       End If
       strSQL = "UPDATE 세금계산서 SET " _
                     & "미수구분 = " & intKijang & " " _
               & "WHERE 사업장코드 = '" & PB_regUserinfoU.UserBranchCode & "' " _
                 & "AND 작성년도 = '" & strMakeYear & "' " _
                 & "AND 책번호 = " & lngLogCnt1 & " " _
                 & "AND 일련번호 = " & lngLogCnt2 & " "
       On Error GoTo ERROR_TABLE_UPDATE
       PB_adoCnnSQL.Execute strSQL
    End If
    'if (거래 AND 영수함) OR (임의 AND 영수함 AND 기장함) then  '미수금입금내역 추가
    If (Trim(Text1(0).Text = "0") And Trim(Text1(8).Text) = "0") Or _
       (Trim(Text1(0).Text = "1") And Trim(Text1(8).Text) = "0" And intKijang = 1) Then
       If Vals(Trim(Text1(7).Text)) = 0 Or Vals(Trim(Text1(7).Text)) = 1 Then '현금 또는 수표
          intSactionWay = 0
       ElseIf _
          Vals(Trim(Text1(7).Text)) = 2 Then '어음
          intSactionWay = 1
       End If
       strJukyo = PB_regUserinfoU.UserBranchCode + "-" + strMakeYear + "-" + CStr(lngLogCnt1) + "-" + CStr(lngLogCnt2)
       strSQL = "INSERT INTO 미수금내역(사업장코드, 매출처코드, " _
                                     & "미수금입금일자, 미수금입금시간," _
                                     & "미수금입금금액, 결제방법, " _
                                     & "만기일자, 어음번호, " _
                                     & "적요, 수정일자, " _
                                     & "사용자코드) VALUES(" _
                        & "'" & PB_regUserinfoU.UserBranchCode & "', '" & Trim(Text1(1).Text) & "', " _
                        & "'" & DTOS(dtpS_Date.Value) & "', '" & strServerTime & "', " _
                        & "" & (Vals(Text1(3).Text) + Vals(Text1(4).Text)) & ", " & intSactionWay & ", " _
                        & "'" & IIf(intSactionWay = 0, "", DTOS(dtpM_Date.Value)) & "', " _
                        & "'" & IIf(intSactionWay = 0, "", Text1(9).Text) & "', " _
                        & "'" & strJukyo & "', '" & PB_regUserinfoU.UserServerDate & "', " _
                        & "'" & PB_regUserinfoU.UserCode & "' )"
       On Error GoTo ERROR_TABLE_INSERT
       PB_adoCnnSQL.Execute strSQL
    End If
    '계산서 발행여부
    If chkPrint.Value = 1 Then
       strSQL = "UPDATE 세금계산서 SET " _
                     & "발행여부 = 1 " _
               & "WHERE 사업장코드 = '" & PB_regUserinfoU.UserBranchCode & "' " _
                 & "AND 작성년도 = '" & strMakeYear & "' " _
                 & "AND 책번호 = " & lngLogCnt1 & " " _
                 & "AND 일련번호 = " & lngLogCnt2 & " "
       On Error GoTo ERROR_TABLE_UPDATE
       PB_adoCnnSQL.Execute strSQL
    End If
    '2.매출세금계산서장부(작성시간:strTime)
    strSQL = "INSERT INTO 매출세금계산서장부 " _
           & "SELECT T1.사업장코드, T1.작성일자, '" & strTime & "', T1.매출처코드, " _
                  & "T1.품목및규격, T1.수량, T1.공급가액, T1.세액, " _
                  & "T1.금액구분, T1.영청구분, T1.발행여부, T1.작성구분, " _
                  & "T1.미수구분, T1.적요, T1.사용구분, T1.수정일자, " _
                  & "T1.사용자코드, T1.작성년도, T1.책번호, T1.일련번호 " _
             & "FROM 세금계산서 T1 " _
            & "WHERE T1.사업장코드 = '" & PB_regUserinfoU.UserBranchCode & "' " _
              & "AND T1.작성년도  = '" & strMakeYear & "' AND 책번호 = " & lngLogCnt1 & " AND T1.일련번호 = " & lngLogCnt2 & " "
    On Error GoTo ERROR_TABLE_INSERT
    PB_adoCnnSQL.Execute strSQL
    
    PB_adoCnnSQL.CommitTrans
    If (chkPrint.Value = 1) And (cboPrinter.ListIndex >= 0) Then                         '세금계산서 출력
       SubPubPrint_TaxBill p, PB_intPrtTypeGbn, PB_regUserinfoU.UserBranchCode, Mid(DTOS(dtpP_Date.Value), 1, 4), lngLogCnt1, lngLogCnt2, _
                           0, "", "", ""
    End If
    Screen.MousePointer = vbDefault
    cmdSave.Enabled = True
    cmdClear_Click
    Exit Sub
ERROR_FORM_ACTIVATE_CONNECTION_SERVER:
    MsgBox Err.Number & Err.Description & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "(서버와의 연결 실패)"
    Unload Me
    Exit Sub
ERROR_TABLE_SELECT:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "검색 실패"
    Unload Me
    Exit Sub
ERROR_TABLE_INSERT:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "추가 실패"
    Unload Me
    Exit Sub
ERROR_TABLE_UPDATE:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "저장 실패"
    Unload Me
    Exit Sub
ERROR_TABLE_DELETE:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "삭제 실패"
    Unload Me
    Exit Sub
ERROR_STORED_PROCEDURE:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "로그 변경 실패"
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
    Set frm계산서건별 = Nothing
    frmMain.SBar.Panels(4).Text = ""
End Sub

'+--------------------+
'/// Sub Procedure ///
'+--------------------+
Private Sub SubOther_FILL()
Dim strSQL        As String
Dim intIndex      As Integer
    For intIndex = 0 To Text1.Count - 1: Text1(intIndex).Text = "": Next intIndex
    Text1(0).Text = "": Text1(1).Text = ""
    dtpP_Date.Value = Format(PB_regUserinfoU.UserClientDate, "0000-00-00")
    dtpF_Date.Value = Format(Mid(PB_regUserinfoU.UserClientDate, 1, 6) & "01", "0000-00-00")
    dtpT_Date.Value = Format(PB_regUserinfoU.UserClientDate, "0000-00-00")
    dtpS_Date.Value = Format(PB_regUserinfoU.UserClientDate, "0000-00-00")
    dtpM_Date.Value = Format(PB_regUserinfoU.UserClientDate, "0000-00-00")
    Exit Sub
ERROR_TABLE_SELECT:
    MsgBox Err.Number & Err.Description & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "사업장 읽기 실패"
    Unload Me
    Exit Sub
End Sub

'+-------------------------+
'/// Clear text1(index) ///
'+-------------------------+
Private Sub SubClearText()
Dim lngC As Long
    For lngC = Text1.LBound To Text1.UBound
        Text1(lngC).Text = ""
    Next lngC
    lblJanMny.Caption = "0"
End Sub

'+-------------------------+
'/// Check text1(index) ///
'+-------------------------+
Private Function FncCheckTextBox(lngC As Long, blnOK As Boolean)
    For lngC = Text1.LBound To Text1.UBound
        Text1(lngC).Text = Trim(Text1(lngC).Text)
        Select Case lngC
               Case 0  '발행구분(0.거래, 1.임의)
                    If Not (Text1(lngC).Text = "0" Or Text1(lngC).Text = "1") Then
                       Exit Function
                    End If
               Case 1  '매출처코드
                    If Len(Trim(Text1(lngC).Text)) < 1 Then
                       Exit Function
                    End If
               Case 3  '계산서금액
                    Text1(lngC).Text = Format(Vals(Text1(lngC).Text), "#,0.00")
                    If (Vals(Trim(Text1(lngC).Text)) = 0) And (Trim(Text1(0).Text) = "0") Then '0.거래
                       lngC = 1
                       Exit Function
                    End If
                    If (Vals(Trim(Text1(lngC).Text)) = 0) And (Trim(Text1(0).Text) = "1") Then '1.임의
                       Exit Function
                    End If
               Case 5  '품목및규격
                    If (Len(Trim(Text1(lngC).Text)) < 1) Or (Len(Trim(Text1(lngC).Text)) > 40) Then
                       Exit Function
                    End If
               Case 6  '수량
                    If Vals(Trim(Text1(lngC).Text)) < 0 Then
                       Exit Function
                    End If
               Case 7  '구분(0.현금, 1.수표, 2.어음, 3.외상미수금)
                    If Not (Text1(lngC).Text = "0" Or Text1(lngC).Text = "1" Or Text1(lngC).Text = "2" Or Text1(lngC).Text = "3") Then
                       Exit Function
                    End If
                    If Text1(lngC).Text = "3" And Trim(Text1(8).Text) = "0" Then '외상미수금 and 영수함
                       Exit Function
                    End If
               Case 8  '영청구분
                    If Not (Text1(lngC).Text = "0" Or Text1(lngC).Text = "1") Then
                       Exit Function
                    End If
                    If Not (Text1(7).Text = "3") And Trim(Text1(lngC).Text) = "1" Then '외상미수금아니고 and 청구함
                       Exit Function
                    End If
               Case 9  '어음번호
                    If Len(Trim(Text1(lngC).Text)) > 20 Then
                       Exit Function
                    End If
               Case 10 '적요
                    If Len(Trim(Text1(lngC).Text)) > 50 Then
                       Exit Function
                    End If
               Case Else
        End Select
    Next lngC
    blnOK = True
End Function

'+-----------------+
'/// 미수금잔액 ///
'+-----------------+
Private Sub SubCompute_JanMny(strBranchCode As String, strT_Date As String, strSupplierCode As String)
Dim strSQL As String
    lblJanMny.Caption = "0"
    Screen.MousePointer = vbHourglass
    P_adoRec.CursorLocation = adUseClient
    strSQL = "SELECT T1.매출처코드 AS 매출처코드, " _
                  & "SUM(T1.미수금누계금액) AS 미수금금액, SUM(T1.미수금입금누계금액) AS 미수금입금금액 " _
             & "FROM 미수금원장마감 T1 " _
             & "LEFT JOIN 사업장 T2 ON T2.사업장코드 = T1.사업장코드 " _
             & "LEFT JOIN 매출처 T3 ON T3.사업장코드 = T1.사업장코드 AND T3.매출처코드 = T1.매출처코드 " _
            & "WHERE T1.사업장코드 = '" & strBranchCode & "' " _
              & "AND T1.매출처코드 = '" & strSupplierCode & "' " _
              & "AND T1.마감년월 >= (SUBSTRING('" & strT_Date & "', 1, 4) + '00') " _
              & "AND T1.마감년월 < SUBSTRING('" & strT_Date & "', 1, 6) " _
            & "GROUP BY T1.매출처코드 " _
            & "UNION ALL "
    If PB_regUserinfoU.UserMSGbn = "1" Then  '미수금발생구분 1.전표이면
       strSQL = strSQL _
           & "SELECT T1.매출처코드 AS 매출처코드, " _
                  & "(SUM(T1.출고수량 * T1.출고단가) * " & (PB_curVatRate + 1) & ") AS 미수금금액, 0 AS 미수금입금금액 " _
             & "FROM 자재입출내역 T1 " _
             & "LEFT JOIN 사업장 T2 ON T2.사업장코드 = T1.사업장코드 " _
             & "LEFT JOIN 매출처 T3 ON T3.사업장코드 = T1.사업장코드 AND T3.매출처코드 = T1.매출처코드 " _
            & "WHERE T1.사업장코드 = '" & strBranchCode & "' AND T1.사용구분 = 0 " _
              & "AND T1.매출처코드 = '" & strSupplierCode & "' " _
              & "AND T1.입출고일자 BETWEEN (SUBSTRING('" & strT_Date & "', 1, 6) + '01') AND '" & strT_Date & "' " _
              & "AND (T1.입출고구분 = 2 OR T1.입출고구분 = 8) " _
            & "GROUP BY T1.매출처코드 "
            '& "AND (T1.현금구분 = 0 OR (T1.현금구분 = 1 AND T1.계산서발행여부 = 1)) "
    Else                                     '미수금발생구분 2.계산서이면
       strSQL = strSQL _
           & "SELECT T1.매출처코드 AS 매출처코드, " _
                 & "(SUM(T1.공급가액 + T1.세액)) AS 미수금금액, 0 AS 미수금입금금액 " _
             & "FROM 매출세금계산서장부 T1 " _
             & "LEFT JOIN 사업장 T2 ON T2.사업장코드 = T1.사업장코드 " _
             & "LEFT JOIN 매출처 T3 ON T3.사업장코드 = T1.사업장코드 AND T3.매출처코드 = T1.매출처코드 " _
            & "WHERE T1.사업장코드 = '" & strBranchCode & "' AND T1.사용구분 = 0 " _
              & "AND T1.매출처코드 = '" & strSupplierCode & "' " _
              & "AND T1.작성일자 BETWEEN (SUBSTRING('" & strT_Date & "', 1, 6) + '01') AND '" & strT_Date & "' " _
              & "AND T1.미수구분 = 1 " _
            & "GROUP BY T1.매출처코드 "
    End If
    strSQL = strSQL + "UNION ALL " _
           & "SELECT T1.매출처코드 AS 매출처코드, " _
                  & "0 AS 미수금금액, " _
                  & "ISNULL(SUM(T1.미수금입금금액), 0) As 미수금입금금액 " _
             & "FROM 미수금내역 T1 " _
             & "LEFT JOIN 사업장 T2 ON T2.사업장코드 = T1.사업장코드 " _
             & "LEFT JOIN 매출처 T3 ON T3.사업장코드 = T1.사업장코드 AND T3.매출처코드 = T1.매출처코드 " _
            & "WHERE T1.사업장코드 = '" & strBranchCode & "' " _
              & "AND T1.매출처코드 = '" & strSupplierCode & "' " _
              & "AND T1.미수금입금일자 BETWEEN (SUBSTRING('" & strT_Date & "', 1, 6) + '01') AND '" & strT_Date & "' " _
            & "GROUP BY T1.매출처코드 " _
            & "ORDER BY T1.매출처코드 "
    On Error GoTo ERROR_TABLE_SELECT
    P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
    If P_adoRec.RecordCount = 0 Then
       P_adoRec.Close
    Else
       Do Until P_adoRec.EOF
          lblJanMny.Caption = Format(Vals(lblJanMny.Caption) + P_adoRec("미수금금액") - P_adoRec("미수금입금금액"), "#,0.00")
          P_adoRec.MoveNext
       Loop
       P_adoRec.Close
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
ERROR_TABLE_SELECT:
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "TABLE 읽기 실패"
    Unload Me
    Exit Sub
End Sub

'+---------------------+
'/// 세금계산서금액 ///
'+---------------------+
Private Sub SubCompute_TaxBillMny(strBranchCode As String, strF_Date As String, strT_Date As String, strSupplierCode As String)
Dim strSQL As String
    Text1(3).Text = "0.00": Text1(4).Text = "0.00": Text1(5).Text = "": Text1(6).Text = ""
    Screen.MousePointer = vbHourglass
    P_adoRec.CursorLocation = adUseClient
    strSQL = "SELECT ISNULL(SUM(ISNULL(출고수량, 0) * ISNULL(출고단가, 0)), 0) AS 공급가액 " _
             & "FROM 자재입출내역 T1 " _
            & "WHERE T1.사업장코드 = '" & strBranchCode & "' AND T1.매출처코드 = '" & strSupplierCode & "' " _
              & "AND T1.입출고일자 BETWEEN '" & strF_Date & "' AND '" & strT_Date & "' " _
              & "AND T1.입출고구분 = 2 AND T1.사용구분 = 0 AND T1.현금구분 = 0 AND T1.계산서발행여부 = 0 "
    On Error GoTo ERROR_TABLE_SELECT
    P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
    If P_adoRec.RecordCount = 0 Then
       P_adoRec.Close
    Else
       Do Until P_adoRec.EOF
          Text1(3).Text = Format(Vals(Text1(3).Text) + P_adoRec("공급가액"), "#,0.00")
          P_adoRec.MoveNext
       Loop
       P_adoRec.Close
    End If
    If Vals(Text1(3).Text) > 1 Then
       Text1(4).Text = Format(Fix(Vals(Text1(3).Text) * (PB_curVatRate)), "#,0.00")
    End If
    '품목및규격 및 수량(종)
    'strSQL = "SELECT ISNULL(SUM(ISNULL(출고단가, 0) * ISNULL(출고수량, 0)), 0) AS 세금계산서금액, "
    strSQL = "SELECT (SELECT TOP 1 (ISNULL(S2.자재명,'') + SPACE(1) + ISNULL(S2.규격, '')) FROM 자재입출내역 S1 " _
                     & "LEFT JOIN 자재 S2 ON S1.분류코드 = S2.분류코드 AND S1.세부코드 = S2.세부코드 " _
                    & "WHERE S1.사업장코드 = T1.사업장코드 " _
                      & "AND S1.입출고일자 BETWEEN '" & strF_Date & "' AND '" & strT_Date & "' " _
                      & "AND S1.입출고구분 = 2 AND S1.사용구분 = 0 AND S1.현금구분 = 0 AND S1.계산서발행여부 = 0 " _
                      & "AND S1.매출처코드 = T1.매출처코드 " _
                    & "ORDER BY S1.입출고일자, S1.입출고시간) AS 품목및규격, " _
                  & "(SELECT (COUNT(S1.세부코드) - 1) FROM 자재입출내역 S1 " _
                    & "WHERE S1.사업장코드 = T1.사업장코드 " _
                      & "AND S1.입출고일자 BETWEEN '" & strF_Date & "' AND '" & strT_Date & "' " _
                      & "AND S1.입출고구분 = 2 AND S1.사용구분 = 0 AND S1.현금구분 = 0 AND S1.계산서발행여부 = 0 " _
                      & "AND S1.매출처코드 = T1.매출처코드) AS 수량 " _
              & "FROM 자재입출내역 T1 " _
              & "LEFT JOIN 자재 T2 ON T2.분류코드 = T1.분류코드 AND T2.세부코드 = T1.세부코드 " _
             & "WHERE T1.사업장코드 = '" & strBranchCode & "' AND T1.매출처코드 = '" & strSupplierCode & "' " _
               & "AND T1.입출고일자 BETWEEN '" & strF_Date & "' AND '" & strT_Date & "' " _
               & "AND T1.입출고구분 = 2 AND T1.사용구분 = 0 AND T1.계산서발행여부 = 0 AND T1.현금구분 = 0 " _
             & "GROUP BY T1.사업장코드, T1.매출처코드 "
    On Error GoTo ERROR_TABLE_SELECT
    P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
    If P_adoRec.RecordCount = 0 Then
       P_adoRec.Close
    Else
       Do Until P_adoRec.EOF
          Text1(5).Text = P_adoRec("품목및규격")
          Text1(6).Text = Format(P_adoRec("수량"), "#,0")
          P_adoRec.MoveNext
       Loop
       P_adoRec.Close
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
ERROR_TABLE_SELECT:
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "TABLE 읽기 실패"
    Unload Me
    Exit Sub
End Sub


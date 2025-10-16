VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm매입세금계산서장부입력 
   BorderStyle     =   1  '단일 고정
   Caption         =   "매입세금계산서장부입력"
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
   Icon            =   "매입세금계산서장부입력.frx":0000
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
         Visible         =   0   'False
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
         Caption         =   "매입세금계산서장부 입력"
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
         Picture         =   "매입세금계산서장부입력.frx":030A
         Style           =   1  '그래픽
         TabIndex        =   43
         Top             =   4560
         Width           =   1095
      End
      Begin VB.CommandButton cmdSave 
         Height          =   390
         Left            =   2880
         Picture         =   "매입세금계산서장부입력.frx":0C58
         Style           =   1  '그래픽
         TabIndex        =   17
         Top             =   4560
         Width           =   1095
      End
      Begin VB.CommandButton cmdClear 
         Height          =   375
         Left            =   1680
         Picture         =   "매입세금계산서장부입력.frx":14DF
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
         Visible         =   0   'False
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
         Enabled         =   0   'False
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
         Format          =   57540609
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
         Format          =   57540609
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
         Format          =   57540609
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
         Format          =   57540609
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
         Format          =   57540609
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
         Caption         =   "지급일자"
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
         Caption         =   "(0.현금, 1.수표, 2.어음, 3.외상미지급금)"
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
         Visible         =   0   'False
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
         Visible         =   0   'False
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
         Caption         =   "매입처코드"
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
         Height          =   240
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
Attribute VB_Name = "frm매입세금계산서장부입력"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'+-------------------------------------------------------------------------------------------------------+
' 프로그램 제 목 : 매입세금계산서장부입력
' 사용된 Control : VideoSoft VSFlexGrid 7.0(OLEDB) = vsflex7.ocx
' 참조된 Table   : 사업장, 매입처, 자재입출내역, 매입세금계산서장부
' 업  무  설  명 :
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

    frmMain.SBar.Panels(4).Text = "3.외상미지급금이 아닌경우 미지급금 지급내역에 저장됩니다."
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
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "매입세금계산서장부입력(서버와의 연결 실패)"
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
    If (Index = 1 And (KeyCode = vbKeyHome Or KeyCode = vbKeyReturn)) Then  '매입처검색
       PB_strSupplierCode = UPPER(Trim(Text1(Index).Text))
       PB_strSupplierName = "" 'Trim(Text1(Index + 1).Text)
       frm매입처검색.Show vbModal
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
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "매입처 읽기 실패"
    Unload Me
    Exit Sub
End Sub

'+-------------+
'/// 매입처 ///
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
                Case 1 '매입처코드
                     .Text = UPPER(Trim(.Text))
                     If Len(Trim(.Text)) < 1 Then
                        .Text = ""
                        Text1(2).Text = ""
                        Exit Sub
                     End If
                     'P_adoRec.CursorLocation = adUseClient
                     'strSQL = "SELECT * FROM 매입처 " _
                             & "WHERE 사업장코드 = '" & PB_regUserinfoU.UserBranchCode & "' " _
                               & "AND 매입처코드 = '" & Trim(.Text) & "' "
                     'On Error GoTo ERROR_TABLE_SELECT
                     'P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
                     'If P_adoRec.RecordCount = 0 Then
                     '   P_adoRec.Close
                     '   .Text = ""
                     '   .SetFocus
                     '   Exit Sub
                     'End If
                     'Text1(2).Text = P_adoRec("매입처명")
                     'P_adoRec.Close
                     SubCompute_JanMny PB_regUserinfoU.UserBranchCode, PB_regUserinfoU.UserClientDate, Text1(1).Text  '미지급금잔액
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
                        dtpS_Date.Enabled = True: dtpM_Date.Enabled = False: Text1(9).Enabled = False
                        Text1(8).Text = "0" '영수함
                        dtpS_Date.SetFocus
                     ElseIf _
                        Text1(7).Text = "2" Then                           '어음
                        Text1(8).Text = "0" '영수함
                        dtpS_Date.Enabled = True: dtpM_Date.Enabled = True: Text1(9).Enabled = True
                        dtpS_Date.SetFocus
                     Else
                        Text1(7).Text = "3" '외상
                        Text1(8).Text = "1" '청구함
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
Dim strJukyo       As String  '적요(자재입출내역. 미지급금내역)
Dim intSactionWay  As Integer '결제방법(미지급금내역)
Dim curInputMoney  As Long
Dim curOutputMoney As Long
Dim strServerTime  As String
Dim strTime        As String
Dim intKijang      As Integer '1.기장
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
    intRetVal = MsgBox("매입세금계산서장부에 저장하시겠습니까 ?", vbQuestion + vbYesNo + vbDefaultButton1, "매입세금계산서장부 저장")
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
    'PB_adoCnnSQL.BeginTrans
    'strMakeYear = Mid(DTOS(dtpP_Date.Value), 1, 4)
    'strSQL = "spLogCounter '세금계산서', '" & PB_regUserinfoU.UserBranchCode + strMakeYear & "', " _
    '                     & "0, 0, '" & PB_regUserinfoU.UserServerDate & "', '" & PB_regUserinfoU.UserCode & "' "
    'On Error GoTo ERROR_STORED_PROCEDURE
    'P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
    'lngLogCnt1 = P_adoRec(0)
    'lngLogCnt2 = P_adoRec(1)
    'P_adoRec.Close
    
    '매입세금계산서장부 추가
    PB_adoCnnSQL.BeginTrans
    strSQL = "INSERT INTO 매입세금계산서장부(사업장코드, 작성일자, 작성시간, " _
                                          & "매입처코드, 품목및규격, 수량, " _
                                          & "공급가액, 세액, 금액구분, 영청구분, " _
                                          & "발행여부, 작성구분, 미지급구분, 적요, 사용구분, " _
                                          & "수정일자, 사용자코드, 작성년도, 책번호, 일련번호) VALUES(" _
    & "'" & PB_regUserinfoU.UserBranchCode & "', '" & DTOS(dtpP_Date.Value) & "', '" & strServerTime & "', " _
    & "'" & Trim(Text1(1).Text) & "', '" & Trim(Text1(5).Text) & "'," & Vals(Text1(6).Text) & ", " _
    & "" & Vals(Text1(3).Text) & ", " & Vals(Text1(4).Text) & ", " & Vals(Text1(7).Text) & "," & Vals(Text1(8).Text) & ", " _
    & "1, 1, 1, '" & Text1(10).Text & "', 0, " _
    & "'" & PB_regUserinfoU.UserServerDate & "','" & PB_regUserinfoU.UserCode & "', '', 0, 0) "
    On Error GoTo ERROR_TABLE_INSERT
    PB_adoCnnSQL.Execute strSQL
    If Trim(Text1(0).Text = "0") Then '거래분 세금계산서이면 (자재입출내역 변경)
       strSQL = "UPDATE 자재입출내역 SET " _
                     & "계산서발행여부 = 1 " _
               & "WHERE 사업장코드 = '" & PB_regUserinfoU.UserBranchCode & "' AND 매입처코드 = '" & Trim(Text1(1).Text) & "' " _
                 & "AND 입출고일자 BETWEEN '" & DTOS(dtpF_Date.Value) & "' AND '" & DTOS(dtpT_Date.Value) & "' " _
                 & "AND 사용구분 = 0 AND 사용구분 = 0 AND 현금구분 = 0 AND 계산서발행여부 = 0 " _
                 & "AND 입출고구분 = 1 "
       On Error GoTo ERROR_TABLE_UPDATE
       PB_adoCnnSQL.Execute strSQL
    End If
    'If 영수함 Then  '미지급금내역 추가
    If (Trim(Text1(8).Text = "0")) Then
       If Vals(Trim(Text1(7).Text)) = 0 Or Vals(Trim(Text1(7).Text)) = 1 Then '현금 또는 수표
          intSactionWay = 0
       ElseIf _
          Vals(Trim(Text1(7).Text)) = 2 Then '어음
          intSactionWay = 1
       End If
       strSQL = "INSERT INTO 미지급금내역(사업장코드, 매입처코드, " _
                                       & "미지급금지급일자, 미지급금지급시간," _
                                       & "미지급금지급금액, 결제방법, " _
                                       & "만기일자, 어음번호, " _
                                       & "적요, 수정일자, " _
                                       & "사용자코드) VALUES(" _
                        & "'" & PB_regUserinfoU.UserBranchCode & "', '" & Trim(Text1(1).Text) & "', " _
                        & "'" & DTOS(dtpS_Date.Value) & "', '" & strServerTime & "', " _
                        & "" & (Vals(Text1(3).Text) + Vals(Text1(4).Text)) & ", " & intSactionWay & ", " _
                        & "'" & IIf(intSactionWay = 0, "", DTOS(dtpM_Date.Value)) & "', " _
                        & "'" & IIf(intSactionWay = 0, "", Text1(9).Text) & "', " _
                        & "'" & Text1(10).Text & "', '" & PB_regUserinfoU.UserServerDate & "', " _
                        & "'" & PB_regUserinfoU.UserCode & "' )"
       On Error GoTo ERROR_TABLE_INSERT
       PB_adoCnnSQL.Execute strSQL
    End If
    '계산서 발행여부
    If chkPrint.Value = 1 Then
       strSQL = "UPDATE 매입세금계산서장부 SET " _
                     & "발행여부 = 1 " _
               & "WHERE 사업장코드 = '" & PB_regUserinfoU.UserBranchCode & "' " _
                 & "AND 작성일자 = '" & DTOS(dtpP_Date.Value) & "' " _
                 & "AND 작성시간 = '" & strServerTime & "' "
       On Error GoTo ERROR_TABLE_UPDATE
       PB_adoCnnSQL.Execute strSQL
    End If
    PB_adoCnnSQL.CommitTrans
    If (chkPrint.Value = 1) And (cboPrinter.ListIndex >= 0) Then                         '세금계산서 출력(Not Used)
       SubPrint_TaxBill PB_regUserinfoU.UserBranchCode, Mid(DTOS(dtpP_Date.Value), 1, 4), lngLogCnt1, lngLogCnt2
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
    Set frm매입세금계산서장부입력 = Nothing
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
    For lngC = Text1.LBound + 1 To Text1.UBound
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
               Case 1  '매입처코드
                    If Len(Trim(Text1(lngC).Text)) < 1 Then
                       Exit Function
                    End If
               Case 3  '계산서금액
                    Text1(lngC).Text = Format(Vals(Text1(lngC).Text), "#,0.00")
                    If (Vals(Trim(Text1(lngC).Text)) < 1) And (Trim(Text1(0).Text) = "0") Then '0.거래
                       lngC = 3
                       Exit Function
                    End If
                    If (Vals(Trim(Text1(lngC).Text)) < 1) And (Trim(Text1(0).Text) = "1") Then '1.임의
                       Exit Function
                    End If
               Case 4  '세액
                    Text1(lngC).Text = Format(Vals(Text1(lngC).Text), "#,0.00")
                    If (Vals(Trim(Text1(lngC).Text)) < 0) Then
                       lngC = 4
                       Exit Function
                    End If
               Case 5  '품목및규격
                    If (Len(Trim(Text1(lngC).Text)) < 0) Or (Len(Trim(Text1(lngC).Text)) > 40) Then
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
                    If Text1(lngC).Text = "3" And Trim(Text1(8).Text) = "0" Then '외상미지급금 and 영수함
                       Exit Function
                    End If
               Case 8  '영청구분
                    If Not (Text1(lngC).Text = "0" Or Text1(lngC).Text = "1") Then
                       lngC = 7
                       Exit Function
                    End If
                    If Not (Text1(7).Text = "3") And Trim(Text1(lngC).Text) = "1" Then '외상미지급금아니고 and 청구함
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

'+-------------------+
'/// 미지급금잔액 ///
'+-------------------+
Private Sub SubCompute_JanMny(strBranchCode As String, strT_Date As String, strSupplierCode As String)
Dim strSQL As String
    lblJanMny.Caption = "0"
    Screen.MousePointer = vbHourglass
    P_adoRec.CursorLocation = adUseClient
    strSQL = "SELECT T1.매입처코드 AS 매입처코드, " _
                  & "SUM(T1.미지급금누계금액) AS 미지급금금액, SUM(T1.미지급금지급누계금액) AS 미지급금지급금액 " _
             & "FROM 미지급금원장마감 T1 " _
             & "LEFT JOIN 사업장 T2 ON T2.사업장코드 = T1.사업장코드 " _
             & "LEFT JOIN 매입처 T3 ON T3.사업장코드 = T1.사업장코드 AND T3.매입처코드 = T1.매입처코드 " _
            & "WHERE T1.사업장코드 = '" & strBranchCode & "' " _
              & "AND T1.매입처코드 = '" & strSupplierCode & "' " _
              & "AND T1.마감년월 >= (SUBSTRING('" & strT_Date & "', 1, 4) + '00') " _
              & "AND T1.마감년월 < SUBSTRING('" & strT_Date & "', 1, 6) " _
            & "GROUP BY T1.매입처코드 " _
            & "UNION ALL "
    If PB_regUserinfoU.UserMJGbn = "1" Then  '미지급금발생구분 1.전표이면
       strSQL = strSQL _
           & "SELECT T1.매입처코드 AS 매입처코드, " _
                  & "(SUM(T1.입고수량 * T1.입고단가) * (PB_curVatRate + 1)) AS 미지급금금액, 0 AS 미지급금지급금액 " _
             & "FROM 자재입출내역 T1 " _
             & "LEFT JOIN 사업장 T2 ON T2.사업장코드 = T1.사업장코드 " _
             & "LEFT JOIN 매입처 T3 ON T3.사업장코드 = T1.사업장코드 AND T3.매입처코드 = T1.매입처코드 " _
            & "WHERE T1.사업장코드 = '" & strBranchCode & "' AND T1.사용구분 = 0 " _
              & "AND T1.매입처코드 = '" & strSupplierCode & "' " _
              & "AND T1.입출고일자 BETWEEN (SUBSTRING('" & strT_Date & "', 1, 6) + '01') AND '" & strT_Date & "' " _
              & "AND T1.입출고구분 = 1 " _
            & "GROUP BY T1.매입처코드 "
    Else                                     '미지급금발생구분 2.계산서이면
       strSQL = strSQL + "SELECT T1.매입처코드 AS 매입처코드, " _
                  & "(SUM(T1.공급가액 + T1.세액)) AS 미지급금금액, 0 AS 미지급금지급금액 " _
             & "FROM 매입세금계산서장부 T1 " _
             & "LEFT JOIN 사업장 T2 ON T2.사업장코드 = T1.사업장코드 " _
             & "LEFT JOIN 매입처 T3 ON T3.사업장코드 = T1.사업장코드 AND T3.매입처코드 = T1.매입처코드 " _
            & "WHERE T1.사업장코드 = '" & strBranchCode & "' AND T1.사용구분 = 0 " _
              & "AND T1.매입처코드 = '" & strSupplierCode & "' " _
              & "AND T1.작성일자 BETWEEN (SUBSTRING('" & strT_Date & "', 1, 6) + '01') AND '" & strT_Date & "' " _
              & "AND T1.미지급구분 = 1 " _
            & "GROUP BY T1.매입처코드 "
    End If
    strSQL = strSQL + "UNION ALL " _
           & "SELECT T1.매입처코드 AS 매입처코드, " _
                  & "0 AS 미지급금금액, " _
                  & "ISNULL(SUM(T1.미지급금지급금액), 0) As 미지급금지급금액 " _
             & "FROM 미지급금내역 T1 " _
             & "LEFT JOIN 사업장 T2 ON T2.사업장코드 = T1.사업장코드 " _
             & "LEFT JOIN 매입처 T3 ON T3.사업장코드 = T1.사업장코드 AND T3.매입처코드 = T1.매입처코드 " _
            & "WHERE T1.사업장코드 = '" & strBranchCode & "' " _
              & "AND T1.매입처코드 = '" & strSupplierCode & "' " _
              & "AND T1.미지급금지급일자 BETWEEN (SUBSTRING('" & strT_Date & "', 1, 6) + '01') AND '" & strT_Date & "' " _
            & "GROUP BY T1.매입처코드 " _
            & "ORDER BY T1.매입처코드 "
    On Error GoTo ERROR_TABLE_SELECT
    P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
    If P_adoRec.RecordCount = 0 Then
       P_adoRec.Close
    Else
       Do Until P_adoRec.EOF
          lblJanMny.Caption = Format(Vals(lblJanMny.Caption) + P_adoRec("미지급금금액") - P_adoRec("미지급금지급금액"), "#,0.00")
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

'+--------------------+
'/// 세금계산서금액 ///
'+--------------------+
Private Sub SubCompute_TaxBillMny(strBranchCode As String, strF_Date As String, strT_Date As String, strSupplierCode As String)
Dim strSQL As String
    Text1(3).Text = "0.00": Text1(4).Text = "0.00": Text1(5).Text = "": Text1(6).Text = ""
    Screen.MousePointer = vbHourglass
    P_adoRec.CursorLocation = adUseClient
    strSQL = "SELECT ISNULL(SUM(ISNULL(입고수량, 0) * ISNULL(입고단가, 0)), 0) AS 세금계산서금액 " _
             & "FROM 자재입출내역 T1 " _
            & "WHERE T1.사업장코드 = '" & strBranchCode & "' AND T1.매입처코드 = '" & strSupplierCode & "' " _
              & "AND T1.입출고일자 BETWEEN '" & strF_Date & "' AND '" & strT_Date & "' " _
              & "AND T1.사용구분 = 0 AND T1.사용구분 = 0 AND T1.현금구분 = 0 AND T1.계산서발행여부 = 0 " _
              & "AND T1.입출고구분 = 1 "
    On Error GoTo ERROR_TABLE_SELECT
    P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
    If P_adoRec.RecordCount = 0 Then
       P_adoRec.Close
    Else
       Do Until P_adoRec.EOF
          Text1(3).Text = Format(Vals(Text1(3).Text) + P_adoRec("세금계산서금액"), "#,0.00")
          P_adoRec.MoveNext
       Loop
       P_adoRec.Close
    End If
    If Vals(Text1(3).Text) > 1 Then
       Text1(4).Text = Format(Fix(Vals(Text1(3).Text) * (PB_curVatRate)), "#,0.00")
    End If
    '품목및규격 및 수량(종)
    'strSQL = "SELECT ISNULL(SUM(ISNULL(입고단가, 0) * ISNULL(입고수량, 0)), 0) AS 세금계산서금액, "
    strSQL = "SELECT (SELECT TOP 1 (ISNULL(S2.자재명,'') + SPACE(1) + ISNULL(S2.규격, '')) FROM 자재입출내역 S1 " _
                    & "LEFT JOIN 자재 S2 ON S1.분류코드 = S2.분류코드 AND S1.세부코드 = S2.세부코드 " _
                   & "WHERE S1.사업장코드 = T1.사업장코드 " _
                     & "AND S1.입출고일자 BETWEEN '" & strF_Date & "' AND '" & strT_Date & "' " _
                     & "AND S1.입출고구분 = 1 AND S1.사용구분 = 0 AND S1.현금구분 = 0 AND S1.계산서발행여부 = 0 " _
                     & "AND S1.매입처코드 = T1.매입처코드 " _
                   & "ORDER BY S1.입출고일자, S1.입출고시간) AS 품목및규격, " _
                 & "(SELECT (COUNT(S1.세부코드) - 1) FROM 자재입출내역 S1 " _
                   & "WHERE S1.사업장코드 = T1.사업장코드 " _
                     & "AND S1.입출고일자 BETWEEN '" & strF_Date & "' AND '" & strT_Date & "' " _
                     & "AND S1.입출고구분 = 1 AND S1.사용구분 = 0 AND S1.현금구분 = 0 AND S1.계산서발행여부 = 0 " _
                     & "AND S1.매입처코드 = T1.매입처코드) AS 수량 " _
             & "FROM 자재입출내역 T1 " _
             & "LEFT JOIN 자재 T2 ON T2.분류코드 = T1.분류코드 AND T2.세부코드 = T1.세부코드 " _
            & "WHERE T1.사업장코드 = '" & strBranchCode & "' AND T1.매입처코드 = '" & strSupplierCode & "' " _
              & "AND T1.입출고일자 BETWEEN '" & strF_Date & "' AND '" & strT_Date & "' " _
              & "AND T1.입출고구분 = 1 AND T1.사용구분 = 0 AND T1.계산서발행여부 = 0 AND T1.현금구분 = 0 " _
            & "GROUP BY T1.사업장코드, T1.매입처코드 "
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

'+--------------------------------+
'/// 세금계산서 출력(Not Used) ///
'+--------------------------------+
Private Sub SubPrint_TaxBill(strBranchCode As String, strMakeYear As String, lngLogCnt1 As Long, lngLogCnt2 As Long)
Dim strSQL               As String
Dim p                    As Printer
Dim strPort              As String
Dim intFile              As Integer
Dim blnEof               As Boolean
Dim intPrtCnt            As Integer
Dim strPrtLine           As String
Dim inti                 As Integer
Dim C_TMargin            As Integer  'Top Margin
Dim C_LMargin            As Integer  'Left Margin
Dim intA                 As Integer
Dim SW_A                 As Integer

Dim C_intCntPerPage      As Integer
Dim intTotCnt            As Integer
Dim strBuyerCode         As String   '매출처코드

Dim A()                  As String   '상

Dim lngR                 As Long
Dim lngC                 As Long

Dim strBookNo            As String   '세금계산서 책번호
Dim lngSeqNo             As Long     '세금계산서 일련번호
    
    For Each p In Printers
        If Trim(p.DeviceName) = Trim(cboPrinter.Text) And lstPort.List(cboPrinter.ListIndex) = p.Port Then
           Set Printer = p
           Exit For
        End If
    Next
    Screen.MousePointer = vbHourglass
    P_adoRec.CursorLocation = adUseClient
    strSQL = "SELECT T1.사업장코드 AS 사업장코드, T1.작성년도 AS 작성년도, T1.책번호 AS 책번호, T1.일련번호 AS 일련번호, " _
                  & "T1.매출처코드 AS 매출처코드, ISNULL(T2.사업자번호, '') AS 등록번호, ISNULL(T2.매출처명, '') AS 상호법인명, " _
                  & "ISNULL(T2.대표자명, '') AS 성명, (ISNULL(T2.주소, '') + SPACE(1) + ISNULL(T2.번지, '')) AS 사업장주소, " _
                  & "ISNULL(업태, '') AS 업태, ISNULL(업종, '') AS 종목, " _
                  & "T1.작성일자 AS 작성일자, T1.공급가액 AS 공급가액, T1.세액 AS 세액, T1.품목및규격 AS 품목및규격, T1.수량 AS 수량, " _
                  & "T1.금액구분 AS 금액구분, T1.영청구분 AS 영청구분, T1.미수구분 AS 미수구분 " _
             & "FROM 세금계산서 T1 " _
             & "LEFT JOIN 매출처 T2 ON T2.매출처코드 = T1.매출처코드 " _
            & "WHERE T1.사업장코드 = '" & strBranchCode & "' AND T1.작성년도 = '" & strMakeYear & "' " _
              & "AND T1.책번호 = " & lngLogCnt1 & " AND T1.일련번호 = " & lngLogCnt2 & " AND T1.사용구분 = 0 " _
            & "ORDER BY T1.사업장코드, T1.작성년도, T1.책번호, T1.일련번호 "
    On Error GoTo ERROR_TABLE_SELECT
    P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
    If P_adoRec.RecordCount = 0 Then
       P_adoRec.Close
       Screen.MousePointer = vbDefault
       Exit Sub
    Else
       intTotCnt = P_adoRec.RecordCount
       C_TMargin = 2
       C_LMargin = 20
       
       ReDim A(intTotCnt, 23)
       
       Do Until P_adoRec.EOF
          A(intA, 0) = P_adoRec("책번호")
          A(intA, 1) = P_adoRec("일련번호")
          A(intA, 2) = P_adoRec("매출처코드")
          A(intA, 3) = P_adoRec("등록번호")
          A(intA, 4) = P_adoRec("상호법인명")
          A(intA, 5) = P_adoRec("성명")
          A(intA, 6) = P_adoRec("사업장주소")
          A(intA, 7) = P_adoRec("업태")
          A(intA, 8) = P_adoRec("종목")
          A(intA, 9) = P_adoRec("작성일자")
          A(intA, 10) = Mid(P_adoRec("작성일자"), 5, 2)         '월
          A(intA, 11) = Mid(P_adoRec("작성일자"), 7, 2)         '일
          A(intA, 12) = PADR(P_adoRec("품목및규격"), 20, "") & " 외"  '품목 및 규격
          A(intA, 13) = Format(P_adoRec("수량"), "#") & "종"    '수량
          A(intA, 14) = ""                                      '단가
          A(intA, 15) = P_adoRec("공급가액")                    '공급가액
          A(intA, 16) = P_adoRec("세액")                        '세액
          A(intA, 17) = P_adoRec("공급가액") + P_adoRec("세액") '합계금액
          Select Case P_adoRec("금액구분")
                 Case 0: A(intA, 18) = "O"                      '현금
                 Case 1: A(intA, 19) = "O"                      '수표
                 Case 2: A(intA, 20) = "O"                      '어음
                 Case 3: A(intA, 21) = "O"                      '외상미수금
          End Select
          A(intA, 22) = P_adoRec("영청구분")                    '0.영수함, 1.청구함
          intA = intA + 1
          P_adoRec.MoveNext
       Loop
       P_adoRec.Close
    End If
    'strPort = x.Port                  '예)\\Gp202\hp 'Print On Printer
    'strPort = "C:\Documents\세금계산서.TXT"
    'intFile = FreeFile
    'Open strPort For Output As #intFile
    Printer.PaperSize = vbPRPSA4          '용지설정
    Printer.Orientation = vbPRORPortrait  '용지방향 [ vbPRORPortrait(세로), vbPRORLandscape(가로) ]
    Printer.FontName = "굴림체"
    Printer.FontUnderline = False
    Printer.FontSize = 8
    Printer.FontBold = False
    For intA = LBound(A, 1) To UBound(A, 1) - 1
        '상
        'HEAD
        SubPrint_TaxBill_HEAD_1 C_TMargin, C_LMargin, A(intA, 0), A(intA, 1), A(intA, 3), A(intA, 4), A(intA, 5), _
                                                      A(intA, 6), A(intA, 7), A(intA, 8), A(intA, 9), A(intA, 15), A(intA, 16)
        'BODY
        '10.월, 11.일, 12.품명및규격, 13.수량, 14.단가, 15.공급가액, 16.세액
        Printer.FontSize = 10
        Printer.Print Space(C_LMargin - 14) & A(intA, 10) & Space(1) & A(intA, 11) & Space(1) _
                     & PADR(A(intA, 12), 35, "") & PADL(A(intA, 13), 8, "") & Space(14) _
                     & PADL(Format(Vals(A(intA, 15)), "#,0"), 13, "") & Space(1) _
                     & PADL(Format(Vals(A(intA, 16)), "#,0"), 13, "")
        Printer.Print ""
        Printer.FontSize = 2: Printer.Print "": Printer.Print ""
        Printer.FontSize = 10
        Printer.Print Space(C_LMargin + 30) & "--- 이 하 여 백 ---"
        For inti = 1 To 4: Printer.Print "": Next inti
        'FOOT(17.합계금액, 18.현금, 19.수표, 20.어음, 21.외상미수금)
        Printer.FontSize = 10
        If A(intA, 22) = "0" Then '영수함
           Printer.Print ""
           Printer.Print Space(C_LMargin - 12) & PADL(Format(Vals(A(intA, 17)), "#,0"), 14, "") & Space(2) _
                                               & PADC(A(intA, 18), 13, "") & PADC(A(intA, 19), 13, "") _
                                               & PADC(A(intA, 20), 13, "") & PADC(A(intA, 21), 13, "") & Space(11) & "****"
        Else                      '청구함
           Printer.Print Space(C_LMargin - 12) & Space(79) & "****"
           Printer.Print Space(C_LMargin - 12) & PADL(Format(Vals(A(intA, 17)), "#,0"), 14, "") & Space(2) _
                                               & PADC(A(intA, 18), 13, "") & PADC(A(intA, 19), 13, "") _
                                               & PADC(A(intA, 20), 13, "") & PADC(A(intA, 21), 13, "")
        End If
        Printer.FontSize = 7
        For inti = 1 To 3: Printer.Print "": Next inti
        Printer.FontSize = 2: Printer.Print ""
        '하
        'HEAD
        SubPrint_TaxBill_HEAD_1 C_TMargin, C_LMargin, A(intA, 0), A(intA, 1), A(intA, 3), A(intA, 4), A(intA, 5), _
                                                      A(intA, 6), A(intA, 7), A(intA, 8), A(intA, 9), A(intA, 15), A(intA, 16)
        'BODY
        '10.월, 11.일, 12.품명및규격, 13.수량, 14.단가, 15.공급가액, 16.세액
        Printer.FontSize = 10
        Printer.Print Space(C_LMargin - 14) & A(intA, 10) & Space(1) & A(intA, 11) & Space(1) _
                     & PADR(A(intA, 12), 35, "") & PADL(A(intA, 13), 8, "") & Space(14) _
                     & PADL(Format(Vals(A(intA, 15)), "#,0"), 13, "") & Space(1) _
                     & PADL(Format(Vals(A(intA, 16)), "#,0"), 13, "")
        Printer.Print ""
        Printer.FontSize = 2: Printer.Print "": Printer.Print ""
        Printer.FontSize = 10
        Printer.Print Space(C_LMargin + 30) & "--- 이 하 여 백 ---"
        For inti = 1 To 4: Printer.Print "": Next inti
        'FOOT(17.합계금액, 18.현금, 19.수표, 20.어음, 21.외상미수금)
        Printer.FontSize = 10
        If A(intA, 22) = "0" Then '영수함
           Printer.Print ""
           Printer.Print Space(C_LMargin - 12) & PADL(Format(Vals(A(intA, 17)), "#,0"), 14, "") & Space(2) _
                                               & PADC(A(intA, 18), 13, "") & PADC(A(intA, 19), 13, "") _
                                               & PADC(A(intA, 20), 13, "") & PADC(A(intA, 21), 13, "") & Space(11) & "****"
        Else                      '청구함
           Printer.Print Space(C_LMargin - 12) & Space(79) & "****"
           Printer.Print Space(C_LMargin - 12) & PADL(Format(Vals(A(intA, 17)), "#,0"), 14, "") & Space(2) _
                                               & PADC(A(intA, 18), 13, "") & PADC(A(intA, 19), 13, "") _
                                               & PADC(A(intA, 20), 13, "") & PADC(A(intA, 21), 13, "")
        End If
        Printer.NewPage
    Next intA
    Erase A
    Printer.EndDoc
    Screen.MousePointer = vbDefault
    Exit Sub
ERROR_CRYSTAL_REPORTS:
    MsgBox Err.Number & Space(1) & Err.Description
    Screen.MousePointer = vbDefault
    Exit Sub
ERROR_TABLE_SELECT:
    MsgBox Err.Number & Err.Description & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "입출내역 읽기 실패"
    Unload Me
    Exit Sub
End Sub

Private Sub SubPrint_TaxBill_HEAD_1(C_TMargin As Integer, C_LMargin As Integer, _
                                    A0 As String, A1 As String, _
                                    A3 As String, A4 As String, A5 As String, _
                                    A6 As String, A7 As String, A8 As String, _
                                    A9 As String, A15 As String, A16 As String)
Dim aryEnterNo12(12) As String
Dim aryEnterNo24(24) As String
Dim strEnterNo23     As String
Dim inti             As Integer
Dim intBlankCnt      As Integer '공란수
Dim aryMny1_11(11)   As String  '공급가액
Dim aryMny1_22(22)   As String
Dim strMny1_21       As String
Dim aryMny2_11(11)   As String  '세액
Dim aryMny2_22(22)   As String
Dim strMny2_21       As String
    Printer.FontSize = 10
    'A0.책번호, A1.일련번호, A3.등록번호, A4.상호, A5.성명, A6.주소, A7.업태, A8.종목, A9.거래일자, A15.공급가액, A16.세액
    'For inti = 1 To C_TMargin: Printer.Print "": Next inti
    '등록번호 정렬
    For inti = 1 To 12
        aryEnterNo12(inti) = Mid(A3, inti, 1)
    Next inti
    For inti = 1 To 12
        If inti = 1 Then
           aryEnterNo24(inti) = aryEnterNo12(inti): aryEnterNo24(inti + 1) = " "
        Else
           aryEnterNo24(inti * 2 - 1) = aryEnterNo12(inti): aryEnterNo24(inti * 2) = " "
        End If
    Next inti
    For inti = 1 To 23
        strEnterNo23 = strEnterNo23 + aryEnterNo24(inti)
    Next inti
    Printer.FontSize = 8
    For inti = 1 To 3: Printer.Print "": Next inti
    Printer.FontSize = 2: Printer.Print ""
    '책번호
    Printer.FontSize = 10
    Printer.Print ""                                      '세금계산서번호 출력(X)
    'Printer.Print Space(C_LMargin + 64) & PADR(A0, 6, "") '세금계산번호 출력(O)
    '일련번호
    Printer.Print ""                                      '세금계산서번호 출력(X)
    'Printer.Print Space(C_LMargin + 64) & PADR(A1, 6, "") '세금계산번호 출력(O)
    '등록번호
    Printer.FontSize = 12
    For inti = 1 To 1: Printer.Print "": Next inti
    Printer.Print Space(C_LMargin + 32) & PADR(strEnterNo23, 23, "")
    For inti = 1 To 1: Printer.Print "": Next inti
    'Printer.Print Space(C_LMargin + 50) & Chr(27) & "W1" & PADC(strEnterNo, 14, "") & Chr(27) & "W0"
    '상호, 성명
    Printer.FontSize = 2: Printer.Print "": Printer.Print "": Printer.Print ""
    Printer.FontSize = 12
    Printer.Print Space(C_LMargin + 30) & PADR(A4, 20, "") & Space(2) & PADR(A5, 16, "")
    '주소(작게)
    Printer.FontSize = 10: Printer.Print "": Printer.Print ""
    Printer.FontSize = 2: Printer.Print "": Printer.Print ""
    Printer.FontSize = 10
    Printer.Print Space(C_LMargin + 40) & PADR(A6, 70, "")
    For inti = 1 To 2: Printer.Print "": Next inti
    '업태, 종목(작게)
    Printer.FontSize = 2: Printer.Print ""
    Printer.FontSize = 8
    Printer.Print Space(C_LMargin + 54) & PADR(A7, 14, "") & Space(7) & PADR(A8, 14, "")
    For inti = 1 To 3: Printer.Print "": Next inti
    '작성년월일, 공란수, 공급가액, 세액
    '공급가액 정렬
    For inti = 1 To 11
        aryMny1_11(inti) = Mid(A15, inti, 1)
    Next inti
    For inti = 1 To 11
        If inti = 1 Then
           aryMny1_22(inti) = aryMny1_11(inti): aryMny1_22(inti + 1) = " "
        Else
           aryMny1_22(inti * 2 - 1) = aryMny1_11(inti): aryMny1_22(inti * 2) = " "
        End If
    Next inti
    For inti = 1 To 21
        strMny1_21 = strMny1_21 + aryMny1_22(inti)
    Next inti
    '세액 정렬
    For inti = 1 To 11
        aryMny2_11(inti) = Mid(A16, inti, 1)
    Next inti
    For inti = 1 To 11
        If inti = 1 Then
           aryMny2_22(inti) = aryMny2_11(inti): aryMny2_22(inti + 1) = " "
        Else
           aryMny2_22(inti * 2 - 1) = aryMny2_11(inti): aryMny2_22(inti * 2) = " "
        End If
    Next inti
    For inti = 1 To 21
        strMny2_21 = strMny2_21 + aryMny2_22(inti)
    Next inti
    Printer.FontSize = 8
    For inti = 1 To 2: Printer.Print "": Next inti
    intBlankCnt = 11 - Len(Trim(A15))
    Printer.FontSize = 12
    Printer.Print Space(C_LMargin - 14) & Mid(A9, 1, 4) & Space(1) & Mid(A9, 5, 2) & Mid(A9, 7, 2) _
                      & " " & PADC(intBlankCnt, 3, "") & PADL(strMny1_21, 22, "") & PADL(strMny2_21, 20, "")
    Printer.FontSize = 8
    For inti = 1 To 3: Printer.Print "": Next inti
End Sub


VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm���԰��� 
   BorderStyle     =   0  '����
   Caption         =   "���԰���"
   ClientHeight    =   10095
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15405
   BeginProperty Font 
      Name            =   "����ü"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   10100
   ScaleMode       =   0  '�����
   ScaleWidth      =   15405
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Appearance      =   0  '���
      ForeColor       =   &H80000008&
      Height          =   675
      Left            =   60
      TabIndex        =   23
      Top             =   0
      Width           =   15195
      Begin VB.OptionButton optPrtChk1 
         Caption         =   "��ü��"
         Height          =   255
         Left            =   6840
         TabIndex        =   41
         Top             =   390
         Width           =   975
      End
      Begin VB.OptionButton optPrtChk0 
         Caption         =   "���ں�"
         Height          =   255
         Left            =   6840
         TabIndex        =   40
         Top             =   150
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.CommandButton cmdPrint 
         Height          =   375
         Left            =   7980
         Picture         =   "���԰���.frx":0000
         Style           =   1  '�׷���
         TabIndex        =   30
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdClear 
         Height          =   375
         Left            =   9120
         Picture         =   "���԰���.frx":0963
         Style           =   1  '�׷���
         TabIndex        =   28
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdExit 
         Height          =   390
         Left            =   13635
         Picture         =   "���԰���.frx":1308
         Style           =   1  '�׷���
         TabIndex        =   27
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         Height          =   390
         Left            =   12510
         Picture         =   "���԰���.frx":1C56
         Style           =   1  '�׷���
         TabIndex        =   26
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdSave 
         Height          =   390
         Left            =   11385
         Picture         =   "���԰���.frx":25DA
         Style           =   1  '�׷���
         TabIndex        =   25
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdFind 
         Height          =   390
         Left            =   10260
         Picture         =   "���԰���.frx":2E61
         Style           =   1  '�׷���
         TabIndex        =   0
         Top             =   195
         Width           =   1095
      End
      Begin Crystal.CrystalReport CrystalReport1 
         Left            =   4320
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
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         BackColor       =   &H00008000&
         BorderStyle     =   1  '���� ����
         Caption         =   "�ŷ����� ��ǰó��"
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   18
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   420
         Left            =   165
         TabIndex        =   24
         Top             =   180
         Width           =   4650
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid vsfg1 
      Height          =   7575
      Left            =   60
      TabIndex        =   12
      Top             =   2460
      Width           =   15195
      _cx             =   26802
      _cy             =   13361
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����ü"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483633
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   -1  'True
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  '���
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1755
      Left            =   60
      TabIndex        =   13
      Top             =   630
      Width           =   15195
      Begin VB.CheckBox chkCash 
         Enabled         =   0   'False
         Height          =   255
         Left            =   13680
         TabIndex        =   42
         Top             =   960
         Width           =   255
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  '���
         Enabled         =   0   'False
         Height          =   285
         IMEMode         =   1  '�Է� ���� ����
         Index           =   4
         Left            =   6510
         TabIndex        =   5
         Top             =   225
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  '���
         Enabled         =   0   'False
         Height          =   285
         IMEMode         =   1  '�Է� ���� ����
         Index           =   6
         Left            =   6510
         MaxLength       =   14
         TabIndex        =   8
         Top             =   1320
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker dtpAppDate 
         Height          =   270
         Left            =   6510
         TabIndex        =   6
         Top             =   600
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   476
         _Version        =   393216
         Format          =   57540609
         UpDown          =   -1  'True
         CurrentDate     =   37763
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  '���
         Height          =   285
         IMEMode         =   10  '�ѱ� 
         Index           =   9
         Left            =   9030
         MaxLength       =   50
         TabIndex        =   11
         Top             =   1320
         Width           =   5655
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  '������ ����
         Appearance      =   0  '���
         Enabled         =   0   'False
         Height          =   285
         IMEMode         =   8  '����
         Index           =   8
         Left            =   11430
         MaxLength       =   14
         TabIndex        =   10
         Top             =   945
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  '������ ����
         Appearance      =   0  '���
         Enabled         =   0   'False
         Height          =   285
         IMEMode         =   8  '����
         Index           =   7
         Left            =   9030
         MaxLength       =   14
         TabIndex        =   9
         Top             =   945
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  '������ ����
         Appearance      =   0  '���
         Height          =   285
         IMEMode         =   8  '����
         Index           =   5
         Left            =   6510
         MaxLength       =   14
         TabIndex        =   7
         Top             =   945
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  '���
         Enabled         =   0   'False
         Height          =   285
         IMEMode         =   10  '�ѱ� 
         Index           =   3
         Left            =   1275
         TabIndex        =   4
         Top             =   1305
         Width           =   3855
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  '���
         Enabled         =   0   'False
         Height          =   285
         IMEMode         =   8  '����
         Index           =   2
         Left            =   1275
         MaxLength       =   50
         TabIndex        =   3
         Top             =   945
         Width           =   2415
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  '���
         Enabled         =   0   'False
         Height          =   285
         IMEMode         =   10  '�ѱ� 
         Index           =   1
         Left            =   1275
         TabIndex        =   2
         Top             =   585
         Width           =   2895
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         Enabled         =   0   'False
         Height          =   285
         IMEMode         =   8  '����
         Index           =   0
         Left            =   1275
         MaxLength       =   8
         TabIndex        =   1
         Top             =   225
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker dtpFDate 
         Height          =   270
         Left            =   10800
         TabIndex        =   34
         Top             =   240
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   476
         _Version        =   393216
         Format          =   57540609
         UpDown          =   -1  'True
         CurrentDate     =   37763
      End
      Begin MSComCtl2.DTPicker dtpTDate 
         Height          =   270
         Left            =   12840
         TabIndex        =   35
         Top             =   240
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   476
         _Version        =   393216
         Format          =   57540609
         UpDown          =   -1  'True
         CurrentDate     =   37763
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "���ݹ�ǰ"
         Height          =   240
         Index           =   16
         Left            =   13920
         TabIndex        =   43
         Top             =   1005
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "����"
         Height          =   240
         Index           =   14
         Left            =   5310
         TabIndex        =   39
         Top             =   1365
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "�԰�"
         Height          =   240
         Index           =   13
         Left            =   5310
         TabIndex        =   38
         Top             =   285
         Width           =   1095
      End
      Begin VB.Line Line1 
         X1              =   8520
         X2              =   14660
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "����"
         Height          =   240
         Index           =   12
         Left            =   14160
         TabIndex        =   37
         Top             =   285
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "����"
         Height          =   240
         Index           =   11
         Left            =   12120
         TabIndex        =   36
         Top             =   285
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "��������"
         Height          =   240
         Index           =   10
         Left            =   9720
         TabIndex        =   33
         Top             =   285
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "�˻�����)"
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   35
         Left            =   8400
         TabIndex        =   32
         Top             =   285
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "[Home]"
         Height          =   240
         Index           =   9
         Left            =   3840
         TabIndex        =   31
         Top             =   1005
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "[Home]"
         Height          =   240
         Index           =   15
         Left            =   2520
         TabIndex        =   29
         Top             =   285
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "����"
         Height          =   240
         Index           =   8
         Left            =   7830
         TabIndex        =   22
         Top             =   1365
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "���ΰ�"
         Height          =   240
         Index           =   7
         Left            =   10230
         TabIndex        =   21
         Top             =   1005
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "���ܰ�"
         Height          =   240
         Index           =   6
         Left            =   7830
         TabIndex        =   20
         Top             =   1005
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "���Լ���"
         Height          =   240
         Index           =   5
         Left            =   5310
         TabIndex        =   19
         Top             =   1005
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "������������"
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   4
         Left            =   5310
         TabIndex        =   18
         Top             =   645
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "ǰ��"
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   3
         Left            =   75
         TabIndex        =   17
         Top             =   1365
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "ǰ���ڵ�"
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   2
         Left            =   75
         TabIndex        =   16
         Top             =   1005
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "����ó��"
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   1
         Left            =   75
         TabIndex        =   15
         Top             =   645
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "����ó�ڵ�"
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   0
         Left            =   75
         TabIndex        =   14
         Top             =   285
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frm���԰���"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'+-------------------------------------------------------------------------------------------------------+
' ���α׷� �� �� : ���԰���
' ���� Control : VideoSoft VSFlexGrid 7.0(OLEDB) = vsflex7.ocx
' ������ Table   :
' ��  ��  ��  �� : ����ó���� ������ ���� -�������� ����
'+-------------------------------------------------------------------------------------------------------+
Option Explicit
Private P_blnActived       As Boolean
Private P_adoRec           As New ADODB.Recordset
Private P_intButton        As Integer
Private Const PC_intRowCnt As Integer = 22  '�׸��� �� ������ �� ���(FixedRows ����)

'+--------------------------------+
'/// LOAD FORM ( �ѹ��� ���� ) ///
'+--------------------------------+
Private Sub Form_Load()
    P_blnActived = False
End Sub

'+-------------------------------------------+
'/// ACTIVATE FORM Ȱ��ȭ ( �ѹ��� ���� ) ///
'+-------------------------------------------+
Private Sub Form_Activate()
Dim SQL     As String

    frmMain.SBar.Panels(4).Text = "�����������ڸ� ��Ȯ�� �Է��Ͽ� �ּ���. "
    If P_blnActived = False Then
       Screen.MousePointer = vbHourglass
       On Error GoTo ERROR_FORM_ACTIVATE_CONNECTION_SERVER
       Subvsfg1_INIT
       Select Case Val(PB_regUserinfoU.UserAuthority)
              Case Is <= 10 '��ȸ
                   cmdPrint.Enabled = False: cmdClear.Enabled = False: cmdFind.Enabled = True
                   cmdSave.Enabled = False: cmdDelete.Enabled = False: cmdExit.Enabled = True
              Case Is <= 20 '�μ�, ��ȸ
                   cmdPrint.Enabled = True: cmdClear.Enabled = False: cmdFind.Enabled = True
                   cmdSave.Enabled = False: cmdDelete.Enabled = False: cmdExit.Enabled = True
              Case Is <= 40 '�߰�, ����, �μ�, ��ȸ
                   cmdPrint.Enabled = True: cmdClear.Enabled = True: cmdFind.Enabled = True
                   cmdSave.Enabled = True: cmdDelete.Enabled = False: cmdExit.Enabled = True
              Case Is <= 50 '����, �߰�, ����, ��ȸ, �μ�
                   cmdPrint.Enabled = True: cmdClear.Enabled = True: cmdFind.Enabled = True
                   cmdSave.Enabled = True: cmdDelete.Enabled = True: cmdExit.Enabled = True
              Case Is <= 99 '����, �߰�, ����, ��ȸ, �μ�
                   cmdPrint.Enabled = True: cmdClear.Enabled = True: cmdFind.Enabled = True
                   cmdSave.Enabled = True: cmdDelete.Enabled = True: cmdExit.Enabled = True
              Case Else
                   cmdClear.Enabled = False: cmdFind.Enabled = False: cmdSave.Enabled = False
                   cmdDelete.Enabled = False: cmdExit.Enabled = True
       End Select
       dtpAppDate.Value = Format(PB_regUserinfoU.UserClientDate, "0000-00-00")
       dtpFDate.Value = Format(Mid(PB_regUserinfoU.UserClientDate, 1, 6) & "01", "0000-00-00")
       dtpTDate.Value = DateAdd("d", -1, DateAdd("m", 1, dtpFDate.Value))
       P_blnActived = True
       Screen.MousePointer = vbDefault
    End If
    Exit Sub
ERROR_FORM_ACTIVATE_CONNECTION_SERVER:
    MsgBox Err.Number & Err.Description & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "����� ����(�������� ���� ����)"
    Unload Me
    Exit Sub
End Sub

'+---------------+
'/// �˻����� ///
'+---------------+
Private Sub dtpFDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub
Private Sub dtpTDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then cmdFind_Click
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
    If (Index = 0 And (KeyCode = vbKeyHome Or KeyCode = vbKeyReturn)) Then  '����ó�˻�
       PB_strSupplierCode = UPPER(Trim(Text1(Index).Text))
       PB_strSupplierName = "" 'Trim(Text1(Index + 1).Text)
       frm����ó�˻�.Show vbModal
       If (Len(PB_strSupplierCode) + Len(PB_strSupplierName)) = 0 Then '�˻����� ���(ESC)
       Else
          Text1(Index).Text = PB_strSupplierCode
          Text1(Index + 1).Text = PB_strSupplierName
       End If
       If PB_strSupplierCode <> "" Then
          SendKeys "{TAB}"
       End If
       PB_strSupplierCode = "": PB_strSupplierName = ""
    ElseIf _
       (Index = 2 And (KeyCode = vbKeyHome Or KeyCode = vbKeyReturn)) Then  '����˻�
       PB_strCallFormName = "frm���԰���"
       PB_strMaterialsCode = Trim(Text1(Index).Text)
       PB_strMaterialsName = ""
       frm����˻�.Show vbModal
       If (Len(PB_strMaterialsCode) + Len(PB_strMaterialsName)) = 0 Then '�˻����� ���(ESC)
       Else
          Text1(2).Text = PB_strMaterialsCode
          Text1(3).Text = PB_strMaterialsName
       End If
       If PB_strMaterialsCode <> "" Then
          SendKeys "{tab}"
       End If
       PB_strCallFormName = "": PB_strMaterialsCode = "": PB_strMaterialsName = ""
    Else
       If KeyCode = vbKeyReturn Then
          Select Case Index
                 Case Text1.UBound
                      If cmdSave.Enabled = True Then
                         cmdSave.SetFocus
                         Exit Sub
                      End If
          End Select
          SendKeys "{tab}"
       End If
    End If
    Exit Sub
ERROR_TABLE_SELECT:
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "����ó �б� ����"
    Unload Me
    Exit Sub
End Sub

Private Sub Text1_LostFocus(Index As Integer)
Dim strSQL As String
    With Text1(Index)
         Select Case Index
                Case 0 '����ó
                     .Text = UPPER(Trim(.Text))
                     If Len(Trim(Text1(Index).Text)) < 1 Then
                        Text1(1).Text = ""
                     End If
                Case 2 '����˻�
                     If Len(Trim(Text1(Index).Text)) = 0 Then
                        Text1(3).Text = ""
                     End If
                Case 5
                     If Vals(.Text) > 0 Then
                        .Text = Vals(.Text) * -1
                     End If
                     .Text = Format(Vals(Trim(.Text)), "#,0")
                Case Else
                     .Text = Trim(.Text)
         End Select
    End With
    Exit Sub
End Sub

Private Sub dtpAppDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub

'+-----------+
'/// Grid ///
'+-----------+
Private Sub vsfg1_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    With vsfg1
         .ToolTipText = ""
         If .MouseRow < .FixedRows Or .MouseCol < 0 Then
            Exit Sub
         End If
         .ToolTipText = .TextMatrix(.MouseRow, .MouseCol)
    End With
End Sub
Private Sub vsfg1_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    P_intButton = Button
End Sub
Private Sub vsfg1_Click()
Dim strData As String
    With vsfg1
         If (.MouseRow >= 0 And .MouseRow < .FixedRows) Then
            .Col = .MouseCol
            '.Cell(flexcpForeColor, 0, 0, 0, .Cols - 1) = vbBlack
            '.Cell(flexcpForeColor, 0, .MouseCol, 0, .MouseCol) = vbRed
            'strData = Trim(.Cell(flexcpData, .Row, 0))
            'Select Case .MouseCol
            '       Case 1
            '            .ColSel = 2
            '            .ColSort(0) = flexSortNone
            '            .ColSort(1) = IIf(P_intButton = 1, flexSortGenericAscending, flexSortGenericDescending)
            '            .ColSort(2) = IIf(P_intButton = 1, flexSortGenericAscending, flexSortGenericDescending)
            '            .Sort = flexSortUseColSort
            '       Case 2
            '            .ColSel = 2
            '            .ColSort(0) = flexSortNone
            '            .ColSort(1) = IIf(P_intButton = 1, flexSortGenericAscending, flexSortGenericDescending)
            '            .ColSort(2) = IIf(P_intButton = 1, flexSortGenericAscending, flexSortGenericDescending)
            '            .Sort = flexSortUseColSort
            '       Case Else
            '            .Sort = IIf(P_intButton = 1, flexSortGenericAscending, flexSortGenericDescending)
            'End Select
            'If .FindRow(strData, , 0) > 0 Then
            '   .Row = .FindRow(strData, , 0)
            'End If
            'If PC_intRowCnt < .Rows Then
            '   .TopRow = .Row
            'End If
         End If
    End With
End Sub
Private Sub vsfg1_EnterCell()
Dim lngR As Long
Dim lngC As Long
    With vsfg1
         Text1(0).Enabled = False: Text1(2).Enabled = False
         If .Row >= .FixedRows Then
            Text1(0).Text = .TextMatrix(.Row, 4): Text1(1).Text = .TextMatrix(.Row, 5) '����ó
            Text1(2).Text = .TextMatrix(.Row, 6): Text1(3).Text = .TextMatrix(.Row, 7) '����
            Text1(4).Text = .TextMatrix(.Row, 8)                      '�԰�
            dtpAppDate.Value = Format(DTOS(.TextMatrix(.Row, 3)), "0000-00-00")      '�������
            Text1(5).Text = Format(.ValueMatrix(.Row, 9), "#,0")      '���Լ���
            Text1(6).Text = .TextMatrix(.Row, 10)                     '����
            Text1(7).Text = Format(.ValueMatrix(.Row, 11), "#,0.00")  '���ܰ�
            Text1(8).Text = Format(.ValueMatrix(.Row, 12), "#,0")     '���ΰ�
            Text1(9).Text = .TextMatrix(.Row, 15)                     '����
            If .Cell(flexcpChecked, .Row, 14) = flexChecked Then
               chkCash.Value = 1
            Else
               chkCash.Value = 0
            End If
         End If
    End With
End Sub
Private Sub vsfg1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strSQL       As String
Dim intRetVal    As Integer
Dim lngCnt       As Long
    With vsfg1
         'If KeyCode = vbKeyInsert Then
         '   SubClearText
         '   .Row = 0
         '   Text1(Text1.LBound).SetFocus
         '   Exit Sub
         'End If
    End With
    Exit Sub
ERROR_TABLE_SELECT:
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "����� �б� ����"
    Unload Me
    Exit Sub

End Sub

'+-----------+
'/// �߰� ///
'+-----------+
Private Sub cmdClear_Click()
Dim strSQL As String
    SubClearText
    vsfg1.Row = 0
    Text1(0).Enabled = True
    Text1(2).Enabled = True
    Text1(0).SetFocus
End Sub
'+-----------+
'/// ��ȸ ///
'+-----------+
Private Sub cmdFind_Click()
    Screen.MousePointer = vbHourglass
    cmdFind.Enabled = False
    SubClearText
    Subvsfg1_FILL
    cmdFind.Enabled = True
    Screen.MousePointer = vbDefault
End Sub
'+-----------+
'/// ���� ///
'+-----------+
Private Sub cmdSave_Click()
Dim strSQL        As String
Dim lngR          As Long
Dim lngC          As Long
Dim lngCnt        As Long
Dim lngLogCnt     As Long
Dim blnOK         As Boolean
Dim intRetVal     As Integer
Dim strServerTime As String
Dim intChkCash    As Integer
    '�Է³��� �˻�
    blnOK = False
    FncCheckTextBox lngC, blnOK
    If blnOK = False Then
       Select Case lngC
              Case 0, 2
                   If Text1(lngC).Enabled = False Then
                      Text1(0).Enabled = True: Text1(2).Enabled = True
                   End If
       End Select
       Text1(lngC).SetFocus
       Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    P_adoRec.CursorLocation = adUseClient
    '�������⳻�� �˻�
    strSQL = "SELECT ISNULL(SUM(T1.������),0) AS ������ " _
             & "FROM �������⳻�� T1 " _
            & "WHERE T1.������ڵ� = '" & PB_regUserinfoU.UserBranchCode & "' " _
              & "AND (T1.�з��ڵ� + T1.�����ڵ�) = '" & Text1(2).Text & "' " _
              & "AND T1.����ó�ڵ� = '" & Text1(0).Text & "' " _
              & "AND ((T1.������� = 2 AND T1.��������� = '" & DTOS(dtpAppDate.Value) & "') " _
               & "OR (T1.������� = 2 AND T1.������������� = '" & DTOS(dtpAppDate.Value) & "')) "
    On Error GoTo ERROR_TABLE_SELECT
    P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
    If P_adoRec.RecordCount = 0 Then
       P_adoRec.Close
       dtpAppDate.SetFocus
       Screen.MousePointer = vbDefault
       Exit Sub
    Else
       lngCnt = P_adoRec("������")
       P_adoRec.Close
       If lngCnt = 0 Then
          MsgBox "������� ���� ����ó���� �� �����ϴ�.", vbCritical, "���� ���� �Ұ�"
          Screen.MousePointer = vbDefault
          Exit Sub
       Else
          If (lngCnt < 1) Or (lngCnt < (Vals(Text1(5).Text) * -1)) Then
             MsgBox "���Լ����� ������(" & Format(lngCnt, "#,#") & ") ���� ���Ƽ� ����ó���� �� �����ϴ�.", vbCritical, "���� ���� �Ұ�"
             Text1(5).SetFocus
             Screen.MousePointer = vbDefault
             Exit Sub
          End If
       End If
    End If
    Screen.MousePointer = vbDefault
    If Text1(Text1.LBound).Enabled = True Then
       intRetVal = MsgBox("�Էµ� �ڷḦ �߰��Ͻðڽ��ϱ� ?", vbQuestion + vbYesNo + vbDefaultButton1, "�ڷ� �߰�")
    Else
       If vsfg1.ValueMatrix(vsfg1.Row, 19) = 1 Then
          MsgBox "���ݰ�꼭 ����� ������ ������ �� �����ϴ�.", vbCritical, "���ݰ�꼭 �����"
          Exit Sub
       End If
       intRetVal = MsgBox("������ �ڷḦ �����Ͻðڽ��ϱ� ?", vbQuestion + vbYesNo + vbDefaultButton1, "�ڷ� ����")
    End If
    If intRetVal = vbNo Then
       vsfg1.SetFocus
       Exit Sub
    End If
    cmdSave.Enabled = False
    With vsfg1
         '�����ð� ���ϱ�
         cmdSave.Enabled = False
         Screen.MousePointer = vbHourglass
         P_adoRec.CursorLocation = adUseClient
         strSQL = "SELECT RIGHT(CONVERT(VARCHAR(23),GETDATE(), 121),12) AS �����ð� "
         On Error GoTo ERROR_FORM_ACTIVATE_CONNECTION_SERVER
         P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
         strServerTime = Mid(P_adoRec("�����ð�"), 1, 2) + Mid(P_adoRec("�����ð�"), 4, 2) + Mid(P_adoRec("�����ð�"), 7, 2) _
                       + Mid(P_adoRec("�����ð�"), 10)
         P_adoRec.Close
         intRetVal = MsgBox("���ݹ������� ó���Ͻðڽ��ϱ� ?", vbQuestion + vbYesNo + vbDefaultButton2, "���ݹ���")
         If intRetVal = vbYes Then
            intChkCash = 1
            chkCash.Value = 1
         Else
            intChkCash = 0
            chkCash.Value = 0
         End If
         '����ü� �˻�(�������⳻�� ���̺���)
         strSQL = "SELECT TOP 1 ISNULL(T1.���ܰ�, 0) AS ���ܰ�, ISNULL(T1.���ΰ�, 0) AS ���ΰ� " _
                  & "FROM �������⳻�� T1 " _
                 & "WHERE (T1.������ڵ� = '" & PB_regUserinfoU.UserBranchCode & "') " _
                   & "AND (T1.�з��ڵ� + T1.�����ڵ�) = '" & Text1(2).Text & "' " _
                   & "AND (T1.����ó�ڵ� = '" & Text1(0).Text & "') " _
                   & "AND (T1.������� = 2 AND T1.��������� = '" & DTOS(dtpAppDate.Value) & "' AND ������ > 0) "
         On Error GoTo ERROR_TABLE_SELECT
         P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
         If P_adoRec.RecordCount = 0 Then
            P_adoRec.Close
            dtpAppDate.SetFocus
            Screen.MousePointer = vbDefault
            cmdSave.Enabled = True
            Exit Sub
         Else
            Text1(7).Text = Format(P_adoRec("���ܰ�"), "#,0.00")
            Text1(8).Text = Format(P_adoRec("���ΰ�"), "#,0")
         End If
         P_adoRec.Close
         PB_adoCnnSQL.BeginTrans
         If Text1(Text1.LBound).Enabled = True Then '��ǰ �߰�
            '�ŷ���ȣ ���ϱ�
            P_adoRec.CursorLocation = adUseClient
            strSQL = "spLogCounter '�������⳻��', '" & PB_regUserinfoU.UserBranchCode + PB_regUserinfoU.UserClientDate + "2" & "', " _
                                   & "0, 0, '" & PB_regUserinfoU.UserServerDate & "', '" & PB_regUserinfoU.UserCode & "' "
            On Error GoTo ERROR_STORED_PROCEDURE
            P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
            lngLogCnt = P_adoRec(0)
            P_adoRec.Close
            strSQL = "INSERT INTO �������⳻��(������ڵ�, �з��ڵ�, �����ڵ�, " _
                                            & "�������, ���������, �����ð�, " _
                                            & "�԰����, �԰�ܰ�, �԰�ΰ�, " _
                                            & "������, ���ܰ�, ���ΰ�, " _
                                            & "����ó�ڵ�, ����ó�ڵ�, �������������, ���۱���, " _
                                            & "�߰�����, �߰߹�ȣ, �ŷ�����, �ŷ���ȣ, " _
                                            & "��꼭���࿩��, ���ݱ���, ��������, ����, å��ȣ, �Ϸù�ȣ, " _
                                            & "��뱸��, ��������, ������ڵ�, ����̵�������ڵ�) Values( " _
                    & "'" & PB_regUserinfoU.UserBranchCode & "','" & Mid(Text1(2).Text, 1, 2) & "','" & Mid(Text1(2).Text, 3) & "', " _
                    & "2, '" & PB_regUserinfoU.UserClientDate & "', '" & strServerTime & "', " _
                    & "0 ,0 ,0, " _
                    & "" & Vals(Text1(5).Text) & ", " & Vals(Text1(7).Text) & ", " & Vals(Text1(8).Text) & ", " _
                    & "'', '" & Text1(0).Text & "', '" & DTOS(dtpAppDate.Value) & "', 0, '', 0, " _
                    & "'" & PB_regUserinfoU.UserClientDate & "', " & lngLogCnt & ", " _
                    & "0, " & intChkCash & ", 0, '" & Text1(9).Text & "', 0, 0, " _
                    & "0, '" & PB_regUserinfoU.UserServerDate & "', '" & PB_regUserinfoU.UserCode & "', '' )"
            On Error GoTo ERROR_TABLE_INSERT
            PB_adoCnnSQL.Execute strSQL
            .AddItem .Rows
            .TextMatrix(.Rows - 1, 0) = Text1(2).Text & "-" & Text1(0) & "-" & "" & "-" & "2" & "-" _
                                      & PB_regUserinfoU.UserClientDate & "-" & strServerTime
            .Cell(flexcpData, .Rows - 1, 0, .Rows - 1, 0) = Trim(.TextMatrix(.Rows - 1, 0))
            .TextMatrix(.Rows - 1, 1) = Format(PB_regUserinfoU.UserClientDate, "0000-00-00") '���������(��������)
            .TextMatrix(.Rows - 1, 2) = Format(Mid(strServerTime, 1, 6), "00:00:00")         '�����ð�
            .TextMatrix(.Rows - 1, 3) = Format(DTOS(dtpAppDate.Value), "0000-00-00")         '�������
            .TextMatrix(.Rows - 1, 4) = Text1(0).Text: .TextMatrix(.Rows - 1, 5) = Text1(1).Text '����ó
            .TextMatrix(.Rows - 1, 6) = Text1(2).Text: .TextMatrix(.Rows - 1, 7) = Text1(3).Text '����
            .TextMatrix(.Rows - 1, 8) = Text1(4).Text                                        '�԰�
            .TextMatrix(.Rows - 1, 9) = Vals(Text1(5).Text)                                  '���Լ���
            .TextMatrix(.Rows - 1, 10) = Text1(6).Text                                       '����
            .TextMatrix(.Rows - 1, 11) = Vals(Text1(7).Text)                                 '���ܰ�
            .TextMatrix(.Rows - 1, 12) = Vals(Text1(8).Text)                                 '���ΰ�
            .TextMatrix(.Rows - 1, 13) = .ValueMatrix(.Rows - 1, 9) * .ValueMatrix(.Rows - 1, 11)  '���ݾ�
            .TextMatrix(.Rows - 1, 14) = intChkCash                                          '���ݱ���
            If intChkCash = 1 Then
               .Cell(flexcpChecked, .Rows - 1, 14) = flexChecked   '1
            Else
               .Cell(flexcpChecked, .Rows - 1, 14) = flexUnchecked '2
            End If
            .Cell(flexcpText, .Rows - 1, 14) = "���ݹ���"
            .TextMatrix(.Rows - 1, 15) = Text1(9).Text                                       '����
            .TextMatrix(.Rows - 1, 16) = PB_regUserinfoU.UserCode                            '������ڵ�
            .TextMatrix(.Rows - 1, 17) = PB_regUserinfoU.UserName                            '����ڸ�
            .TextMatrix(.Rows - 1, 18) = strServerTime                                       '�����ð�
            If .Rows > PC_intRowCnt Then
               .ScrollBars = flexScrollBarBoth
               .TopRow = .Rows - 1
            End If
            Text1(Text1.LBound).Enabled = False
            .Row = .Rows - 1          '�ڵ����� vsfg1_EnterCell Event �߻�
         Else                                          '�������⳻�� ����
            strSQL = "UPDATE �������⳻�� SET " _
                          & "������������� = '" & DTOS(dtpAppDate.Value) & "', " _
                          & "������ = " & Vals(Text1(5).Text) & ", " _
                          & "���ܰ� = " & Vals(Text1(7).Text) & ", " _
                          & "���ΰ� = " & Vals(Text1(8).Text) & ", " _
                          & "���ݱ��� = " & intChkCash & ", " _
                          & "���� = '" & Trim(Text1(9).Text) & "', " _
                          & "�������� = '" & PB_regUserinfoU.UserServerDate & "', " _
                          & "������ڵ� = '" & PB_regUserinfoU.UserCode & "' " _
                    & "WHERE ������ڵ� = '" & PB_regUserinfoU.UserBranchCode & "' " _
                      & "AND �з��ڵ� = '" & Mid(.TextMatrix(.Row, 6), 1, 2) & "' " _
                      & "AND �����ڵ� = '" & Mid(.TextMatrix(.Row, 6), 3) & "' " _
                      & "AND ����ó�ڵ� = '' " _
                      & "AND ����ó�ڵ� = '" & .TextMatrix(.Row, 4) & "' " _
                      & "AND ������� = 2 " _
                      & "AND ��������� = '" & DTOS(.TextMatrix(.Row, 1)) & "' " _
                      & "AND �����ð� = '" & .TextMatrix(.Row, 18) & "' "
            On Error GoTo ERROR_TABLE_UPDATE
            PB_adoCnnSQL.Execute strSQL
            .TextMatrix(.Row, 3) = Format(DTOS(dtpAppDate.Value), "0000-00-00")     '�������������
            .TextMatrix(.Row, 9) = Vals(Text1(5).Text)                              '���Լ���
            .TextMatrix(.Row, 11) = Vals(Text1(7).Text)                             '���ܰ�
            .TextMatrix(.Row, 12) = Vals(Text1(8).Text)                             '���ΰ�
            .TextMatrix(.Row, 13) = .ValueMatrix(.Row, 9) * (.ValueMatrix(.Row, 11))  '���ݾ�
            .TextMatrix(.Row, 14) = intChkCash
            If intChkCash = 1 Then                                                  '���ݱ���
               .Cell(flexcpChecked, .Row, 14) = flexChecked    '1
            Else
               .Cell(flexcpChecked, .Row, 14) = flexUnchecked  '2
            End If
            .Cell(flexcpText, .Row, 14) = "���ݹ���"
            .TextMatrix(.Row, 15) = Text1(9).Text                                   '����
            .TextMatrix(.Row, 16) = PB_regUserinfoU.UserCode                        '������ڵ�
            .TextMatrix(.Row, 17) = PB_regUserinfoU.UserName                        '����ڸ�
            'if x then '���ݰ�꼭������̸�
            'end if
         End If
         PB_adoCnnSQL.CommitTrans
         .SetFocus
         Screen.MousePointer = vbDefault
    End With
    cmdSave.Enabled = True
    Exit Sub
ERROR_FORM_ACTIVATE_CONNECTION_SERVER:
    MsgBox Err.Number & Err.Description & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "(�������� ���� ����)"
    Unload Me
    Exit Sub
ERROR_TABLE_SELECT:
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "���Գ��� �б� ����"
    Unload Me
    Exit Sub
ERROR_TABLE_INSERT:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "���Գ��� �߰� ����"
    Unload Me
    Exit Sub
ERROR_TABLE_UPDATE:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "���Գ��� ���� ����"
    Unload Me
    Exit Sub
ERROR_STORED_PROCEDURE:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "�α� ���� ����"
    Unload Me
    Exit Sub
End Sub

Private Sub cmdDelete_Click()
Dim strSQL       As String
Dim intRetVal    As Integer
Dim lngCnt       As Long
    With vsfg1
         If .Row >= .FixedRows Then
            intRetVal = MsgBox("��ϵ� �ڷḦ �����Ͻðڽ��ϱ� ?", vbCritical + vbYesNo + vbDefaultButton2, "�ڷ� ����")
            If vsfg1.ValueMatrix(vsfg1.Row, 19) = 1 Then
               MsgBox "���ݰ�꼭 ����� ������ ������ �� �����ϴ�.", vbCritical, "���ݰ�꼭 �����"
               Exit Sub
            End If
            If intRetVal = vbYes Then
               Screen.MousePointer = vbHourglass
               cmdDelete.Enabled = False
               '������ �������̺� �˻�
               'P_adoRec.CursorLocation = adUseClient
               'strSQL = "SELECT Count(*) AS �ش�Ǽ� FROM TableName " _
               '        & "WHERE ����屸�� = " & .TextMatrix(.Row, 0) & " "
               'On Error GoTo ERROR_TABLE_SELECT
               'P_adoRec.Open SQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
               'lngCnt = P_adoRec("�ش�Ǽ�")
               'P_adoRec.Close
               If lngCnt <> 0 Then
                  MsgBox "XXX(" & Format(lngCnt, "#,#") & "��)�� �����Ƿ� ������ �� �����ϴ�.", vbCritical, "���Գ��� ���� �Ұ�"
                  cmdDelete.Enabled = True
                  Screen.MousePointer = vbDefault
                  Exit Sub
               End If
               PB_adoCnnSQL.BeginTrans
               P_adoRec.CursorLocation = adUseClient
               strSQL = "UPDATE �������⳻�� SET " _
                             & "��뱸�� = 9, " _
                             & "���� = '" & Trim(Text1(9).Text) & "', " _
                             & "�������� = '" & PB_regUserinfoU.UserServerDate & "', " _
                             & "������ڵ� = '" & PB_regUserinfoU.UserCode & "' " _
                       & "WHERE ������ڵ� = '" & PB_regUserinfoU.UserBranchCode & "' " _
                         & "AND �з��ڵ� = '" & Mid(.TextMatrix(.Row, 6), 1, 2) & "' " _
                         & "AND �����ڵ� = '" & Mid(.TextMatrix(.Row, 6), 3) & "' " _
                         & "AND ����ó�ڵ� = '" & .TextMatrix(.Row, 4) & "' " _
                         & "AND ����ó�ڵ� = '' " _
                         & "AND ������� = 2 " _
                         & "AND ��������� = '" & DTOS(.TextMatrix(.Row, 1)) & "' " _
                         & "AND �����ð� = '" & .TextMatrix(.Row, 18) & "' "
               On Error GoTo ERROR_TABLE_DELETE
               PB_adoCnnSQL.Execute strSQL
               PB_adoCnnSQL.CommitTrans
               .RemoveItem .Row
               Text1(Text1.LBound).Enabled = False
               If .Rows <= PC_intRowCnt Then
                  .ScrollBars = flexScrollBarHorizontal
               End If
               Screen.MousePointer = vbDefault
               If .Rows = 1 Then
                  SubClearText
                  .Row = 0
                  Text1(Text1.LBound).Enabled = True
                  Text1(Text1.LBound).SetFocus
                  Exit Sub
               End If
               vsfg1_EnterCell
               cmdDelete.Enabled = True
            End If
            .SetFocus
         End If
    End With
    Screen.MousePointer = vbDefault
    Exit Sub
ERROR_TABLE_DELETE:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "���Գ��� ���� ����"
    Unload Me
    Screen.MousePointer = vbDefault
    Exit Sub
End Sub

'+-----------+
'/// ���� ///
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
    Set frm���԰��� = Nothing
    frmMain.SBar.Panels(4).Text = ""
End Sub

'+--------------------+
'/// Sub Procedure ///
'+--------------------+
'+----------------------------------+
'/// VsFlexGrid(vsfg1) �ʱ�ȭ ///
'+----------------------------------+
Private Sub Subvsfg1_INIT()
Dim lngR    As Long
Dim lngC    As Long
    Text1(Text1.LBound).Enabled = False                '����ó�ڵ� FLASE
    With vsfg1              'Rows 1, Cols 22, RowHeightMax(Min) 300
         .AllowBigSelection = False
         .AllowSelection = False
         .AllowUserResizing = flexResizeColumns
         .BackColorBkg = &H8000000F
         .BackColorSel = &H8000&
         .ExtendLastCol = True
         .FocusRect = flexFocusHeavy
         .ScrollBars = flexScrollBarHorizontal
         .ScrollTrack = True
         .SelectionMode = flexSelectionByRow
         .FixedRows = 1
         .FixedCols = 3
         .Rows = 1             'Subvsfg1_Fill����ÿ� ����
         .Cols = 22
         .RowHeightMax = 300
         .RowHeightMin = 300
         .ColWidth(0) = 1000   'H, KEY(�����ڵ�-����ó�ڵ�-����ó�ڵ�-�������-���������-�����ð�)
         .ColWidth(1) = 1200   '��������
         .ColWidth(2) = 1200   '���Խð�(���ʺ�)
         .ColWidth(3) = 1200   '�������(�����������)
         .ColWidth(4) = 1000   'H, ����ó�ڵ�
         .ColWidth(5) = 2000   '����ó��
         .ColWidth(6) = 1900   '�����ڵ�
         .ColWidth(7) = 2500   '�����
         .ColWidth(8) = 2000   '����԰�
         .ColWidth(9) = 800    '���Լ���
         .ColWidth(10) = 600   '�������
         .ColWidth(11) = 1500  '���ܰ�
         .ColWidth(12) = 1300  'H, ���ΰ�
         .ColWidth(13) = 1600  '���ݾ�
         .ColWidth(14) = 1200  '���ⱸ��
         .ColWidth(15) = 5000  '����
         .ColWidth(16) = 1000  '������ڵ�
         .ColWidth(17) = 1000  '����ڸ�
         .ColWidth(18) = 1000  '�����ð�
         
         .ColWidth(19) = 1000  '��꼭���࿩��
         .ColWidth(20) = 1000  '���ݰ�꼭(å��ȣ)
         .ColWidth(21) = 1000  '���ݰ�꼭(�Ϸù�ȣ)
         
         .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = False
         .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
         .TextMatrix(0, 0) = "KEY"         'H
         .TextMatrix(0, 1) = "��������"
         .TextMatrix(0, 2) = "���Խð�"
         .TextMatrix(0, 3) = "��������"
         .TextMatrix(0, 4) = "����ó�ڵ�"  'H
         .TextMatrix(0, 5) = "����ó��"
         .TextMatrix(0, 6) = "ǰ���ڵ�"
         .TextMatrix(0, 7) = "ǰ��"
         .TextMatrix(0, 8) = "�԰�"
         .TextMatrix(0, 9) = "����"
         .TextMatrix(0, 10) = "����"
         .TextMatrix(0, 11) = "����ܰ�"
         .TextMatrix(0, 12) = "����ΰ�"   'H
         .TextMatrix(0, 13) = "����ݾ�"
         .TextMatrix(0, 14) = "���ⱸ��"
         .TextMatrix(0, 15) = "����"
         .TextMatrix(0, 16) = "������ڵ�" 'H
         .TextMatrix(0, 17) = "����ڸ�"
         .TextMatrix(0, 18) = "�����ð�" 'H
         .TextMatrix(0, 19) = "��꼭����" 'H
         .TextMatrix(0, 20) = "å��ȣ"     'H
         .TextMatrix(0, 21) = "�Ϸù�ȣ"   'H
         .ColHidden(0) = True: .ColHidden(4) = True:  .ColHidden(12) = True
         .ColHidden(16) = True: .ColHidden(18) = True: .ColHidden(19) = True: .ColHidden(20) = True: .ColHidden(21) = True
         .ColFormat(9) = ","
         For lngC = 0 To .Cols - 1
             Select Case lngC
                    Case 0, 5, 6, 7, 8, 10, 14, 15
                         .ColAlignment(lngC) = flexAlignLeftCenter
                    Case 1, 2, 3, 4, 16, 17, 18, 19
                         .ColAlignment(lngC) = flexAlignCenterCenter
                    Case Else
                         .ColAlignment(lngC) = flexAlignRightCenter
             End Select
         Next lngC
         For lngC = 0 To .Cols - 1
             Select Case lngC
                    Case 11, 13
                         .ColFormat(lngC) = "#,#.00"
             End Select
         Next lngC
    End With
End Sub

'+---------------------------------+
'/// VsFlexGrid(vsfg1) ä���///
'+---------------------------------+
Private Sub Subvsfg1_FILL()
Dim strSQL     As String
Dim strWhere   As String
Dim strOrderBy As String
Dim lngR       As Long
Dim lngC       As Long
Dim lngRR      As Long
    Text1(0).Enabled = False: Text1(2).Enabled = False
    If dtpFDate > dtpTDate Then
       dtpFDate.SetFocus
       Exit Sub
    End If
    P_adoRec.CursorLocation = adUseClient
    strSQL = "SELECT (T1.�з��ڵ� + T1.�����ڵ�) AS �����ڵ�, ISNULL(T2.�����,'') AS �����, " _
                  & "ISNULL(T1.����ó�ڵ�,'') AS ����ó�ڵ�, ISNULL(T3.����ó��,'') AS ����ó��, " _
                  & "ISNULL(T1.����ó�ڵ�,'') AS ����ó�ڵ�, ISNULL(T4.����ó��,'') AS ����ó��, " _
                  & "T1.�������,  T1.���������, T1.�����ð�, " _
                  & "T2.�԰� AS ����԰�,  T2.���� AS �������, " _
                  & "T1.�԰���� AS �԰����, T1.�԰�ܰ�, �԰�ΰ�, " _
                  & "T1.������ AS ������, T1.���ܰ�, ���ΰ�, " _
                  & "T1.�������������, T1.���ݱ���, T1.����, T1.������ڵ�, ISNULL(T5.����ڸ�, '') AS ����ڸ�, " _
                  & "T1.��꼭���࿩��, T1.å��ȣ, T1.�Ϸù�ȣ " _
             & "FROM �������⳻�� T1 " _
             & "LEFT JOIN ���� T2 " _
                    & "ON T2.�з��ڵ� = T1.�з��ڵ� AND T2.�����ڵ� = T1.�����ڵ� " _
             & "LEFT JOIN ����ó T3 " _
                    & "ON T3.������ڵ� = T1.������ڵ� AND T3.����ó�ڵ�= T1.����ó�ڵ� " _
             & "LEFT JOIN ����ó T4 " _
                    & "ON T4.������ڵ� = T1.������ڵ� AND T4.����ó�ڵ�= T1.����ó�ڵ� " _
             & "LEFT JOIN ����� T5 " _
                    & "ON T5.������ڵ� = T1.������ڵ� AND T5.������ڵ�= T1.������ڵ� " _
            & "WHERE T1.������ڵ� = '" & PB_regUserinfoU.UserBranchCode & "' AND T1.��뱸�� = 0 " _
              & "AND T1.������� = 2 AND ������ < 0 " _
              & "AND T1.��������� BETWEEN '" & DTOS(dtpFDate.Value) & "' AND '" & DTOS(dtpTDate.Value) & "' " _
            & "ORDER BY T1.���������, T1.�����ð�, T1.������������� "
    On Error GoTo ERROR_TABLE_SELECT
    P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
    vsfg1.Rows = P_adoRec.RecordCount + 1
    If P_adoRec.RecordCount = 0 Then
       P_adoRec.Close
    Else
       With vsfg1
            .Cell(flexcpForeColor, 0, 0, 0, .Cols - 1) = vbBlack
            If .Rows <= PC_intRowCnt Then
               .ScrollBars = flexScrollBarHorizontal
            Else
               .ScrollBars = flexScrollBarBoth
            End If
            Do Until P_adoRec.EOF
               lngR = lngR + 1
               .TextMatrix(lngR, 0) = P_adoRec("�����ڵ�") & "-" & P_adoRec("����ó�ڵ�") & "-" _
                                    & P_adoRec("����ó�ڵ�") & "-" & P_adoRec("�������") & "-" _
                                    & P_adoRec("���������") & "-" & P_adoRec("�����ð�") & "-"
               .Cell(flexcpData, lngR, 0, lngR, 0) = Trim(.TextMatrix(lngR, 0)) 'FindRow ����� ����
               .TextMatrix(lngR, 1) = Format(P_adoRec("���������"), "0000-00-00")
               .TextMatrix(lngR, 2) = Format(Mid(P_adoRec("�����ð�"), 1, 6), "00:00:00")
               .TextMatrix(lngR, 3) = Format(P_adoRec("�������������"), "0000-00-00")
               .TextMatrix(lngR, 4) = P_adoRec("����ó�ڵ�")
               .TextMatrix(lngR, 5) = P_adoRec("����ó��")
               .TextMatrix(lngR, 6) = P_adoRec("�����ڵ�")
               .TextMatrix(lngR, 7) = P_adoRec("�����")
               .TextMatrix(lngR, 8) = P_adoRec("����԰�")
               .TextMatrix(lngR, 9) = P_adoRec("������")
               .TextMatrix(lngR, 10) = P_adoRec("�������")
               .TextMatrix(lngR, 11) = P_adoRec("���ܰ�")
               .TextMatrix(lngR, 12) = P_adoRec("���ΰ�")
               .TextMatrix(lngR, 13) = .ValueMatrix(lngR, 9) * (.ValueMatrix(lngR, 11))
               .TextMatrix(lngR, 14) = IIf(IsNull(P_adoRec("���ݱ���")), 0, P_adoRec("���ݱ���"))
               If P_adoRec("���ݱ���") = 1 Then
                  .Cell(flexcpChecked, lngR, 14) = flexChecked    '1
               Else
                  .Cell(flexcpChecked, lngR, 14) = flexUnchecked  '2
               End If
               .Cell(flexcpText, lngR, 14) = "���ݹ���"
               .TextMatrix(lngR, 15) = P_adoRec("����")
               .TextMatrix(lngR, 16) = P_adoRec("������ڵ�")
               .TextMatrix(lngR, 17) = P_adoRec("����ڸ�")
               .TextMatrix(lngR, 18) = P_adoRec("�����ð�")
               .TextMatrix(lngR, 19) = P_adoRec("��꼭���࿩��")
               .TextMatrix(lngR, 20) = P_adoRec("å��ȣ")
               .TextMatrix(lngR, 21) = P_adoRec("�Ϸù�ȣ")
               'If .TextMatrix(lngR, 0) = PB_regUserinfoU.UserBranchCode Then
               '   lngRR = lngR
               'End If
               P_adoRec.MoveNext
            Loop
            P_adoRec.Close
            If lngRR = 0 Then
               .Row = lngRR       'vsfg1_EnterCell �ڵ�����(���� �Ѱ� �϶��� �ڵ����� ����)
               If .Rows > PC_intRowCnt Then
                  '.TopRow = .Rows - PC_intRowCnt + .FixedRows
                  .TopRow = 1
               End If
            Else
               .Row = lngRR       'vsfg1_EnterCell �ڵ�����(���� �Ѱ� �϶��� �ڵ����� ����)
               If .Rows > PC_intRowCnt Then
                  .TopRow = .Row
               End If
            End If
            vsfg1_EnterCell       'vsfg1_EnterCell �ڵ�����(���� �Ѱ� �϶��� ������ �ڵ�����)
            .SetFocus
       End With
    End If
    Exit Sub
ERROR_TABLE_SELECT:
    MsgBox Err.Number & Err.Description & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "��ǰ���� �б� ����"
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
End Sub

'+-------------------------+
'/// Check text1(index) ///
'+-------------------------+
Private Function FncCheckTextBox(lngC As Long, blnOK As Boolean)
    For lngC = Text1.LBound To Text1.UBound
        Select Case lngC
               Case 0  '����ó�ڵ�
                    If Not (Len(Text1(lngC).Text) > 0) Then
                       Exit Function
                    End If
               Case 1  '����ó��
                    If Len(Trim(Text1(lngC).Text)) = 0 Then
                       lngC = 0
                       Exit Function
                    End If
               Case 2  '�����ڵ�
                    If Not (Len(Text1(lngC).Text) > 0) Then
                       Exit Function
                    End If
               Case 3  '�����
                    If Len(Trim(Text1(lngC).Text)) = 0 Then
                       lngC = 2
                       Exit Function
                    End If
               Case 5  '���Լ���
                    If Not (Vals(Text1(lngC).Text) < 0) Then
                       Exit Function
                    End If
               Case 9  '����
                    If Not (LenH(Text1(lngC).Text) <= 50) Then
                       Exit Function
                    End If
               Case Else
        End Select
    Next lngC
    blnOK = True
End Function

'+---------------------------+
'/// ũ����Ż ������ ��� ///
'+---------------------------+
Private Sub cmdPrint_Click()
Dim strSQL                 As String
Dim strWhere               As String
Dim strOrderBy             As String

Dim varRetVal              As Variant '������ ����
Dim strExeFile             As String
Dim strExeMode             As String
Dim intRetCHK              As Integer '���࿩��

Dim lngR                   As Long
Dim lngC                   As Long
Dim strForPrtDateTime      As String  '����Ͻ�           (Formula)

    If dtpFDate.Value > dtpTDate.Value Then
       dtpFDate.SetFocus
       Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    '�����Ͻ�(����Ͻ�)
    strSQL = "SELECT CONVERT(VARCHAR(19),GETDATE(), 120) AS �����ð� "
    On Error GoTo ERROR_FORM_ACTIVATE_CONNECTION_SERVER
    P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
    strForPrtDateTime = Format(PB_regUserinfoU.UserServerDate, "0000-00-00") & Space(1) _
                      & Format(Right(P_adoRec("�����ð�"), 8), "hh:mm:ss")
    P_adoRec.Close
    
    intRetCHK = 99
    With CrystalReport1
         If PB_Test = 0 Then
            strExeFile = App.Path & IIf(optPrtChk0.Value = True, ".\���Գ���(���ں�).rpt", ".\���Գ���(��ü��).rpt")
         Else
            strExeFile = App.Path & IIf(optPrtChk0.Value = True, ".\���Գ���(���ں�)T.rpt", ".\���Գ���(��ü��)T.rpt")
         End If
         varRetVal = Dir(strExeFile)
         If Len(varRetVal) = 0 Then
            intRetCHK = 0
         Else
            .ReportFileName = strExeFile
            On Error GoTo ERROR_CRYSTAL_REPORTS
            '--- Formula Fields ---
            .Formulas(0) = "ForPgDate = '" & Format(PB_regUserinfoU.UserClientDate, "0000-00-00") & "'" '���α׷���������
            .Formulas(1) = "ForBranchName= '" & PB_regUserinfoU.UserBranchName & "'"                    '������
            .Formulas(2) = "ForPrtDateTime = '" & strForPrtDateTime & "'"                               '����Ͻ�
            .Formulas(3) = "ForAppDate = '�������� : ' & '" & Format(DTOS(dtpFDate.Value), "0000-00-00") & "' & ' ���� ' & '" & Format(DTOS(dtpTDate.Value), "0000-00-00") & "' & ' ����' "
            '--- Parameter Fields ---
            .StoredProcParam(0) = PB_regUserinfoU.UserBranchCode  '�����ڵ�
            .StoredProcParam(1) = DTOS(dtpFDate.Value)            '��������(��������)
            .StoredProcParam(2) = DTOS(dtpTDate.Value)            '��������(��������)
            .StoredProcParam(3) = " "                             '�����
            .StoredProcParam(4) = " "                             '��ü��
         End If
         If intRetCHK = 99 Then
            .Connect = PB_adoCnnSQL.ConnectionString
            .Destination = crptToWindow
            .DiscardSavedData = True
            .ProgressDialog = True
            .ReportSource = crptReport
            .WindowAllowDrillDown = False
            .WindowShowProgressCtls = True
            .WindowShowCloseBtn = True
            .WindowShowExportBtn = False
            .WindowShowGroupTree = True
            .WindowShowPrintSetupBtn = True
            .WindowTop = 0: .WindowTop = 0: .WindowHeight = 11100: .WindowWidth = 15405
            .WindowState = crptMaximized
            .WindowTitle = PB_regUserinfoU.UserBranchName & " - " & IIf(optPrtChk0.Value = True, "���Գ���(���ں�).rpt", "���Գ���(��ü��)")
            .Action = 1
            .Reset
         End If
    End With
    Screen.MousePointer = vbDefault
    Exit Sub
ERROR_FORM_ACTIVATE_CONNECTION_SERVER:
    MsgBox Err.Number & Err.Description & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "(�������� ���� ����)"
    Unload Me
    Exit Sub
ERROR_CRYSTAL_REPORTS:
    MsgBox Err.Number & Space(1) & Err.Description
    Screen.MousePointer = vbDefault
    Exit Sub
End Sub



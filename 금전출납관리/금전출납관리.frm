VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm�����ⳳ���� 
   BorderStyle     =   0  '����
   Caption         =   "��������"
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
      TabIndex        =   18
      Top             =   0
      Width           =   15195
      Begin VB.CommandButton cmdPrint 
         Height          =   375
         Left            =   7980
         Picture         =   "�����ⳳ����.frx":0000
         Style           =   1  '�׷���
         TabIndex        =   22
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdClear 
         Height          =   375
         Left            =   9120
         Picture         =   "�����ⳳ����.frx":0963
         Style           =   1  '�׷���
         TabIndex        =   7
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdExit 
         Height          =   390
         Left            =   13635
         Picture         =   "�����ⳳ����.frx":1308
         Style           =   1  '�׷���
         TabIndex        =   21
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         Height          =   390
         Left            =   12510
         Picture         =   "�����ⳳ����.frx":1C56
         Style           =   1  '�׷���
         TabIndex        =   20
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdSave 
         Height          =   390
         Left            =   11385
         Picture         =   "�����ⳳ����.frx":25DA
         Style           =   1  '�׷���
         TabIndex        =   6
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdFind 
         Height          =   390
         Left            =   10260
         Picture         =   "�����ⳳ����.frx":2E61
         Style           =   1  '�׷���
         TabIndex        =   12
         Top             =   195
         Width           =   1095
      End
      Begin Crystal.CrystalReport CrystalReport1 
         Left            =   4320
         Top             =   195
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin VB.Label Label2 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         BackColor       =   &H00008000&
         BorderStyle     =   1  '���� ����
         Caption         =   "�� �� �� �� �� ��"
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
         TabIndex        =   19
         Top             =   180
         Width           =   4650
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid vsfg1 
      Height          =   7896
      Left            =   60
      TabIndex        =   13
      Top             =   2055
      Width           =   15195
      _cx             =   26802
      _cy             =   13928
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
      Height          =   1395
      Left            =   60
      TabIndex        =   14
      Top             =   630
      Width           =   15195
      Begin VB.ComboBox cboIOGbn 
         Height          =   300
         Left            =   915
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   4
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox txtAccName 
         Appearance      =   0  '���
         Enabled         =   0   'False
         Height          =   285
         Left            =   4995
         MaxLength       =   30
         TabIndex        =   2
         Top             =   240
         Width           =   2895
      End
      Begin VB.TextBox txtFindAccName 
         Appearance      =   0  '���
         Enabled         =   0   'False
         Height          =   285
         Left            =   12480
         MaxLength       =   30
         TabIndex        =   9
         Top             =   240
         Width           =   2295
      End
      Begin VB.TextBox txtFindAccCode 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         Height          =   285
         IMEMode         =   8  '����
         Left            =   10440
         MaxLength       =   4
         TabIndex        =   8
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtAccCode 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         Height          =   285
         IMEMode         =   8  '����
         Left            =   3315
         MaxLength       =   4
         TabIndex        =   1
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtMoney 
         Alignment       =   1  '������ ����
         Appearance      =   0  '���
         Height          =   285
         IMEMode         =   10  '�ѱ� 
         Left            =   4995
         MaxLength       =   60
         TabIndex        =   5
         Top             =   1000
         Width           =   1575
      End
      Begin VB.TextBox txtJukyo 
         Appearance      =   0  '���
         Height          =   285
         IMEMode         =   10  '�ѱ� 
         Left            =   915
         MaxLength       =   60
         TabIndex        =   3
         Top             =   620
         Width           =   5655
      End
      Begin MSComCtl2.DTPicker dtpFDate 
         Height          =   270
         Left            =   10440
         TabIndex        =   10
         Top             =   600
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
         Left            =   12480
         TabIndex        =   11
         Top             =   600
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   476
         _Version        =   393216
         Format          =   57540609
         UpDown          =   -1  'True
         CurrentDate     =   37763
      End
      Begin MSComCtl2.DTPicker dtpCDate 
         Height          =   270
         Left            =   915
         TabIndex        =   0
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
         Caption         =   "�ݾ�"
         Height          =   240
         Index           =   6
         Left            =   4080
         TabIndex        =   34
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "�����ڵ�"
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   5
         Left            =   2040
         TabIndex        =   33
         Top             =   285
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "��(��/��)�հ�"
         Height          =   240
         Index           =   4
         Left            =   8880
         TabIndex        =   32
         Top             =   1020
         Width           =   1335
      End
      Begin VB.Label lblTotOut 
         Alignment       =   1  '������ ����
         Caption         =   "0"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   12480
         TabIndex        =   31
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label lblTotIn 
         Alignment       =   1  '������ ����
         Caption         =   "0"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   10440
         TabIndex        =   30
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "[Home]"
         Height          =   240
         Index           =   3
         Left            =   11760
         TabIndex        =   29
         Top             =   265
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "[Home]"
         Height          =   240
         Index           =   15
         Left            =   4320
         TabIndex        =   28
         Top             =   270
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "�����ڵ�"
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   2
         Left            =   9120
         TabIndex        =   27
         Top             =   285
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "����"
         Height          =   240
         Index           =   12
         Left            =   13800
         TabIndex        =   26
         Top             =   660
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "����"
         Height          =   240
         Index           =   11
         Left            =   11760
         TabIndex        =   25
         Top             =   660
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "����"
         Height          =   240
         Index           =   10
         Left            =   9360
         TabIndex        =   24
         Top             =   660
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
         Left            =   8040
         TabIndex        =   23
         Top             =   285
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "����"
         Height          =   240
         Index           =   8
         Left            =   75
         TabIndex        =   17
         Top             =   1020
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "����"
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   1
         Left            =   75
         TabIndex        =   16
         Top             =   660
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "����"
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   0
         Left            =   75
         TabIndex        =   15
         Top             =   285
         Width           =   735
      End
   End
End
Attribute VB_Name = "frm�����ⳳ����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'+-------------------------------------------------------------------------------------------------------+
' ���α׷� �� �� : ��������
' ���� Control : VideoSoft VSFlexGrid 7.0(OLEDB) = vsflex7.ocx
' ������ Table   :
' ��  ��  ��  �� :
'+-------------------------------------------------------------------------------------------------------+
Option Explicit
Private P_blnActived       As Boolean
Private P_adoRec           As New ADODB.Recordset
Private P_intButton        As Integer
Private Const PC_intRowCnt As Integer = 24  '�׸��� �� ������ �� ���(FixedRows ����)

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
    If P_blnActived = False Then
       Screen.MousePointer = vbHourglass
       On Error GoTo ERROR_FORM_ACTIVATE_CONNECTION_SERVER
       Subvsfg1_INIT
       Select Case Val(PB_regUserinfoU.UserAuthority)
              Case Is <= 10 '��ȸ
                   cmdPrint.Enabled = False: cmdClear.Enabled = False: cmdFind.Enabled = True
                   cmdSave.Enabled = False: cmdDelete.Enabled = False: cmdExit.Enabled = True
              Case Is <= 20 '�μ�, ��ȸ
                   cmdPrint.Enabled = False: cmdClear.Enabled = False: cmdFind.Enabled = True
                   cmdSave.Enabled = False: cmdDelete.Enabled = False: cmdExit.Enabled = True
              Case Is <= 40 '�߰�, ����, �μ�, ��ȸ
                   cmdPrint.Enabled = False: cmdClear.Enabled = True: cmdFind.Enabled = True
                   cmdSave.Enabled = True: cmdDelete.Enabled = False: cmdExit.Enabled = True
              Case Is <= 50 '����, �߰�, ����, ��ȸ, �μ�
                   cmdPrint.Enabled = False: cmdClear.Enabled = True: cmdFind.Enabled = True
                   cmdSave.Enabled = True: cmdDelete.Enabled = True: cmdExit.Enabled = True
              Case Is <= 99 '����, �߰�, ����, ��ȸ, �μ�
                   cmdPrint.Enabled = False: cmdClear.Enabled = True: cmdFind.Enabled = True
                   cmdSave.Enabled = True: cmdDelete.Enabled = True: cmdExit.Enabled = True
              Case Else
                   cmdClear.Enabled = False: cmdFind.Enabled = False: cmdSave.Enabled = False
                   cmdDelete.Enabled = False: cmdExit.Enabled = True
       End Select
       dtpCDate.Value = Format(PB_regUserinfoU.UserClientDate, "0000-00-00")
       dtpFDate.Value = Format(PB_regUserinfoU.UserClientDate, "0000-00-00")
       dtpTDate.Value = Format(PB_regUserinfoU.UserClientDate, "0000-00-00")
       'dtpFDate.Value = Format(Mid(PB_regUserinfoU.UserClientDate, 1, 6) & "01", "0000-00-00")
       'dtpTDate.Value = DateAdd("d", -1, DateAdd("m", 1, dtpFDate.Value))
       SubOther_FILL
       P_blnActived = True
       Screen.MousePointer = vbDefault
    End If
    Exit Sub
ERROR_FORM_ACTIVATE_CONNECTION_SERVER:
    MsgBox Err.Number & Err.Description & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "(�������� ���� ����)"
    Unload Me
    Exit Sub
End Sub

'+--------------------+
'/// �Է�/�������� ///
'+--------------------+
'�ۼ�����
Private Sub dtpCDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub
'�����ڵ�
Private Sub txtAccCode_GotFocus()
    With txtAccCode
         .SelStart = 0
         .SelLength = Len(Trim(.Text))
    End With
End Sub
Private Sub txtAccCode_KeyDown(KeyCode As Integer, Shift As Integer)
    If (KeyCode = vbKeyHome Or KeyCode = vbKeyReturn) Then  '�����ڵ�˻�
       PB_strAccCode = UPPER(Trim(txtAccCode.Text))
       PB_strAccName = ""  'Trim(Text1(Index + 1).Text)
       frm�����ڵ�˻�.Show vbModal
       If (Len(PB_strAccCode) + Len(PB_strAccName)) = 0 Then '�˻����� ���(ESC)
       Else
          txtAccCode.Text = PB_strAccCode
          txtAccName.Text = PB_strAccName
       End If
       If PB_strAccCode <> "" Then
          SendKeys "{TAB}"
       End If
       PB_strAccCode = "": PB_strAccName = ""
    ElseIf _
       KeyCode = vbKeyDelete Then
       If Len(txtAccCode) = 0 Then
          txtAccName.Text = ""
          Exit Sub
       End If
    Else
       If KeyCode = vbKeyReturn Then
          SendKeys "{tab}"
       End If
    End If
    Exit Sub
End Sub
Private Sub txtAccCode_LostFocus()
    With txtAccCode
         .Text = Trim(.Text)
         If Len(.Text) < 1 Then
            txtAccName.Text = ""
         End If
    End With
End Sub

Private Sub txtJukyo_GotFocus()
    With txtJukyo
         .SelStart = 0
         .SelLength = Len(Trim(.Text))
    End With
End Sub
Private Sub txtJukyo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
       SendKeys "{tab}"
    End If
    Exit Sub
End Sub
'���ⱸ��
Private Sub cboIOGbn_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
       SendKeys "{tab}"
    End If
    Exit Sub
End Sub
'�ݾ�
Private Sub txtMoney_GotFocus()
    With txtMoney
         .SelStart = 0
         .SelLength = Len(Trim(.Text))
    End With
End Sub
Private Sub txtMoney_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
       SendKeys "{tab}"
    End If
    Exit Sub
End Sub
Private Sub txtMoney_LostFocus()
    With txtMoney
         .Text = Format(Vals(Trim(.Text)), "#,0")
    End With
End Sub
'+---------------+
'/// �˻����� ///
'+---------------+
Private Sub txtFindAccCode_GotFocus()
    With txtFindAccCode
         .SelStart = 0
         .SelLength = Len(Trim(.Text))
    End With
End Sub
Private Sub txtFindAccCode_KeyDown(KeyCode As Integer, Shift As Integer)
    If (KeyCode = vbKeyHome Or KeyCode = vbKeyReturn) Then  '�����ڵ�˻�
       PB_strAccCode = UPPER(Trim(txtFindAccCode.Text))
       PB_strAccName = ""  'Trim(Text1(Index + 1).Text)
       frm�����ڵ�˻�.Show vbModal
       If (Len(PB_strAccCode) + Len(PB_strAccName)) = 0 Then '�˻����� ���(ESC)
       Else
          txtFindAccCode.Text = PB_strAccCode
          txtFindAccName.Text = PB_strAccName
       End If
       If PB_strAccCode <> "" Then
          SendKeys "{TAB}"
       Else
          SendKeys "{TAB}"
       End If
       PB_strAccCode = "": PB_strAccName = ""
    ElseIf _
       KeyCode = vbKeyDelete Then
       If Len(txtFindAccCode) = 0 Then
          txtFindAccName.Text = ""
          Exit Sub
       End If
    Else
       If KeyCode = vbKeyReturn Then
          SendKeys "{tab}"
       End If
    End If
    Exit Sub
End Sub
Private Sub txtFindAccCode_LostFocus()
    With txtFindAccCode
         .Text = Trim(.Text)
         If Len(.Text) < 1 Then
            txtFindAccName.Text = ""
         End If
    End With
End Sub
Private Sub dtpFDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub
Private Sub dtpTDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then cmdFind.SetFocus
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
            .Cell(flexcpForeColor, 0, 0, 0, .Cols - 1) = vbBlack
            .Cell(flexcpForeColor, 0, .MouseCol, 0, .MouseCol) = vbRed
            strData = Trim(.Cell(flexcpData, .Row, 0))
            Select Case .MouseCol
                   Case 1 '�ۼ�����
                        .ColSel = 2
                        .ColSort(0) = flexSortNone
                        .ColSort(1) = IIf(P_intButton = 1, flexSortGenericAscending, flexSortGenericDescending)
                        .ColSort(2) = IIf(P_intButton = 1, flexSortGenericAscending, flexSortGenericDescending)
                        .Sort = flexSortUseColSort
                   Case 2 '�ۼ��ð�
                        .ColSel = 3
                        .ColSort(0) = flexSortNone
                        .ColSort(1) = IIf(P_intButton = 1, flexSortGenericAscending, flexSortGenericDescending)
                        .ColSort(2) = IIf(P_intButton = 1, flexSortGenericAscending, flexSortGenericDescending)
                        .Sort = flexSortUseColSort
                   Case 3 '�����ڵ�
                        .ColSel = 4
                        .ColSort(0) = flexSortNone
                        .ColSort(1) = flexSortNone
                        .ColSort(2) = flexSortNone
                        .ColSort(3) = IIf(P_intButton = 1, flexSortGenericAscending, flexSortGenericDescending)
                        .Sort = flexSortUseColSort
                   Case 4 '������
                        .ColSel = 5
                        .ColSort(0) = flexSortNone
                        .ColSort(1) = IIf(P_intButton = 1, flexSortGenericAscending, flexSortGenericDescending)
                        .ColSort(2) = flexSortNone
                        .ColSort(3) = flexSortNone
                        .ColSort(4) = IIf(P_intButton = 1, flexSortGenericAscending, flexSortGenericDescending)
                        .Sort = flexSortUseColSort
                   Case Else
                        .Sort = IIf(P_intButton = 1, flexSortGenericAscending, flexSortGenericDescending)
            End Select
            If .FindRow(strData, , 0) > 0 Then
               .Row = .FindRow(strData, , 0)
            End If
            If PC_intRowCnt < .Rows Then
               .TopRow = .Row
            End If
         End If
    End With
End Sub
Private Sub vsfg1_EnterCell()
Dim lngR As Long
Dim lngC As Long
    With vsfg1
         If .Row >= .FixedRows Then
            dtpCDate.Value = Format(DTOS(.TextMatrix(.Row, 1)), "0000-00-00") '����
            txtAccCode.Text = .TextMatrix(.Row, 3): txtAccName.Text = .TextMatrix(.Row, 4)  '�����ڵ�, ������
            txtJukyo.Text = .TextMatrix(.Row, 5) '����
            If .ValueMatrix(.Row, 6) > 0 Then
               cboIOGbn.ListIndex = 0
               cboIOGbn.Text = "1. �Ա�"
               txtMoney.Text = Format(.ValueMatrix(.Row, 6), "#,0")
            Else
               cboIOGbn.ListIndex = 1
               cboIOGbn.Text = "2. ���"
               txtMoney.Text = Format(.ValueMatrix(.Row, 7), "#,0")
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
    dtpCDate.SetFocus
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
Dim blnOK         As Boolean
Dim intRetVal     As Integer
Dim strServerTime As String
Dim strTime       As String
Dim intAddMode    As Integer '1.�߰�, Etc.����
    '�Է³��� �˻�
    blnOK = False
    FncCheckTextBox blnOK
    If blnOK = False Then
       Exit Sub
    End If
    If vsfg1.Row < vsfg1.FixedRows Then
       intAddMode = 1
       intRetVal = MsgBox("�Էµ� �ڷḦ �߰��Ͻðڽ��ϱ� ?", vbQuestion + vbYesNo + vbDefaultButton1, "�ڷ� �߰�")
    Else
       intRetVal = MsgBox("������ �ڷḦ �����Ͻðڽ��ϱ� ?", vbQuestion + vbYesNo + vbDefaultButton1, "�ڷ� ����")
    End If
    If intRetVal = vbNo Then
       vsfg1.SetFocus
       Exit Sub
    End If
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
         PB_adoCnnSQL.BeginTrans
         If intAddMode = 1 Then 'ȸ����ǥ���� �߰�
            strSQL = "INSERT INTO ȸ����ǥ����(������ڵ�, �ۼ�����, �ۼ��ð�, �����ڵ�, " _
                                            & "���ⱸ��, �Աݱݾ�, ��ݱݾ�, ����, ��뱸��, �ۼ����ڵ�, " _
                                            & "��������,������ڵ�) Values( " _
                    & "'" & PB_regUserinfoU.UserBranchCode & "','" & DTOS(dtpCDate.Value) & "', " _
                    & "'" & strServerTime & "', '" & txtAccCode.Text & "', " _
                    & "" & cboIOGbn.ListIndex + 1 & ", " & IIf(cboIOGbn.ListIndex = 0, Vals(txtMoney.Text), 0) & ", " _
                    & "" & IIf(cboIOGbn.ListIndex = 1, Vals(txtMoney.Text), 0) & ", '" & txtJukyo.Text & "', " _
                    & "0, '" & PB_regUserinfoU.UserCode & "', " _
                    & "'" & PB_regUserinfoU.UserServerDate & "', '" & PB_regUserinfoU.UserCode & "' )"
            On Error GoTo ERROR_TABLE_INSERT
            .AddItem .Rows
            .TextMatrix(.Rows - 1, 0) = PB_regUserinfoU.UserBranchCode & DTOS(dtpCDate.Value) & strServerTime & txtAccCode.Text
            .Cell(flexcpData, .Rows - 1, 0, .Rows - 1, 0) = Trim(.TextMatrix(.Rows - 1, 0))
            .TextMatrix(.Rows - 1, 1) = Format(DTOS(dtpCDate.Value), "0000-00-00")            '�ۼ�����
            .TextMatrix(.Rows - 1, 2) = Format(Mid(strServerTime, 1, 6), "00:00:00")          '�ۼ��ð�
            .TextMatrix(.Rows - 1, 3) = txtAccCode.Text                                       '�����ڵ�
            .TextMatrix(.Rows - 1, 4) = txtAccName.Text                                       '������
            .TextMatrix(.Rows - 1, 5) = txtJukyo.Text                                         '����
            .TextMatrix(.Rows - 1, 6) = IIf(cboIOGbn.ListIndex = 0, Format(Vals(txtMoney.Text), "#,0"), 0) '�Աݱݾ�
            .TextMatrix(.Rows - 1, 7) = IIf(cboIOGbn.ListIndex = 1, Format(Vals(txtMoney.Text), "#,0"), 0) '��ݱݾ�
            .TextMatrix(.Rows - 1, 8) = PB_regUserinfoU.UserCode                              '�ۼ����ڵ�
            .TextMatrix(.Rows - 1, 9) = PB_regUserinfoU.UserName                              '�ۼ��ڸ�
            .TextMatrix(.Rows - 1, 10) = "��    ��"                                           '��뱸��
            .TextMatrix(.Rows - 1, 11) = Format(PB_regUserinfoU.UserServerDate, "0000-00-00") '��������
            .TextMatrix(.Rows - 1, 12) = PB_regUserinfoU.UserCode                             '������ڵ�
            .TextMatrix(.Rows - 1, 13) = PB_regUserinfoU.UserName                             '����ڸ�
            .TextMatrix(.Rows - 1, 14) = strServerTime                                        '�ۼ��ð�
            If (dtpCDate.Value >= dtpFDate.Value And dtpCDate.Value <= dtpTDate.Value) Then
               If cboIOGbn.ListIndex = 0 Then
                  lblTotIn.Caption = Format(Vals(lblTotIn.Caption) + Vals(txtMoney.Text), "#,0")
               Else
                  lblTotOut.Caption = Format(Vals(lblTotOut.Caption) + Vals(txtMoney.Text), "#,0")
               End If
            End If
            If .Rows > PC_intRowCnt Then
               .ScrollBars = flexScrollBarBoth
               .TopRow = .Rows - 1
            End If
            .Row = .Rows - 1          '�ڵ����� vsfg1_EnterCell Event �߻�
         Else                                         'ȸ����ǥ���� ����
            strSQL = "UPDATE ȸ����ǥ���� SET " _
                          & "�ۼ����� = '" & DTOS(dtpCDate.Value) & "', " _
                          & "�����ڵ� = '" & txtAccCode.Text & "', " _
                          & "���� = '" & txtJukyo.Text & "', " _
                          & "�Աݱݾ� = " & IIf(cboIOGbn.ListIndex = 0, Vals(txtMoney.Text), 0) & ", " _
                          & "��ݱݾ� = " & IIf(cboIOGbn.ListIndex = 1, Vals(txtMoney.Text), 0) & ", " _
                          & "�������� = '" & PB_regUserinfoU.UserServerDate & "', " _
                          & "������ڵ� = '" & PB_regUserinfoU.UserCode & "' " _
                    & "WHERE ������ڵ� = '" & PB_regUserinfoU.UserBranchCode & "' " _
                      & "AND �ۼ����� = '" & DTOS(.TextMatrix(.Row, 1)) & "' " _
                      & "AND �ۼ��ð� = '" & .TextMatrix(.Row, 14) & "' " _
                      & "AND �����ڵ� = '" & .TextMatrix(.Row, 3) & "' "
            On Error GoTo ERROR_TABLE_UPDATE
            '�ش�Ⱓ�� ���ԵǸ� ��������ݾ� �հ迡�� ����.
            If (DTOS(.TextMatrix(.Row, 1)) >= DTOS(dtpFDate.Value) And DTOS(.TextMatrix(.Row, 1)) <= DTOS(dtpTDate.Value)) Then
               lblTotIn.Caption = Format(Vals(lblTotIn.Caption) - .ValueMatrix(.Row, 6), "#,0")
               lblTotOut.Caption = Format(Vals(lblTotOut.Caption) - .ValueMatrix(.Row, 7), "#,0")
            End If
            .TextMatrix(.Row, 0) = PB_regUserinfoU.UserBranchCode & DTOS(dtpCDate.Value) & strServerTime & txtAccCode.Text
            .TextMatrix(.Row, 1) = Format(DTOS(dtpCDate.Value), "0000-00-00")            '�ۼ�����
            .TextMatrix(.Row, 2) = Format(Mid(strServerTime, 1, 6), "00:00:00")          '�ۼ��ð�
            .TextMatrix(.Row, 3) = txtAccCode.Text                                       '�����ڵ�
            .TextMatrix(.Row, 4) = txtAccName.Text                                       '������
            .TextMatrix(.Row, 5) = txtJukyo.Text                                         '����
            If (dtpCDate.Value >= dtpFDate.Value And dtpCDate.Value <= dtpTDate.Value) Then
               If cboIOGbn.ListIndex = 0 Then
                  lblTotIn.Caption = Format(Vals(lblTotIn.Caption) + Vals(txtMoney.Text), "#,0")
               Else
                  lblTotOut.Caption = Format(Vals(lblTotOut.Caption) + Vals(txtMoney.Text), "#,0")
               End If
            End If
            .TextMatrix(.Row, 6) = 0: .TextMatrix(.Row, 7) = 0
            .TextMatrix(.Row, 6) = IIf(cboIOGbn.ListIndex = 0, Format(Vals(txtMoney.Text), "#,0"), 0) '�Աݱݾ�
            .TextMatrix(.Row, 7) = IIf(cboIOGbn.ListIndex = 1, Format(Vals(txtMoney.Text), "#,0"), 0) '��ݱݾ�
            .TextMatrix(.Row, 10) = "��    ��"                                           '��뱸��
            .TextMatrix(.Row, 11) = Format(PB_regUserinfoU.UserServerDate, "0000-00-00") '��������
            .TextMatrix(.Row, 12) = PB_regUserinfoU.UserCode                             '������ڵ�
            .TextMatrix(.Row, 13) = PB_regUserinfoU.UserName                             '����ڸ�
         End If
         PB_adoCnnSQL.Execute strSQL
         PB_adoCnnSQL.CommitTrans
         Screen.MousePointer = vbDefault
    End With
    cmdSave.Enabled = True
    If intAddMode = 1 Then
       cmdClear_Click '�߰����� �ٷΰ���
    Else
       vsfg1.SetFocus '�׸���� �ٷΰ���
    End If
    Exit Sub
ERROR_FORM_ACTIVATE_CONNECTION_SERVER:
    MsgBox Err.Number & Err.Description & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "(�������� ���� ����)"
    Unload Me
    Exit Sub
ERROR_TABLE_SELECT:
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "ȸ����ǥ���� �б� ����"
    Unload Me
    Exit Sub
ERROR_TABLE_INSERT:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "ȸ����ǥ���� �߰� ����"
    Unload Me
    Exit Sub
ERROR_TABLE_UPDATE:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "ȸ����ǥ���� ���� ����"
    Unload Me
    Exit Sub
End Sub

'+-----------+
'/// ���� ///
'+-----------+
Private Sub cmdDelete_Click()
Dim strSQL        As String
Dim intRetVal     As Integer
Dim lngCnt        As Long
Dim strServerTime As String
Dim strTime       As String
    With vsfg1
         If .Row >= .FixedRows Then
            intRetVal = MsgBox("��ϵ� �ڷḦ �����Ͻðڽ��ϱ� ?", vbCritical + vbYesNo + vbDefaultButton2, "�ڷ� ����")
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
                  MsgBox "XXX(" & Format(lngCnt, "#,#") & "��)�� �����Ƿ� ������ �� �����ϴ�.", vbCritical, "ȸ����ǥ���� ���� �Ұ�"
                  cmdDelete.Enabled = True
                  Screen.MousePointer = vbDefault
                  Exit Sub
               End If
               '�����ð�
               P_adoRec.CursorLocation = adUseClient
               strSQL = "SELECT RIGHT(CONVERT(VARCHAR(23),GETDATE(), 121),12) AS �����ð� "
               On Error GoTo ERROR_FORM_ACTIVATE_CONNECTION_SERVER
               P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
               strServerTime = Mid(P_adoRec("�����ð�"), 1, 2) + Mid(P_adoRec("�����ð�"), 4, 2) + Mid(P_adoRec("�����ð�"), 7, 2) _
                             + Mid(P_adoRec("�����ð�"), 10)
               P_adoRec.Close
               PB_adoCnnSQL.BeginTrans
               P_adoRec.CursorLocation = adUseClient
               strSQL = "UPDATE ȸ����ǥ���� SET " _
                             & "��뱸�� = 9, " _
                             & "�������� = '" & Mid(strServerTime, 1, 6) & "', " _
                             & "������ڵ� = '" & PB_regUserinfoU.UserCode & "' " _
                       & "WHERE ������ڵ� = '" & PB_regUserinfoU.UserBranchCode & "' " _
                         & "AND �ۼ����� = '" & DTOS(.TextMatrix(.Row, 1)) & "' " _
                         & "AND �ۼ��ð� = '" & .TextMatrix(.Row, 14) & "' " _
                         & "AND �����ڵ� = '" & .TextMatrix(.Row, 3) & "' "
               On Error GoTo ERROR_TABLE_DELETE
               PB_adoCnnSQL.Execute strSQL
               PB_adoCnnSQL.CommitTrans
               '���հ�����
               If (DTOS(.TextMatrix(.Row, 1)) >= DTOS(dtpFDate.Value) And DTOS(.TextMatrix(.Row, 1)) <= DTOS(dtpTDate.Value)) Then
                  If .ValueMatrix(.Row, 6) > 0 Then lblTotIn.Caption = Format(Vals(lblTotIn.Caption) - .ValueMatrix(.Row, 6), "#,0")
                  If .ValueMatrix(.Row, 7) > 0 Then lblTotOut.Caption = Format(Vals(lblTotOut.Caption) - .ValueMatrix(.Row, 7), "#,0")
               End If
               .RemoveItem .Row
               If .Rows <= PC_intRowCnt Then
                  '.ScrollBars = flexScrollBarHorizontal
               End If
               cmdDelete.Enabled = True
               Screen.MousePointer = vbDefault
               If .Rows = 1 Then
                  SubClearText
                  .Row = 0
                  Exit Sub
               End If
               vsfg1_EnterCell
            End If
            .SetFocus
         End If
    End With
    Screen.MousePointer = vbDefault
    Exit Sub
ERROR_FORM_ACTIVATE_CONNECTION_SERVER:
    MsgBox Err.Number & Err.Description & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "(�������� ���� ����)"
    Unload Me
    Exit Sub
ERROR_TABLE_DELETE:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "ȸ����ǥ���� ���� ����"
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
    Set frm�����ⳳ���� = Nothing
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
    'cboCode(0).Enabled = False               '�����ڵ� FLASE
    With vsfg1              'Rows 1, Cols 15, RowHeightMax(Min) 300
         .AllowBigSelection = False
         .AllowSelection = False
         .AllowUserResizing = flexResizeColumns
         .BackColorBkg = &H8000000F
         .BackColorSel = &H8000&
         .ExtendLastCol = True
         .ScrollBars = flexScrollBarVertical
         .ScrollTrack = True
         .SelectionMode = flexSelectionByRow
         .FixedRows = 1
         .FixedCols = 0
         .Rows = 1             'Subvsfg1_Fill����ÿ� ����
         .Cols = 15
         .RowHeightMax = 300
         .RowHeightMin = 300
         .ColWidth(0) = 1000   'KEY(������ڵ�+�ۼ�����+�ۼ��ð�+�����ڵ�)
         .ColWidth(1) = 1100   '�ۼ�����(0000-00-00)
         .ColWidth(2) = 1000   '�ۼ��ð�(00:00:00)
         .ColWidth(3) = 500    '�ڵ�
         .ColWidth(4) = 2800   '������
         .ColWidth(5) = 5500   '����
         .ColWidth(6) = 1500   '�Աݱݾ�
         .ColWidth(7) = 1500   '��ݱݾ�
         .ColWidth(8) = 1      '�ۼ����ڵ�
         .ColWidth(9) = 1000   '�ۼ��ڸ�
         .ColWidth(10) = 1     '��뱸��
         .ColWidth(11) = 1200  '��������
         .ColWidth(12) = 1     '������ڵ�
         .ColWidth(13) = 1000  '����ڸ�
         .ColWidth(14) = 1000  '�簣(�и��� ����)
         
         .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = False
         .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
         .TextMatrix(0, 0) = "KEY"         'H
         .TextMatrix(0, 1) = "����"
         .TextMatrix(0, 2) = "�ð�"
         .TextMatrix(0, 3) = "�ڵ�"
         .TextMatrix(0, 4) = "������"
         .TextMatrix(0, 5) = "����"
         .TextMatrix(0, 6) = "�Աݱݾ�"
         .TextMatrix(0, 7) = "��ݱݾ�"
         .TextMatrix(0, 8) = "�ۼ����ڵ�"  'H
         .TextMatrix(0, 9) = "�ۼ��ڸ�"
         .TextMatrix(0, 10) = "��뱸��"   'H
         .TextMatrix(0, 11) = "��������"   'H
         .TextMatrix(0, 12) = "������ڵ�" 'H
         .TextMatrix(0, 13) = "����ڸ�"   'H
         .TextMatrix(0, 14) = "�ð�"       'H
         .ColHidden(0) = True: .ColHidden(8) = True: .ColHidden(10) = True
         .ColHidden(11) = True: .ColHidden(12) = True: .ColHidden(13) = True: .ColHidden(14) = True
         For lngC = 0 To .Cols - 1
             Select Case lngC
                    Case 0, 4, 5
                         .ColAlignment(lngC) = flexAlignLeftCenter
                    Case 1, 2, 3, 8, 9, 10, 11, 12, 13, 14
                         .ColAlignment(lngC) = flexAlignCenterCenter
                    Case Else
                         .ColAlignment(lngC) = flexAlignRightCenter
             End Select
         Next lngC
         For lngC = 0 To .Cols - 1
             Select Case lngC
                    Case 6, 7
                         .ColFormat(lngC) = "#,#"
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
    lblTotIn.Caption = "0": lblTotOut.Caption = "0"
    If dtpFDate.Value > dtpTDate.Value Then
       dtpFDate.SetFocus
       Exit Sub
    End If
    vsfg1.Rows = 1
    With vsfg1
         '�˻����� �����ڵ�
         txtFindAccCode.Text = Trim(txtFindAccCode.Text)
         Select Case txtFindAccCode.Text
                Case ""         '�����ڵ� ��ü
                     strWhere = strWhere
                Case Else        '�����ڵ� ��ü �ƴϸ�
                     strWhere = IIf(Trim(strWhere) = "", "WHERE ", "AND ") _
                              & "T1.�����ڵ� = '" & txtFindAccCode.Text & "' "
         End Select
         strWhere = strWhere + IIf(Trim(strWhere) = "", "WHERE ", "AND ") _
                  & "T1.������ڵ� = '" & PB_regUserinfoU.UserBranchCode & "' "
         strWhere = strWhere + IIf(Trim(strWhere) = "", "WHERE ", "AND ") _
                  & "(T1.�ۼ����� BETWEEN '" & DTOS(dtpFDate.Value) & "' AND '" & DTOS(dtpTDate.Value) & "') " _
                                     & "AND T1.��뱸�� = 0 "
         strOrderBy = "ORDER BY T1.������ڵ�, T1.�ۼ�����, T1.�ۼ��ð� "
    End With
    P_adoRec.CursorLocation = adUseClient
    strSQL = "SELECT T1.������ڵ� AS ������ڵ�, T1.�ۼ����� AS �ۼ�����, " _
                  & "T1.�ۼ��ð� AS �ۼ��ð�, T1.�����ڵ� AS �����ڵ�, " _
                  & "ISNULL(T2.������, '') AS ������, ISNULL(T1.����, '') AS ����, " _
                  & "ISNULL(T1.�Աݱݾ�, 0) AS �Աݱݾ�, ISNULL(T1.��ݱݾ�, 0) AS ��ݱݾ�, " _
                  & "T1.�ۼ����ڵ� AS �ۼ����ڵ�, " _
                  & "ISNULL(T3.����ڸ�, '') AS �ۼ��ڸ�, T1.��뱸�� AS ��뱸��, " _
                  & "T1.�������� AS ��������, T1.������ڵ� AS ������ڵ�, " _
                  & "T4.����ڸ� AS ����ڸ� " _
             & "FROM ȸ����ǥ���� T1 " _
            & "INNER JOIN �������� T2 " _
                    & "ON T2.�����ڵ� = T1.�����ڵ� " _
             & "LEFT JOIN ����� T3 " _
                    & "ON T3.������ڵ� = T1.�ۼ����ڵ� " _
             & "LEFT JOIN ����� T4 " _
                    & "ON T4.������ڵ� = T1.������ڵ� "
    strSQL = strSQL _
           & "" & strWhere & " " _
           & "" & strOrderBy & " "
    On Error GoTo ERROR_TABLE_SELECT
    P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
    vsfg1.Rows = P_adoRec.RecordCount + 1
    If P_adoRec.RecordCount = 0 Then
       P_adoRec.Close
       cmdExit.SetFocus
    Else
       With vsfg1
            .Cell(flexcpForeColor, 0, 0, 0, .Cols - 1) = vbBlack
            If .Rows <= PC_intRowCnt Then
               '.ScrollBars = flexScrollBarHorizontal
            Else
               '.ScrollBars = flexScrollBarBoth
            End If
            Do Until P_adoRec.EOF
               lngR = lngR + 1
               .TextMatrix(lngR, 0) = P_adoRec("������ڵ�") & P_adoRec("�ۼ�����") _
                                    & P_adoRec("�ۼ��ð�") & P_adoRec("�����ڵ�")
               .Cell(flexcpData, lngR, 0, lngR, 0) = Trim(.TextMatrix(lngR, 0)) 'FindRow ����� ����
               .TextMatrix(lngR, 1) = Format(P_adoRec("�ۼ�����"), "0000-00-00")
               .TextMatrix(lngR, 2) = Format(Mid(P_adoRec("�ۼ��ð�"), 1, 6), "00:00:00")
               .TextMatrix(lngR, 3) = P_adoRec("�����ڵ�")
               .TextMatrix(lngR, 4) = P_adoRec("������")
               .TextMatrix(lngR, 5) = P_adoRec("����")
               .TextMatrix(lngR, 6) = P_adoRec("�Աݱݾ�")
               .TextMatrix(lngR, 7) = P_adoRec("��ݱݾ�")
               .TextMatrix(lngR, 8) = P_adoRec("�ۼ����ڵ�")
               .TextMatrix(lngR, 9) = P_adoRec("�ۼ��ڸ�")
               If P_adoRec("��뱸��") = 0 Then
                  .TextMatrix(lngR, 10) = "��    ��"
               ElseIf _
                  P_adoRec("��뱸��") = 9 Then
                  .TextMatrix(lngR, 10) = "��    ��"
               Else
                  .TextMatrix(lngR, 10) = "�ڵ����"
               End If
               .TextMatrix(lngR, 11) = Format(P_adoRec("��������"), "0000-00-00")
               .TextMatrix(lngR, 12) = P_adoRec("������ڵ�")
               .TextMatrix(lngR, 13) = P_adoRec("����ڸ�")
               .TextMatrix(lngR, 14) = P_adoRec("�ۼ��ð�")
               lblTotIn.Caption = Format(Vals(lblTotIn.Caption) + .ValueMatrix(lngR, 6), "#,0")
               lblTotOut.Caption = Format(Vals(lblTotOut.Caption) + .ValueMatrix(lngR, 7), "#,0")
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
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "ȸ����ǥ���� �б� ����"
    Unload Me
    Exit Sub
End Sub

Private Sub SubOther_FILL()
Dim strSQL        As String
Dim lngI          As Long
Dim lngJ          As Long
Dim intIndex      As Integer
    P_adoRec.CursorLocation = adUseClient
    With cboIOGbn
         .AddItem "1. �Ա�"
         .AddItem "2. ���"
         .ListIndex = 0
    End With
    Exit Sub
ERROR_TABLE_SELECT:
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "�����ڵ� �б� ����"
    Unload Me
    Exit Sub
End Sub

'+-------------------------+
'/// Clear text1(index) ///
'+-------------------------+
Private Sub SubClearText()
Dim lngC As Long
    txtAccCode.Text = "": txtAccName.Text = "": txtJukyo.Text = "": txtMoney.Text = ""
End Sub

'+-------------------------+
'/// Check text1(index) ///
'+-------------------------+
Private Function FncCheckTextBox(blnOK As Boolean)
    txtAccCode.Text = Trim(txtAccCode.Text) '��������
    If Not (Vals(txtAccCode.Text) > 0) Then
       txtAccCode.SetFocus
       Exit Function
    End If
    txtMoney.Text = Trim(txtMoney.Text) '����
    If Not (LenH(txtJukyo.Text) <= 60) Then
       txtJukyo.SetFocus
       Exit Function
    End If
    txtMoney.Text = Trim(txtMoney.Text) '����ݾ�
    If Not (Vals(txtMoney.Text) <> 0) Then
       txtMoney.SetFocus
       Exit Function
    End If
    blnOK = True
End Function

'+---------------------------+
'/// ũ����Ż ������ ��� ///
'+---------------------------+
Private Sub cmdPrint_Click()
    '
End Sub



VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm���Һ� 
   BorderStyle     =   0  '����
   Caption         =   "���Һ�"
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
      TabIndex        =   11
      Top             =   0
      Width           =   15195
      Begin VB.OptionButton optPrtChk0 
         Caption         =   "�ڵ��"
         Height          =   255
         Left            =   6960
         TabIndex        =   27
         Top             =   150
         Width           =   975
      End
      Begin VB.OptionButton optPrtChk1 
         Caption         =   "�̸���"
         Height          =   255
         Left            =   6960
         TabIndex        =   26
         Top             =   390
         Value           =   -1  'True
         Width           =   975
      End
      Begin Crystal.CrystalReport CrystalReport1 
         Left            =   4920
         Top             =   200
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin VB.CommandButton cmdPrint 
         Height          =   375
         Left            =   7980
         Picture         =   "���Һ�.frx":0000
         Style           =   1  '�׷���
         TabIndex        =   18
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdClear 
         Height          =   375
         Left            =   9120
         Picture         =   "���Һ�.frx":0963
         Style           =   1  '�׷���
         TabIndex        =   16
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdExit 
         Height          =   390
         Left            =   13635
         Picture         =   "���Һ�.frx":1308
         Style           =   1  '�׷���
         TabIndex        =   15
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         Height          =   390
         Left            =   12510
         Picture         =   "���Һ�.frx":1C56
         Style           =   1  '�׷���
         TabIndex        =   14
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdSave 
         Height          =   390
         Left            =   11385
         Picture         =   "���Һ�.frx":25DA
         Style           =   1  '�׷���
         TabIndex        =   13
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdFind 
         Height          =   390
         Left            =   10260
         Picture         =   "���Һ�.frx":2E61
         Style           =   1  '�׷���
         TabIndex        =   6
         Top             =   195
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         BackColor       =   &H00008000&
         BorderStyle     =   1  '���� ����
         Caption         =   "�� �� ��"
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
         TabIndex        =   12
         Top             =   180
         Width           =   4650
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid vsfg1 
      Height          =   8471
      Left            =   60
      TabIndex        =   7
      Top             =   1619
      Width           =   15195
      _cx             =   26802
      _cy             =   14937
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
      Height          =   970
      Left            =   60
      TabIndex        =   8
      Top             =   630
      Width           =   15195
      Begin VB.CheckBox chkTotal 
         Caption         =   "��ü ǰ��"
         Height          =   255
         Left            =   5265
         TabIndex        =   25
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  '���
         Enabled         =   0   'False
         Height          =   285
         IMEMode         =   10  '�ѱ� 
         Index           =   3
         Left            =   12480
         MaxLength       =   50
         TabIndex        =   3
         Top             =   585
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  '���
         Enabled         =   0   'False
         Height          =   285
         IMEMode         =   10  '�ѱ� 
         Index           =   2
         Left            =   8040
         MaxLength       =   50
         TabIndex        =   2
         Top             =   585
         Width           =   3375
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  '���
         Enabled         =   0   'False
         Height          =   285
         IMEMode         =   10  '�ѱ� 
         Index           =   1
         Left            =   2475
         MaxLength       =   50
         TabIndex        =   1
         Top             =   585
         Width           =   3855
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  '���
         Height          =   285
         IMEMode         =   8  '����
         Index           =   0
         Left            =   2475
         TabIndex        =   0
         Top             =   225
         Width           =   1695
      End
      Begin MSComCtl2.DTPicker dtpF_Date 
         Height          =   270
         Left            =   10440
         TabIndex        =   4
         Top             =   240
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   476
         _Version        =   393216
         Format          =   56557569
         UpDown          =   -1  'True
         CurrentDate     =   37763
      End
      Begin MSComCtl2.DTPicker dtpT_Date 
         Height          =   270
         Left            =   12480
         TabIndex        =   5
         Top             =   240
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   476
         _Version        =   393216
         Format          =   56557569
         UpDown          =   -1  'True
         CurrentDate     =   37763
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "����"
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   3
         Left            =   11400
         TabIndex        =   24
         Top             =   645
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "�԰�"
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   2
         Left            =   6960
         TabIndex        =   23
         Top             =   645
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "����"
         Height          =   240
         Index           =   12
         Left            =   13800
         TabIndex        =   22
         Top             =   285
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "����"
         Height          =   240
         Index           =   11
         Left            =   11760
         TabIndex        =   21
         Top             =   285
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "��������"
         Height          =   240
         Index           =   10
         Left            =   9360
         TabIndex        =   20
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
         Left            =   120
         TabIndex        =   19
         Top             =   285
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "[Home]"
         Height          =   240
         Index           =   15
         Left            =   4320
         TabIndex        =   17
         Top             =   285
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "ǰ��"
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   1
         Left            =   1275
         TabIndex        =   10
         Top             =   645
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "ǰ���ڵ�"
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   0
         Left            =   1275
         TabIndex        =   9
         Top             =   285
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frm���Һ�"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'+-------------------------------------------------------------------------------------------------------+
' ���α׷� �� �� : ���Һ�
' ���� Control : VideoSoft VSFlexGrid 7.0(OLEDB) = vsflex7.ocx
' ������ Table   :
' ��  ��  ��  �� :
'+-------------------------------------------------------------------------------------------------------+
Option Explicit
Private P_blnActived       As Boolean
Private P_adoRec           As New ADODB.Recordset
Private P_intButton        As Integer
Private Const PC_intRowCnt As Integer = 25  '�׸��� �� ������ �� ���(FixedRows ����)

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
                   cmdPrint.Enabled = True: cmdClear.Enabled = False: cmdFind.Enabled = True
                   cmdSave.Enabled = False: cmdDelete.Enabled = False: cmdExit.Enabled = True
              Case Is <= 40 '�߰�, ����, �μ�, ��ȸ
                   cmdPrint.Enabled = True: cmdClear.Enabled = False: cmdFind.Enabled = True
                   cmdSave.Enabled = False: cmdDelete.Enabled = False: cmdExit.Enabled = True
              Case Is <= 50 '����, �߰�, ����, ��ȸ, �μ�
                   cmdPrint.Enabled = True: cmdClear.Enabled = False: cmdFind.Enabled = True
                   cmdSave.Enabled = False: cmdDelete.Enabled = False: cmdExit.Enabled = True
              Case Is <= 99 '����, �߰�, ����, ��ȸ, �μ�
                   cmdPrint.Enabled = True: cmdClear.Enabled = False: cmdFind.Enabled = True
                   cmdSave.Enabled = False: cmdDelete.Enabled = False: cmdExit.Enabled = True
              Case Else
                   cmdClear.Enabled = False: cmdFind.Enabled = False: cmdSave.Enabled = False
                   cmdDelete.Enabled = False: cmdExit.Enabled = True
       End Select
       dtpF_Date.Value = Format(Mid(PB_regUserinfoU.UserClientDate, 1, 6) & "01", "0000-00-00")
       dtpT_Date.Value = Format(PB_regUserinfoU.UserClientDate, "0000-00-00")
       'dtpT_Date.Value = DateAdd("d", -1, DateAdd("m", 1, dtpF_Date.Value))
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
Private Sub dtpF_Date_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub
Private Sub dtpT_Date_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then cmdFind.SetFocus
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
    If (Index = 0 And (KeyCode = vbKeyHome Or KeyCode = vbKeyReturn)) Then '����˻�
       Text1(0).Text = Trim(Text1(0).Text)
       PB_strCallFormName = "frm���Һ�"
       PB_strMaterialsCode = Trim(Text1(0).Text)
       PB_strMaterialsName = "" 'Trim(Text1(1).Text)
       frm����˻�.Show vbModal
       If (Len(PB_strMaterialsCode) + Len(PB_strMaterialsName)) = 0 Then '�˻����� ���(ESC)
       Else
          Text1(0).Text = PB_strMaterialsCode
          Text1(1).Text = PB_strMaterialsName
       End If
       If PB_strMaterialsCode <> "" Then
          SendKeys "{tab}"
       End If
       PB_strCallFormName = "": PB_strMaterialsCode = "": PB_strMaterialsName = ""
    ElseIf _
       KeyCode = vbKeyDelete Then
       If Len(Text1(0).Text) = 0 Then
          Text1(1).Text = "": Text1(2).Text = "": Text1(3).Text = ""
          Exit Sub
       End If
    Else
       If KeyCode = vbKeyReturn Then
          Select Case Index
              Case Text1.UBound
                   If cmdFind.Enabled = True Then
                      cmdFind.SetFocus
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
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "���� �б� ����"
    Unload Me
    Exit Sub
End Sub

Private Sub Text1_LostFocus(Index As Integer)
Dim strSQL As String
    With Text1(Index)
         Select Case Index
                Case 0 '����˻�
                     .Text = Trim(.Text)
                     If Len(.Text) < 1 Then
                        Text1(Index).Text = ""
                        Text1(Index + 1).Text = ""
                        Text1(Index + 2).Text = ""
                        Text1(Index + 3).Text = ""
                        Exit Sub
                     End If
                Case Else
                     .Text = Trim(.Text)
         End Select
    End With
    Exit Sub
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
         End If
    End With
End Sub
Private Sub vsfg1_EnterCell()
Dim lngR As Long
Dim lngC As Long
    With vsfg1
         'Text1(0).Enabled = False: Text1(2).Enabled = False
         If .Row >= .FixedRows Then
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
End Sub

'+-----------+
'/// �߰� ///
'+-----------+
Private Sub cmdClear_Click()
Dim strSQL As String
    '
End Sub
'+-----------+
'/// ��ȸ ///
'+-----------+
Private Sub cmdFind_Click()
    Screen.MousePointer = vbHourglass
    cmdFind.Enabled = False
    'SubClearText
    Subvsfg1_FILL
    cmdFind.Enabled = True
    Screen.MousePointer = vbDefault
End Sub
'+-----------+
'/// ���� ///
'+-----------+
Private Sub cmdSave_Click()
    '
End Sub

'+-----------+
'/// ���� ///
'+-----------+
Private Sub cmdDelete_Click()
    '
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
    Set frm���Һ� = Nothing
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
    With vsfg1              'Rows 2, Cols 21, RowHeightMax(Min) 300
         .AllowBigSelection = False
         .AllowSelection = False
         .AllowUserResizing = flexResizeColumns
         .BackColorBkg = &H8000000F
         .BackColorSel = &H8000&
         .ExtendLastCol = True
         .ScrollBars = flexScrollBarHorizontal
         .ScrollTrack = True
         .SelectionMode = flexSelectionByRow
         .FixedRows = 2
         .FixedCols = 6
         .Rows = 2             'Subvsfg1_Fill����ÿ� ����
         .Cols = 13
         .RowHeightMax = 300
         .RowHeightMin = 300
         .ColWidth(0) = 1000   '������ڵ�
         .ColWidth(1) = 1200   '�����ڵ�(�з��ڵ�+�����ڵ�)
         .ColWidth(2) = 3000   '�����
         .ColWidth(3) = 1200   '�԰�
         .ColWidth(4) = 800    '����
         .ColWidth(5) = 1200   '����
         .ColWidth(6) = 3000   '�԰�����
         .ColWidth(7) = 1600   '�԰����
         .ColWidth(8) = 3000   '�������
         .ColWidth(9) = 1600   '������
         .ColWidth(10) = 1600  '������          'H
         .ColWidth(11) = 1600  '��������      'H
         .ColWidth(12) = 2000  '���
         
         .Cell(flexcpFontBold, 0, 0, 1, .Cols - 1) = False
         .Cell(flexcpAlignment, 0, 0, 1, .Cols - 1) = flexAlignCenterCenter
         .TextMatrix(0, 0) = "������ڵ�"  'H
         .TextMatrix(0, 1) = "ǰ���ڵ�"    'H
         .TextMatrix(0, 2) = "ǰ��"        'H
         .TextMatrix(0, 3) = "�԰�"        'H
         .TextMatrix(0, 4) = "����"        'H
         .TextMatrix(0, 5) = "��¥"
         .TextMatrix(0, 6) = "����"
         .TextMatrix(0, 7) = "����"
         .TextMatrix(0, 8) = "����"
         .TextMatrix(0, 9) = "����"
         .TextMatrix(0, 10) = "���"       'H
         .TextMatrix(0, 11) = "���"       'H
         .TextMatrix(0, 12) = "���"
         .TextMatrix(1, 0) = "������ڵ�"  'H
         .TextMatrix(1, 1) = "ǰ���ڵ�"    'H
         .TextMatrix(1, 2) = "ǰ��"        'H
         .TextMatrix(1, 3) = "�԰�"        'H
         .TextMatrix(1, 4) = "����"        'H
         .TextMatrix(1, 5) = "��¥"
         .TextMatrix(1, 6) = "����"
         .TextMatrix(1, 7) = "����"
         .TextMatrix(1, 8) = "����"
         .TextMatrix(1, 9) = "����"
         .TextMatrix(1, 10) = "���"       'H
         .TextMatrix(1, 11) = "���"       'H
         .TextMatrix(1, 12) = "���"
         .ColHidden(0) = True: .ColHidden(1) = True: .ColHidden(2) = True: .ColHidden(3) = True
         .ColHidden(4) = True: .ColHidden(10) = True: .ColHidden(11) = True
         For lngC = 0 To .Cols - 1
             Select Case lngC
                    Case 2, 3, 4, 6, 8, 12
                         .ColAlignment(lngC) = flexAlignLeftCenter
                    Case 0, 1, 5
                         .ColAlignment(lngC) = flexAlignCenterCenter
                    Case Else
                         .ColAlignment(lngC) = flexAlignRightCenter
             End Select
         Next lngC
         For lngC = 0 To .Cols - 1
             Select Case lngC
                    Case 7, 9, 10, 11
                         .ColFormat(lngC) = "#,#.00"
             End Select
         Next lngC
         .MergeCells = flexMergeFixedOnly
         .MergeRow(0) = True: .MergeRow(1) = True
         For lngC = 0 To .Cols - 1
             .MergeCol(lngC) = True
         Next lngC
    End With
End Sub

'+---------------------------------+
'/// VsFlexGrid(vsfg1) ä���///
'+---------------------------------+
Private Sub Subvsfg1_FILL()
Dim strSQL      As String
Dim strJoin     As String
Dim strGroupBy  As String
Dim strHaving   As String
Dim strWhere    As String
Dim strOrderBy  As String
Dim lngR        As Long
Dim lngC        As Long
Dim lngRR       As Long
Dim lngRRR      As Long
Dim curStockAmt As Currency  '����������
Dim StrDate     As String    '�ش�����
Dim curMonIAmt  As Currency  '�ش���԰���
Dim curMonOAmt  As Currency  '�ش�������
Dim curTotIAmt  As Currency  '�ش������԰���
Dim curTotOAmt  As Currency  '�ش����������
    vsfg1.Rows = 2
    If DTOS(dtpF_Date.Value) > DTOS(dtpT_Date.Value) Then
       Exit Sub
    End If
    With vsfg1
         '�˻����� ����з�
         If chkTotal.Value = 0 Then '�Ǻ� ��ȸ
            If Len(Text1(0).Text) > 0 Then
               strWhere = "WHERE T1.�з��ڵ� = '" & Mid(Trim(Text1(0).Text), 1, 2) & "' " _
                          & "AND T1.�����ڵ� = '" & Mid(Trim(Text1(0).Text), 3) & "' "
            Else
               Text1(0).SetFocus
               Exit Sub
            End If
         End If
         strWhere = strWhere + IIf(Trim(strWhere) = "", "WHERE ", "AND ") _
                  & "T1.������ڵ� = '" & PB_regUserinfoU.UserBranchCode & "' "
    End With
    strJoin = "LEFT JOIN ����� T2 ON T2.������ڵ� = T1.������ڵ� " _
            & "LEFT JOIN ����з� T3 " _
                   & "ON T3.�з��ڵ� = T1.�з��ڵ� " _
            & "LEFT JOIN ���� T4 " _
                   & "ON T4.�з��ڵ� = T1.�з��ڵ� AND T4.�����ڵ� = T1.�����ڵ� "
    If optPrtChk0.Value = True Then '�ڵ��
       strOrderBy = "ORDER BY T1.������ڵ�, T1.�з��ڵ�, T1.�����ڵ�, ���� "
    Else
       strOrderBy = "ORDER BY T1.������ڵ�, T4.�����, T4.�԰�, ���� "
    End If
    P_adoRec.CursorLocation = adUseClient
    strSQL = "SELECT T1.������ڵ� AS ������ڵ�, T2.������ AS ������, " _
                  & "T1.�з��ڵ� AS �з��ڵ�, T3.�з��� AS �з���, " _
                  & "T1.�����ڵ� AS �����ڵ�, ISNULL(T4.�����,'ERROR!') AS �����, " _
                  & "ISNULL(T4.�԰�,'') AS �԰�, ISNULL(T4.����,'') AS ����, " _
                  & "(T1.������� + '00') AS ����, '(�� �̿�)' AS ����, " _
                  & "(T1.�԰������) AS �԰����, (T1.��������) AS ������, " _
                  & "(T1.�԰������ - T1.��������) AS ������ " _
             & "FROM ������帶�� T1 " _
             & "" & strJoin & " " _
             & "" & strWhere & " AND SUBSTRING(T1.�������,5,2) = '00' " _
              & "AND T1.������� BETWEEN '" & (Mid(DTOS(dtpF_Date.Value), 1, 4) + "00") & "' " _
                                  & "AND '" & (Mid(DTOS(dtpT_Date.Value), 1, 4) + "00") & "' "
    strSQL = strSQL & "UNION ALL " _
           & "SELECT T1.������ڵ� AS ������ڵ�, T2.������ AS ������, " _
                  & "T1.�з��ڵ� AS �з��ڵ�, T3.�з��� AS �з���, " _
                  & "T1.�����ڵ� AS �����ڵ�, ISNULL(T4.�����,'ERROR!') AS �����, " _
                  & "ISNULL(T4.�԰�,'') AS �԰�, ISNULL(T4.����,'') AS ����, " _
                  & "(T1.������� + '00') AS ����, '(��   ��)' AS ����, " _
                  & "(T1.�԰������) AS �԰����, (T1.��������) AS ������, " _
                  & "(T1.�԰������ - T1.��������) AS ������ " _
             & "FROM ������帶�� T1 " _
             & "" & strJoin & " " _
             & "" & strWhere & " AND SUBSTRING(T1.�������,5,2) <> '00' " _
              & "AND T1.������� > '" & (Mid(DTOS(dtpF_Date.Value), 1, 4) + "00") & "' " _
              & "AND T1.������� < '" & Mid(DTOS(dtpF_Date.Value), 1, 6) & "' "
    strSQL = strSQL & "UNION ALL " _
           & "SELECT T1.������ڵ� AS ������ڵ�, T2.������ AS ������, " _
                  & "T1.�з��ڵ� AS �з��ڵ�, T3.�з��� AS �з���, " _
                  & "T1.�����ڵ� AS �����ڵ�, ISNULL(T4.�����,'ERROR!') AS �����, " _
                  & "ISNULL(T4.�԰�,'') AS �԰�, ISNULL(T4.����,'') AS ����, " _
                  & "(T1.���������) AS ����, '' AS ����, " _
                  & "SUM(T1.�԰����) AS �԰����, SUM(T1.������) AS ������, " _
                  & "SUM(T1.�԰���� - T1.������) As ������ " _
             & "FROM �������⳻�� T1 " _
             & "" & strJoin & " " _
             & "" & strWhere & " AND (T1.��뱸�� = 0) AND (T1.������� BETWEEN 1 AND 2) " _
              & "AND T1.��������� BETWEEN '" & (Mid(DTOS(dtpF_Date.Value), 1, 6) + "01") & "' AND '" & DTOS(dtpT_Date.Value) & "' " _
            & "GROUP BY T1.������ڵ�, T2.������, T1.�з��ڵ�, T3.�з���, " _
                     & "T1.�����ڵ�, T4.�����, T4.�԰�, T4.����, " _
                     & "T1.��������� "
    strSQL = strSQL _
           & "" & strOrderBy & " "
    On Error GoTo ERROR_TABLE_SELECT
    P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
    vsfg1.Rows = P_adoRec.RecordCount + vsfg1.FixedRows
    If P_adoRec.RecordCount = 0 Then
       P_adoRec.Close
       Exit Sub
    Else
       With vsfg1
            .Cell(flexcpForeColor, 0, 0, .FixedRows - 1, .Cols - 1) = vbBlack
            lngR = 1
            .AddItem ""
            lngR = lngR + 1
            .TextMatrix(lngR, 6) = P_adoRec("�����")
            .TextMatrix(lngR, 8) = "�԰�(" & P_adoRec("�԰�") & "), (" & P_adoRec("����") & ")"
            .Cell(flexcpBackColor, lngR, 0, lngR, .Cols - 1) = vbYellow
            Do Until P_adoRec.EOF
               lngR = lngR + 1
               If lngR <> 3 Then    'ó�� ���ڵ� �ƴϸ�
                  If .TextMatrix(lngR - 1, 1) <> (P_adoRec("�з��ڵ�") & P_adoRec("�����ڵ�")) Then '�����ڵ�ٸ���
                     curStockAmt = 0   '����������
                     If .TextMatrix(lngR - 1, 5) <> "" Then '�� �����
                        '�ش����(��)����
                         .AddItem ""
                         '.TextMatrix(lngR, 6) = "(" & Format(Mid(StrDate, 1, 6), "0000-00") & Space(1) & "����)"
                         .TextMatrix(lngR, 6) = "(��  ��)"
                         .TextMatrix(lngR, 7) = curMonIAmt: .TextMatrix(lngR, 9) = curMonOAmt
                         curMonIAmt = 0: curMonOAmt = 0
                         lngR = lngR + 1
                     End If
                     '�ش���(��)����
                     .TextMatrix(lngR, 6) = "(��  ��)"
                     .TextMatrix(lngR, 7) = curTotIAmt: .TextMatrix(lngR, 9) = curTotOAmt
                     curTotIAmt = 0: curTotOAmt = 0
                     .AddItem ""
                     lngR = lngR + 1
                     .AddItem ""
                     lngR = lngR + 1
                     '�����
                     .TextMatrix(lngR, 6) = P_adoRec("�����")
                     .TextMatrix(lngR, 8) = "�԰�(" & P_adoRec("�԰�") & "), (" & P_adoRec("����") & ")"
                     .TextMatrix(lngR, 12) = P_adoRec("�з��ڵ�") & "-" & P_adoRec("�����ڵ�")
                     .Cell(flexcpBackColor, lngR, 0, lngR, .Cols - 1) = vbYellow
                     .AddItem ""
                     lngR = lngR + 1
                  Else
                     If .TextMatrix(lngR - 1, 5) <> "" And _
                         Mid(StrDate, 1, 6) <> Mid(P_adoRec("����"), 1, 6) Then '�� ����� And ���� �ٸ���
                         '�ش����(��)����
                         .AddItem ""
                         '.TextMatrix(lngR, 6) = "(" & Format(Mid(StrDate, 1, 6), "0000-00") & Space(1) & "����)"
                         .TextMatrix(lngR, 6) = "(��  ��)"
                         .TextMatrix(lngR, 7) = curMonIAmt: .TextMatrix(lngR, 9) = curMonOAmt
                         curMonIAmt = 0: curMonOAmt = 0
                         lngR = lngR + 1
                     End If
                  End If
               End If
               .TextMatrix(lngR, 0) = P_adoRec("������ڵ�")
               .TextMatrix(lngR, 1) = P_adoRec("�з��ڵ�") & P_adoRec("�����ڵ�")
               .TextMatrix(lngR, 2) = P_adoRec("�����")
               .TextMatrix(lngR, 3) = P_adoRec("�԰�")
               .TextMatrix(lngR, 4) = P_adoRec("����")
               '5. ����
               If Mid(P_adoRec("����"), 7, 2) = "00" Then
                  .TextMatrix(lngR, 5) = ""
               Else
                  .TextMatrix(lngR, 5) = Format(P_adoRec("����"), "0000-00-00")
               End If
               '6. �԰�����
               If Mid(P_adoRec("����"), 7, 2) = "00" Then
                  If Mid(P_adoRec("����"), 5, 2) = "00" Then
                     .TextMatrix(lngR, 6) = "(" & Mid(P_adoRec("����"), 1, 4) & Space(1) & "�� �̿�)"
                  Else
                     .TextMatrix(lngR, 6) = "(" & Format(Mid(P_adoRec("����"), 1, 6), "0000-00") & Space(1) & "����)"
                  End If
               End If
               .TextMatrix(lngR, 7) = P_adoRec("�԰����")
               '8. �������
               If Mid(P_adoRec("����"), 5, 2) = "00" Then
                  .TextMatrix(lngR, 8) = .TextMatrix(lngR, 5)
               Else
                  .TextMatrix(lngR, 8) = ""
               End If
               .TextMatrix(lngR, 9) = P_adoRec("������")
               .TextMatrix(lngR, 10) = P_adoRec("������")
               .TextMatrix(lngR, 11) = curStockAmt + P_adoRec("������")
               .TextMatrix(lngR, 12) = ""
               If Mid(P_adoRec("����"), 7, 2) <> 0 Then '�ش���԰��� , �ش�������
                  curMonIAmt = curMonIAmt + P_adoRec("�԰����")
                  curMonOAmt = curMonOAmt + P_adoRec("������")
               End If
               '�ش����紩��
               curTotIAmt = curTotIAmt + P_adoRec("�԰����")
               curTotOAmt = curTotOAmt + P_adoRec("������")
               curStockAmt = curStockAmt + P_adoRec("������")
               StrDate = P_adoRec("����")
               'FindRow ����� ����
               '.TextMatrix(lngR, 4) = .TextMatrix(lngR, 0) & P_adoRec("���з��ڵ�")
               '.Cell(flexcpData, lngR, 4, lngR, 4) = .TextMatrix(lngR, 4)
               P_adoRec.MoveNext
               If P_adoRec.EOF = True Then '������ ���ڵ��
                  If .TextMatrix(lngR, 5) <> "" Then  '�� �����
                     lngR = lngR + 1
                     '�ش����(��)����
                     .AddItem ""
                     .TextMatrix(lngR, 6) = "(��  ��)"
                     .TextMatrix(lngR, 7) = curMonIAmt: .TextMatrix(lngR, 9) = curMonOAmt
                     curMonIAmt = 0: curMonOAmt = 0
                  End If
                  '�ش���(��)����
                  lngR = lngR + 1
                  .AddItem ""
                  .TextMatrix(lngR, 6) = "(��  ��)"
                  .TextMatrix(lngR, 7) = curTotIAmt: .TextMatrix(lngR, 9) = curTotOAmt
                   curTotIAmt = 0: curTotOAmt = 0
               End If
            Loop
            P_adoRec.Close
            If .Rows <= PC_intRowCnt Then
               .ScrollBars = flexScrollBarHorizontal
            Else
               .ScrollBars = flexScrollBarBoth
            End If
            If lngRR = 0 Then
               .Row = lngRR       'vsfg1_EnterCell �ڵ�����(���� �Ѱ� �϶��� �ڵ����� ����)
               If .Rows > PC_intRowCnt Then
                  .TopRow = .Rows - PC_intRowCnt + .FixedRows
               End If
            Else
               .Row = lngRR       'vsfg1_EnterCell �ڵ�����(���� �Ѱ� �϶��� �ڵ����� ����)
               If .Rows > PC_intRowCnt Then
                  .TopRow = .Row
               End If
            End If
            .TopRow = 2
            vsfg1_EnterCell       'vsfg1_EnterCell �ڵ�����(���� �Ѱ� �϶��� ������ �ڵ�����)
            .SetFocus
       End With
    End If
    Exit Sub
ERROR_TABLE_SELECT:
    MsgBox Err.Number & Err.Description & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "���ҳ��� �б� ����"
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
               Case 1  '�����ڵ�
                    If Not (Text1(lngC).Text > 0) Then
                       Exit Function
                    End If
               Case 2  '�����
                    If Len(Trim(Text1(lngC).Text)) = 0 Then
                       lngC = 1
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
    
    If Len(Trim(Text1(0).Text)) = 0 And (chkTotal.Value = 0) Then
       Exit Sub
    End If
    If DTOS(dtpF_Date.Value) > DTOS(dtpT_Date.Value) Then
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
            strExeFile = App.Path & ".\���Һ�.rpt"
         Else
            strExeFile = App.Path & ".\���Һ�T.rpt"
         End If
         varRetVal = Dir(strExeFile)
         If Len(varRetVal) = 0 Then
            intRetCHK = 0
         Else
            .ReportFileName = strExeFile
            On Error GoTo ERROR_CRYSTAL_REPORTS
            '--- Formula Fileds ---
            .Formulas(0) = "ForAppPgDate = '" & Format(PB_regUserinfoU.UserClientDate, "0000-00-00") & "'" '���α׷���������
            .Formulas(1) = "ForBranchName= '" & PB_regUserinfoU.UserBranchName & "'"                       '������
            .Formulas(2) = "ForPrtDateTime = '" & strForPrtDateTime & "'"                                  '����Ͻ�
            '�����������
            .Formulas(3) = "ForAppDate = '�������� : ' & '" & Format(DTOS(dtpF_Date.Value), "0000-00-00") & "' & ' ���� ' & '" & Format(DTOS(dtpT_Date.Value), "0000-00-00") & "' & ' ����' "
            .Formulas(4) = "ForSortWay = '" & IIf(optPrtChk0.Value = True, "1", "2") & "' "                '���ļ���
            '--- Parameter Fileds ---
            '���α׷���������
            .StoredProcParam(0) = PB_regUserinfoU.UserClientDate
            '����(����)���ڽ���
            .StoredProcParam(1) = DTOS(dtpF_Date.Value)
            '����(����)��������
            .StoredProcParam(2) = DTOS(dtpT_Date.Value)
            '�����ڵ�
            If (Len(Text1(0).Text) = 0) Or (chkTotal.Value) = 1 Then
               .StoredProcParam(3) = " "
            Else
               .StoredProcParam(3) = Trim(Text1(0).Text)
            End If
            .StoredProcParam(4) = PB_regUserinfoU.UserBranchCode '�����ڵ�
            '--- SortFields
            If optPrtChk0.Value = True Then
               .SortFields(0) = "+{sp1.�з��ڵ�}"
               .SortFields(1) = "+{sp1.�����ڵ�}"
               .SortFields(2) = "+{sp1.����}"
            Else
               .SortFields(0) = "+{sp1.�з��ڵ�}"
               .SortFields(1) = "+{sp1.�����}"
               .SortFields(2) = "+{sp1.�԰�}"
               .SortFields(3) = "+{sp1.����}"
            End If
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
            .WindowTitle = PB_regUserinfoU.UserBranchName & " - " & "���Һ�"
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


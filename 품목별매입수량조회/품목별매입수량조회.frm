VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmǰ�񺰸��Լ�����ȸ 
   BorderStyle     =   0  '����
   Caption         =   "ǰ�񺰸��Լ�����ȸ"
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
      TabIndex        =   12
      Top             =   0
      Width           =   15195
      Begin VB.OptionButton optPrtChk0 
         Caption         =   "�ڵ��"
         Height          =   255
         Left            =   6840
         TabIndex        =   26
         Top             =   150
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton optPrtChk1 
         Caption         =   "�̸���"
         Height          =   255
         Left            =   6840
         TabIndex        =   25
         Top             =   390
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
         Picture         =   "ǰ�񺰸��Լ�����ȸ.frx":0000
         Style           =   1  '�׷���
         TabIndex        =   19
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdClear 
         Height          =   375
         Left            =   9120
         Picture         =   "ǰ�񺰸��Լ�����ȸ.frx":0963
         Style           =   1  '�׷���
         TabIndex        =   17
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdExit 
         Height          =   390
         Left            =   13635
         Picture         =   "ǰ�񺰸��Լ�����ȸ.frx":1308
         Style           =   1  '�׷���
         TabIndex        =   16
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         Height          =   390
         Left            =   12510
         Picture         =   "ǰ�񺰸��Լ�����ȸ.frx":1C56
         Style           =   1  '�׷���
         TabIndex        =   15
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdSave 
         Height          =   390
         Left            =   11385
         Picture         =   "ǰ�񺰸��Լ�����ȸ.frx":25DA
         Style           =   1  '�׷���
         TabIndex        =   14
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdFind 
         Height          =   390
         Left            =   10260
         Picture         =   "ǰ�񺰸��Լ�����ȸ.frx":2E61
         Style           =   1  '�׷���
         TabIndex        =   7
         Top             =   195
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         BackColor       =   &H00008000&
         BorderStyle     =   1  '���� ����
         Caption         =   "ǰ�� ���Լ��� ��ȸ"
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
         TabIndex        =   13
         Top             =   180
         Width           =   4650
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid vsfg1 
      Height          =   7896
      Left            =   60
      TabIndex        =   8
      Top             =   1650
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
      Height          =   970
      Left            =   60
      TabIndex        =   9
      Top             =   630
      Width           =   15195
      Begin VB.TextBox txt_TMt 
         Appearance      =   0  '���
         Height          =   285
         IMEMode         =   8  '����
         Left            =   11160
         MaxLength       =   50
         TabIndex        =   6
         Top             =   585
         Width           =   1695
      End
      Begin VB.TextBox txt_FMt 
         Appearance      =   0  '���
         Height          =   285
         IMEMode         =   8  '����
         Left            =   8040
         MaxLength       =   50
         TabIndex        =   5
         Top             =   585
         Width           =   1695
      End
      Begin VB.CheckBox chkTotalMt 
         Caption         =   "��üǰ��"
         Height          =   255
         Left            =   6840
         TabIndex        =   4
         Top             =   600
         Width           =   1215
      End
      Begin VB.CheckBox chkTotal 
         Caption         =   "��ü ����ó"
         Height          =   255
         Left            =   5030
         TabIndex        =   24
         Top             =   240
         Width           =   1335
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
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         Height          =   285
         IMEMode         =   8  '����
         Index           =   0
         Left            =   2475
         MaxLength       =   8
         TabIndex        =   0
         Top             =   225
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker dtpF_Date 
         Height          =   270
         Left            =   8040
         TabIndex        =   2
         Top             =   240
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
         Left            =   11160
         TabIndex        =   3
         Top             =   240
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   476
         _Version        =   393216
         Format          =   56623105
         UpDown          =   -1  'True
         CurrentDate     =   37763
      End
      Begin VB.Label Label1 
         Caption         =   "[Home]"
         Height          =   240
         Index           =   5
         Left            =   12960
         TabIndex        =   30
         Top             =   645
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "[Home]"
         Height          =   240
         Index           =   4
         Left            =   9840
         TabIndex        =   29
         Top             =   645
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "����"
         Height          =   240
         Index           =   3
         Left            =   13560
         TabIndex        =   28
         Top             =   645
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "����"
         Height          =   240
         Index           =   2
         Left            =   10440
         TabIndex        =   27
         Top             =   645
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "����"
         Height          =   240
         Index           =   12
         Left            =   13560
         TabIndex        =   23
         Top             =   285
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "����"
         Height          =   240
         Index           =   11
         Left            =   10440
         TabIndex        =   22
         Top             =   285
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "��������"
         Height          =   240
         Index           =   10
         Left            =   6960
         TabIndex        =   21
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
         TabIndex        =   20
         Top             =   285
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "[Home]"
         Height          =   240
         Index           =   15
         Left            =   4080
         TabIndex        =   18
         Top             =   285
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "����ó��"
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   1
         Left            =   1275
         TabIndex        =   11
         Top             =   645
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "����ó�ڵ�"
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   0
         Left            =   1275
         TabIndex        =   10
         Top             =   285
         Width           =   1095
      End
   End
   Begin VB.Label lblTotIMny 
      Alignment       =   1  '������ ����
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
      Left            =   6960
      TabIndex        =   32
      Top             =   9720
      Width           =   2775
   End
   Begin VB.Label Label1 
      Alignment       =   1  '������ ����
      Caption         =   "[ �� ���Աݾ� ]"
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
      Index           =   6
      Left            =   5040
      TabIndex        =   31
      Top             =   9720
      Width           =   1695
   End
End
Attribute VB_Name = "frmǰ�񺰸��Լ�����ȸ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'+-------------------------------------------------------------------------------------------------------+
' ���α׷� �� �� : ǰ�񺰸��Լ�����ȸ
' ���� Control : VideoSoft VSFlexGrid 7.0(OLEDB) = vsflex7.ocx
' ������ Table   :
' ��  ��  ��  �� :
'+-------------------------------------------------------------------------------------------------------+
Option Explicit
Private P_blnActived       As Boolean
Private P_adoRec           As New ADODB.Recordset
Private P_intButton        As Integer
Private Const PC_intRowCnt As Integer = 26  '�׸��� �� ������ �� ���(FixedRows ����)

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
                   cmdPrint.Enabled = False: cmdClear.Enabled = False: cmdFind.Enabled = True
                   cmdSave.Enabled = False: cmdDelete.Enabled = False: cmdExit.Enabled = True
              Case Is <= 50 '����, �߰�, ����, ��ȸ, �μ�
                   cmdPrint.Enabled = False: cmdClear.Enabled = False: cmdFind.Enabled = True
                   cmdSave.Enabled = False: cmdDelete.Enabled = False: cmdExit.Enabled = True
              Case Is <= 99 '����, �߰�, ����, ��ȸ, �μ�
                   cmdPrint.Enabled = False: cmdClear.Enabled = False: cmdFind.Enabled = True
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
    If KeyCode = vbKeyReturn Then SendKeys "{tab}"
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
       PB_strSupplierName = ""  'Trim(Text1(Index + 1).Text)
       frm����ó�˻�.Show vbModal
       If (Len(PB_strSupplierCode) + Len(PB_strSupplierName)) = 0 Then '�˻����� ���(ESC)
       Else
          Text1(Index).Text = PB_strSupplierCode
          Text1(Index + 1).Text = PB_strSupplierName
       End If
       If PB_strSupplierCode <> "" Then
          dtpF_Date.SetFocus
       End If
       PB_strSupplierCode = "": PB_strSupplierName = ""
    ElseIf _
       KeyCode = vbKeyDelete Then
       If Len(Text1(0).Text) = 0 Then
          Text1(1).Text = ""
          Exit Sub
       End If
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
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "����ó �б� ����"
    Unload Me
    Exit Sub
End Sub

Private Sub Text1_LostFocus(Index As Integer)
Dim strSQL As String
    With Text1(Index)
         .Text = Trim(.Text)
         Select Case Index
                Case 0 '����ó
                     .Text = UPPER(Trim(.Text))
                     If Len(.Text) < 1 Then
                        Text1(1).Text = ""
                     End If
                Case Else
         End Select
    End With
    Exit Sub
End Sub

Private Sub chkTotalMt_KeyDown(KeyCode As Integer, Shift As Integer)
    If (KeyCode = vbKeyHome Or KeyCode = vbKeyReturn) Then
       txt_FMt.SetFocus
    End If
End Sub
Private Sub txt_FMt_GotFocus()
    With txt_FMt
         .SelStart = 0
         .SelLength = Len(Trim(.Text))
    End With
End Sub
Private Sub txt_FMt_Keydown(KeyCode As Integer, Shift As Integer)
    If (KeyCode = vbKeyHome Or KeyCode = vbKeyReturn) And Len(Trim(txt_FMt.Text)) > 0 Then  '����˻�
       PB_strCallFormName = "frmǰ�񺰸��Լ�����ȸ"
       PB_strMaterialsCode = Trim(txt_FMt.Text)
       PB_strMaterialsName = ""
       frm����˻�.Show vbModal
       If (Len(PB_strMaterialsCode) + Len(PB_strMaterialsName)) = 0 Then '�˻����� ���(ESC)
       Else
          txt_FMt.Text = Mid(PB_strMaterialsCode, 3)
       End If
       If PB_strMaterialsCode <> "" Then
          SendKeys "{tab}"
       End If
       PB_strCallFormName = "": PB_strMaterialsCode = "": PB_strMaterialsName = ""
    Else
       If KeyCode = vbKeyReturn Then
          SendKeys "{tab}"
       End If
    End If
    Exit Sub
End Sub
Private Sub txt_FMt_LostFocus()
    With txt_FMt
         .Text = Trim(.Text)
         If Len(.Text) = 0 Then
         End If
    End With
End Sub
Private Sub txt_TMt_GotFocus()
    With txt_TMt
         .SelStart = 0
         .SelLength = Len(Trim(.Text))
    End With
End Sub
Private Sub txt_TMt_Keydown(KeyCode As Integer, Shift As Integer)
    If (KeyCode = vbKeyHome Or KeyCode = vbKeyReturn) And Len(Trim(txt_TMt.Text)) > 0 Then  '����˻�
       PB_strCallFormName = "frmǰ�񺰸��Լ�����ȸ"
       PB_strMaterialsCode = Trim(txt_TMt.Text)
       PB_strMaterialsName = ""
       frm����˻�.Show vbModal
       If (Len(PB_strMaterialsCode) + Len(PB_strMaterialsName)) = 0 Then '�˻����� ���(ESC)
       Else
          txt_TMt.Text = Mid(PB_strMaterialsCode, 3)
       End If
       If PB_strMaterialsCode <> "" Then
          SendKeys "{tab}"
          cmdFind_Click
       End If
       PB_strCallFormName = "": PB_strMaterialsCode = "": PB_strMaterialsName = ""
    Else
       If KeyCode = vbKeyReturn Then
          'SendKeys "{tab}"
          cmdFind_Click
       End If
    End If
    Exit Sub
End Sub
Private Sub txt_TMt_LostFocus()
    With txt_TMt
         .Text = Trim(.Text)
         If Len(.Text) = 0 Then
         End If
    End With
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
    Set frmǰ�񺰸��Լ�����ȸ = Nothing
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
    With vsfg1              'Rows 1, Cols 11, RowHeightMax(Min) 300
         .AllowBigSelection = False
         .AllowSelection = False
         .AllowUserResizing = flexResizeColumns
         .BackColorBkg = &H8000000F
         .BackColorSel = &H8000&
         .ExtendLastCol = True
         .FocusRect = flexFocusHeavy
         .ScrollBars = flexScrollBarVertical
         .ScrollTrack = True
         .SelectionMode = flexSelectionByRow
         .FixedRows = 1
         .FixedCols = 4
         '.FrozenCols = 5
         .Rows = 1             'Subvsfg1_Fill����ÿ� ����
         .Cols = 11
         .RowHeightMax = 300
         .RowHeightMin = 300
         .ColWidth(0) = 1000   '������ڵ�          'H
         .ColWidth(1) = 2000   '������            'H
         .ColWidth(2) = 1000   '�з��ڵ�
         .ColWidth(3) = 2000   '�з���
         .ColWidth(4) = 3000   '�����ڵ�
         .ColWidth(5) = 3500   '�����ڵ�(�з�+����) 'H
         .ColWidth(6) = 3000   '�����
         .ColWidth(7) = 3000   '�԰�
         .ColWidth(8) = 2000   '����
         .ColWidth(9) = 1000   '����
         .ColWidth(10) = 1400  '���� * �԰�ܰ�(���ް�)
         
         .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = False
         .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
         .TextMatrix(0, 0) = "������ڵ�"  'H
         .TextMatrix(0, 1) = "������"    'H
         .TextMatrix(0, 2) = "�з��ڵ�"    'H
         .TextMatrix(0, 3) = "�з���"      'H
         .TextMatrix(0, 4) = "ǰ���ڵ�"    '�����ڵ�
         .TextMatrix(0, 5) = "ǰ���ڵ�"    '�з��ڵ�+�����ڵ�
         .TextMatrix(0, 6) = "ǰ��"
         .TextMatrix(0, 7) = "�԰�"
         .TextMatrix(0, 8) = "����"
         .TextMatrix(0, 9) = "����"
         .TextMatrix(0, 10) = "���Աݾ�"   '���� * �ܰ�
         .ColHidden(0) = True: .ColHidden(1) = True: .ColHidden(5) = True: .ColHidden(10) = True
         For lngC = 0 To .Cols - 1
             Select Case lngC
                    Case 1, 3, 4, 5, 6, 7, 9
                         .ColAlignment(lngC) = flexAlignLeftCenter
                    Case 0, 2
                         .ColAlignment(lngC) = flexAlignCenterCenter
                    Case Else
                         .ColAlignment(lngC) = flexAlignRightCenter
             End Select
         Next lngC
         For lngC = 0 To .Cols - 1
             Select Case lngC
                    Case 8
                         .ColFormat(lngC) = "#,#"
                    Case 10
                         .ColFormat(lngC) = "#,#.00"
             End Select
         Next lngC
         '.MergeCells = flexMergeFixedOnly
         '.MergeRow(0) = True: .MergeRow(1) = True
         'For lngC = 0 To 2
         '    .MergeCol(lngC) = True
         'Next lngC
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

    lblTotIMny.Caption = ""
    vsfg1.Rows = 1
    With vsfg1
         '�˻����� ����ó
         If chkTotal.Value = 0 Then  '�Ǻ� ��ȸ
            If Len(Text1(0).Text) > 0 Then
               strWhere = "WHERE T1.����ó�ڵ� = '" & Trim(Text1(0).Text) & "' "
            Else
               Text1(0).SetFocus
               Exit Sub
            End If
         End If
         If chkTotalMt.Value = 0 Then '�Ǻ� ��ȸ
            If Len(txt_TMt.Text) > 0 Then
               strWhere = strWhere + IIf(Trim(strWhere) = "", "WHERE ", "AND ") _
                         & "T1.�����ڵ� BETWEEN '" & txt_FMt.Text & "' AND '" & txt_TMt.Text & "' "
            Else
               txt_TMt.SetFocus
               Exit Sub
            End If
         End If
         strWhere = strWhere + IIf(Trim(strWhere) = "", "WHERE ", "AND ") _
                  & "T1.������ڵ� = '" & PB_regUserinfoU.UserBranchCode & "' "
    End With
    strOrderBy = "ORDER BY T1.������ڵ�, " & IIf(optPrtChk0.Value = True, "T1.�з��ڵ�, T1.�����ڵ� ", "T5.����� ") & " "
    P_adoRec.CursorLocation = adUseClient
    strSQL = "SELECT T1.������ڵ� AS ������ڵ�, T2.������ AS ������, " _
                  & "T1.�з��ڵ� AS �з��ڵ�, T4.�з��� AS �з���, " _
                  & "T1.�����ڵ� AS �����ڵ�, ISNULL(T5.�����,'ERROR!') AS �����, " _
                  & "ISNULL(T5.�԰�,'') AS �԰�, ISNULL(T5.����,'') AS ����, " _
                  & "SUM(T1.�԰����) AS �԰����, SUM(T1.�԰���� * T1.�԰�ܰ�) AS �԰�ݾ� " _
             & "FROM �������⳻�� T1 " _
             & "LEFT JOIN ����� T2 ON T2.������ڵ� = T1.������ڵ� " _
             & "LEFT JOIN ����ó T3 ON T3.������ڵ� = T1.������ڵ� AND T3.����ó�ڵ� = T1.����ó�ڵ� " _
             & "LEFT JOIN ����з� T4 ON T4.�з��ڵ� = T1.�з��ڵ� " _
             & "LEFT JOIN ���� T5 ON T5.�з��ڵ� = T1.�з��ڵ� AND T5.�����ڵ� = T1.�����ڵ� " _
             & "" & strWhere & " AND (T1.������� = 1) AND T1.��뱸�� = 0 " _
              & "AND T1.��������� BETWEEN '" & DTOS(dtpF_Date.Value) & "' AND '" & DTOS(dtpT_Date.Value) & "' "
    strSQL = strSQL _
           & "GROUP BY T1.������ڵ�, T2.������," _
                    & "T1.�з��ڵ�, T4.�з���, T1.�����ڵ�, T5.�����, " _
                    & "T5.�԰�, T5.���� "
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
            lngR = 0
            Do Until P_adoRec.EOF
               lngR = lngR + 1
               .TextMatrix(lngR, 0) = P_adoRec("������ڵ�")
               .TextMatrix(lngR, 1) = P_adoRec("������")
               .TextMatrix(lngR, 2) = P_adoRec("�з��ڵ�")
               .TextMatrix(lngR, 3) = P_adoRec("�з���")
               .TextMatrix(lngR, 4) = P_adoRec("�����ڵ�")
               .TextMatrix(lngR, 5) = P_adoRec("�з��ڵ�") & P_adoRec("�����ڵ�")
               .TextMatrix(lngR, 6) = P_adoRec("�����")
               .TextMatrix(lngR, 7) = P_adoRec("�԰�")
               .TextMatrix(lngR, 8) = P_adoRec("�԰����")
               .TextMatrix(lngR, 9) = P_adoRec("����")
               .TextMatrix(lngR, 10) = P_adoRec("�԰�ݾ�")
               'FindRow ����� ����
               '.TextMatrix(lngR, 5) = .TextMatrix(lngR, 0) & .TextMatrix(lngR, 5)
               '.Cell(flexcpData, lngR, 5, lngR, 5) = .TextMatrix(lngR, 5)
               lblTotIMny.Caption = Format((Vals(lblTotIMny.Caption) + .ValueMatrix(lngR, 10)), "#,#.00")
               P_adoRec.MoveNext
            Loop
            P_adoRec.Close
            '��ü �հ�
            If .Rows <= PC_intRowCnt Then
               .ScrollBars = flexScrollBarNone
            Else
               .ScrollBars = flexScrollBarVertical
            End If
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
            .TopRow = 1
            vsfg1_EnterCell       'vsfg1_EnterCell �ڵ�����(���� �Ѱ� �϶��� ������ �ڵ�����)
            .SetFocus
       End With
    End If
    Exit Sub
ERROR_TABLE_SELECT:
    MsgBox Err.Number & Err.Description & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "ǰ�񺰸��Լ��� �б� ����"
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
               Case Else
        End Select
    Next lngC
    blnOK = True
End Function


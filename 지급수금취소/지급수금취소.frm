VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm���޼������ 
   BorderStyle     =   0  '����
   Caption         =   "���޼������"
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
      TabIndex        =   10
      Top             =   0
      Width           =   15195
      Begin VB.CommandButton cmdPrint 
         Height          =   375
         Left            =   7980
         Picture         =   "���޼������.frx":0000
         Style           =   1  '�׷���
         TabIndex        =   16
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdClear 
         Height          =   375
         Left            =   9120
         Picture         =   "���޼������.frx":0963
         Style           =   1  '�׷���
         TabIndex        =   15
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdExit 
         Height          =   390
         Left            =   13635
         Picture         =   "���޼������.frx":1308
         Style           =   1  '�׷���
         TabIndex        =   14
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         Height          =   390
         Left            =   12510
         Picture         =   "���޼������.frx":1C56
         Style           =   1  '�׷���
         TabIndex        =   13
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdSave 
         Height          =   390
         Left            =   11385
         Picture         =   "���޼������.frx":25DA
         Style           =   1  '�׷���
         TabIndex        =   12
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdFind 
         Height          =   390
         Left            =   10260
         Picture         =   "���޼������.frx":2E61
         Style           =   1  '�׷���
         TabIndex        =   7
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
         Caption         =   "����/�������"
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
         TabIndex        =   11
         Top             =   180
         Width           =   4650
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid vsfg1 
      Height          =   8476
      Left            =   60
      TabIndex        =   8
      Top             =   1620
      Width           =   15195
      _cx             =   26802
      _cy             =   14951
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
      Height          =   960
      Left            =   60
      TabIndex        =   9
      Top             =   630
      Width           =   15195
      Begin VB.TextBox Text1 
         Appearance      =   0  '���
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   6000
         TabIndex        =   4
         Top             =   570
         Width           =   3375
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         Height          =   285
         IMEMode         =   8  '����
         Index           =   0
         Left            =   6000
         MaxLength       =   8
         TabIndex        =   3
         Top             =   240
         Width           =   1095
      End
      Begin VB.CheckBox chkTotal 
         Caption         =   "��ü ��ü"
         Height          =   255
         Left            =   3600
         TabIndex        =   2
         Top             =   360
         Value           =   1  'Ȯ��
         Width           =   1215
      End
      Begin VB.OptionButton optJSGbn 
         Caption         =   "�� �� �� �Ա�"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   1440
         TabIndex        =   1
         Top             =   550
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.OptionButton optJSGbn 
         Caption         =   "�����ޱ� ����"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   1440
         TabIndex        =   0
         Top             =   270
         Width           =   1575
      End
      Begin MSComCtl2.DTPicker dtpFDate 
         Height          =   270
         Left            =   10800
         TabIndex        =   5
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
         TabIndex        =   6
         Top             =   240
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   476
         _Version        =   393216
         Format          =   57540609
         UpDown          =   -1  'True
         CurrentDate     =   37763
      End
      Begin VB.Label lblTotMny 
         Alignment       =   1  '������ ����
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   10660
         TabIndex        =   29
         Top             =   630
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "��ü�ݾ�"
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   2
         Left            =   9720
         TabIndex        =   28
         Top             =   630
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "��ü��"
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   1
         Left            =   4920
         TabIndex        =   27
         Top             =   630
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "[Home]"
         Height          =   240
         Index           =   15
         Left            =   7200
         TabIndex        =   26
         Top             =   285
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   ")"
         Height          =   255
         Index           =   3
         Left            =   9480
         TabIndex        =   25
         Top             =   405
         Width           =   135
      End
      Begin VB.Label Label3 
         Caption         =   "("
         Height          =   255
         Index           =   2
         Left            =   4920
         TabIndex        =   24
         Top             =   405
         Width           =   135
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "��ü�ڵ�"
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   0
         Left            =   4920
         TabIndex        =   23
         Top             =   285
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   ")"
         Height          =   255
         Index           =   1
         Left            =   3000
         TabIndex        =   22
         Top             =   405
         Width           =   135
      End
      Begin VB.Label Label3 
         Caption         =   "("
         Height          =   255
         Index           =   0
         Left            =   1200
         TabIndex        =   21
         Top             =   405
         Width           =   135
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "����"
         Height          =   240
         Index           =   12
         Left            =   14160
         TabIndex        =   20
         Top             =   285
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "����"
         Height          =   240
         Index           =   11
         Left            =   12120
         TabIndex        =   19
         Top             =   285
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "�ŷ�����"
         Height          =   240
         Index           =   10
         Left            =   9720
         TabIndex        =   18
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
         TabIndex        =   17
         Top             =   405
         Width           =   975
      End
   End
End
Attribute VB_Name = "frm���޼������"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'+-------------------------------------------------------------------------------------------------------+
' ���α׷� �� �� : �����ޱ����� ���, �̼����Ա� ���
' ���� Control : VideoSoft VSFlexGrid 7.0(OLEDB) = vsflex7.ocx
' ������ Table   :
' ��  ��  ��  �� : �����ޱݳ��� �Ǵ� �̼��ݳ����� ����(UPDATE)
'+-------------------------------------------------------------------------------------------------------+
Option Explicit
Private P_blnActived       As Boolean
Private P_adoRec           As New ADODB.Recordset
Private P_intButton        As Integer
Private P_intBeforeOptGbn  As Integer
Private Const PC_intRowCnt As Integer = 28  '�׸��� �� ������ �� ���(FixedRows ����)

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

    frmMain.SBar.Panels(4).Text = "�ܼ��� �����ޱݳ���(����)/�̼��ݳ���(����)�� ���(����) �Ǵ� �����մϴ�. "
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
                   cmdSave.Enabled = True: cmdDelete.Enabled = True: cmdExit.Enabled = True
              Case Is <= 99 '����, �߰�, ����, ��ȸ, �μ�
                   cmdPrint.Enabled = False: cmdClear.Enabled = False: cmdFind.Enabled = True
                   cmdSave.Enabled = True: cmdDelete.Enabled = True: cmdExit.Enabled = True
              Case Else
                   cmdClear.Enabled = False: cmdFind.Enabled = False: cmdSave.Enabled = False
                   cmdDelete.Enabled = False: cmdExit.Enabled = True
       End Select
       P_intBeforeOptGbn = 1
       optJSGbn(0).ForeColor = vbBlue: optJSGbn(1).ForeColor = vbRed
       Label1(0).ForeColor = vbRed: Label1(1).ForeColor = vbRed
       dtpFDate.Value = Format(Mid(PB_regUserinfoU.UserClientDate, 1, 6) & "01", "0000-00-00")
       dtpTDate.Value = DateAdd("d", -1, DateAdd("m", 1, dtpFDate.Value))
       P_blnActived = True
       Screen.MousePointer = vbDefault
    End If
    Exit Sub
ERROR_FORM_ACTIVATE_CONNECTION_SERVER:
    MsgBox Err.Number & Err.Description & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "�ŷ��������(�������� ���� ����)"
    Unload Me
    Exit Sub
End Sub

'+---------------+
'/// �˻����� ///
'+---------------+
Private Sub optJSGbn_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    With optJSGbn
         If KeyCode = vbKeyReturn Then
            chkTotal.SetFocus
         End If
    End With
End Sub
Private Sub optJSGbn_Click(Index As Integer)
    With optJSGbn
         If optJSGbn(0).Value = True Then
            Label1(0).ForeColor = vbBlue: Label1(1).ForeColor = vbBlue
         Else
            Label1(0).ForeColor = vbRed: Label1(1).ForeColor = vbRed
         End If
         If Index <> P_intBeforeOptGbn Then
            P_intBeforeOptGbn = Index
            Text1(0).Text = "": Text1(1).Text = "": lblTotMny.Caption = "0"
            vsfg1.Rows = 1
         End If
    End With
End Sub

Private Sub chkTotal_KeyDown(KeyCode As Integer, Shift As Integer)
    With chkTotal
         If KeyCode = vbKeyReturn Then
            If chkTotal.Value = 1 Then
               dtpFDate.SetFocus
            Else
               Text1(0).SetFocus
            End If
         End If
    End With
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
       If optJSGbn(0).Value = True Then
          frm����ó�˻�.Show vbModal
       Else
          frm����ó�˻�.Show vbModal
       End If
       If (Len(PB_strSupplierCode) + Len(PB_strSupplierName)) = 0 Then '�˻����� ���(ESC)
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
                 Case Text1.UBound
                      'If chkTotal.Value = 0 And cmdSave.Enabled = True And vsfg1.Rows > 1 Then
                      '   cmdSave.SetFocus
                      '   Exit Sub
                      'End If
           End Select
           SendKeys "{tab}"
       End If
    End If
    Exit Sub
ERROR_TABLE_SELECT:
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "����/����ó �б� ����"
    Unload Me
    Exit Sub
End Sub

Private Sub Text1_LostFocus(Index As Integer)
Dim strSQL As String
    With Text1(Index)
         .Text = Trim(.Text)
         Select Case Index
                Case 0 '��ü
                     .Text = UPPER(Trim(.Text))
                     If Len(Trim(.Text)) < 1 Then
                        Text1(1).Text = ""
                     End If
                Case Else
                     .Text = Trim(.Text)
         End Select
    End With
    Exit Sub
End Sub

Private Sub dtpFDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub
Private Sub dtpTDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then cmdFind_Click
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
Private Sub vsfg1_Click()
Dim strData As String
    With vsfg1
         If (.MouseRow >= 0 And .MouseRow < .FixedRows) Then
            .Col = .MouseCol
            .Cell(flexcpForeColor, 0, 0, 0, .Cols - 1) = vbBlack
            .Cell(flexcpForeColor, 0, .MouseCol, 0, .MouseCol) = vbRed
            strData = Trim(.Cell(flexcpData, .Row, 0))
            Select Case .MouseCol
                   Case 1
                        .ColSel = 2
                        .ColSort(0) = flexSortNone
                        .ColSort(1) = IIf(P_intButton = 1, flexSortGenericAscending, flexSortGenericDescending)
                        .ColSort(2) = IIf(P_intButton = 1, flexSortGenericAscending, flexSortGenericDescending)
                        .Sort = flexSortUseColSort
                   Case 2
                        .ColSel = 2
                        .ColSort(0) = flexSortNone
                        .ColSort(1) = IIf(P_intButton = 1, flexSortGenericAscending, flexSortGenericDescending)
                        .ColSort(2) = IIf(P_intButton = 1, flexSortGenericAscending, flexSortGenericDescending)
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
Private Sub vsfg1_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    With vsfg1
         P_intButton = Button
         If .MouseRow >= .FixedRows Then
            If (.MouseCol = 8) Then   '�ŷ��ݾ�
               If Button = vbLeftButton Then
                  .Select .MouseRow, IIf(.MouseCol = .Col, .MouseCol, .Col)
                  .EditCell
               End If
            ElseIf _
               (.MouseCol = 9 And .TextMatrix(.Row, 7) = "����") Then    '��������, col(7).����
               If Button = vbLeftButton Then
                  .Select .MouseRow, IIf(.MouseCol = .Col, .MouseCol, .Col)
                  .EditCell
               End If
            ElseIf _
               (.MouseCol = 10 And .TextMatrix(.Row, 7) = "����") Then   '������ȣ, col(7).����
               If Button = vbLeftButton Then
                  .Select .MouseRow, IIf(.MouseCol = .Col, .MouseCol, .Col)
                  .EditCell
               End If
            ElseIf _
               (.MouseCol = 11) Then   '����
               If Button = vbLeftButton Then
                  .Select .MouseRow, IIf(.MouseCol = .Col, .MouseCol, .Col)
                  .EditCell
               End If
            End If
         End If
    End With
End Sub
Sub vsfg1_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
Dim blnModify As Boolean
Dim curTmpMny As Currency
    With vsfg1
         If Row >= .FixedRows Then
            If (Col = 8) Then         '�ŷ��ݾ�
               If .TextMatrix(Row, Col) <> .EditText Then
                  If (IsNumeric(.EditText) = False Or _
                     Fix(Vals(.EditText)) < Vals(.EditText) Or IsNumeric(Right(.EditText, 1)) = False) Then
                     Beep
                     Cancel = True
                  Else
                     blnModify = True
                     curTmpMny = .ValueMatrix(Row, 8)
                     .TextMatrix(Row, 8) = Vals(.EditText)   '�ŷ��ݾ�
                  End If
               End If
            ElseIf _
               (Col = 9) Then '��������
               If .TextMatrix(Row, 9) <> .EditText Then
                  .EditText = Format(Replace(.EditText, "-", ""), "0000-00-00")
                  If Not ((Len(Trim(.EditText)) = 10) And IsDate(.EditText) And Val(.EditText) > 2000) Then
                     Beep
                     Cancel = True
                  Else
                     blnModify = True
                  End If
               End If
            ElseIf _
               (Col = 10) Then '������ȣ
               If .TextMatrix(Row, Col) <> .EditText Then
                  If Not (LenH(Trim(.EditText)) <= 20) Then
                     Beep
                     .TextMatrix(Row, Col) = .EditText
                     Cancel = True
                  Else
                     blnModify = True
                  End If
               End If
            ElseIf _
               (Col = 11) Then '���� ���� �˻�
               If .TextMatrix(Row, Col) <> .EditText Then
                  If Not (LenH(Trim(.EditText)) <= 50) Then
                     Beep
                     .TextMatrix(Row, Col) = .EditText
                     Cancel = True
                  Else
                     blnModify = True
                  End If
               End If
            End If
            '����ǥ�� + �ݾ�����
            If blnModify = True Then
               .Cell(flexcpBackColor, Row, Col, Row, Col) = vbRed
               .Cell(flexcpForeColor, Row, Col, Row, Col) = vbWhite
               Select Case Col
                      Case 8
                           lblTotMny.Caption = Format(Vals(lblTotMny.Caption) - curTmpMny + .ValueMatrix(Row, 8), "#,0")
                      Case Else
                      
               End Select
            End If
         End If
    End With
End Sub

Private Sub vsfg1_EnterCell()
Dim lngR As Long
Dim lngC As Long
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
End Sub
'+-----------+
'/// ��ȸ ///
'+-----------+
Private Sub cmdFind_Click()
    Screen.MousePointer = vbHourglass
    cmdFind.Enabled = False
    Subvsfg1_FILL
    cmdFind.Enabled = True
    Screen.MousePointer = vbDefault
End Sub
'+-----------+
'/// ���� ///
'+-----------+
Private Sub cmdSave_Click()
Dim strSQL       As String
Dim intRetVal    As Integer
Dim lngCnt       As Long
    With vsfg1
         If .Row >= .FixedRows Then
            If .Cell(flexcpBackColor, .Row, 8, .Row, 8) = vbRed Or .Cell(flexcpBackColor, .Row, 9, .Row, 9) = vbRed Or _
               .Cell(flexcpBackColor, .Row, 10, .Row, 10) = vbRed Or .Cell(flexcpBackColor, .Row, 11, .Row, 11) = vbRed Then
               intRetVal = MsgBox("����� �ڷḦ �����Ͻðڽ��ϱ� ?", vbCritical + vbYesNo + vbDefaultButton1, "�ڷ� ����")
            Else
               .SetFocus
               Exit Sub
            End If
            If intRetVal = vbYes Then
               Screen.MousePointer = vbHourglass
               cmdSave.Enabled = False
               PB_adoCnnSQL.BeginTrans
               P_adoRec.CursorLocation = adUseClient
               If (.ValueMatrix(.Row, 3) Mod 2) = 1 Then '1.����
                  strSQL = "UPDATE �����ޱݳ��� SET " _
                                & "�����ޱ����ޱݾ� = " & .ValueMatrix(.Row, 8) & ", " _
                                & "�������� = '" & DTOS(.TextMatrix(.Row, 9)) & "', " _
                                & "������ȣ = '" & .TextMatrix(.Row, 10) & "', " _
                                & "���� = '" & .TextMatrix(.Row, 10) & "', " _
                                & "�������� = '" & PB_regUserinfoU.UserServerDate & "', " _
                                & "������ڵ� = '" & PB_regUserinfoU.UserCode & "' " _
                          & "WHERE ������ڵ� = '" & PB_regUserinfoU.UserBranchCode & "' " _
                            & "AND ����ó�ڵ� = '" & .TextMatrix(.Row, 5) & "' " _
                            & "AND �����ޱ��������� = '" & DTOS(.TextMatrix(.Row, 1)) & "' " _
                            & "AND �����ޱ����޽ð� = '" & .TextMatrix(.Row, 14) & "' "
               Else                                      '2.����
                  strSQL = "UPDATE �̼��ݳ��� SET " _
                                & "�̼����Աݱݾ� = " & .ValueMatrix(.Row, 8) & ", " _
                                & "�������� = '" & DTOS(.TextMatrix(.Row, 9)) & "', " _
                                & "������ȣ = '" & .TextMatrix(.Row, 10) & "', " _
                                & "���� = '" & .TextMatrix(.Row, 10) & "', " _
                                & "�������� = '" & PB_regUserinfoU.UserServerDate & "', " _
                                & "������ڵ� = '" & PB_regUserinfoU.UserCode & "' " _
                          & "WHERE ������ڵ� = '" & PB_regUserinfoU.UserBranchCode & "' " _
                            & "AND ����ó�ڵ� = '" & .TextMatrix(.Row, 5) & "' " _
                            & "AND �̼����Ա����� = '" & DTOS(.TextMatrix(.Row, 1)) & "' " _
                            & "AND �̼����Աݽð� = '" & .TextMatrix(.Row, 14) & "' "
               End If
               On Error GoTo ERROR_TABLE_UPDATE
               PB_adoCnnSQL.Execute strSQL
               PB_adoCnnSQL.CommitTrans
               .TextMatrix(.Row, 12) = PB_regUserinfoU.UserCode
               .TextMatrix(.Row, 13) = PB_regUserinfoU.UserName
               'lblTotMny.Caption = Format(Vals(lblTotMny.Caption) - .ValueMatrix(.Row, 8), "#,0")
               .Cell(flexcpBackColor, .Row, .FixedCols, .Row, .Cols - 1) = vbWhite
               .Cell(flexcpForeColor, .Row, .FixedCols, .Row, .Cols - 1) = vbBlack
               If .Rows <= PC_intRowCnt Then
                  .ScrollBars = flexScrollBarVertical
               End If
               Screen.MousePointer = vbDefault
               vsfg1_EnterCell
               cmdSave.Enabled = True
            End If
            .SetFocus
         End If
    End With
    Screen.MousePointer = vbDefault
    Exit Sub
ERROR_TABLE_SELECT:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "�б� ����"
    Unload Me
    Screen.MousePointer = vbDefault
    Exit Sub
ERROR_TABLE_UPDATE:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "���� ����"
    Unload Me
    Screen.MousePointer = vbDefault
    Exit Sub
ERROR_TABLE_DELETE:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "���� ����"
    Unload Me
    Screen.MousePointer = vbDefault
    Exit Sub
End Sub

'+-----------+
'/// ���� ///
'+-----------+
Private Sub cmdDelete_Click()
Dim strSQL       As String
Dim intRetVal    As Integer
Dim lngCnt       As Long
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
                  MsgBox "XXX(" & Format(lngCnt, "#,#") & "��)�� �����Ƿ� ������ �� �����ϴ�.", vbCritical, "�����ޱ�/���ݳ��� ���� �Ұ�"
                  cmdDelete.Enabled = True
                  Screen.MousePointer = vbDefault
                  Exit Sub
               End If
               PB_adoCnnSQL.BeginTrans
               P_adoRec.CursorLocation = adUseClient
               If (.ValueMatrix(.Row, 3) Mod 2) = 1 Then '1.����
                  strSQL = "DELETE FROM �����ޱݳ��� " _
                          & "WHERE ������ڵ� = '" & PB_regUserinfoU.UserBranchCode & "' " _
                            & "AND ����ó�ڵ� = '" & .TextMatrix(.Row, 5) & "' " _
                            & "AND �����ޱ��������� = '" & DTOS(.TextMatrix(.Row, 1)) & "' " _
                            & "AND �����ޱ����޽ð� = '" & .TextMatrix(.Row, 14) & "' "
               Else                                      '2.����
                  strSQL = "DELETE FROM �̼��ݳ��� " _
                          & "WHERE ������ڵ� = '" & PB_regUserinfoU.UserBranchCode & "' " _
                            & "AND ����ó�ڵ� = '" & .TextMatrix(.Row, 5) & "' " _
                            & "AND �̼����Ա����� = '" & DTOS(.TextMatrix(.Row, 1)) & "' " _
                            & "AND �̼����Աݽð� = '" & .TextMatrix(.Row, 14) & "' "
               End If
               On Error GoTo ERROR_TABLE_DELETE
               PB_adoCnnSQL.Execute strSQL
               PB_adoCnnSQL.CommitTrans
               lblTotMny.Caption = Format(Vals(lblTotMny.Caption) - .ValueMatrix(.Row, 8), "#,0")
               .RemoveItem .Row
               If .Rows <= PC_intRowCnt Then
                  .ScrollBars = flexScrollBarHorizontal
               End If
               Screen.MousePointer = vbDefault
               If .Rows = 1 Then
                  .Row = 0
                  cmdFind.SetFocus
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
ERROR_TABLE_SELECT:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "�б� ����"
    Unload Me
    Screen.MousePointer = vbDefault
    Exit Sub
ERROR_TABLE_UPDATE:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "���� ����"
    Unload Me
    Screen.MousePointer = vbDefault
    Exit Sub
ERROR_TABLE_DELETE:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "���� ����"
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
    Set frm���޼������ = Nothing
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
    With vsfg1              'Rows 1, Cols 15, RowHeightMax(Min) 300
         .AllowBigSelection = False
         .AllowSelection = False
         .AllowUserResizing = flexResizeColumns
         .BackColorBkg = &H8000000F
         .BackColorSel = &H8000&
         .ExtendLastCol = True
         .ScrollBars = flexScrollBarHorizontal
         .ScrollTrack = True
         .SelectionMode = flexSelectionByRow
         .FixedRows = 1
         .FixedCols = 5
         .Rows = 1             'Subvsfg1_Fill����ÿ� ����
         .Cols = 15
         .RowHeightMax = 300
         .RowHeightMin = 300
         .ColWidth(0) = 1000   'KEY(�����ڵ�-����ó�ڵ�-����ó�ڵ�-�������-���������-�����ð�)
         .ColWidth(1) = 1000   '�ŷ�����
         .ColWidth(2) = 900    '�ŷ��ð�
         .ColWidth(3) = 1000   '�ŷ�����
         .ColWidth(4) = 600    '�ŷ����и�
         .ColWidth(5) = 1200   '��ü�ڵ�
         .ColWidth(6) = 2000   '��ü��
         .ColWidth(7) = 600    '�������
         .ColWidth(8) = 1300   '�ŷ��ݾ�
         .ColWidth(9) = 1000   '��������
         .ColWidth(10) = 2000  '������ȣ
         .ColWidth(11) = 4500  '����
         .ColWidth(12) = 1000  '������ڵ�
         .ColWidth(13) = 900   '����ڸ�
         .ColWidth(14) = 1000  '�ð�(�и���)
         
         .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = False
         .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
         .TextMatrix(0, 0) = "KEY"         'H
         .TextMatrix(0, 1) = "�ŷ�����"
         .TextMatrix(0, 2) = "�ŷ��ð�"
         .TextMatrix(0, 3) = "����"        'H
         .TextMatrix(0, 4) = "����"        '1.����, 2.����
         .TextMatrix(0, 5) = "��ü�ڵ�"  'H
         .TextMatrix(0, 6) = "��ü��"
         .TextMatrix(0, 7) = "����"
         .TextMatrix(0, 8) = "�ŷ��ݾ�"
         .TextMatrix(0, 9) = "��������"
         .TextMatrix(0, 10) = "������ȣ"
         .TextMatrix(0, 11) = "����"
         .TextMatrix(0, 12) = "������ڵ�" 'H
         .TextMatrix(0, 13) = "����ڸ�"
         .TextMatrix(0, 14) = "�ð�"
         .ColHidden(0) = True: .ColHidden(3) = True: .ColHidden(5) = True
         .ColHidden(12) = True: .ColHidden(14) = True
         For lngC = 0 To .Cols - 1
             Select Case lngC
                    Case 0, 6, 10, 11
                         .ColAlignment(lngC) = flexAlignLeftCenter
                    Case 1, 2, 3, 4, 5, 7, 9, 12, 13, 14
                         .ColAlignment(lngC) = flexAlignCenterCenter
                    Case Else
                         .ColAlignment(lngC) = flexAlignRightCenter
             End Select
         Next lngC
         For lngC = 0 To .Cols - 1
             Select Case lngC
                    Case 8
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
    If dtpFDate > dtpTDate Then
       dtpFDate.SetFocus
       Exit Sub
    End If
    lblTotMny.Caption = "0" '��ü�ݾ�
    vsfg1.Rows = 1
    With vsfg1
         '�˻����� ��ü
         strWhere = "T1.������ڵ� = '" & PB_regUserinfoU.UserBranchCode & "' "
         If chkTotal.Value = 0 Then '�Ǻ� ��ȸ
            If Len(Text1(0).Text) > 0 Then
               strWhere = strWhere & "AND " & IIf(optJSGbn(0).Value = True, "T1.����ó�ڵ�", "T1.����ó�ڵ�") & " = '" & Trim(Text1(0).Text) & "' "
            Else
               Text1(0).SetFocus
               Exit Sub
            End If
         End If
         If optJSGbn(0).Value = True Then
            strWhere = strWhere & "AND T1.�����ޱ��������� BETWEEN '" & DTOS(dtpFDate.Value) & "' AND '" & DTOS(dtpTDate.Value) & "' "
         Else
            strWhere = strWhere & "AND T1.�̼����Ա����� BETWEEN '" & DTOS(dtpFDate.Value) & "' AND '" & DTOS(dtpTDate.Value) & "' "
         End If
    End With
    P_adoRec.CursorLocation = adUseClient
    If optJSGbn(0).Value = True Then
       strSQL = "SELECT T1.������ڵ� AS ������ڵ�, T2.������ AS ������, " _
                     & "T1.����ó�ڵ� AS ��ü�ڵ�, T3.����ó�� AS ��ü��, " _
                     & "T1.�����ޱ��������� AS �ŷ�����, T1.�����ޱ����޽ð� AS �ŷ��ð�, 1 AS �ŷ�����, " _
                     & "������� = CASE WHEN T1.������� = 0 THEN '����' " _
                                     & "WHEN T1.������� = 1 THEN '��ǥ' " _
                                     & "WHEN T1.������� = 2 THEN '����' " _
                                     & "ELSE '��  ��' " _
                                 & "END, " _
                     & "ISNULL(T1.�����ޱ����ޱݾ�,0) AS �ŷ��ݾ�, " _
                     & "�������� = CASE WHEN T1.������� = 0 THEN '' " _
                                     & "WHEN T1.������� = 2 THEN T1.�������� " _
                                & "END, " _
                     & "������ȣ = CASE WHEN T1.������� = 0 THEN '' " _
                                     & "WHEN T1.������� = 2 THEN T1.������ȣ " _
                                & "END, " _
                     & "T1.���� AS ����, T1.������ڵ�, T4.����ڸ� " _
                & "FROM �����ޱݳ��� T1 " _
                & "LEFT JOIN ����� T2 ON T2.������ڵ� = T1.������ڵ� " _
                & "LEFT JOIN ����ó T3 ON T3.������ڵ� = T1.������ڵ� AND T3.����ó�ڵ� = T1.����ó�ڵ� " _
                & "LEFT JOIN ����� T4 " _
                       & "ON T4.������ڵ� = T1.������ڵ� AND T4.������ڵ�= T1.������ڵ� " _
               & "WHERE " & strWhere & " " _
               & "ORDER BY T1.������ڵ�, T1.�ŷ�����, T1.�ŷ��ð�, T1.�ŷ�����, T1.������� "
    Else
       strSQL = "SELECT T1.������ڵ� AS ������ڵ�, T2.������ AS ������, " _
                     & "T1.����ó�ڵ� AS ��ü�ڵ�, T3.����ó�� AS ��ü��, " _
                     & "T1.�̼����Ա����� AS �ŷ�����, T1.�̼����Աݽð� AS �ŷ��ð�, 2 AS �ŷ�����, " _
                     & "������� = CASE WHEN T1.������� = 0 THEN '����' " _
                                     & "WHEN T1.������� = 1 THEN '��ǥ' " _
                                     & "WHEN T1.������� = 2 THEN '����' " _
                                     & "Else '��  ��' " _
                                & "END, " _
                     & "ISNULL(T1.�̼����Աݱݾ�,0) AS �ŷ��ݾ�, " _
                     & "�������� = CASE WHEN T1.������� = 0 THEN '' " _
                                     & "WHEN T1.������� = 2 THEN T1.�������� " _
                                & "END, " _
                     & "������ȣ = CASE WHEN T1.������� = 0 THEN '' " _
                                     & "WHEN T1.������� = 2 THEN T1.������ȣ " _
                                & "END, " _
                     & "T1.���� AS ����, T1.������ڵ�, T4.����ڸ� " _
                & "FROM �̼��ݳ��� T1 " _
                & "LEFT JOIN ����� T2 ON T2.������ڵ� = T1.������ڵ� " _
                & "LEFT JOIN ����ó T3 ON T3.������ڵ� = T1.������ڵ� AND T3.����ó�ڵ� = T1.����ó�ڵ� " _
                & "LEFT JOIN ����� T4 " _
                       & "ON T4.������ڵ� = T1.������ڵ� AND T4.������ڵ�= T1.������ڵ� " _
               & "WHERE " & strWhere & " " _
               & "ORDER BY T1.������ڵ�, T1.�ŷ�����, T1.�ŷ��ð�, T1.�ŷ�����, T1.������� "
    End If
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
               .TextMatrix(lngR, 0) = P_adoRec("������ڵ�") & "-" & P_adoRec("��ü�ڵ�") & "-" _
                                    & P_adoRec("�ŷ�����") & "-" & P_adoRec("�������") & "-" _
                                    & P_adoRec("�ŷ�����") & "-" & P_adoRec("�ŷ��ð�") & "-"
               .Cell(flexcpData, lngR, 0, lngR, 0) = Trim(.TextMatrix(lngR, 0)) 'FindRow ����� ����
               .TextMatrix(lngR, 1) = Format(P_adoRec("�ŷ�����"), "0000-00-00")
               .TextMatrix(lngR, 2) = Format(Mid(P_adoRec("�ŷ��ð�"), 1, 6), "00:00:00")
               .TextMatrix(lngR, 3) = P_adoRec("�ŷ�����")
               If P_adoRec("�ŷ�����") = 1 Then
                  .TextMatrix(lngR, 4) = "����"
                  .Cell(flexcpForeColor, lngR, 4, lngR, 4) = vbBlue
                  '.Cell(flexcpForeColor, lngR, 8, lngR, 8) = vbBlue
               ElseIf _
                  P_adoRec("�ŷ�����") = 2 Then
                  .Cell(flexcpForeColor, lngR, 4, lngR, 4) = vbRed
                  '.Cell(flexcpForeColor, lngR, 8, lngR, 8) = vbRed
                  .TextMatrix(lngR, 4) = "����"
               End If
               .TextMatrix(lngR, 5) = P_adoRec("��ü�ڵ�")
               .TextMatrix(lngR, 6) = P_adoRec("��ü��")
               .TextMatrix(lngR, 7) = P_adoRec("�������")
               .TextMatrix(lngR, 8) = P_adoRec("�ŷ��ݾ�")
               lblTotMny.Caption = Format(Vals(lblTotMny.Caption) + P_adoRec("�ŷ��ݾ�"), "#,0")
               If P_adoRec("�������") = "����" Then
                  .TextMatrix(lngR, 9) = Format(P_adoRec("��������"), "0000-00-00")
                  .TextMatrix(lngR, 10) = P_adoRec("������ȣ")
               End If
               .TextMatrix(lngR, 11) = P_adoRec("����")
               .TextMatrix(lngR, 12) = P_adoRec("������ڵ�")
               .TextMatrix(lngR, 13) = P_adoRec("����ڸ�")
               .TextMatrix(lngR, 14) = P_adoRec("�ŷ��ð�")
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
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "����/���� �б� ����"
    Unload Me
    Exit Sub
End Sub


VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm�����ۼ�2 
   BorderStyle     =   0  '����
   Caption         =   "�����ۼ�2"
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
      TabIndex        =   6
      Top             =   0
      Width           =   15195
      Begin Crystal.CrystalReport CrystalReport1 
         Left            =   4320
         Top             =   195
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin VB.CommandButton cmdPrint 
         Height          =   375
         Left            =   7980
         Picture         =   "�����ۼ�2.frx":0000
         Style           =   1  '�׷���
         TabIndex        =   13
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdClear 
         Height          =   375
         Left            =   9120
         Picture         =   "�����ۼ�2.frx":0963
         Style           =   1  '�׷���
         TabIndex        =   12
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdExit 
         Height          =   390
         Left            =   13635
         Picture         =   "�����ۼ�2.frx":1308
         Style           =   1  '�׷���
         TabIndex        =   11
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         Height          =   390
         Left            =   12510
         Picture         =   "�����ۼ�2.frx":1C56
         Style           =   1  '�׷���
         TabIndex        =   10
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdSave 
         Height          =   390
         Left            =   11385
         Picture         =   "�����ۼ�2.frx":25DA
         Style           =   1  '�׷���
         TabIndex        =   9
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdFind 
         Height          =   390
         Left            =   10260
         Picture         =   "�����ۼ�2.frx":2E61
         Style           =   1  '�׷���
         TabIndex        =   8
         Top             =   195
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         BackColor       =   &H00008000&
         BorderStyle     =   1  '���� ����
         Caption         =   "�� �� �� ǥ �� ��"
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
         TabIndex        =   7
         Top             =   180
         Width           =   4650
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid vsfg1 
      Height          =   8441
      Left            =   60
      TabIndex        =   2
      Top             =   1644
      Width           =   15195
      _cx             =   26802
      _cy             =   14889
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
      Rows            =   100
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
      Height          =   1035
      Left            =   60
      TabIndex        =   3
      Top             =   600
      Width           =   15195
      Begin VB.TextBox txtTelNo 
         Appearance      =   0  '���
         Enabled         =   0   'False
         Height          =   285
         IMEMode         =   10  '�ѱ� 
         Left            =   7920
         MaxLength       =   50
         TabIndex        =   21
         Top             =   600
         Width           =   2295
      End
      Begin VB.TextBox txtAddress 
         Appearance      =   0  '���
         Enabled         =   0   'False
         Height          =   285
         IMEMode         =   10  '�ѱ� 
         Left            =   1275
         MaxLength       =   50
         TabIndex        =   20
         Top             =   600
         Width           =   5535
      End
      Begin MSComCtl2.DTPicker dtpJ_Date 
         Height          =   270
         Left            =   7920
         TabIndex        =   17
         ToolTipText     =   "���� �۾�(����)���ڸ� ����"
         Top             =   240
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   476
         _Version        =   393216
         CalendarBackColor=   16777215
         Format          =   19857409
         CurrentDate     =   38301
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  '���
         Enabled         =   0   'False
         Height          =   285
         IMEMode         =   10  '�ѱ� 
         Index           =   1
         Left            =   4035
         MaxLength       =   50
         TabIndex        =   1
         Top             =   225
         Width           =   2775
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         Height          =   285
         IMEMode         =   8  '����
         Index           =   0
         Left            =   1275
         MaxLength       =   8
         TabIndex        =   0
         Top             =   225
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "�ּ�"
         Height          =   240
         Index           =   5
         Left            =   75
         TabIndex        =   23
         Top             =   645
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "��ȭ��ȣ"
         Height          =   240
         Index           =   4
         Left            =   7000
         TabIndex        =   22
         Top             =   660
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "��ü�ݾ�"
         Height          =   240
         Index           =   2
         Left            =   9480
         TabIndex        =   19
         Top             =   285
         Width           =   735
      End
      Begin VB.Label lblJDate 
         Caption         =   "( ��������                )"
         Height          =   240
         Left            =   6960
         TabIndex        =   18
         ToolTipText     =   "���� �۾�(����)���ڸ� ����"
         Top             =   285
         Width           =   3135
      End
      Begin VB.Label lblTotMny 
         Alignment       =   1  '������ ����
         Caption         =   "0.00"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   10560
         TabIndex        =   16
         Top             =   285
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "��ü�ݾ�"
         Height          =   240
         Index           =   3
         Left            =   6960
         TabIndex        =   15
         Top             =   285
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "[Home]"
         Height          =   240
         Index           =   15
         Left            =   2520
         TabIndex        =   14
         Top             =   285
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "����ó��"
         Height          =   240
         Index           =   1
         Left            =   3075
         TabIndex        =   5
         Top             =   285
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "����ó�ڵ�"
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   0
         Left            =   75
         TabIndex        =   4
         Top             =   285
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frm�����ۼ�2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'+-------------------------------------------------------------------------------------------------------+
' ���α׷� �� �� : �����ۼ�2
' ���� Control : VideoSoft VSFlexGrid 7.0(OLEDB) = vsflex7.ocx
' ������ Table   : ����ó, ����ó, �������⳻��
' ��  ��  ��  �� : ���־��� �ٷ� ����
'+-------------------------------------------------------------------------------------------------------+
Option Explicit
Private P_blnActived         As Boolean
Private P_adoRec             As New ADODB.Recordset
Private P_adoRecW            As New ADODB.Recordset
Private P_intButton          As Integer
Private P_strFindString2     As String
Private Const PC_intRowCnt   As Integer = 27  '�׸���1�� �� ������ �� ���(FixedRows ����)

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
       Subvsfg1_INIT  '���Գ���
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
                   cmdSave.Enabled = True: cmdDelete.Enabled = False: cmdExit.Enabled = True
              Case Is <= 99 '����, �߰�, ����, ��ȸ, �μ�
                   cmdPrint.Enabled = True: cmdClear.Enabled = True: cmdFind.Enabled = True
                   cmdSave.Enabled = True: cmdDelete.Enabled = True: cmdExit.Enabled = True
              Case Else
                   cmdClear.Enabled = False: cmdFind.Enabled = False: cmdSave.Enabled = False
                   cmdDelete.Enabled = False: cmdExit.Enabled = True
       End Select
       '����(����)����
       cmdFind.Enabled = False: cmdDelete.Enabled = False
       SubOther_FILL
       P_blnActived = True
       Screen.MousePointer = vbDefault
    End If
    Exit Sub
ERROR_FORM_ACTIVATE_CONNECTION_SERVER:
    MsgBox Err.Number & Err.Description & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "�����ۼ�2(�������� ���� ����)"
    Unload Me
    Exit Sub
End Sub

'+--------------------+
'/// OtherControls ///
'+--------------------+
Private Sub dtpHopeDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
       SendKeys "{tab}"
    End If
End Sub
Private Sub dtpSactionDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
       SendKeys "{tab}"
    End If
End Sub
Private Sub cboSactionWay_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
       SendKeys "{tab}"
    End If
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
       PB_strFMCCallFormName = "frm�����ۼ�2"
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

'+-------------+
'/// ����ó ///
'+-------------+
Private Sub Text1_LostFocus(Index As Integer)
Dim strSQL As String
Dim lngR   As Long
    With Text1(Index)
         Select Case Index
                Case 0
                     .Text = UPPER(Trim(.Text))
                     If Len(.Text) < 1 Then
                        Text1(1).Text = "": txtAddress.Text = "": txtTelNo.Text = ""
                     End If
         End Select
    End With
    Exit Sub
ERROR_TABLE_SELECT:
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "TABLE �б� ����"
    Unload Me
    Exit Sub
End Sub

Private Sub chkCash_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{tab}"
    Exit Sub
End Sub

'+-------------------+
'/// �۾����ڼ��� ///
'+-------------------+
Private Sub dtpJ_Date_Change()
    If PB_regUserinfoU.UserClientDate = DTOS(dtpJ_Date.Value) Then
       lblJDate.ForeColor = vbBlack
       With dtpJ_Date
            .CalendarBackColor = vbWhite
            .CalendarForeColor = vbBlack
       End With
    Else
       lblJDate.ForeColor = vbRed
       With dtpJ_Date
            .CalendarBackColor = vbRed
            .CalendarForeColor = vbWhite
       End With
    End If
End Sub
Private Sub dtpJ_Date_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
       vsfg1.SetFocus
    End If
End Sub

'+-----------+
'/// Grid ///
'+-----------+
Private Sub vsfg1_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    With vsfg1
         P_intButton = Button
         If .MouseRow >= .FixedRows Then
            'If Button = vbLeftButton Then
             '  .Select .MouseRow, .MouseCol
             '  .EditCell
            'End If
         End If
    End With
End Sub
Private Sub vsfg1_Click()
Dim strData As String
Dim lngR1   As Long
Dim lngRH1  As Long
Dim lngR2   As Long
Dim lngRR2  As Long
    With vsfg1
         If (.MouseRow >= 0 And .MouseRow < .FixedRows) Then
            .Col = .MouseCol
            '.Cell(flexcpForeColor, 0, 0, 0, .Cols - 1) = vbBlack
            '.Cell(flexcpForeColor, 0, .MouseCol, 0, .MouseCol) = vbRed
            'strData = Trim(.Cell(flexcpData, .Row, 0))
            'Select Case .MouseCol
            '       Case 0
            '            .ColSel = 0
            '            .ColSort(0) = IIf(P_intButton = 1, flexSortGenericAscending, flexSortGenericDescending)
            '            .Sort = flexSortUseColSort
            '       Case 1
            '            .ColSel = 1
            '            .ColSort(0) = flexSortNone
            '            .ColSort(1) = IIf(P_intButton = 1, flexSortGenericAscending, flexSortGenericDescending)
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
Private Sub vsfg1_DblClick()
    With vsfg1
         If .Row >= .FixedRows Then
             vsfg1_KeyDown vbKeyF1, 0  '����ü��˻� OR ����ó�˻����� �ٷ� ��.
         End If
    End With
End Sub
Private Sub vsfg1_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    With vsfg1
         .ToolTipText = ""
         If .MouseRow < .FixedRows Or .MouseCol < 0 Then
            Exit Sub
         End If
         .ToolTipText = .TextMatrix(.MouseRow, .MouseCol)
    End With
End Sub
Private Sub vsfg1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    With vsfg1
         If .Row >= .FixedRows Then
            If Len(.TextMatrix(.Row, 0)) <> 0 Then '0.�����ڵ�, 3.���ַ�, 7.����
               If (.Col = 3) Then   '3. ���ַ�
                  If Button = vbLeftButton Then
                     .Select .Row, .Col
                     .EditCell
                  End If
               ElseIf _
                  (.Col = 7) Then   '7. ����
                  If Not (.ValueMatrix(.Row, 3) = 0 Or Len(.TextMatrix(.Row, 5)) = 0) Then
                     If Button = vbLeftButton Then
                        .Select .Row, .Col
                        .EditCell
                     End If
                  End If
               ElseIf _
                  (.Col = 8 Or .Col = 9) Then   '8.�԰�ܰ�, 9.�԰�ΰ�
                  If Button = vbLeftButton Then
                     .Select .Row, .Col
                     .EditCell
                  End If
               ElseIf _
                  (.Col = 16) Then   '16.����
                  If Button = vbLeftButton Then
                     .Select .Row, .Col
                     .EditCell
                  End If
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
            If Len(.TextMatrix(Row, 0)) <> 0 Then
               If (Col = 3) Then         '����
                  If .TextMatrix(Row, Col) <> .EditText Then
                     If (IsNumeric(.EditText) = False Or _
                        Fix(Vals(.EditText)) < Vals(.EditText) Or IsNumeric(Right(.EditText, 1)) = False) Then
                        Beep
                        Cancel = True
                     Else
                        blnModify = True
                        curTmpMny = .ValueMatrix(Row, 11)
                        .TextMatrix(Row, 11) = Vals(.EditText) * .ValueMatrix(Row, 8)   '�հ�ݾ� = ���� * �ܰ�
                     End If
                  End If
                  If Vals(.EditText) = 0 Or Len(.TextMatrix(Row, 5)) = 0 Then
                     .Cell(flexcpChecked, Row, 7, Row, 7) = flexUnchecked
                  End If
               ElseIf _
                  (Col = 8) Then '8.�԰�ܰ�
                  If .TextMatrix(Row, Col) <> .EditText Then                            '������ ��� �Է±ݾ� �˻�
                     If (IsNumeric(.EditText) = False Or Vals(.EditText) < 0 Or _
                        IsNumeric(Right(.EditText, 1)) = False) Then
                        Beep
                        Cancel = True
                     Else
                        blnModify = True
                        .TextMatrix(Row, 9) = Fix(Vals(.EditText) * (PB_curVatRate))
                        curTmpMny = .ValueMatrix(Row, 11)
                        .TextMatrix(Row, 10) = Vals(.EditText) + .ValueMatrix(Row, 9)
                        .TextMatrix(Row, 11) = .ValueMatrix(Row, 3) * Vals(.EditText)
                     End If
                  End If
               ElseIf _
                  (Col = 9) Then '9.�԰�ΰ�
                  If .TextMatrix(Row, Col) <> .EditText Then                            '������ ��� �Է±ݾ� �˻�
                     If (IsNumeric(.EditText) = False Or Vals(.EditText) < 0 Or _
                        Fix(Vals(.EditText)) < Vals(.EditText) Or IsNumeric(Right(.EditText, 1)) = False) Or _
                        (Vals(.EditText) > .ValueMatrix(Row, 8)) Then
                        Beep
                        Cancel = True
                     Else
                        blnModify = True
                        curTmpMny = .ValueMatrix(Row, 11)
                        .TextMatrix(Row, 10) = .ValueMatrix(Row, 8) + Vals(.EditText)
                        .TextMatrix(Row, 11) = .ValueMatrix(Row, 3) * .ValueMatrix(Row, 8)
                     End If
                  End If
               ElseIf _
                  (Col = 16) Then '���� ���� �˻�
                  If .TextMatrix(Row, Col) <> .EditText Then
                     If Not (LenH(Trim(.EditText)) <= 50) Then
                        Beep
                        Cancel = True
                     Else
                        blnModify = True
                     End If
                  End If
               End If
            End If
            '����ǥ�� + �ݾ�����
            If blnModify = True Then
               Select Case Col
                      Case 3, 8, 9
                           lblTotMny.Caption = Format(Vals(lblTotMny.Caption) - curTmpMny + .ValueMatrix(Row, 11), "#,#.00")
                      Case Else
                      
               End Select
            End If
         End If
    End With
End Sub
Private Sub vsfg1_EnterCell()
Dim lngR As Long
Dim lngC As Long
    With vsfg1
         .Editable = flexEDNone
         If .Row >= .FixedRows Then
             Select Case .Col
                    Case 3, 8, 16
                         .Editable = flexEDKbdMouse
                         vsfg1_MouseDown vbLeftButton, 0, 0, 0
             End Select
         End If
    End With
End Sub
Private Sub vsfg1_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    With vsfg1
         If KeyCode = vbKeyReturn Then
            If Col = 3 Then
               .Col = 8
            ElseIf _
               Col = 8 Then
               .Col = 16
            ElseIf _
               Col = 16 And Row < (.Rows - 1) Then
               .Col = 3: .Row = .Row + 1
               If .Row >= PC_intRowCnt Then
                  .TopRow = .TopRow + 1
               End If
            End If
         End If
    End With
End Sub
Private Sub vsfg1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim lngR     As Long
Dim blnDupOK As Boolean
Dim intRetVal As Integer
    With vsfg1
         If .Row >= .FixedRows Then     '�����ü��˻�
            If KeyCode = vbKeyF2 And (Len(Text1(0).Text) > 0) And (Len(.TextMatrix(.Row, 0)) > 0) Then
               PB_strFMCCallFormName = "frm�����ۼ�2"
               PB_strMaterialsCode = .TextMatrix(.Row, 0)
               PB_strMaterialsName = .TextMatrix(.Row, 1)
               If Len(Trim(Text1(0).Text)) = 0 Then
                  PB_strSupplierCode = ""
               Else
                  PB_strSupplierCode = Trim(Text1(0).Text)
               End If
               frm�����ü��˻�.Show vbModal
               If Len(PB_strMaterialsCode) <> 0 Then
                  PB_strFMCCallFormName = "": PB_strMaterialsCode = "": PB_strMaterialsName = ""
               End If
            End If
         End If
    End With
    With vsfg1
         If .Row >= .FixedRows Then
            If KeyCode = vbKeyF1 And Len(Trim(Text1(0).Text)) > 0 Then  '����ü��˻�
               PB_strFMCCallFormName = "frm�����ۼ�2"
               PB_strMaterialsCode = .TextMatrix(.Row, 0)
               PB_strMaterialsName = .TextMatrix(.Row, 1)
               If Len(Trim(Text1(0).Text)) = 0 Then
                  PB_strSupplierCode = ""
               Else
                  PB_strSupplierCode = Trim(Text1(0).Text)
               End If
               frm����ü��˻�.Show vbModal
               If Len(PB_strMaterialsCode) <> 0 Then
                  PB_strFMCCallFormName = "": PB_strMaterialsCode = "": PB_strMaterialsName = ""
               End If
            ElseIf _
               KeyCode = vbKeyDelete And .Col = 6 Then
               .TextMatrix(.Row, 5) = "": .TextMatrix(.Row, 6) = ""
               .Cell(flexcpChecked, .Row, 7, .Row, 7) = flexUnchecked
            ElseIf _
               KeyCode = vbKeyDelete And (.Col <> 6) And (Len(.TextMatrix(.Row, 0)) > 0) Then 'And (.MouseRow > 0) Then
               intRetVal = MsgBox("�Է��� ���Գ����� �����Ͻðڽ��ϱ� ?", vbQuestion + vbYesNo + vbDefaultButton2, "���Գ�������")
               If intRetVal = vbYes Then
                  lblTotMny.Caption = Format(Vals(lblTotMny.Caption) - .ValueMatrix(.Row, 11), "#,#.00") '��ü�ݾ׿��� ����
                  .Cell(flexcpText, .Row, 0, .Row, .Cols - 1) = ""
               End If
            End If
         End If
    End With
End Sub

'+-----------+
'/// ��� ///
'+-----------+
Private Sub cmdPrint_Click()
    If (vsfg1.Rows) = 1 Then
       Exit Sub
    End If
End Sub
'+-----------+
'/// ��ȸ ///
'+-----------+
Private Sub cmdFind_Click()
    Screen.MousePointer = vbHourglass
    cmdFind.Enabled = False
    cmdFind.Enabled = True
    Screen.MousePointer = vbDefault
End Sub
'+-----------+
'/// ���� ///
'+-----------+
Private Sub cmdSave_Click()
Dim strSQL         As String
Dim lngR           As Long
Dim lngRR          As Long
Dim lngRRR         As Long
Dim lngC           As Long
Dim blnOK          As Boolean
Dim intRetVal      As Integer
Dim lngChkCnt      As Long
Dim lngDelCntS     As Long
Dim lngDelCntE     As Long
Dim strServerTime  As String
Dim strTime        As String
Dim strHH          As String
Dim strMM          As String
Dim strSS          As String
Dim strMS          As String
Dim lngLogCnt      As Long    '�α�ī����
Dim intChkCash     As Integer '1.���ݸ���
Dim strJ_Date      As String  '�۾�����
    If (vsfg1.Rows) = 1 Then
       Exit Sub
    End If
    If Len(Text1(0).Text) < 1 Then '����ó�ڵ�
       Text1(0).SetFocus
       Exit Sub
    End If
    With vsfg1
         For lngR = 1 To .Rows - 1
             If Len(.TextMatrix(lngR, 0)) > 0 Then 'And .ValueMatrix(lngR, 3) <> 0 Then
                lngChkCnt = lngChkCnt + 1
             End If
         Next lngR
         If lngChkCnt = 0 Then
            Exit Sub
         End If
    End With
    intRetVal = MsgBox("�Էµ� �ڷ�" & lngChkCnt & "(��)�� �����Ͻðڽ��ϱ� ?", vbQuestion + vbYesNo + vbDefaultButton1, "�ڷ� ����")
    If intRetVal = vbNo Then
       vsfg1.SetFocus
       Exit Sub
    End If
    intRetVal = MsgBox("���ݸ����� �Ͻðڽ��ϱ� ?", vbQuestion + vbYesNoCancel + vbDefaultButton2, "���ݸ���")
    If intRetVal = vbYes Then
       intChkCash = 1
    ElseIf _
       intRetVal = vbCancel Then
       vsfg1.SetFocus
       Exit Sub
    End If
    '�۾����� ���ϱ�
    strJ_Date = DTOS(dtpJ_Date.Value)
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
    strTime = strServerTime
    '�ŷ���ȣ ���ϱ�
    PB_adoCnnSQL.BeginTrans
    P_adoRec.CursorLocation = adUseClient
    strSQL = "spLogCounter '�������⳻��', '" & PB_regUserinfoU.UserBranchCode + strJ_Date + "1" & "', " _
                            & "0, 0, '" & PB_regUserinfoU.UserServerDate & "', '" & PB_regUserinfoU.UserCode & "' "
    On Error GoTo ERROR_STORED_PROCEDURE
    P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
    lngLogCnt = P_adoRec(0)
    P_adoRec.Close
    
    lngChkCnt = 0
    With vsfg1
         For lngR = 1 To .Rows - 1
             If Len(.TextMatrix(lngR, 0)) > 0 Then 'And .ValueMatrix(lngR, 3) <> 0 Then
                '�������⳻��
                lngChkCnt = lngChkCnt + 1
                If lngChkCnt = 1 Then
                   strTime = strServerTime
                Else
                   strTime = Format((Val(strTime) + 10000), "000000000")
                   strHH = Mid(strTime, 1, 2): strMM = Mid(strTime, 3, 2): strSS = Mid(strTime, 5, 2): strMS = Mid(strTime, 7, 3)
                   If Val(strMS) > 999 Then
                      strMS = Format(0, "000")
                      strSS = Format(Val(strMM) + 1, "00")
                   End If
                   If Val(strSS) > 59 Then
                      strSS = Format(Val(strSS) - 60, "00")
                      strMM = Format(Val(strMM) + 1, "00")
                   End If
                   If Val(strMM) > 59 Then
                      strMM = Format(Val(strMM) - 60, "00")
                      strHH = Format(Val(strHH) + 1, "00")
                   End If
                   strTime = strHH & strMM & strSS & strMS
                End If
                strSQL = "INSERT INTO �������⳻��(������ڵ�, �з��ڵ�, " _
                                                & "�����ڵ�, �������, " _
                                                & "���������, �����ð�, " _
                                                & "�԰����, �԰�ܰ�, " _
                                                & "�԰�ΰ�, ������, " _
                                                & "���ܰ�, ���ΰ�, " _
                                                & "����ó�ڵ�, ����ó�ڵ�, " _
                                                & "�������������, ���۱���, " _
                                                & "�߰�����, �߰߹�ȣ, �ŷ�����, �ŷ���ȣ, " _
                                                & "��꼭���࿩��, ���ݱ���, ��������, ����, �ۼ��⵵, å��ȣ, �Ϸù�ȣ, " _
                                                & "��뱸��, ��������, " _
                                                & "������ڵ�, ����̵�������ڵ�) VALUES( " _
                             & "'" & PB_regUserinfoU.UserBranchCode & "', '" & Mid(.TextMatrix(lngR, 0), 1, 2) & "', " _
                             & "'" & Mid(.TextMatrix(lngR, 0), 3) & "', 1, " _
                             & "'" & strJ_Date & "', '" & strTime & "', " _
                             & "" & .ValueMatrix(lngR, 3) & ", " & .ValueMatrix(lngR, 8) & ", " _
                             & "" & .ValueMatrix(lngR, 9) & ", 0, " _
                             & "" & .ValueMatrix(lngR, 12) & ", " & .ValueMatrix(lngR, 13) & ", " _
                             & "'" & Trim(Text1(0).Text) & "' , '" & .TextMatrix(lngR, 5) & "', " _
                             & "'" & strJ_Date & "'," & IIf(.Cell(flexcpChecked, lngR, 7, lngR, 7) = flexUnchecked, 0, 1) & ", " _
                             & "'', 0, " _
                             & "'" & strJ_Date & "' , " & lngLogCnt & ", " _
                             & "0, " & intChkCash & ", 0, '" & .TextMatrix(lngR, 16) & "', '', 0, 0, " _
                             & "0, '" & PB_regUserinfoU.UserServerDate & "', " _
                             & "'" & PB_regUserinfoU.UserCode & "', '' ) "
                On Error GoTo ERROR_TABLE_INSERT
                PB_adoCnnSQL.Execute strSQL
                '������������ڰ���(������ڵ�, �з��ڵ�, �����ڵ�, �������)
                strSQL = "sp����������������ڰ��� '" & PB_regUserinfoU.UserBranchCode & "', " _
                       & "'" & Mid(.TextMatrix(lngR, 0), 1, 2) & "', '" & Mid(.TextMatrix(lngR, 0), 3) & "', 1 "
                On Error GoTo ERROR_STORED_PROCEDURE
                PB_adoCnnSQL.Execute strSQL
                '���������ܰ�����(������ڵ�, �з��ڵ�, �����ڵ�, �������, ��ü�ڵ�, �ܰ�, �ŷ�����)
                If .ValueMatrix(lngR, 8) > 0 And PB_intIAutoPriceGbn = 1 Then
                   strSQL = "sp���������ܰ����� '" & PB_regUserinfoU.UserBranchCode & "', " _
                          & "'" & Mid(.TextMatrix(lngR, 0), 1, 2) & "', '" & Mid(.TextMatrix(lngR, 0), 3) & "', 1, " _
                          & "'" & Trim(Text1(0).Text) & "', " _
                          & "" & .ValueMatrix(lngR, 8) & ", '" & strJ_Date & "' "
                   On Error GoTo ERROR_STORED_PROCEDURE
                   PB_adoCnnSQL.Execute strSQL
                End If
             End If
         Next lngR
    End With
    PB_adoCnnSQL.CommitTrans
    SubClearText
    Text1(0).SetFocus
    Screen.MousePointer = vbDefault
    cmdSave.Enabled = True
    'If chkPrint = 1 Then '������ ���
       '���
    'End If
    Exit Sub
ERROR_FORM_ACTIVATE_CONNECTION_SERVER:
    MsgBox Err.Number & Err.Description & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "(�������� ���� ����)"
    Unload Me
    Exit Sub
ERROR_TABLE_SELECT:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "�˻� ����"
    Unload Me
    Exit Sub
ERROR_TABLE_INSERT:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "�߰� ����"
    Unload Me
    Exit Sub
ERROR_TABLE_UPDATE:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "���� ����"
    Unload Me
    Exit Sub
ERROR_TABLE_DELETE:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "���� ����"
    Unload Me
    Exit Sub
ERROR_STORED_PROCEDURE:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "���� ���� ����"
    Unload Me
    Exit Sub
End Sub

'+-----------+
'/// �߰� ///
'+-----------+
Private Sub cmdClear_Click()
Dim strSQL As String
    vsfg1.Row = 0
    SubClearText
    Text1(Text1.LBound).Enabled = True
    Text1(Text1.LBound + 1).Enabled = False
    Text1(Text1.LBound).SetFocus
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
    Set frm�����ۼ�2 = Nothing
End Sub

'+--------------------+
'/// Sub Procedure ///
'+--------------------+
Private Sub SubOther_FILL()
Dim strSQL        As String
Dim intIndex      As Integer
    dtpJ_Date.Value = Format(PB_regUserinfoU.UserClientDate, "0000-00-00")
    Exit Sub
ERROR_TABLE_SELECT:
    MsgBox Err.Number & Err.Description & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "TABLE �б� ����"
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
    txtAddress.Text = "": txtTelNo.Text = ""
    lblTotMny.Caption = "0.00"
    With vsfg1
         .Rows = 1: .Rows = 101
         .Row = 1: .Col = 3
         .TopRow = 1: .LeftCol = 3
         .Cell(flexcpChecked, 1, 7, .Rows - 1, 7) = flexUnchecked
         .Cell(flexcpText, 1, 7, .Rows - 1, 7) = "�� ��"
    End With
End Sub
'+----------------------------------+
'/// VsFlexGrid(vsfg1) �ʱ�ȭ ///
'+----------------------------------+
Private Sub Subvsfg1_INIT()
Dim lngR    As Long
Dim lngC    As Long
    With vsfg1              'Rows 101, Cols 17, RowHeightMax(Min) 300
         .AllowBigSelection = False
         .AllowSelection = False
         .AllowUserResizing = flexResizeColumns
         .BackColorBkg = &H8000000F
         .BackColorSel = &H8000&
         .ExtendLastCol = True
         .FocusRect = flexFocusHeavy
         .ScrollBars = flexScrollBarBoth
         .ScrollTrack = True
         .SelectionMode = flexSelectionByRow
         .FixedRows = 1
         .FixedCols = 3
         .Rows = 101
         .Cols = 17
         .RowHeightMax = 300
         .RowHeightMin = 300
         .ColWidth(0) = 1900   '�����ڵ�(�з��ڵ�+�����ڵ�)
         .ColWidth(1) = 2500   '�����
         .ColWidth(2) = 2200   '����԰�
         .ColWidth(3) = 1000   '����
         .ColWidth(4) = 800    '�������
         .ColWidth(5) = 1200   '����ó�ڵ�   'H
         .ColWidth(6) = 2500   '����ó��     'H
         .ColWidth(7) = 800    '����
         .ColWidth(8) = 1500   '�԰�ܰ�
         .ColWidth(9) = 1300   '�԰�ΰ�     'H
         .ColWidth(10) = 1500  '�԰���(�ܰ� + �ΰ�)
         .ColWidth(11) = 2000  '�԰�ݾ�(���ַ� * �԰���)
         .ColWidth(12) = 1500  '���ܰ�     'H
         .ColWidth(13) = 1300  '���ΰ�     'H
         .ColWidth(14) = 1500  '�����(�ܰ�+�ΰ�) 'H
         .ColWidth(15) = 2000  '���ݾ�(���ַ� * �����) 'H
         .ColWidth(16) = 4500  '����
         
         .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = False
         .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
         .TextMatrix(0, 0) = "�ڵ�"
         .TextMatrix(0, 1) = "ǰ��"
         .TextMatrix(0, 2) = "�԰�"
         .TextMatrix(0, 3) = "����"
         .TextMatrix(0, 4) = "����"          '���Դ���
         .TextMatrix(0, 5) = "����ó�ڵ�"    'H
         .TextMatrix(0, 6) = "����ó��"      'H
         .TextMatrix(0, 7) = "����"          'H
         .TextMatrix(0, 8) = "���Դܰ�"
         .TextMatrix(0, 9) = "���Ժΰ�"      'H
         .TextMatrix(0, 10) = "���԰���"     'H
         .TextMatrix(0, 11) = "���Աݾ�"
         .TextMatrix(0, 12) = "����ܰ�"     'H
         .TextMatrix(0, 13) = "����ΰ�"     'H
         .TextMatrix(0, 14) = "����ܰ�"     'H
         .TextMatrix(0, 15) = "����ݾ�"     'H
         .TextMatrix(0, 16) = "����"
         .ColHidden(5) = True: .ColHidden(6) = True: .ColHidden(7) = True
         .ColHidden(9) = True: .ColHidden(10) = True
         .ColHidden(12) = True: .ColHidden(13) = True: .ColHidden(14) = True: .ColHidden(15) = True
         .ColFormat(3) = "#,#"
         For lngC = 8 To 15
             .ColFormat(lngC) = "#,#.00"
         Next lngC
         For lngC = 0 To .Cols - 1
             Select Case lngC
                    Case 0, 1, 2, 4, 6, 7, 16
                         .ColAlignment(lngC) = flexAlignLeftCenter
                    Case 5
                         .ColAlignment(lngC) = flexAlignCenterCenter
                    Case Else
                         .ColAlignment(lngC) = flexAlignRightCenter
             End Select
         Next lngC
         '.MergeCells = flexMergeRestrictRows 'flexMergeFixedOnly
         '.MergeRow(0) = True
         'For lngC = 0 To 2
         '    .MergeCol(lngC) = True
         'Next lngC
         .Cell(flexcpChecked, 1, 7, .Rows - 1, 7) = flexUnchecked
         .Cell(flexcpText, 1, 7, .Rows - 1, 7) = "�� ��"
         .Cell(flexcpAlignment, 1, 7, .Rows - 1, 7) = flexAlignLeftCenter
         
         vsfg1_EnterCell
    End With
End Sub


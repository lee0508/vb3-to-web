VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Begin VB.Form frm����ü��˻� 
   Appearance      =   0  '���
   BackColor       =   &H00008000&
   BorderStyle     =   1  '���� ����
   Caption         =   "���� �ü� �˻�"
   ClientHeight    =   4620
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13320
   BeginProperty Font 
      Name            =   "����ü"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "����ü��˻�.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   13320
   StartUpPosition =   2  'ȭ�� ���
   Begin VB.Frame Frame1 
      BackColor       =   &H00008000&
      Height          =   1095
      Left            =   120
      TabIndex        =   8
      Top             =   0
      Width           =   13095
      Begin VB.TextBox txtUnit 
         Appearance      =   0  '���
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   1  '�Է� ���� ����
         Left            =   1320
         MaxLength       =   20
         TabIndex        =   3
         Top             =   600
         Width           =   1215
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "�߰�(&A)"
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   11160
         TabIndex        =   7
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox txtSize 
         Appearance      =   0  '���
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8520
         MaxLength       =   30
         TabIndex        =   2
         Top             =   240
         Width           =   2355
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "��ȸ(&F)"
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   11160
         TabIndex        =   5
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox txtBarCode 
         Appearance      =   0  '���
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   8  '����
         Left            =   4680
         MaxLength       =   8
         TabIndex        =   4
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox txtName 
         Appearance      =   0  '���
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4680
         MaxLength       =   30
         TabIndex        =   1
         Top             =   240
         Width           =   2475
      End
      Begin VB.TextBox txtCode 
         Appearance      =   0  '���
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   8  '����
         Left            =   1320
         MaxLength       =   18
         TabIndex        =   0
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         BackStyle       =   0  '����
         Caption         =   "�԰�"
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   12
         Left            =   7440
         TabIndex        =   14
         Top             =   300
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         BackStyle       =   0  '����
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   3
         Left            =   240
         TabIndex        =   13
         Top             =   660
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         BackStyle       =   0  '����
         Caption         =   "���ڵ�"
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   8
         Left            =   3360
         TabIndex        =   12
         Top             =   645
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         BackStyle       =   0  '����
         Caption         =   "ǰ��"
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   7
         Left            =   3600
         TabIndex        =   11
         Top             =   300
         Width           =   855
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '����
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   2
         Left            =   3240
         TabIndex        =   10
         Top             =   300
         Width           =   375
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         BackStyle       =   0  '����
         Caption         =   "�ڵ�"
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   0
         Left            =   0
         TabIndex        =   9
         Top             =   300
         Width           =   1095
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid vsfg1 
      Height          =   3345
      Left            =   120
      TabIndex        =   6
      Top             =   1155
      Width           =   13095
      _cx             =   23098
      _cy             =   5900
      _ConvInfo       =   1
      Appearance      =   0
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
      BackColor       =   16777215
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483638
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483638
      BackColorAlternate=   16777215
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
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
End
Attribute VB_Name = "frm����ü��˻�"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'+-------------------------------------------------------------------------------------------------------+
' ���α׷� �� �� : ����ü��˻�
' ���� Control :
' ������ Table   : ����, �������
' ��  ��  ��  �� :
'
'+-------------------------------------------------------------------------------------------------------+
Private P_blnActived       As Boolean
Private P_intFindGbn       As Integer      '0.�����˻�, 1.�ڵ�, 2.�̸�(��), 3.�ڵ� + �̸�(��)
Private P_intButton        As Integer
Private P_strFindString2   As String
Private P_strFirstCode     As String
Private P_intIOGbn         As Integer      '������� : 1.�԰�, 2.���
Private P_blnSelect        As Boolean
Private P_adoRec           As New ADODB.Recordset
Private Const PC_intRowCnt As Integer = 10  '�׸��� �� ������ �� ���(FixedRows ����)

'+--------------------------------+
'/// LOAD FORM ( �ѹ��� ���� ) ///
'+--------------------------------+
Private Sub Form_Load()
    P_blnActived = False
    P_strFirstCode = PB_strMaterialsCode '���� ã������ �����ڵ�
    txtCode.Text = PB_strMaterialsCode: txtName.Text = PB_strMaterialsName
    If (PB_strFMCCallFormName = "frm���ּ��ۼ�") Then
       With frm���ּ��ۼ�
            'txtSize.Text = .vsfg1.TextMatrix(.vsfg1.Row, 2)
            'txtUnit.Text = .vsfg1.TextMatrix(.vsfg1.Row, 4)
       End With
    End If
    If (PB_strFMCCallFormName = "frm���ּ�����") Then
       With frm���ּ�����
            'txtSize.Text = .vsfg2.TextMatrix(.vsfg2.Row, 15)
            'txtUnit.Text = .vsfg2.TextMatrix(.vsfg2.Row, 17)
       End With
    End If
    If (PB_strFMCCallFormName = "frm�����ۼ�1") Then
       With frm�����ۼ�1
            'txtSize.Text = .vsfg2.TextMatrix(.vsfg2.Row, 15)
            'txtUnit.Text = .vsfg2.TextMatrix(.vsfg2.Row, 17)
       End With
    End If
    If (PB_strFMCCallFormName = "frm�����ۼ�2") Then
       With frm�����ۼ�2
            'txtSize.Text = .vsfg1.TextMatrix(.vsfg1.Row, 2)
            'txtUnit.Text = .vsfg1.TextMatrix(.vsfg1.Row, 4)
       End With
    End If
    If (PB_strFMCCallFormName = "frm���Լ���") Then
       With frm���Լ���
            'txtSize.Text = .vsfg2.TextMatrix(.vsfg2.Row, 15)
            'txtUnit.Text = .vsfg2.TextMatrix(.vsfg2.Row, 17)
       End With
    End If
    
    If (PB_strFMCCallFormName = "frm�������ۼ�") Then
       With frm�������ۼ�
            'txtSize.Text = .vsfg1.TextMatrix(.vsfg1.Row, 2)
            'txtUnit.Text = .vsfg1.TextMatrix(.vsfg1.Row, 4)
       End With
    End If
    If (PB_strFMCCallFormName = "frm����������") Then
       With frm����������
            'txtSize.Text = .vsfg2.TextMatrix(.vsfg2.Row, 15)
            'txtUnit.Text = .vsfg2.TextMatrix(.vsfg2.Row, 17)
       End With
    End If
    If (PB_strFMCCallFormName = "frm�����ۼ�1") Then
       With frm�����ۼ�1
            'txtSize.Text = .vsfg2.TextMatrix(.vsfg2.Row, 15)
            'txtUnit.Text = .vsfg2.TextMatrix(.vsfg2.Row, 17)
       End With
    End If
    If (PB_strFMCCallFormName = "frm�����ۼ�2") Then
       With frm�����ۼ�2
            'txtSize.Text = .vsfg1.TextMatrix(.vsfg1.Row, 2)
            'txtUnit.Text = .vsfg1.TextMatrix(.vsfg1.Row, 4)
       End With
    End If
    If (PB_strFMCCallFormName = "frm�������") Then
       With frm�������
            'txtSize.Text = .vsfg2.TextMatrix(.vsfg2.Row, 15)
            'txtUnit.Text = .vsfg2.TextMatrix(.vsfg2.Row, 17)
       End With
    End If
    If Len(PB_strMaterialsCode) = 0 And Len(PB_strMaterialsName) = 0 Then
       P_intFindGbn = 0    '�����˻�
    ElseIf _
       Len(PB_strMaterialsCode) <> 0 And Len(PB_strMaterialsName) = 0 Then
       P_intFindGbn = 1    '�ڵ�θ� �ڵ��˻�
    ElseIf _
       Len(PB_strMaterialsCode) = 0 And Len(PB_strMaterialsName) <> 0 Then
       P_intFindGbn = 2    '�̸�(��)���θ� �ڵ��˻�
    ElseIf _
       Len(PB_strMaterialsCode) <> 0 And Len(PB_strMaterialsName) <> 0 Then
       P_intFindGbn = 3    '�ڵ�� �̸�(��)�� ���ÿ� �ڵ��˻�
    Else
       P_intFindGbn = 0    '�����˻�
    End If
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
              'Case Is <= 10 '��ȸ
              '     txtCode.Enabled = False: txtSupplierCode.Enabled = False: cmdFind.Enabled = True
              'Case Is <= 20 '�μ�, ��ȸ
              '     cmdPrint.Enabled = True: cmdClear.Enabled = False: cmdFind.Enabled = True
              '     cmdSave.Enabled = False: cmdDelete.Enabled = False: cmdExit.Enabled = True
              'Case Is <= 40 '�߰�, ����, �μ�, ��ȸ
              '     cmdPrint.Enabled = True: cmdClear.Enabled = True: cmdFind.Enabled = True
              '     cmdSave.Enabled = True: cmdDelete.Enabled = False: cmdExit.Enabled = True
              'Case Is <= 50 '����, �߰�, ����, ��ȸ, �μ�
              '     cmdPrint.Enabled = True: cmdClear.Enabled = True: cmdFind.Enabled = True
              '     cmdSave.Enabled = True: cmdDelete.Enabled = True: cmdExit.Enabled = True
              'Case Is <= 99 '����, �߰�, ����, ��ȸ, �μ�
              '     cmdPrint.Enabled = True: cmdClear.Enabled = True: cmdFind.Enabled = True
              '     cmdSave.Enabled = True: cmdDelete.Enabled = True: cmdExit.Enabled = True
              'Case Else
              '     cmdClear.Enabled = False: cmdFind.Enabled = False: cmdSave.Enabled = False
              '     cmdDelete.Enabled = False: cmdExit.Enabled = True
       End Select
       SubOther_FILL
       If P_intFindGbn <> 0 Then
          Subvsfg1_FILL
       End If
       P_blnActived = True
       Screen.MousePointer = vbDefault
    End If
    Exit Sub
ERROR_FORM_ACTIVATE_CONNECTION_SERVER:
    MsgBox Err.Number & Err.Description & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "����ü� (�������� ���� ����)"
    Unload Me
    Exit Sub
End Sub

'+--------------+
'/// txtCode ///
'+--------------+
Private Sub txtCode_GotFocus()
    With txtCode
         .SelStart = 0
         .SelLength = Len(Trim(.Text))
    End With
End Sub
Private Sub txtCode_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strSQL     As String
    If KeyCode = vbKeyEscape Then
       SubEscape
    End If
    If KeyCode = vbKeyF1 Then
       'PB_strCallFormName = "frm����ü��˻�"
       'PB_strMaterialsCode = Trim(txtCode.Text)
       'PB_strMaterialsName = Trim(txtName.Text)
       'frm����˻�.Show vbModal
       'If (Len(PB_strMaterialsCode) + Len(PB_strMaterialsName)) = 0 Then '�˻����� ���(ESC)
       'Else
       '   txtCode.Text = PB_strMaterialsCode
       '   txtName.Text = PB_strMaterialsName
       'End If
       'If PB_strMaterialsCode <> "" Then
       '   SendKeys "{tab}"
       'End If
       'PB_strCallFormName = "": PB_strMaterialsCode = "": PB_strMaterialsName = ""
    End If
    txtCode.Text = Trim(txtCode.Text)
    If KeyCode = vbKeyReturn Then
       If IsNumeric(Mid(txtCode.Text, 1, 2)) And Mid(UPPER(txtCode.Text), 3, 4) = "CODE" And Len(Mid(txtCode, 7)) = 0 Then
          SendKeys "{tab}"
       ElseIf _
          Len(Trim(txtCode.Text)) > 1 Then
          cmdFind_Click
          vsfg1.SetFocus
       Else
          SendKeys "{tab}"
       End If
    End If
End Sub

'+--------------+
'/// txtName ///
'+--------------+
Private Sub txtName_GotFocus()
    With txtName
         .SelStart = 0
         .SelLength = Len(Trim(.Text))
    End With
End Sub
Private Sub txtName_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strSQL     As String
    If KeyCode = vbKeyEscape Then
       SubEscape
    End If
    'If KeyCode = vbKeyF1 Then
    '   PB_strCallFormName = "frm����ü��˻�"
    '   PB_strMaterialsCode = Trim(txtCode.Text)
    '   PB_strMaterialsName = Trim(txtName.Text)
    '   frm����˻�.Show vbModal
    '   If (Len(PB_strMaterialsCode) + Len(PB_strMaterialsName)) = 0 Then '�˻����� ���(ESC)
    '   Else
    '      txtCode.Text = PB_strMaterialsCode
    '      txtName.Text = PB_strMaterialsName
    '   End If
    '   If PB_strMaterialsCode <> "" Then
    '      SendKeys "{tab}"
    '   End If
    '   PB_strCallFormName = "": PB_strMaterialsCode = "": PB_strMaterialsName = ""
    'End If
    If KeyCode = vbKeyReturn Then
       SendKeys "{tab}"
    End If
End Sub

'+--------------+
'/// txtSize ///
'+--------------+
Private Sub txtSize_GotFocus()
    With txtSize
         .SelStart = 0
         .SelLength = Len(Trim(.Text))
    End With
End Sub
Private Sub txtSize_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
       txtUnit.SetFocus
    ElseIf _
       KeyCode = vbKeyEscape Then
       SubEscape
    End If
End Sub

'+--------------+
'/// txtUnit ///
'+--------------+
Private Sub txtUnit_GotFocus()
    With txtUnit
         .SelStart = 0
         .SelLength = Len(Trim(.Text))
    End With
End Sub
Private Sub txtUnit_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
       SendKeys "{tab}"
    ElseIf _
       KeyCode = vbKeyEscape Then
       SubEscape
    End If
End Sub

'+----------------+
'/// txtBarCode ///
'+----------------+
Private Sub txtBarCode_GotFocus()
    With txtBarCode
         .SelStart = 0
         .SelLength = Len(Trim(.Text))
    End With
End Sub
Private Sub txtBarCode_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strSQL     As String
    If KeyCode = vbKeyEscape Then
       SubEscape
    End If
    If KeyCode = vbKeyReturn Then
       SendKeys "{tab}"
    End If
End Sub
Private Sub txtBarCode_LostFocus()
    With txtBarCode
         .Text = Trim(txtBarCode.Text)
    End With
End Sub

'+------------+
'/// vsfg1 ///
'+------------+
Private Sub vsfg1_DblClick()
    vsfg1_KeyDown vbKeyReturn, 0
End Sub
Private Sub vsfg1_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
       SubEscape
    End If
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
Private Sub vsfg1_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    P_intButton = Button
End Sub
Private Sub vsfg1_Click()
Dim strData As String
    With vsfg1
         If (.MouseRow >= 0 And .MouseRow < .FixedRows) Then
            .Col = .MouseCol
            .Cell(flexcpForeColor, 0, 0, .FixedRows - 1, .Cols - 1) = vbBlack '.ForeColorSel
            .Cell(flexcpForeColor, .MouseRow, .MouseCol, .MouseRow, .MouseCol) = vbRed
            strData = .TextMatrix(.Row, 6)
            Select Case .MouseCol
                   'Case 3           '(1.��������, 2.����ó��)
                   '     .ColSel = 5
                   '     .ColSort(0) = flexSortNone
                   '     .ColSort(1) = flexSortNone
                   '     .ColSort(2) = flexSortNone
                   '     .ColSort(3) = IIf(P_intButton = 1, flexSortGenericAscending, flexSortGenericDescending)
                   '     .ColSort(4) = flexSortNone
                   '     .ColSort(5) = flexSortGenericAscending
                   '     .Sort = flexSortUseColSort
                   Case Else
                        .Sort = IIf(P_intButton = 1, flexSortGenericAscending, flexSortGenericDescending)
            End Select
            If .FindRow(strData, , 6) > 0 Then
               .Row = .FindRow(strData, , 6)
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
         If .Row < .FixedRows Then
         '   Text1(2).Enabled = True
         Else
         '   Text1(2).Enabled = False
         End If
         If .Row >= .FixedRows Then
            'For lngC = 0 To .Cols - 1
            '    Select Case lngC
            '           Case 0 '�����ڵ�
            '                txtCode.Text = .TextMatrix(.Row, lngC)
            '           Case 1 '�����
            '                txtName.Text = .TextMatrix(.Row, lngC)
            '           Case 7 '�԰�
            '                txtSize.Text = .TextMatrix(.Row, lngC)
            '           'Case 4 '�ָ���ó�ڵ�
            '           '     txtSupplierCode.Text = .TextMatrix(.Row, lngC)
            '           Case 8 '����
            '                txtUnit.Text = .TextMatrix(.Row, lngC)
            '           Case Else
            '    End Select
            'Next lngC
         End If
    End With
End Sub
'+---------------------------------------+
'/// �ü� �˻��� Return
'+---------------------------------------+
Private Sub vsfg1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim SQL         As String
Dim intRetVal   As Integer
Dim lngR        As Long
Dim blnDupOK    As Boolean
Dim varFormName As Variant
    With vsfg1
         If KeyCode = vbKeyReturn Then
            If .Row <= 0 Then
               PB_strMaterialsCode = ""
               PB_strMaterialsName = ""
            Else
               PB_strMaterialsCode = .TextMatrix(.Row, 0)
               PB_strMaterialsName = .TextMatrix(.Row, 1)
            End If
            varFormName = PB_strFMCCallFormName
            If .Row >= .FixedRows Then
               If (PB_strFMCCallFormName = "frm���ּ��ۼ�") Then
                  With frm���ּ��ۼ�
                       'If Trim(.Text1(0).Text) = vsfg1.TextMatrix(vsfg1.Row, 4) Then                '����ó�ڵ� ������
                          'For lngR = 1 To .vsfg1.Rows - 1
                          '    If .vsfg1.Row <> lngR Then
                          '       If vsfg1.TextMatrix(vsfg1.Row, 0) = .vsfg1.TextMatrix(lngR, 0) And _
                          '          .vsfg1.TextMatrix(.vsfg1.Row, 5) = .vsfg1.TextMatrix(lngR, 5) Then '�����ڵ� + ����ó�ڵ�
                          '          blnDupOK = True
                          '          Exit For
                          '       End If
                          '    End If
                          'Next lngR
                          If (blnDupOK = False) Then 'And _
                             '(.vsfg1.TextMatrix(.vsfg1.Row, 0) <> vsfg1.TextMatrix(vsfg1.Row, 0)) Then '�����ڵ� �ٸ���
                             .vsfg1.TextMatrix(.vsfg1.Row, 0) = vsfg1.TextMatrix(vsfg1.Row, 0)      '�����ڵ�
                             .vsfg1.TextMatrix(.vsfg1.Row, 1) = vsfg1.TextMatrix(vsfg1.Row, 1)      '�����
                             .vsfg1.TextMatrix(.vsfg1.Row, 2) = vsfg1.TextMatrix(vsfg1.Row, 7)      '����԰�
                             .vsfg1.TextMatrix(.vsfg1.Row, 4) = vsfg1.TextMatrix(vsfg1.Row, 8)      '�������
                             .vsfg1.TextMatrix(.vsfg1.Row, 8) = vsfg1.ValueMatrix(vsfg1.Row, 10)    '�԰�ܰ�
                             .vsfg1.TextMatrix(.vsfg1.Row, 9) = vsfg1.ValueMatrix(vsfg1.Row, 11)    '�԰�ΰ�
                             .vsfg1.TextMatrix(.vsfg1.Row, 10) = vsfg1.ValueMatrix(vsfg1.Row, 12)   '�԰���
                             '�հ�ݾ� ���(�԰�ݾ��� �ٸ���)
                             If .vsfg1.ValueMatrix(.vsfg1.Row, 11) <> _
                               (.vsfg1.ValueMatrix(.vsfg1.Row, 3) * vsfg1.ValueMatrix(vsfg1.Row, 10)) Then
                                .lblTotMny.Caption = Format(Vals(.lblTotMny.Caption) - .vsfg1.ValueMatrix(.vsfg1.Row, 11) _
                                                   + (.vsfg1.ValueMatrix(.vsfg1.Row, 3) * vsfg1.ValueMatrix(vsfg1.Row, 10)), "#,#.00")
                             End If
                             .vsfg1.TextMatrix(.vsfg1.Row, 11) = .vsfg1.ValueMatrix(.vsfg1.Row, 3) _
                                                               * vsfg1.ValueMatrix(vsfg1.Row, 10)   '�԰�ݾ�
                          
                             .vsfg1.TextMatrix(.vsfg1.Row, 12) = vsfg1.ValueMatrix(vsfg1.Row, 13)   '���ܰ�
                             .vsfg1.TextMatrix(.vsfg1.Row, 13) = vsfg1.ValueMatrix(vsfg1.Row, 14)   '���ΰ�
                             .vsfg1.TextMatrix(.vsfg1.Row, 14) = vsfg1.ValueMatrix(vsfg1.Row, 15)   '�����
                             P_blnSelect = True
                          Else '�ߺ��̸�
                             PB_strFMCCallFormName = "": PB_strMaterialsCode = "": PB_strMaterialsName = ""
                          End If
                       'Else '����ó�ڵ尡 �ٸ���
                       '   PB_strFMCCallFormName = "": PB_strMaterialsCode = "": PB_strMaterialsName = ""
                       'End If
                  End With
               ElseIf _
                  (PB_strFMCCallFormName = "frm���ּ�����") Then
                  With frm���ּ�����
                       'If .vsfg1.TextMatrix(.vsfg1.Row, 5) = vsfg1.TextMatrix(vsfg1.Row, 4) Then                  '����ó�ڵ� ������
                          'For lngR = .vsfg1.ValueMatrix(.vsfg1.Row, 15) To .vsfg1.ValueMatrix(.vsfg1.Row, 16)
                          '    If .vsfg2.Row <> lngR Then
                          '       If vsfg1.TextMatrix(vsfg1.Row, 0) = .vsfg2.TextMatrix(lngR, 4) And _
                          '          .vsfg2.TextMatrix(.vsfg2.Row, 8) = .vsfg2.TextMatrix(lngR, 8) Then '�����ڵ� + ����ó�ڵ�
                          '          blnDupOK = True
                          '          Exit For
                          '       End If
                          '    End If
                          'Next lngR
                          If (blnDupOK = False) And _
                             (.vsfg2.TextMatrix(.vsfg2.Row, 4) <> vsfg1.TextMatrix(vsfg1.Row, 0)) Or _
                             (.vsfg2.ValueMatrix(.vsfg2.Row, 19) <> vsfg1.ValueMatrix(vsfg1.Row, 10)) Then '�����ڵ�, �ܰ� �ٸ���
                             If .vsfg2.TextMatrix(.vsfg2.Row, 4) <> vsfg1.TextMatrix(vsfg1.Row, 0) Then
                                .vsfg2.Cell(flexcpBackColor, .vsfg2.Row, 5, .vsfg2.Row, 5) = vbRed
                                .vsfg2.Cell(flexcpForeColor, .vsfg2.Row, 5, .vsfg2.Row, 5) = vbWhite
                             End If
                             If .vsfg2.ValueMatrix(.vsfg2.Row, 19) <> vsfg1.ValueMatrix(vsfg1.Row, 10) Then
                                .vsfg2.Cell(flexcpBackColor, .vsfg2.Row, 19, .vsfg2.Row, 19) = vbRed
                                .vsfg2.Cell(flexcpForeColor, .vsfg2.Row, 19, .vsfg2.Row, 19) = vbWhite
                             End If
                             If .vsfg2.TextMatrix(.vsfg2.Row, 28) = "" Then
                                .vsfg2.TextMatrix(.vsfg2.Row, 28) = "U"
                             End If
                             .vsfg2.TextMatrix(.vsfg2.Row, 4) = vsfg1.TextMatrix(vsfg1.Row, 0)      '�����ڵ�
                             .vsfg2.TextMatrix(.vsfg2.Row, 5) = vsfg1.TextMatrix(vsfg1.Row, 1)      '�����
                             'Grid Key(������ڵ�+��/��/��/��+���ֹ�ȣ+�����ڵ�+����ó�ڵ�
                             .vsfg2.TextMatrix(.vsfg2.Row, 12) = .vsfg1.TextMatrix(.vsfg1.Row, 1) _
                                       & "-" & .vsfg2.TextMatrix(.vsfg2.Row, 4) & "-" & .vsfg2.TextMatrix(.vsfg2.Row, 8)
                             .vsfg2.TextMatrix(.vsfg2.Row, 15) = vsfg1.TextMatrix(vsfg1.Row, 7)     '����԰�
                             .vsfg2.TextMatrix(.vsfg2.Row, 17) = vsfg1.TextMatrix(vsfg1.Row, 8)     '�������
                             
                             .vsfg2.TextMatrix(.vsfg2.Row, 19) = vsfg1.ValueMatrix(vsfg1.Row, 10)   '�԰�ܰ�
                             .vsfg2.TextMatrix(.vsfg2.Row, 20) = vsfg1.ValueMatrix(vsfg1.Row, 11)   '�԰�ΰ�
                             .vsfg2.TextMatrix(.vsfg2.Row, 21) = vsfg1.ValueMatrix(vsfg1.Row, 12)   '�԰���
                             '�հ�ݾ� ���(�԰�ݾ��� �ٸ���)
                             If .vsfg2.ValueMatrix(.vsfg2.Row, 22) <> _
                               (.vsfg2.ValueMatrix(.vsfg2.Row, 16) * vsfg1.ValueMatrix(vsfg1.Row, 10)) Then
                                .vsfg1.TextMatrix(.vsfg1.Row, 7) = .vsfg1.ValueMatrix(.vsfg1.Row, 7) - .vsfg2.ValueMatrix(.vsfg2.Row, 22) _
                                                                + (.vsfg2.ValueMatrix(.vsfg2.Row, 16) * vsfg1.ValueMatrix(vsfg1.Row, 10))
                                .lblTotMny.Caption = Format(Vals(.lblTotMny.Caption) - .vsfg2.ValueMatrix(.vsfg2.Row, 22) _
                                                   + (.vsfg2.ValueMatrix(.vsfg2.Row, 16) * vsfg1.ValueMatrix(vsfg1.Row, 10)), "#,#.00")
                             End If
                             .vsfg2.TextMatrix(.vsfg2.Row, 22) = .vsfg2.ValueMatrix(.vsfg2.Row, 16) _
                                                               * vsfg1.ValueMatrix(vsfg1.Row, 10)   '�԰�ݾ�
                                                               
                             .vsfg2.TextMatrix(.vsfg2.Row, 23) = vsfg1.ValueMatrix(vsfg1.Row, 13)   '���ܰ�
                             .vsfg2.TextMatrix(.vsfg2.Row, 24) = vsfg1.ValueMatrix(vsfg1.Row, 14)   '���ΰ�
                             .vsfg2.TextMatrix(.vsfg2.Row, 25) = vsfg1.ValueMatrix(vsfg1.Row, 15)   '�����
                             .vsfg2.TextMatrix(.vsfg2.Row, 26) = .vsfg2.ValueMatrix(.vsfg2.Row, 16) _
                                                               * vsfg1.ValueMatrix(vsfg1.Row, 13)   '���ݾ�
                             P_blnSelect = True
                          Else '�ߺ��̸�
                             PB_strFMCCallFormName = "": PB_strMaterialsCode = "": PB_strMaterialsName = ""
                          End If
                       'Else '����ó�ڵ尡 �ٸ���
                       '   PB_strFMCCallFormName = "": PB_strMaterialsCode = "": PB_strMaterialsName = ""
                       'End If
                  End With
               ElseIf _
                  (PB_strFMCCallFormName = "frm�����ۼ�1") Then
                  With frm�����ۼ�1
                       'If .vsfg1.TextMatrix(.vsfg1.Row, 5) = vsfg1.TextMatrix(vsfg1.Row, 4) Then                '����ó�ڵ�
                          'For lngR = .vsfg1.ValueMatrix(.vsfg1.Row, 15) To .vsfg1.ValueMatrix(.vsfg1.Row, 16)
                          '    If .vsfg2.Row <> lngR Then
                          '       If vsfg1.TextMatrix(vsfg1.Row, 0) = .vsfg2.TextMatrix(lngR, 4) And _
                          '          .vsfg2.TextMatrix(.vsfg2.Row, 8) = .vsfg2.TextMatrix(lngR, 8) Then '�����ڵ� + ����ó�ڵ�
                          '          blnDupOK = True
                          '          Exit For
                          '       End If
                          '    End If
                          'Next lngR
                          'if �ߺ� = false And (0.���������ڵ� <> �����ڵ�)           '����
                          If (blnDupOK = False) And _
                             (.vsfg2.TextMatrix(.vsfg2.Row, 4) <> vsfg1.TextMatrix(vsfg1.Row, 0)) Or _
                             (.vsfg2.ValueMatrix(.vsfg2.Row, 19) <> vsfg1.ValueMatrix(vsfg1.Row, 10)) Then '�����ڵ�, �ܰ� �ٸ���
                             If .vsfg2.TextMatrix(.vsfg2.Row, 4) <> vsfg1.TextMatrix(vsfg1.Row, 0) Then
                                .vsfg2.Cell(flexcpBackColor, .vsfg2.Row, 5, .vsfg2.Row, 5) = vbRed
                                .vsfg2.Cell(flexcpForeColor, .vsfg2.Row, 5, .vsfg2.Row, 5) = vbWhite
                             End If
                             If .vsfg2.ValueMatrix(.vsfg2.Row, 19) <> vsfg1.ValueMatrix(vsfg1.Row, 10) Then
                                .vsfg2.Cell(flexcpBackColor, .vsfg2.Row, 19, .vsfg2.Row, 19) = vbRed
                                .vsfg2.Cell(flexcpForeColor, .vsfg2.Row, 19, .vsfg2.Row, 19) = vbWhite
                             End If
                             If .vsfg2.TextMatrix(.vsfg2.Row, 28) = "" Then
                                .vsfg2.TextMatrix(.vsfg2.Row, 28) = "U"
                             End If
                             .vsfg2.TextMatrix(.vsfg2.Row, 4) = vsfg1.TextMatrix(vsfg1.Row, 0)      '�����ڵ�
                             .vsfg2.TextMatrix(.vsfg2.Row, 5) = vsfg1.TextMatrix(vsfg1.Row, 1)      '�����
                             'Grid Key(������ڵ�+��/��/��/��+���ֹ�ȣ+�����ڵ�+����ó�ڵ�
                             .vsfg2.TextMatrix(.vsfg2.Row, 12) = .vsfg1.TextMatrix(.vsfg1.Row, 1) _
                                       & "-" & .vsfg2.TextMatrix(.vsfg2.Row, 4) & "-" & .vsfg2.TextMatrix(.vsfg2.Row, 8)
                             .vsfg2.TextMatrix(.vsfg2.Row, 15) = vsfg1.TextMatrix(vsfg1.Row, 7)     '����԰�
                             .vsfg2.TextMatrix(.vsfg2.Row, 17) = vsfg1.TextMatrix(vsfg1.Row, 8)     '�������
                             
                             .vsfg2.TextMatrix(.vsfg2.Row, 19) = vsfg1.ValueMatrix(vsfg1.Row, 10)   '�԰�ܰ�
                             .vsfg2.TextMatrix(.vsfg2.Row, 20) = vsfg1.ValueMatrix(vsfg1.Row, 11)   '�԰�ΰ�
                             .vsfg2.TextMatrix(.vsfg2.Row, 21) = vsfg1.ValueMatrix(vsfg1.Row, 12)   '�԰���
                             '�հ�ݾ� ���(�԰�ݾ��� �ٸ���)
                             If .vsfg2.ValueMatrix(.vsfg2.Row, 22) <> _
                               (.vsfg2.ValueMatrix(.vsfg2.Row, 16) * vsfg1.ValueMatrix(vsfg1.Row, 10)) Then
                                .vsfg1.TextMatrix(.vsfg1.Row, 7) = .vsfg1.ValueMatrix(.vsfg1.Row, 7) - .vsfg2.ValueMatrix(.vsfg2.Row, 22) _
                                                                + (.vsfg2.ValueMatrix(.vsfg2.Row, 16) * vsfg1.ValueMatrix(vsfg1.Row, 10))
                                .lblTotMny.Caption = Format(Vals(.lblTotMny.Caption) - .vsfg2.ValueMatrix(.vsfg2.Row, 22) _
                                                   + (.vsfg2.ValueMatrix(.vsfg2.Row, 16) * vsfg1.ValueMatrix(vsfg1.Row, 10)), "#,#.00")
                             End If
                             .vsfg2.TextMatrix(.vsfg2.Row, 22) = .vsfg2.ValueMatrix(.vsfg2.Row, 16) _
                                                               * vsfg1.ValueMatrix(vsfg1.Row, 10)   '�԰�ݾ�
                                                               
                             .vsfg2.TextMatrix(.vsfg2.Row, 23) = vsfg1.ValueMatrix(vsfg1.Row, 13)   '���ܰ�
                             .vsfg2.TextMatrix(.vsfg2.Row, 24) = vsfg1.ValueMatrix(vsfg1.Row, 14)   '���ΰ�
                             .vsfg2.TextMatrix(.vsfg2.Row, 25) = vsfg1.ValueMatrix(vsfg1.Row, 15)   '�����
                             .vsfg2.TextMatrix(.vsfg2.Row, 26) = .vsfg2.ValueMatrix(.vsfg2.Row, 16) _
                                                               * vsfg1.ValueMatrix(vsfg1.Row, 13)   '���ݾ�
                             P_blnSelect = True
                          Else '�ߺ��̸�
                             PB_strFMCCallFormName = "": PB_strMaterialsCode = "": PB_strMaterialsName = ""
                          End If
                       'Else '����ó�ڵ尡 �ٸ���
                       '   PB_strFMCCallFormName = "": PB_strMaterialsCode = "": PB_strMaterialsName = ""
                       'End If
                  End With
               ElseIf _
                  (PB_strFMCCallFormName = "frm�����ۼ�2") Then
                  With frm�����ۼ�2
                       'If Trim(.Text1(0).Text) = vsfg1.TextMatrix(vsfg1.Row, 4) Then                '����ó�ڵ�
                          'For lngR = 1 To .vsfg1.Rows - 1
                          '    If .vsfg1.Row <> lngR Then
                          '       If vsfg1.TextMatrix(vsfg1.Row, 0) = .vsfg1.TextMatrix(lngR, 0) And _
                          '          .vsfg1.TextMatrix(.vsfg1.Row, 5) = .vsfg1.TextMatrix(lngR, 5) Then '�����ڵ� + ����ó�ڵ�
                          '          blnDupOK = True
                          '          Exit For
                          '       End If
                          '    End If
                          'Next lngR
                          If (blnDupOK = False) And _
                             (.vsfg1.TextMatrix(.vsfg1.Row, 0) <> vsfg1.TextMatrix(vsfg1.Row, 0)) Then '�����ڵ� �ٸ���
                             .vsfg1.TextMatrix(.vsfg1.Row, 0) = vsfg1.TextMatrix(vsfg1.Row, 0)      '�����ڵ�
                             .vsfg1.TextMatrix(.vsfg1.Row, 1) = vsfg1.TextMatrix(vsfg1.Row, 1)      '�����
                             .vsfg1.TextMatrix(.vsfg1.Row, 2) = vsfg1.TextMatrix(vsfg1.Row, 7)      '����԰�
                             .vsfg1.TextMatrix(.vsfg1.Row, 4) = vsfg1.TextMatrix(vsfg1.Row, 8)      '�������
                             .vsfg1.TextMatrix(.vsfg1.Row, 8) = vsfg1.ValueMatrix(vsfg1.Row, 10)    '�԰�ܰ�
                             .vsfg1.TextMatrix(.vsfg1.Row, 9) = vsfg1.ValueMatrix(vsfg1.Row, 11)    '�԰�ΰ�
                             .vsfg1.TextMatrix(.vsfg1.Row, 10) = vsfg1.ValueMatrix(vsfg1.Row, 12)   '�԰���
                             '�հ�ݾ� ���(�԰�ݾ��� �ٸ���)
                             If .vsfg1.ValueMatrix(.vsfg1.Row, 11) <> _
                               (.vsfg1.ValueMatrix(.vsfg1.Row, 3) * vsfg1.ValueMatrix(vsfg1.Row, 10)) Then
                                .lblTotMny.Caption = Format(Vals(.lblTotMny.Caption) - .vsfg1.ValueMatrix(.vsfg1.Row, 11) _
                                                   + (.vsfg1.ValueMatrix(.vsfg1.Row, 3) * vsfg1.ValueMatrix(vsfg1.Row, 10)), "#,#.00")
                             End If
                             .vsfg1.TextMatrix(.vsfg1.Row, 11) = .vsfg1.ValueMatrix(.vsfg1.Row, 3) _
                                                               * vsfg1.ValueMatrix(vsfg1.Row, 10)   '�԰�ݾ�
                             .vsfg1.TextMatrix(.vsfg1.Row, 12) = vsfg1.ValueMatrix(vsfg1.Row, 13)   '���ܰ�
                             .vsfg1.TextMatrix(.vsfg1.Row, 13) = vsfg1.ValueMatrix(vsfg1.Row, 14)   '���ΰ�
                             .vsfg1.TextMatrix(.vsfg1.Row, 14) = vsfg1.ValueMatrix(vsfg1.Row, 15)   '�����
                             P_blnSelect = True
                          Else '�ߺ��̸�
                             PB_strFMCCallFormName = "": PB_strMaterialsCode = "": PB_strMaterialsName = ""
                          End If
                       'Else '����ó�ڵ尡 �ٸ���
                       '   PB_strFMCCallFormName = "": PB_strMaterialsCode = "": PB_strMaterialsName = ""
                       'End If
                  End With
               ElseIf _
                  (PB_strFMCCallFormName = "frm���Լ���") Then
                  With frm���Լ���
                       'For lngR = .vsfg1.ValueMatrix(.vsfg1.Row, 14) To .vsfg1.ValueMatrix(.vsfg1.Row, 15)
                       '    If .vsfg2.Row <> lngR Then
                       '       If vsfg1.TextMatrix(vsfg1.Row, 0) = .vsfg2.TextMatrix(lngR, 4) And _
                       '          .vsfg2.TextMatrix(.vsfg2.Row, 8) = .vsfg2.TextMatrix(lngR, 8) Then '�����ڵ� + ����ó�ڵ�
                       '          blnDupOK = True
                       '          Exit For
                       '       End If
                       '    End If
                       'Next lngR
                       If (blnDupOK = False) And _
                          (.vsfg2.TextMatrix(.vsfg2.Row, 4) <> vsfg1.TextMatrix(vsfg1.Row, 0)) Then
                          .vsfg2.Cell(flexcpBackColor, .vsfg2.Row, 5, .vsfg2.Row, 5) = vbRed
                          .vsfg2.Cell(flexcpForeColor, .vsfg2.Row, 5, .vsfg2.Row, 5) = vbWhite
                          If .vsfg2.TextMatrix(.vsfg2.Row, 28) = "" Then
                             .vsfg2.TextMatrix(.vsfg2.Row, 28) = "U"
                          End If
                          .vsfg2.TextMatrix(.vsfg2.Row, 4) = vsfg1.TextMatrix(vsfg1.Row, 0)      '�����ڵ�
                          .vsfg2.TextMatrix(.vsfg2.Row, 5) = vsfg1.TextMatrix(vsfg1.Row, 1)      '�����
                          'Grid Key(������ڵ�+��/��/��/��+�ŷ���ȣ+�����ڵ�+����ó�ڵ�
                          .vsfg2.TextMatrix(.vsfg2.Row, 12) = .vsfg1.TextMatrix(.vsfg1.Row, 3) _
                                    & "-" & .vsfg2.TextMatrix(.vsfg2.Row, 4) & "-" & .vsfg2.TextMatrix(.vsfg2.Row, 8)
                          .vsfg2.TextMatrix(.vsfg2.Row, 15) = vsfg1.TextMatrix(vsfg1.Row, 7)     '����԰�
                          .vsfg2.TextMatrix(.vsfg2.Row, 17) = vsfg1.TextMatrix(vsfg1.Row, 8)     '�������
                             
                          .vsfg2.TextMatrix(.vsfg2.Row, 19) = vsfg1.ValueMatrix(vsfg1.Row, 10)   '�԰�ܰ�
                          .vsfg2.TextMatrix(.vsfg2.Row, 20) = vsfg1.ValueMatrix(vsfg1.Row, 11)   '�԰�ΰ�
                          .vsfg2.TextMatrix(.vsfg2.Row, 21) = vsfg1.ValueMatrix(vsfg1.Row, 12)   '�԰���
                          '�հ�ݾ� ���(�԰�ݾ��� �ٸ���)
                          If .vsfg2.ValueMatrix(.vsfg2.Row, 22) <> _
                            (.vsfg2.ValueMatrix(.vsfg2.Row, 16) * vsfg1.ValueMatrix(vsfg1.Row, 10)) Then
                             .vsfg1.TextMatrix(.vsfg1.Row, 6) = .vsfg1.ValueMatrix(.vsfg1.Row, 6) - .vsfg2.ValueMatrix(.vsfg2.Row, 22) _
                                                             + (.vsfg2.ValueMatrix(.vsfg2.Row, 16) * vsfg1.ValueMatrix(vsfg1.Row, 10))
                             .lblTotMny.Caption = Format(Vals(.lblTotMny.Caption) - .vsfg2.ValueMatrix(.vsfg2.Row, 22) _
                                                + (.vsfg2.ValueMatrix(.vsfg2.Row, 16) * vsfg1.ValueMatrix(vsfg1.Row, 10)), "#,#.00")
                          End If
                          .vsfg2.TextMatrix(.vsfg2.Row, 22) = .vsfg2.ValueMatrix(.vsfg2.Row, 16) _
                                                            * vsfg1.ValueMatrix(vsfg1.Row, 10)   '�԰�ݾ�
                                                               
                          .vsfg2.TextMatrix(.vsfg2.Row, 23) = vsfg1.ValueMatrix(vsfg1.Row, 13)   '���ܰ�
                          .vsfg2.TextMatrix(.vsfg2.Row, 24) = vsfg1.ValueMatrix(vsfg1.Row, 14)   '���ΰ�
                          .vsfg2.TextMatrix(.vsfg2.Row, 25) = vsfg1.ValueMatrix(vsfg1.Row, 15)   '�����
                          .vsfg2.TextMatrix(.vsfg2.Row, 26) = .vsfg2.ValueMatrix(.vsfg2.Row, 16) _
                                                            * vsfg1.ValueMatrix(vsfg1.Row, 13)   '���ݾ�
                          P_blnSelect = True
                       Else '�ߺ��̸�
                          PB_strFMCCallFormName = "": PB_strMaterialsCode = "": PB_strMaterialsName = ""
                       End If
                  End With
               ElseIf _
                  (PB_strFMCCallFormName = "frm�������ۼ�") Then
                  With frm�������ۼ�
                       If (blnDupOK = False) Then 'And _
                          (.vsfg1.TextMatrix(.vsfg1.Row, 0) <> vsfg1.TextMatrix(vsfg1.Row, 0)) Then '�����ڵ� �ٸ���
                          .vsfg1.TextMatrix(.vsfg1.Row, 0) = vsfg1.TextMatrix(vsfg1.Row, 0)      '�����ڵ�
                          .vsfg1.TextMatrix(.vsfg1.Row, 1) = vsfg1.TextMatrix(vsfg1.Row, 1)      '�����
                          .vsfg1.TextMatrix(.vsfg1.Row, 2) = vsfg1.TextMatrix(vsfg1.Row, 7)      '����԰�
                          .vsfg1.TextMatrix(.vsfg1.Row, 4) = vsfg1.TextMatrix(vsfg1.Row, 8)      '�������
                          .vsfg1.TextMatrix(.vsfg1.Row, 8) = vsfg1.ValueMatrix(vsfg1.Row, 10)    '�԰�ܰ�
                          .vsfg1.TextMatrix(.vsfg1.Row, 9) = vsfg1.ValueMatrix(vsfg1.Row, 11)    '�԰�ΰ�
                          .vsfg1.TextMatrix(.vsfg1.Row, 10) = vsfg1.ValueMatrix(vsfg1.Row, 12)   '�԰���
                          .vsfg1.TextMatrix(.vsfg1.Row, 12) = vsfg1.ValueMatrix(vsfg1.Row, 13)   '���ܰ�
                          .vsfg1.TextMatrix(.vsfg1.Row, 13) = vsfg1.ValueMatrix(vsfg1.Row, 14)   '���ΰ�
                          .vsfg1.TextMatrix(.vsfg1.Row, 14) = vsfg1.ValueMatrix(vsfg1.Row, 15)   '�����
                          '�հ�ݾ� ���(���ݾ��� �ٸ���)
                          If .vsfg1.ValueMatrix(.vsfg1.Row, 15) <> _
                            (.vsfg1.ValueMatrix(.vsfg1.Row, 3) * vsfg1.ValueMatrix(vsfg1.Row, 13)) Then
                             .lblTotMny.Caption = Format(Vals(.lblTotMny.Caption) - .vsfg1.ValueMatrix(.vsfg1.Row, 15) _
                                                + (.vsfg1.ValueMatrix(.vsfg1.Row, 3) * vsfg1.ValueMatrix(vsfg1.Row, 13)), "#,#.00")
                          End If
                          .vsfg1.TextMatrix(.vsfg1.Row, 15) = .vsfg1.ValueMatrix(.vsfg1.Row, 3) _
                                                            * vsfg1.ValueMatrix(vsfg1.Row, 13)   '���ݾ�
                          P_blnSelect = True
                       Else '�ߺ��̸�
                          PB_strFMCCallFormName = "": PB_strMaterialsCode = "": PB_strMaterialsName = ""
                       End If
                  End With
               ElseIf _
                  (PB_strFMCCallFormName = "frm����������") Then
                  With frm����������
                       If (blnDupOK = False) And _
                          (.vsfg2.TextMatrix(.vsfg2.Row, 4) <> vsfg1.TextMatrix(vsfg1.Row, 0)) Or _
                          (.vsfg2.ValueMatrix(.vsfg2.Row, 23) <> vsfg1.ValueMatrix(vsfg1.Row, 13)) Then '�����ڵ�, �ܰ� �ٸ���
                          If .vsfg2.TextMatrix(.vsfg2.Row, 4) <> vsfg1.TextMatrix(vsfg1.Row, 0) Then
                             .vsfg2.Cell(flexcpBackColor, .vsfg2.Row, 5, .vsfg2.Row, 5) = vbRed
                             .vsfg2.Cell(flexcpForeColor, .vsfg2.Row, 5, .vsfg2.Row, 5) = vbWhite
                          End If
                          If .vsfg2.ValueMatrix(.vsfg2.Row, 23) <> vsfg1.ValueMatrix(vsfg1.Row, 13) Then
                             .vsfg2.Cell(flexcpBackColor, .vsfg2.Row, 23, .vsfg2.Row, 23) = vbRed
                             .vsfg2.Cell(flexcpForeColor, .vsfg2.Row, 23, .vsfg2.Row, 23) = vbWhite
                          End If
                          If .vsfg2.TextMatrix(.vsfg2.Row, 28) = "" Then
                             .vsfg2.TextMatrix(.vsfg2.Row, 28) = "U"
                          End If
                          .vsfg2.TextMatrix(.vsfg2.Row, 4) = vsfg1.TextMatrix(vsfg1.Row, 0)      '�����ڵ�
                          .vsfg2.TextMatrix(.vsfg2.Row, 5) = vsfg1.TextMatrix(vsfg1.Row, 1)      '�����
                          'Grid Key(������ڵ�+��/��/��/��+������ȣ+�����ڵ�+����ó�ڵ�
                          .vsfg2.TextMatrix(.vsfg2.Row, 12) = .vsfg1.TextMatrix(.vsfg1.Row, 1) _
                                    & "-" & .vsfg2.TextMatrix(.vsfg2.Row, 4) & "-" & .vsfg2.TextMatrix(.vsfg2.Row, 8)
                          .vsfg2.TextMatrix(.vsfg2.Row, 15) = vsfg1.TextMatrix(vsfg1.Row, 7)     '����԰�
                          .vsfg2.TextMatrix(.vsfg2.Row, 17) = vsfg1.TextMatrix(vsfg1.Row, 8)     '�������
                             
                          .vsfg2.TextMatrix(.vsfg2.Row, 19) = vsfg1.ValueMatrix(vsfg1.Row, 10)   '�԰�ܰ�
                          .vsfg2.TextMatrix(.vsfg2.Row, 20) = vsfg1.ValueMatrix(vsfg1.Row, 11)   '�԰�ΰ�
                          .vsfg2.TextMatrix(.vsfg2.Row, 21) = vsfg1.ValueMatrix(vsfg1.Row, 12)   '�԰���
                          .vsfg2.TextMatrix(.vsfg2.Row, 22) = .vsfg2.ValueMatrix(.vsfg2.Row, 16) _
                                                            * vsfg1.ValueMatrix(vsfg1.Row, 10)   '�԰�ݾ�
                                                               
                          .vsfg2.TextMatrix(.vsfg2.Row, 23) = vsfg1.ValueMatrix(vsfg1.Row, 13)   '���ܰ�
                          .vsfg2.TextMatrix(.vsfg2.Row, 24) = vsfg1.ValueMatrix(vsfg1.Row, 14)   '���ΰ�
                          .vsfg2.TextMatrix(.vsfg2.Row, 25) = vsfg1.ValueMatrix(vsfg1.Row, 15)   '�����
                          '�հ�ݾ� ���(���ݾ��� �ٸ���)
                          If .vsfg2.ValueMatrix(.vsfg2.Row, 26) <> _
                            (.vsfg2.ValueMatrix(.vsfg2.Row, 16) * vsfg1.ValueMatrix(vsfg1.Row, 13)) Then
                            .vsfg1.TextMatrix(.vsfg1.Row, 7) = .vsfg1.ValueMatrix(.vsfg1.Row, 7) - .vsfg2.ValueMatrix(.vsfg2.Row, 26) _
                                                             + (.vsfg2.ValueMatrix(.vsfg2.Row, 16) * vsfg1.ValueMatrix(vsfg1.Row, 13))
                            .lblTotMny.Caption = Format(Vals(.lblTotMny.Caption) - .vsfg2.ValueMatrix(.vsfg2.Row, 26) _
                                               + (.vsfg2.ValueMatrix(.vsfg2.Row, 16) * vsfg1.ValueMatrix(vsfg1.Row, 13)), "#,#.00")
                          End If
                          .vsfg2.TextMatrix(.vsfg2.Row, 26) = .vsfg2.ValueMatrix(.vsfg2.Row, 16) _
                                                            * vsfg1.ValueMatrix(vsfg1.Row, 13)   '���ݾ�
                          P_blnSelect = True
                       Else '�ߺ��̸�
                          PB_strFMCCallFormName = "": PB_strMaterialsCode = "": PB_strMaterialsName = ""
                       End If
                  End With
               ElseIf _
                  (PB_strFMCCallFormName = "frm�����ۼ�1") Then
                  With frm�����ۼ�1
                       'For lngR = .vsfg1.ValueMatrix(.vsfg1.Row, 15) To .vsfg1.ValueMatrix(.vsfg1.Row, 16)
                       '    If .vsfg2.Row <> lngR Then
                       '       If vsfg1.TextMatrix(vsfg1.Row, 0) = .vsfg2.TextMatrix(lngR, 4) And _
                       '          .vsfg2.TextMatrix(.vsfg2.Row, 8) = .vsfg2.TextMatrix(lngR, 8) Then '�����ڵ� + ����ó�ڵ�
                       '          blnDupOK = True
                       '          Exit For
                       '       End If
                       '    End If
                       'Next lngR
                       If (blnDupOK = False) And _
                          (.vsfg2.TextMatrix(.vsfg2.Row, 4) <> vsfg1.TextMatrix(vsfg1.Row, 0)) Or _
                          (.vsfg2.ValueMatrix(.vsfg2.Row, 23) <> vsfg1.ValueMatrix(vsfg1.Row, 13)) Then '�����ڵ�, �ܰ� �ٸ���
                          If .vsfg2.TextMatrix(.vsfg2.Row, 4) <> vsfg1.TextMatrix(vsfg1.Row, 0) Then
                             .vsfg2.Cell(flexcpBackColor, .vsfg2.Row, 5, .vsfg2.Row, 5) = vbRed
                             .vsfg2.Cell(flexcpForeColor, .vsfg2.Row, 5, .vsfg2.Row, 5) = vbWhite
                          End If
                          If .vsfg2.ValueMatrix(.vsfg2.Row, 23) <> vsfg1.ValueMatrix(vsfg1.Row, 13) Then
                             .vsfg2.Cell(flexcpBackColor, .vsfg2.Row, 23, .vsfg2.Row, 23) = vbRed
                             .vsfg2.Cell(flexcpForeColor, .vsfg2.Row, 23, .vsfg2.Row, 23) = vbWhite
                          End If
                          .vsfg2.TextMatrix(.vsfg2.Row, 4) = vsfg1.TextMatrix(vsfg1.Row, 0)      '�����ڵ�
                          .vsfg2.TextMatrix(.vsfg2.Row, 5) = vsfg1.TextMatrix(vsfg1.Row, 1)      '�����
                          'Grid Key(������ڵ�+��/��/��/��+������ȣ+�����ڵ�+����ó�ڵ�
                          .vsfg2.TextMatrix(.vsfg2.Row, 12) = .vsfg1.TextMatrix(.vsfg1.Row, 3) _
                                    & "-" & .vsfg2.TextMatrix(.vsfg2.Row, 4) & "-" & .vsfg2.TextMatrix(.vsfg2.Row, 8)
                          .vsfg2.TextMatrix(.vsfg2.Row, 15) = vsfg1.TextMatrix(vsfg1.Row, 7)     '����԰�
                          .vsfg2.TextMatrix(.vsfg2.Row, 17) = vsfg1.TextMatrix(vsfg1.Row, 8)     '�������
                             
                          .vsfg2.TextMatrix(.vsfg2.Row, 19) = vsfg1.ValueMatrix(vsfg1.Row, 10)   '�԰�ܰ�
                          .vsfg2.TextMatrix(.vsfg2.Row, 20) = vsfg1.ValueMatrix(vsfg1.Row, 11)   '�԰�ΰ�
                          .vsfg2.TextMatrix(.vsfg2.Row, 21) = vsfg1.ValueMatrix(vsfg1.Row, 12)   '�԰���
                          .vsfg2.TextMatrix(.vsfg2.Row, 22) = .vsfg2.ValueMatrix(.vsfg2.Row, 16) _
                                                            * vsfg1.ValueMatrix(vsfg1.Row, 12)   '�԰�ݾ�
                                                               
                          .vsfg2.TextMatrix(.vsfg2.Row, 23) = vsfg1.ValueMatrix(vsfg1.Row, 13)   '���ܰ�
                          .vsfg2.TextMatrix(.vsfg2.Row, 24) = vsfg1.ValueMatrix(vsfg1.Row, 14)   '���ΰ�
                          .vsfg2.TextMatrix(.vsfg2.Row, 25) = vsfg1.ValueMatrix(vsfg1.Row, 15)   '�����
                          '�հ�ݾ� ���(���ݾ��� �ٸ���)
                          If .vsfg2.ValueMatrix(.vsfg2.Row, 26) <> _
                            (.vsfg2.ValueMatrix(.vsfg2.Row, 16) * vsfg1.ValueMatrix(vsfg1.Row, 13)) Then
                            .vsfg1.TextMatrix(.vsfg1.Row, 7) = .vsfg1.ValueMatrix(.vsfg1.Row, 7) - .vsfg2.ValueMatrix(.vsfg2.Row, 26) _
                                                             + (.vsfg2.ValueMatrix(.vsfg2.Row, 16) * vsfg1.ValueMatrix(vsfg1.Row, 13))
                            .lblTotMny.Caption = Format(Vals(.lblTotMny.Caption) - .vsfg2.ValueMatrix(.vsfg2.Row, 26) _
                                               + (.vsfg2.ValueMatrix(.vsfg2.Row, 16) * vsfg1.ValueMatrix(vsfg1.Row, 13)), "#,#.00")
                          End If
                          .vsfg2.TextMatrix(.vsfg2.Row, 26) = .vsfg2.ValueMatrix(.vsfg2.Row, 16) _
                                                            * vsfg1.ValueMatrix(vsfg1.Row, 13)   '���ݾ�
                          P_blnSelect = True
                       Else '�ߺ��̸�
                          PB_strFMCCallFormName = "": PB_strMaterialsCode = "": PB_strMaterialsName = ""
                       End If
                  End With
               ElseIf _
                  (PB_strFMCCallFormName = "frm�����ۼ�2") Then
                  With frm�����ۼ�2
                       'For lngR = 1 To .vsfg1.Rows - 1
                       '    If .vsfg1.Row <> lngR Then
                       '       If vsfg1.TextMatrix(vsfg1.Row, 0) = .vsfg1.TextMatrix(lngR, 0) Then '�����ڵ�
                       '          blnDupOK = True
                       '          Exit For
                       '       End If
                       '    End If
                       'Next lngR
                       If (blnDupOK = False) And _
                          (.vsfg1.TextMatrix(.vsfg1.Row, 0) <> vsfg1.TextMatrix(vsfg1.Row, 0)) Then
                          .vsfg1.TextMatrix(.vsfg1.Row, 0) = vsfg1.TextMatrix(vsfg1.Row, 0)      '�����ڵ�
                          .vsfg1.TextMatrix(.vsfg1.Row, 1) = vsfg1.TextMatrix(vsfg1.Row, 1)      '�����
                          .vsfg1.TextMatrix(.vsfg1.Row, 2) = vsfg1.TextMatrix(vsfg1.Row, 7)      '����԰�
                          .vsfg1.TextMatrix(.vsfg1.Row, 4) = vsfg1.TextMatrix(vsfg1.Row, 8)      '�������
                          .vsfg1.TextMatrix(.vsfg1.Row, 8) = vsfg1.ValueMatrix(vsfg1.Row, 10)    '�԰�ܰ�
                          .vsfg1.TextMatrix(.vsfg1.Row, 9) = vsfg1.ValueMatrix(vsfg1.Row, 11)    '�԰�ΰ�
                          .vsfg1.TextMatrix(.vsfg1.Row, 10) = vsfg1.ValueMatrix(vsfg1.Row, 12)   '�԰���
                          .vsfg1.TextMatrix(.vsfg1.Row, 12) = vsfg1.ValueMatrix(vsfg1.Row, 13)   '���ܰ�
                          .vsfg1.TextMatrix(.vsfg1.Row, 13) = vsfg1.ValueMatrix(vsfg1.Row, 14)   '���ΰ�
                          .vsfg1.TextMatrix(.vsfg1.Row, 14) = vsfg1.ValueMatrix(vsfg1.Row, 15)   '�����
                          '�հ�ݾ� ���(���ݾ��� �ٸ���)
                          If .vsfg1.ValueMatrix(.vsfg1.Row, 15) <> _
                            (.vsfg1.ValueMatrix(.vsfg1.Row, 3) * vsfg1.ValueMatrix(vsfg1.Row, 13)) Then
                            .lblTotMny.Caption = Format(Vals(.lblTotMny.Caption) - .vsfg1.ValueMatrix(.vsfg1.Row, 15) _
                                               + (.vsfg1.ValueMatrix(.vsfg1.Row, 3) * vsfg1.ValueMatrix(vsfg1.Row, 13)), "#,#.00")
                          End If
                          .vsfg1.TextMatrix(.vsfg1.Row, 15) = .vsfg1.ValueMatrix(.vsfg1.Row, 3) _
                                                            * vsfg1.ValueMatrix(vsfg1.Row, 13)   '���ݾ�
                          P_blnSelect = True
                       Else '�ߺ��̸�
                          PB_strFMCCallFormName = "": PB_strMaterialsCode = "": PB_strMaterialsName = ""
                       End If
                  End With
               ElseIf _
                  (PB_strFMCCallFormName = "frm�������") Then
                  With frm�������
                       'For lngR = .vsfg1.ValueMatrix(.vsfg1.Row, 15) To .vsfg1.ValueMatrix(.vsfg1.Row, 16)
                       '    If .vsfg2.Row <> lngR Then
                       '       If vsfg1.TextMatrix(vsfg1.Row, 0) = .vsfg2.TextMatrix(lngR, 4) And _
                       '          .vsfg2.TextMatrix(.vsfg2.Row, 8) = .vsfg2.TextMatrix(lngR, 8) Then '�����ڵ� + ����ó�ڵ�
                       '          blnDupOK = True
                       '          Exit For
                       '       End If
                       '    End If
                       'Next lngR
                       'if �ߺ� = false And (0.���������ڵ� <> �����ڵ�)           '����
                       If (blnDupOK = False) And _
                          (.vsfg2.TextMatrix(.vsfg2.Row, 4) <> vsfg1.TextMatrix(vsfg1.Row, 0)) Then
                          .vsfg2.Cell(flexcpBackColor, .vsfg2.Row, 5, .vsfg2.Row, 5) = vbRed
                          .vsfg2.Cell(flexcpForeColor, .vsfg2.Row, 5, .vsfg2.Row, 5) = vbWhite
                          If .vsfg2.TextMatrix(.vsfg2.Row, 28) = "" Then
                             .vsfg2.TextMatrix(.vsfg2.Row, 28) = "U"
                          End If
                          .vsfg2.TextMatrix(.vsfg2.Row, 4) = vsfg1.TextMatrix(vsfg1.Row, 0)      '�����ڵ�
                          .vsfg2.TextMatrix(.vsfg2.Row, 5) = vsfg1.TextMatrix(vsfg1.Row, 1)      '�����
                          'Grid Key(������ڵ�+��/��/��/��+�ŷ���ȣ+�����ڵ�+����ó�ڵ�
                          .vsfg2.TextMatrix(.vsfg2.Row, 12) = .vsfg1.TextMatrix(.vsfg1.Row, 3) _
                                    & "-" & .vsfg2.TextMatrix(.vsfg2.Row, 4) & "-" & .vsfg2.TextMatrix(.vsfg2.Row, 8)
                          .vsfg2.TextMatrix(.vsfg2.Row, 15) = vsfg1.TextMatrix(vsfg1.Row, 7)     '����԰�
                          .vsfg2.TextMatrix(.vsfg2.Row, 17) = vsfg1.TextMatrix(vsfg1.Row, 8)     '�������
                             
                          .vsfg2.TextMatrix(.vsfg2.Row, 19) = vsfg1.ValueMatrix(vsfg1.Row, 10)   '�԰�ܰ�
                          .vsfg2.TextMatrix(.vsfg2.Row, 20) = vsfg1.ValueMatrix(vsfg1.Row, 11)   '�԰�ΰ�
                          .vsfg2.TextMatrix(.vsfg2.Row, 21) = vsfg1.ValueMatrix(vsfg1.Row, 12)   '�԰���
                          .vsfg2.TextMatrix(.vsfg2.Row, 22) = .vsfg2.ValueMatrix(.vsfg2.Row, 16) _
                                                            * vsfg1.ValueMatrix(vsfg1.Row, 12)   '�԰�ݾ�
                                                               
                          .vsfg2.TextMatrix(.vsfg2.Row, 23) = vsfg1.ValueMatrix(vsfg1.Row, 13)   '���ܰ�
                          .vsfg2.TextMatrix(.vsfg2.Row, 24) = vsfg1.ValueMatrix(vsfg1.Row, 14)   '���ΰ�
                          .vsfg2.TextMatrix(.vsfg2.Row, 25) = vsfg1.ValueMatrix(vsfg1.Row, 15)   '�����
                          '�հ�ݾ� ���(���ݾ��� �ٸ���)
                          If .vsfg2.ValueMatrix(.vsfg2.Row, 26) <> _
                            (.vsfg2.ValueMatrix(.vsfg2.Row, 16) * vsfg1.ValueMatrix(vsfg1.Row, 13)) Then
                            .vsfg1.TextMatrix(.vsfg1.Row, 6) = .vsfg1.ValueMatrix(.vsfg1.Row, 6) - .vsfg2.ValueMatrix(.vsfg2.Row, 26) _
                                                             + (.vsfg2.ValueMatrix(.vsfg2.Row, 16) * vsfg1.ValueMatrix(vsfg1.Row, 13))
                            .lblTotMny.Caption = Format(Vals(.lblTotMny.Caption) - .vsfg2.ValueMatrix(.vsfg2.Row, 26) _
                                               + (.vsfg2.ValueMatrix(.vsfg2.Row, 16) * vsfg1.ValueMatrix(vsfg1.Row, 13)), "#,#.00")
                          End If
                          .vsfg2.TextMatrix(.vsfg2.Row, 26) = .vsfg2.ValueMatrix(.vsfg2.Row, 16) _
                                                            * vsfg1.ValueMatrix(vsfg1.Row, 13)   '���ݾ�
                          P_blnSelect = True
                       Else '�ߺ��̸�
                          PB_strFMCCallFormName = "": PB_strMaterialsCode = "": PB_strMaterialsName = ""
                       End If
                  End With
               End If
            End If
            Unload Me
         End If
    End With
    Exit Sub
End Sub

'+-----------+
'/// ��ȸ ///
'+-----------+
Private Sub cmdFind_Click()
    Screen.MousePointer = vbHourglass
    cmdFind.Enabled = False
    'P_strFindString2 = Trim(Text1(1).Text)  '��ȸ�� ��� �˻��� ����� ����
    Subvsfg1_FILL
    cmdFind.Enabled = True
    Screen.MousePointer = vbDefault
End Sub

'+-----------+
'/// �߰� ///
'+-----------+
Private Sub cmdAdd_Click()
Dim strSQL     As String
Dim strMtCode  As String
Dim lngCodeSeq As Long
Dim lngR       As Long
Dim blnExist   As Boolean
    If LenH(txtCode.Text) > 18 Then
       MsgBox "�ڵ���̰� �ִ����(18��)�� �ʰ��մϴ�. �ٽ� Ȯ���� �Է��Ͽ� �ּ���.", vbCritical, "�ڵ�"
       txtCode.SetFocus
       Exit Sub
    End If
    If LenH(txtName.Text) > 30 Then
       MsgBox "ǰ����̰� �ִ����(30��)�� �ʰ��մϴ�. �ٽ� Ȯ���� �Է��Ͽ� �ּ���.", vbCritical, "ǰ��"
       txtName.SetFocus
       Exit Sub
    End If
    If LenH(txtSize.Text) > 30 Then
       MsgBox "�԰ݱ��̰� �ִ����(30��)�� �ʰ��մϴ�. �ٽ� Ȯ���� �Է��Ͽ� �ּ���.", vbCritical, "�԰�"
       txtSize.SetFocus
       Exit Sub
    End If
    If LenH(txtUnit.Text) > 20 Then
       MsgBox "�������̰� �ִ����(20��)�� �ʰ��մϴ�. �ٽ� Ȯ���� �Է��Ͽ� �ּ���.", vbCritical, "����"
       txtUnit.SetFocus
       Exit Sub
    End If
    If LenH(txtBarCode.Text) > 13 Then
       MsgBox "���ڵ���̰� �ִ����(13��)�� �ʰ��մϴ�. �ٽ� Ȯ���� �Է��Ͽ� �ּ���.", vbCritical, "���ڵ�"
       txtUnit.SetFocus
       Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    cmdAdd.Enabled = False
    If (Len(Trim(txtCode.Text)) > 5) And (Len(Trim(txtName.Text)) > 0) And (Len(Trim(txtUnit.Text)) > 0) Then
       PB_adoCnnSQL.BeginTrans
       strSQL = "SELECT �з��ڵ� FROM ����з� " _
               & "WHERE �з��ڵ� = '" & Mid(Trim(txtCode.Text), 1, 2) & "' "
       On Error GoTo ERROR_TABLE_SELECT
       P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
       If P_adoRec.RecordCount = 1 Then
          P_adoRec.Close
          strMtCode = Mid(Trim(txtCode.Text), 3)
          strSQL = "SELECT �з��ڵ�, �����ڵ� FROM ���� " _
                  & "WHERE �з��ڵ� = '" & Mid(Trim(txtCode.Text), 1, 2) & "' AND ����� = '" & Trim(txtName.Text) & "' " _
                    & "AND �԰� = '" & Trim(txtSize.Text) & "' AND ���� = '" & Trim(txtUnit.Text) & "' "
          On Error GoTo ERROR_TABLE_SELECT
          P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
          If P_adoRec.RecordCount = 0 Then
             P_adoRec.Close
             strSQL = "SELECT �з��ڵ�, �����ڵ� FROM ���� " _
                     & "WHERE �з��ڵ� = '" & Mid(Trim(txtCode.Text), 1, 2) & "' AND �����ڵ� = '" & strMtCode & "' "
             On Error GoTo ERROR_TABLE_SELECT
             P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
             If P_adoRec.RecordCount = 1 Then
                txtCode.Text = P_adoRec("�з��ڵ�") + "CODE"
                strMtCode = "CODE"
             End If
             P_adoRec.Close
             If UPPER(strMtCode) = "CODE" Then
                strMtCode = UPPER(strMtCode)
                strSQL = "SELECT ISNULL(MAX(�����ڵ�),0) AS �����ڵ� FROM ���� " _
                        & "WHERE �з��ڵ� = '" & Mid(Trim(txtCode.Text), 1, 2) & "' AND �����ڵ� LIKE 'CODE_____' " _
                          & "AND DATALENGTH(�����ڵ�) = 9 " _
                          & "AND ISNUMERIC(SUBSTRING(�����ڵ�, 5, 5)) = 1 "
                        '& "WHERE �з��ڵ� = '" & Mid(Trim(txtCode.Text), 1, 2) & "' AND �����ڵ� LIKE'CODE%' "
                On Error GoTo ERROR_TABLE_SELECT
                P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
                If P_adoRec.RecordCount = 0 Then
                  lngCodeSeq = 0
                Else
                  lngCodeSeq = Val(Mid(P_adoRec("�����ڵ�"), 5))
                End If
                P_adoRec.Close
                lngCodeSeq = lngCodeSeq + 1
                strMtCode = strMtCode + Format(lngCodeSeq, "00000")
                txtCode.Text = Mid(Trim(txtCode.Text), 1, 2) + strMtCode
             End If
             strSQL = "INSERT INTO ���� VALUES(" _
                                & "'" & Mid(Trim(txtCode.Text), 1, 2) & "', '" & strMtCode & "', " _
                                & "'" & Trim(txtName.Text) & "', '', '" & Trim(txtSize.Text) & "', '" & Trim(txtUnit.Text) & "', " _
                                & "0, 1, '', 0, '', '" & PB_regUserinfoU.UserCode & "') "
             On Error GoTo ERROR_TABLE_INSERT
             PB_adoCnnSQL.Execute strSQL
             '������� �߰�
             strSQL = "INSERT INTO ������� VALUES(" _
                                & "'" & PB_regUserinfoU.UserBranchCode & "', " _
                                & "'" & Mid(Trim(txtCode.Text), 1, 2) & "', '" & strMtCode & "', 0, 0, '', '', " _
                                & "0, '" & PB_regUserinfoU.UserClientDate & "', '" & PB_regUserinfoU.UserCode & "', " _
                                & "'', '', 0, 0, 0, 0, 0, 0 )"
             On Error GoTo ERROR_TABLE_INSERT
             PB_adoCnnSQL.Execute strSQL
          ElseIf _
             P_adoRec.RecordCount = 1 Then
             txtCode.Text = P_adoRec("�з��ڵ�") + P_adoRec("�����ڵ�")
             strMtCode = P_adoRec("�����ڵ�")
             P_adoRec.Close
             For lngR = 1 To vsfg1.Rows - 1
                 If vsfg1.TextMatrix(lngR, 0) = Trim(txtCode.Text) Then
                    vsfg1.Row = lngR
                    blnExist = True
                    Exit For
                 End If
             Next lngR
          Else
             P_adoRec.Close
          End If
          With vsfg1
               If blnExist = False Then
                  .AddItem ""
                  lngR = .Rows - 1
                  .TextMatrix(lngR, 0) = Trim(txtCode.Text): .TextMatrix(lngR, 1) = Trim(txtName.Text)
                  .TextMatrix(lngR, 2) = 0: .TextMatrix(lngR, 3) = 0
                  .TextMatrix(lngR, 5) = ""
                  .TextMatrix(lngR, 6) = .TextMatrix(lngR, 0)
                  .Cell(flexcpData, lngR, 6, lngR, 6) = .TextMatrix(lngR, 6)
                  .TextMatrix(lngR, 7) = Trim(txtSize.Text): .TextMatrix(lngR, 8) = Trim(txtUnit.Text)
                  .TextMatrix(lngR, 9) = 0: .TextMatrix(lngR, 10) = 0
                  .TextMatrix(lngR, 11) = 0: .TextMatrix(lngR, 12) = 0
                  .TextMatrix(lngR, 13) = 0: .TextMatrix(lngR, 14) = 0
                  .TextMatrix(lngR, 15) = 0: .TextMatrix(lngR, 16) = 0
                  .TextMatrix(lngR, 17) = 1: .TextMatrix(lngR, 18) = "��  ��"
                  .Row = lngR
                  .SetFocus
               End If
          End With
       Else
          P_adoRec.Close
       End If
       P_adoRec.CursorLocation = adUseClient
       strSQL = "SELECT ������ڵ� FROM ����� " _
               & "WHERE ������ڵ� <> '" & PB_regUserinfoU.UserBranchCode & "' "
       On Error GoTo ERROR_TABLE_SELECT
       P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
       If P_adoRec.RecordCount = 0 Then
          P_adoRec.Close
       Else
          Do Until P_adoRec.EOF
             strSQL = "INSERT INTO ������� " _
                    & "SELECT '" & P_adoRec("������ڵ�") & "' AS ������ڵ�, " _
                            & "ISNULL(T1.�з��ڵ� , '') AS �з��ڵ�, ISNULL(T1.�����ڵ� , '') AS �����ڵ�, " _
                            & "0 AS �������, 0 AS �������, '' AS �����԰�����, '' AS �����������, " _
                            & "T1.��뱸�� AS ��뱸��, T1.�������� AS ��������, T1.������ڵ� AS ������ڵ�, " _
                            & "'' AS ����, '' AS �ָ���ó�ڵ�, " _
                            & "T1.�԰�ܰ�1 AS �԰�ܰ�1, T1.�԰�ܰ�2 AS �԰�ܰ�2, T1.�԰�ܰ�3 AS �԰�ܰ�3, " _
                            & "T1.���ܰ�1 AS ���ܰ�1, T1.���ܰ�2 AS ���ܰ�2, T1.���ܰ�3 AS ���ܰ�3 " _
                       & "FROM ���� T1 " _
                      & "WHERE T1.�з��ڵ� = '" & Mid(Trim(txtCode.Text), 1, 2) & "' AND T1.�����ڵ� = '" & strMtCode & "' "
              On Error GoTo ERROR_TABLE_INSERT
              PB_adoCnnSQL.Execute strSQL
              P_adoRec.MoveNext
           Loop
           P_adoRec.Close
       End If
       PB_adoCnnSQL.CommitTrans
    End If
    cmdAdd.Enabled = True
    vsfg1.SetFocus
    Screen.MousePointer = vbDefault
    Exit Sub
ERROR_TABLE_SELECT:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "�б� ����"
    Screen.MousePointer = vbDefault
    Unload Me
    Exit Sub
ERROR_TABLE_INSERT:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "�߰� ����"
    Screen.MousePointer = vbDefault
    Unload Me
    Exit Sub
ERROR_TABLE_UPDATE:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "���� ����"
    Screen.MousePointer = vbDefault
    Unload Me
    Exit Sub
End Sub

'+-----------+
'/// ���� ///
'+-----------+
Private Sub Form_Unload(Cancel As Integer)
    If P_blnSelect = False Then
       SubEscape
    End If
    Screen.MousePointer = vbDefault
    If P_adoRec.State <> adStateClosed Then
       P_adoRec.Close
    End If
    Set P_adoRec = Nothing
    Set frm����ü��˻� = Nothing
End Sub

'+--------------------+
'/// Sub Procedure ///
'+--------------------+
'+----------------------------------+
'/// VsFlexGrid(vsfgGrid) �ʱ�ȭ ///
'+----------------------------------+
Private Sub Subvsfg1_INIT()
Dim lngR    As Long
Dim lngC    As Long
    With vsfg1           'Rows 1, Cols 19, RowHeightMax(Min) 300
         .AllowBigSelection = False
         .AllowSelection = False
         .AllowUserResizing = flexResizeColumns
         '.BackColorBkg = &H8000000F
         .BackColorSel = &H8000&
         .Ellipsis = flexEllipsisEnd
         '.ExplorerBar = flexExSortShow
         .ExtendLastCol = True
         .FocusRect = flexFocusHeavy
         .FontSize = 9
         .ScrollBars = flexScrollBarHorizontal
         .ScrollTrack = True
         .SelectionMode = flexSelectionByRow
         .FixedRows = 1
         .FixedCols = 6
         .Rows = 1             'SubvsfgUpGrid_Fill����ÿ� ����
         .Cols = 19
         .RowHeightMax = 300
         .RowHeightMin = 300
         .ColWidth(0) = 1700   '�����ڵ�(�з��ڵ�+�����ڵ�) '1900
         .ColWidth(1) = 2300   '�����                      '2500
         .ColWidth(2) = 850    '�������
         .ColWidth(3) = 1000   '�������
         .ColWidth(4) = 1000   '�ָ���ó�ڵ�                'Hidden
         .ColWidth(5) = 1800   '�ָ���ó��
         .ColWidth(6) = 1000   '�����ڵ�                    'Hidden
         .ColWidth(7) = 2200   '�԰�
         .ColWidth(8) = 550    '����
         .ColWidth(9) = 1000   '�����
         .ColWidth(10) = 1200  '�԰�ܰ�
         .ColWidth(11) = 1200  '�԰�ΰ�
         .ColWidth(12) = 1500  '�԰���
         .ColWidth(13) = 1200  '���ܰ�
         .ColWidth(14) = 1200  '���ΰ�
         .ColWidth(15) = 1500  '�����
         .ColWidth(16) = 800   '������
         .ColWidth(17) = 1     '��������
         .ColWidth(18) = 500   '��������
         .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
         .TextMatrix(0, 0) = "�ڵ�"
         .TextMatrix(0, 1) = "ǰ��"
         .TextMatrix(0, 2) = "�������"
         .TextMatrix(0, 3) = "�������"
         .TextMatrix(0, 4) = "�ָ���ó�ڵ�"    'H
         .TextMatrix(0, 5) = "�ָ���ó��"
         .TextMatrix(0, 6) = "KEY"             'H
         .TextMatrix(0, 7) = "�԰�"
         .TextMatrix(0, 8) = "����"
         .TextMatrix(0, 9) = "�����"          'H
         .TextMatrix(0, 10) = "���Դܰ�"
         .TextMatrix(0, 11) = "���Ժΰ�"       'H
         .TextMatrix(0, 12) = "���԰���"       'H
         .TextMatrix(0, 13) = "����ܰ�"
         .TextMatrix(0, 14) = "����ΰ�"       'H
         .TextMatrix(0, 15) = "���Ⱑ��"       'H
         .TextMatrix(0, 16) = "������"         'H
         .TextMatrix(0, 17) = "��������"       'H
         .TextMatrix(0, 18) = "����"           'H
         .ColHidden(4) = True: .ColHidden(6) = True: .ColHidden(9) = True: .ColHidden(12) = True
         .ColHidden(11) = True: .ColHidden(14) = True: .ColHidden(15) = True: .ColHidden(16) = True
         .ColHidden(17) = True: .ColHidden(18) = True
         For lngC = 0 To .Cols - 1
             Select Case lngC
                    Case 2
                         .ColFormat(lngC) = "#,#"
                    Case 9 To 16
                         .ColFormat(lngC) = "#,#.00"
             End Select
         Next lngC
         For lngC = 0 To .Cols - 1
             Select Case lngC
                    Case 0, 1, 5, 7, 8
                        .ColAlignment(lngC) = flexAlignLeftCenter
                    Case 4, 18
                        .ColAlignment(lngC) = flexAlignCenterCenter
                    Case Else
                        .ColAlignment(lngC) = flexAlignRightCenter
             End Select
         Next lngC
         .MergeCells = flexMergeRestrictAll
         For lngC = 0 To 5
             .MergeCol(lngC) = True
         Next lngC
         If (PB_strFMCCallFormName = "frm���ּ��ۼ�") Then
            .ColHidden(13) = True
         End If
         If (PB_strFMCCallFormName = "frm���ּ�����") Then
            .ColHidden(13) = True
         End If
         If (PB_strFMCCallFormName = "frm�����ۼ�1") Then
            .ColHidden(13) = True
         End If
         If (PB_strFMCCallFormName = "frm�����ۼ�2") Then
            .ColHidden(13) = True
         End If
         If (PB_strFMCCallFormName = "��ǰ����") Then
            .ColHidden(13) = True
         End If
         If (PB_strFMCCallFormName = "frm���Լ���") Then
            .ColHidden(13) = True
         End If
         If (PB_strFMCCallFormName = "frm�������ۼ�") Then
            '.ColHidden(10) = True
         End If
         If (PB_strFMCCallFormName = "frm����������") Then
            '.ColHidden(10) = True
         End If
         If (PB_strFMCCallFormName = "frm�����ۼ�1") Then
            '.ColHidden(10) = True
         End If
         If (PB_strFMCCallFormName = "frm�����ۼ�2") Then
            '.ColHidden(10) = True
         End If
         If (PB_strFMCCallFormName = "frm���԰���") Then
            '.ColHidden(10) = True
         End If
         If (PB_strFMCCallFormName = "frm�������") Then
            '.ColHidden(10) = True
         End If
    End With
End Sub

'+---------------------------------+
'/// VsFlexGrid(vsfgGrid) ä���///
'+---------------------------------+
Private Sub Subvsfg1_FILL()
Dim strSQL     As String
Dim strSelect  As String
Dim strJoin    As String
Dim strWhere   As String
Dim strOrderBy As String
Dim lngR       As Long
Dim lngC       As Long
Dim lngRR      As Long
Dim lngRRR     As Long
Dim strAppDate As Long
    Screen.MousePointer = vbHourglass
    txtCode.Text = Trim(txtCode.Text): txtName.Text = Trim(txtName.Text)
    If Len(txtCode.Text) = 0 And Len(txtName.Text) = 0 Then
       P_intFindGbn = 0    '�����˻�
    ElseIf _
       Len(txtCode.Text) <> 0 And Len(txtName.Text) = 0 Then
       P_intFindGbn = 1    '�ڵ�θ� �ڵ��˻�
    ElseIf _
       Len(txtCode.Text) = 0 And Len(txtName.Text) <> 0 Then
       P_intFindGbn = 2    '�̸�(��)���θ� �ڵ��˻�
    ElseIf _
       Len(txtCode.Text) <> 0 And Len(txtName.Text) <> 0 Then
       P_intFindGbn = 3    '�ڵ�� �̸�(��)�� ���ÿ� �ڵ��˻�
    Else
       P_intFindGbn = 0    '�����˻�
    End If
    If P_intFindGbn = 1 Then '�ڵ�� �˻�
       strWhere = strWhere + IIf(Trim(strWhere) = "", "WHERE ", "AND ") _
                & "(T1.�з��ڵ� + T1.�����ڵ�) LIKE '%" & Trim(txtCode.Text) & "%' " _
                & "AND T2.���� LIKE '%" & Trim(txtUnit.Text) & "%' " _
                & "AND T1.������ڵ� = '" & PB_regUserinfoU.UserBranchCode & "' "
       strOrderBy = "ORDER BY T1.�з��ڵ�, T1.�����ڵ� "
    ElseIf _
       P_intFindGbn = 2 Then '�̸����� �˻�
       strWhere = strWhere + IIf(Trim(strWhere) = "", "WHERE ", "AND ") _
                & "T2.����� LIKE '%" & Trim(txtName.Text) & "%' " _
                & "AND T2.�԰� LIKE '%" & Trim(txtSize.Text) & "%' " _
                & "AND T2.���� LIKE '%" & Trim(txtUnit.Text) & "%' " _
                & "AND T1.������ڵ� = '" & PB_regUserinfoU.UserBranchCode & "' "
       strOrderBy = "ORDER BY T2.����� "
    ElseIf _
       P_intFindGbn = 3 Then '�ڵ�� �̸����� �˻�
       strWhere = strWhere + IIf(Trim(strWhere) = "", "WHERE ", "AND ") _
                & "(T1.�з��ڵ� + T1.�����ڵ�) LIKE '%" & Trim(txtCode.Text) & "%' " _
                & "AND T2.����� LIKE '%" & Trim(txtName.Text) & "%' " _
                & "AND T2.�԰� LIKE '%" & Trim(txtSize.Text) & "%' " _
                & "AND T2.���� LIKE '%" & Trim(txtUnit.Text) & "%' " _
                & "AND T1.������ڵ� = '" & PB_regUserinfoU.UserBranchCode & "' "
       strOrderBy = "ORDER BY T1.�з��ڵ�, T1.�����ڵ� "
    Else
       strWhere = strWhere + IIf(Trim(strWhere) = "", "WHERE ", "AND ") _
                & "T2.����� LIKE '%" & Trim(txtName.Text) & "%' " _
                & "AND T2.�԰� LIKE '%" & Trim(txtSize.Text) & "%' " _
                & "AND T2.���� LIKE '%" & Trim(txtUnit.Text) & "%' " _
                & "AND T1.������ڵ� = '" & PB_regUserinfoU.UserBranchCode & "' "
       strOrderBy = "ORDER BY T1.�з��ڵ�, T1.�����ڵ� "
    End If
    If PB_strFMCCallFormName = "frm���ּ��ۼ�" Or PB_strFMCCallFormName = "frm���ּ�����" Or _
       PB_strFMCCallFormName = "frm�����ۼ�1" Or PB_strFMCCallFormName = "frm�����ۼ�2" Or PB_strFMCCallFormName = "frm���Լ���" Or _
       PB_strFMCCallFormName = "frm��ǰ����" Then
       P_intIOGbn = 1
       strSelect = "�԰�ܰ� = CASE WHEN T4.�ܰ����� = 1 THEN T1.�԰�ܰ�1 " _
                                 & "WHEN T4.�ܰ����� = 2 THEN T1.�԰�ܰ�2 " _
                                 & "WHEN T4.�ܰ����� = 3 THEN T1.�԰�ܰ�3 ELSE 0 END, " _
                 & "�԰�ΰ� = CASE WHEN T4.�ܰ����� = 1 THEN ROUND(T1.�԰�ܰ�1 * (" & PB_curVatRate & "), 0, 1) " _
                                 & "WHEN T4.�ܰ����� = 2 THEN ROUND(T1.�԰�ܰ�2 * (" & PB_curVatRate & "), 0, 1) " _
                                 & "WHEN T4.�ܰ����� = 3 THEN ROUND(T1.�԰�ܰ�3 * (" & PB_curVatRate & "), 0, 1) ELSE 0 END, " _
                 & "�԰��� = CASE WHEN T4.�ܰ����� = 1 THEN T1.�԰�ܰ�1 + ROUND(T1.�԰�ܰ�1 * (" & PB_curVatRate & "), 0, 1) " _
                                 & "WHEN T4.�ܰ����� = 2 THEN T1.�԰�ܰ�2 + ROUND(T1.�԰�ܰ�2 * (" & PB_curVatRate & "), 0, 1) " _
                                 & "WHEN T4.�ܰ����� = 3 THEN T1.�԰�ܰ�3 + ROUND(T1.�԰�ܰ�3 * (" & PB_curVatRate & "), 0, 1) ELSE 0 END, " _
                 & "0 AS ���ܰ�, 0 AS ���ΰ�, 0 AS �����, "
       strJoin = "INNER JOIN ����ó T4 ON T4.������ڵ� = T1.������ڵ� AND T4.����ó�ڵ� = '" & Trim(PB_strSupplierCode) & "' "
       strWhere = strWhere + IIf(Trim(strWhere) = "", "WHERE ", "AND ") _
                & "T4.����ó�ڵ� = '" & Trim(PB_strSupplierCode) & "' "
    Else
       P_intIOGbn = 2
       strSelect = "T1.�԰�ܰ�1 AS �԰�ܰ�, (T1.�԰�ܰ�1 * 0.1) AS �԰�ΰ�, (T1.�԰�ܰ�1 + (T1.�԰�ܰ�1 * 0.1)) AS �԰���, " _
                 & "���ܰ� = CASE WHEN T4.�ܰ����� = 1 THEN T1.���ܰ�1 " _
                                 & "WHEN T4.�ܰ����� = 2 THEN T1.���ܰ�2 " _
                                 & "WHEN T4.�ܰ����� = 3 THEN T1.���ܰ�3 ELSE 0 END, " _
                 & "���ΰ� = CASE WHEN T4.�ܰ����� = 1 THEN ROUND(T1.���ܰ�1 * (" & PB_curVatRate & "), 0, 1) " _
                                 & "WHEN T4.�ܰ����� = 2 THEN ROUND(T1.���ܰ�2 * (" & PB_curVatRate & "), 0, 1) " _
                                 & "WHEN T4.�ܰ����� = 3 THEN ROUND(T1.���ܰ�3 * (" & PB_curVatRate & "), 0, 1) ELSE 0 END, " _
                 & "����� = CASE WHEN T4.�ܰ����� = 1 THEN T1.���ܰ�1 + ROUND(T1.���ܰ�1 * (" & PB_curVatRate & "), 0, 1) " _
                                 & "WHEN T4.�ܰ����� = 2 THEN T1.���ܰ�2 + ROUND(T1.���ܰ�2 * (" & PB_curVatRate & "), 0, 1) " _
                                 & "WHEN T4.�ܰ����� = 3 THEN T1.���ܰ�3 + ROUND(T1.���ܰ�3 * (" & PB_curVatRate & "), 0, 1) ELSE 0 END, "
       strJoin = "INNER JOIN ����ó T4 ON T4.������ڵ� = T1.������ڵ� AND T4.����ó�ڵ� = '" & Trim(PB_strSupplierCode) & "' "
       strWhere = strWhere + IIf(Trim(strWhere) = "", "WHERE ", "AND ") _
                & "T4.����ó�ڵ� = '" & Trim(PB_strSupplierCode) & "' "
    End If
    strWhere = strWhere + IIf(Trim(strWhere) = "", "WHERE ", "AND ") & "T2.��뱸�� = 0 "
    P_adoRec.CursorLocation = adUseClient
    strSQL = "SELECT T1.�з��ڵ�, T1.�����ڵ�, T2.����� AS �����, " _
                  & "T1.�ָ���ó�ڵ� AS �ָ���ó�ڵ�, ISNULL(T3.����ó��, '') AS �ָ���ó��, T2.�԰� AS �԰�, T2.���� AS ����," _
                  & "T2.����� AS �����, T2.�������� AS ��������, T2.��뱸�� AS ��뱸��, ISNULL(T1.�������, 0) AS �������, "
    strSQL = strSQL + strSelect
    strSQL = strSQL _
                  & "(SELECT ISNULL(SUM(�԰������ - ��������), 0) FROM ������帶�� " _
                    & "WHERE �з��ڵ� = T1.�з��ڵ� AND �����ڵ� = T1.�����ڵ� " _
                      & "AND ������ڵ� = T1.������ڵ� " _
                      & "AND ������� >= '" & Left(PB_regUserinfoU.UserClientDate, 4) + "00" & "' " _
                      & "AND ������� <  '" & Left(PB_regUserinfoU.UserClientDate, 6) & "') AS �̿����,"
    strSQL = strSQL _
                 & "(SELECT ISNULL(SUM(�԰���� - ������), 0) FROM �������⳻�� " _
                   & "WHERE �з��ڵ� = T1.�з��ڵ� AND �����ڵ� = T1.�����ڵ� " _
                     & "AND ������ڵ� = T1.������ڵ� " _
                     & "AND ��������� BETWEEN '" & Left(PB_regUserinfoU.UserClientDate, 6) + "01" & "' " _
                                            & "AND '" & PB_regUserinfoU.UserClientDate & "') AS �ݿ����, "
    strSQL = strSQL _
                 & "(SELECT ISNULL(SUM(���ַ�), 0) FROM ���ֳ��� " _
                   & "WHERE �����ڵ� = (T1.�з��ڵ� + T1.�����ڵ�) " _
                     & "AND ������ڵ� = T1.������ڵ� AND �����ڵ� = 1 AND ��뱸�� = 0 " _
                     & "AND �������� >= '" & PB_regUserinfoU.UserClientDate & "') AS �԰���, "
    strSQL = strSQL _
                 & "(SELECT ISNULL(SUM(����), 0) FROM �������� " _
                   & "WHERE �����ڵ� = (T1.�з��ڵ� + T1.�����ڵ�) " _
                     & "AND ������ڵ� = T1.������ڵ� AND �����ڵ� = 1 AND ��뱸�� = 0 " _
                     & "AND �������� >= '" & PB_regUserinfoU.UserClientDate & "') AS ����� "
    strSQL = strSQL _
             & "FROM ������� T1 " _
            & "INNER JOIN ���� T2 " _
                    & "ON T2.�з��ڵ� = T1.�з��ڵ� AND T2.�����ڵ� = T1.�����ڵ� " _
             & "LEFT JOIN ����ó T3 ON T3.������ڵ� = T1.������ڵ� AND T3.����ó�ڵ� = T1.�ָ���ó�ڵ� " _
            & "" & strJoin & " " _
            & "" & strWhere & " " _
            & "" & strOrderBy & " "
    On Error GoTo ERROR_TABLE_SELECT
    P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
    vsfg1.Rows = P_adoRec.RecordCount + vsfg1.FixedRows
    If P_adoRec.RecordCount = 0 Then
       P_adoRec.Close
       Screen.MousePointer = vbDefault
       Exit Sub
    Else
       With vsfg1
            .Cell(flexcpForeColor, 0, 0, .FixedRows - 1, .Cols - 1) = vbBlack
            .Rows = P_adoRec.RecordCount + 1
            If .Rows <= PC_intRowCnt Then
               .ScrollBars = flexScrollBarHorizontal
            Else
               .ScrollBars = flexScrollBarBoth
            End If
            Do Until P_adoRec.EOF
               lngR = lngR + 1
               .TextMatrix(lngR, 0) = P_adoRec("�з��ڵ�") + P_adoRec("�����ڵ�")
               .TextMatrix(lngR, 1) = P_adoRec("�����")
               .TextMatrix(lngR, 2) = P_adoRec("�̿����") + P_adoRec("�ݿ����") + P_adoRec("�԰���") + P_adoRec("�����")
               If P_adoRec("�������") <> 0 And _
                  .ValueMatrix(lngR, 2) < P_adoRec("�������") Then
                  .Cell(flexcpForeColor, lngR, 2, lngR, 2) = vbRed
                  .Cell(flexcpFontBold, lngR, 2, lngR, 2) = True
               End If
               .TextMatrix(lngR, 3) = P_adoRec("�̿����") + P_adoRec("�ݿ����")
               .TextMatrix(lngR, 4) = P_adoRec("�ָ���ó�ڵ�")
               .TextMatrix(lngR, 5) = P_adoRec("�ָ���ó��")
               .TextMatrix(lngR, 6) = .TextMatrix(lngR, 0)
               .Cell(flexcpData, lngR, 6, lngR, 6) = .TextMatrix(lngR, 6)
               .TextMatrix(lngR, 7) = P_adoRec("�԰�")
               .TextMatrix(lngR, 8) = P_adoRec("����")
               
               .TextMatrix(lngR, 9) = IIf(IsNull(P_adoRec("�����")), 0, P_adoRec("�����"))
               .TextMatrix(lngR, 10) = P_adoRec("�԰�ܰ�")
               .TextMatrix(lngR, 11) = P_adoRec("�԰�ΰ�")
               .TextMatrix(lngR, 12) = P_adoRec("�԰���")
               .TextMatrix(lngR, 13) = P_adoRec("���ܰ�")
               .TextMatrix(lngR, 14) = P_adoRec("���ΰ�")
               .TextMatrix(lngR, 15) = P_adoRec("�����")
               '.TextMatrix(lngR, 16) = P_adoRec("������")
               .TextMatrix(lngR, 17) = IIf(IsNull(P_adoRec("��������")), 0, P_adoRec("��������"))
               If .ValueMatrix(lngR, 17) = 0 Then
                  .TextMatrix(lngR, 18) = "�����"
               Else
                  .TextMatrix(lngR, 18) = "��  ��"
               End If
               lngRR = 1
               'If P_intFindGbn = 1 Then
               '   If txtCode.Text = Trim(.TextMatrix(lngR, 0)) And _
               '      txtSupplierCode.Text = Trim(.TextMatrix(lngR, 4)) Then
               '      lngRR = lngR
               '   End If
               'ElseIf _
               '   P_intFindGbn = 2 Then
               '   If txtName.Text = Trim(.TextMatrix(lngR, 1)) And _
               '      txtSupplierCode.Text = Trim(.TextMatrix(lngR, 4)) And _
               '      txtSize.Text = Trim(.TextMatrix(lngR, 7)) And _
               '      txtUnit.Text = Trim(.TextMatrix(lngR, 8)) Then
               '      lngRR = lngR
               '   End If
               'ElseIf _
               '   P_intFindGbn = 3 Then
               '   If txtCode.Text = Trim(.TextMatrix(lngR, 0)) And _
               '      txtName.Text = Trim(.TextMatrix(lngR, 1)) And _
               '      txtSupplierCode.Text = Trim(.TextMatrix(lngR, 4)) And _
               '      txtSize.Text = Trim(.TextMatrix(lngR, 7)) And _
               '      txtUnit.Text = Trim(.TextMatrix(lngR, 8)) Then
               '      lngRR = lngR
               '   End If
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
    Screen.MousePointer = vbDefault
    Exit Sub
ERROR_TABLE_SELECT:
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "����ü� �б� ����"
    Unload Me
    Exit Sub
End Sub

Private Sub SubOther_FILL()
Dim strSQL        As String
Dim intIndex      As Integer
    Exit Sub
ERROR_TABLE_SELECT:
    MsgBox Err.Number & Err.Description & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "����ü� �б� ����"
    Unload Me
    Exit Sub
End Sub

'+----------+
'/// ESC ///
'+----------+
Private Sub SubEscape()
    PB_strFMCCallFormName = ""
    PB_strMaterialsCode = ""
    PB_strMaterialsName = ""
    PB_strSupplierCode = ""
    Unload Me
End Sub


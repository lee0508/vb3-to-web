VERSION 5.00
Begin VB.Form Form1 
   ClientHeight    =   3435
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4425
   LinkTopic       =   "Form2"
   ScaleHeight     =   3435
   ScaleWidth      =   4425
   StartUpPosition =   3  'Windows 기본값
   Begin VB.CommandButton cmdOpen 
      Height          =   405
      Left            =   2835
      Picture         =   "Form1.frx":0000
      Style           =   1  '그래픽
      TabIndex        =   17
      Top             =   1890
      Width           =   1350
   End
   Begin VB.CommandButton cmdCalc 
      Height          =   405
      Left            =   2835
      Picture         =   "Form1.frx":0D4A
      Style           =   1  '그래픽
      TabIndex        =   16
      Top             =   1440
      Width           =   1350
   End
   Begin VB.CommandButton cmdOk 
      Height          =   405
      Left            =   45
      Picture         =   "Form1.frx":1AF1
      Style           =   1  '그래픽
      TabIndex        =   15
      Top             =   540
      Width           =   1350
   End
   Begin VB.CommandButton cmdExcel 
      Height          =   405
      Left            =   45
      Picture         =   "Form1.frx":28B6
      Style           =   1  '그래픽
      TabIndex        =   14
      Top             =   990
      Width           =   1350
   End
   Begin VB.CommandButton cmdEsc 
      Height          =   405
      Left            =   1440
      Picture         =   "Form1.frx":3835
      Style           =   1  '그래픽
      TabIndex        =   13
      Top             =   540
      Width           =   1350
   End
   Begin VB.CommandButton cmdPrint 
      Height          =   405
      Left            =   45
      Picture         =   "Form1.frx":4506
      Style           =   1  '그래픽
      TabIndex        =   12
      Top             =   1440
      Width           =   1350
   End
   Begin VB.CommandButton cmdExit 
      DisabledPicture =   "Form1.frx":529A
      Height          =   405
      Left            =   2835
      Picture         =   "Form1.frx":60A4
      Style           =   1  '그래픽
      TabIndex        =   11
      Top             =   540
      Width           =   1350
   End
   Begin VB.CommandButton cmdDel 
      Height          =   405
      Left            =   2835
      Picture         =   "Form1.frx":6F07
      Style           =   1  '그래픽
      TabIndex        =   10
      Top             =   90
      Width           =   1350
   End
   Begin VB.CommandButton cmdSave 
      Height          =   405
      Left            =   1440
      Picture         =   "Form1.frx":7D80
      Style           =   1  '그래픽
      TabIndex        =   9
      Top             =   1440
      Width           =   1350
   End
   Begin VB.CommandButton cmdAdd 
      Height          =   405
      Left            =   45
      Picture         =   "Form1.frx":8A8E
      Style           =   1  '그래픽
      TabIndex        =   8
      Top             =   90
      Width           =   1350
   End
   Begin VB.CommandButton cmdModify 
      Height          =   405
      Left            =   1440
      Picture         =   "Form1.frx":98B5
      Style           =   1  '그래픽
      TabIndex        =   7
      Top             =   90
      Width           =   1350
   End
   Begin VB.CommandButton cmdFind 
      Height          =   570
      Left            =   1890
      Picture         =   "Form1.frx":A63E
      Style           =   1  '그래픽
      TabIndex        =   6
      Top             =   2025
      Width           =   465
   End
   Begin VB.CommandButton cmdDown 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   945
      Picture         =   "Form1.frx":B680
      Style           =   1  '그래픽
      TabIndex        =   5
      Top             =   2115
      Width           =   420
   End
   Begin VB.CommandButton cmdUp 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   1395
      Picture         =   "Form1.frx":B883
      Style           =   1  '그래픽
      TabIndex        =   4
      Top             =   2115
      Width           =   420
   End
   Begin VB.CommandButton cmdRefresh 
      Height          =   405
      Left            =   1440
      Picture         =   "Form1.frx":BA88
      Style           =   1  '그래픽
      TabIndex        =   3
      Top             =   990
      Width           =   1350
   End
   Begin VB.CommandButton cmdSearch 
      Height          =   405
      Left            =   2835
      Picture         =   "Form1.frx":C96A
      Style           =   1  '그래픽
      TabIndex        =   2
      Top             =   990
      Width           =   1350
   End
   Begin VB.CommandButton cmdRight 
      Height          =   465
      Left            =   45
      Picture         =   "Form1.frx":D6F5
      Style           =   1  '그래픽
      TabIndex        =   1
      Top             =   2115
      Width           =   420
   End
   Begin VB.CommandButton cmdLeft 
      Height          =   465
      Left            =   495
      Picture         =   "Form1.frx":D8F4
      Style           =   1  '그래픽
      TabIndex        =   0
      Top             =   2115
      Width           =   420
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Command1"
      Height          =   285
      Left            =   90
      TabIndex        =   5
      Top             =   1035
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   330
      Left            =   1485
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   1035
      Width           =   2760
   End
   Begin VB.TextBox Text1 
      Height          =   330
      Left            =   1485
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   270
      Width           =   2760
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   285
      Left            =   90
      TabIndex        =   1
      Top             =   270
      Width           =   1095
   End
   Begin proj_LineEX.LineEX LineEX5 
      Height          =   30
      Left            =   0
      TabIndex        =   7
      Top             =   2205
      Width           =   4650
      _ExtentX        =   8202
      _ExtentY        =   53
      line_style      =   1
   End
   Begin proj_LineEX.LineEX LineEX3 
      Height          =   1500
      Left            =   2250
      TabIndex        =   6
      Top             =   1620
      Width           =   30
      _ExtentX        =   53
      _ExtentY        =   0
      line_orientation=   1
      line_style      =   1
   End
   Begin proj_LineEX.LineEX LineEX2 
      Height          =   60
      Left            =   45
      TabIndex        =   3
      Top             =   1530
      Width           =   4605
      _ExtentX        =   8123
      _ExtentY        =   106
      line_type       =   3
   End
   Begin proj_LineEX.LineEX LineEX1 
      Height          =   90
      Left            =   45
      TabIndex        =   0
      Top             =   855
      Width           =   4605
      _ExtentX        =   8123
      _ExtentY        =   159
      line_type       =   5
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


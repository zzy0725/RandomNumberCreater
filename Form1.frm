VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Random 1 to x (Beta)"
   ClientHeight    =   3225
   ClientLeft      =   9180
   ClientTop       =   2745
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "不公平？"
      Height          =   780
      Left            =   3600
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2280
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   855
      IMEMode         =   2  'OFF
      Left            =   1440
      TabIndex        =   0
      TabStop         =   0   'False
      Text            =   "50"
      Top             =   2280
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "选出"
      Height          =   855
      Left            =   2280
      MaskColor       =   &H00808080&
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "候选个数"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "看看谁是""LuckyDog"""
      BeginProperty Font 
         Name            =   "楷体_GB2312"
         Size            =   42
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1815
      Left            =   240
      OLEDropMode     =   1  'Manual
      TabIndex        =   2
      Top             =   240
      Width           =   4095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Function Rndz(a As Long, b As Long)
'Randomize
 '   Rndz = Int((n - 1 + 1) * Rnd() + 1)
End Function
Private Sub Command1_Click()
'm = Int(Rnd * (n - 1 + 1)) + 1
Randomize
Label1.Caption = Int((Text1.Text - 1 + 0) * Rnd() + 1)
End Sub

Private Sub Command2_Click()
Dialog.Visible = True
End Sub


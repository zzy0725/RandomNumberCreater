VERSION 5.00
Begin VB.Form Dialog 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "关于 与 源代码"
   ClientHeight    =   1635
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6060
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1635
   ScaleWidth      =   6060
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "源代码如下。更多请联系技术人员。"
      BeginProperty Font 
         Name            =   "楷体_GB2312"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   5655
   End
   Begin VB.Label Label1 
      Caption         =   $"Dialog.frx":0000
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   5775
   End
End
Attribute VB_Name = "Dialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub OKButton_Click()

End Sub


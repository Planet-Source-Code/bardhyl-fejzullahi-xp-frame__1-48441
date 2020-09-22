VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H8000000E&
   Caption         =   "Form1"
   ClientHeight    =   4695
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7125
   LinkTopic       =   "Form1"
   ScaleHeight     =   4695
   ScaleWidth      =   7125
   StartUpPosition =   3  'Windows Default
   Begin ButoonWizard.BciFrame BciFrame1 
      Height          =   3495
      Left            =   600
      TabIndex        =   0
      Top             =   600
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   6165
      Caption         =   "vera && bardhi"
      BackColor       =   12171705
      BackHeadColor   =   9474192
      BorderColor     =   -2147483640
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Command1"
         Height          =   375
         Left            =   2040
         MaskColor       =   &H00808080&
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   2640
         Width           =   1335
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
MsgBox "hahha!"
End
End Sub

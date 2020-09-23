VERSION 5.00
Begin VB.Form frmmain 
   Caption         =   "Test Form to stop password unmaskers"
   ClientHeight    =   4020
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmmain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4020
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   1815
      Left            =   120
      TabIndex        =   6
      Top             =   2040
      Width           =   4455
      Begin VB.TextBox txtread 
         Height          =   1455
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Text            =   "frmmain.frx":0442
         Top             =   240
         Width           =   4215
      End
   End
   Begin VB.CommandButton Cmd_E 
      Caption         =   "End Tick Function"
      Enabled         =   0   'False
      Height          =   735
      Left            =   1560
      TabIndex        =   5
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CommandButton Cmd_s 
      Caption         =   "Start Tick Function"
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CommandButton Cmd_tmr1d 
      Caption         =   "Disable Timer Function"
      Enabled         =   0   'False
      Height          =   615
      Left            =   1560
      TabIndex        =   3
      Top             =   480
      Width           =   1335
   End
   Begin VB.CommandButton cmd_tmr1 
      Caption         =   "Enable timer function"
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3600
      Top             =   480
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Click to Exit"
      Height          =   615
      Left            =   3120
      TabIndex        =   1
      Top             =   960
      Width           =   1455
   End
   Begin VB.TextBox txt_pw 
      Height          =   325
      IMEMode         =   3  'DISABLE
      Left            =   240
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   120
      Width           =   3975
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cmd_E_Click()
    bEnd = True
    Cmd_s.Enabled = True
    Cmd_E.Enabled = False
End Sub

Private Sub Cmd_s_Click()
    bEnd = False
    Cmd_s.Enabled = False
    Cmd_E.Enabled = True
    TickLoopreset
End Sub

Private Sub cmd_tmr1_Click()
    If Timer1.Enabled = False Then Timer1.Enabled = True
    If Cmd_tmr1d.Enabled = False Then Cmd_tmr1d.Enabled = True
    cmd_tmr1.Enabled = False
End Sub

Private Sub Cmd_tmr1d_Click()
    If Timer1.Enabled = True Then Timer1.Enabled = False
    If cmd_tmr1.Enabled = False Then cmd_tmr1.Enabled = True
    Cmd_tmr1d.Enabled = False
End Sub

Private Sub cmdExit_Click()
        End
End Sub

Private Sub Form_Load()
    bEnd = False
End Sub

Private Sub Timer1_Timer()
    tmrreset
End Sub

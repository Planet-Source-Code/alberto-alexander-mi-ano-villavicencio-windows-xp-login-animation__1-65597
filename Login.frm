VERSION 5.00
Begin VB.Form frmLogin 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   3330
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6195
   Icon            =   "Login.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Login.frx":000C
   ScaleHeight     =   222
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   413
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Login.Progress Progress1 
      Height          =   90
      Left            =   0
      Top             =   1320
      Width           =   6195
      _ExtentX        =   10927
      _ExtentY        =   159
      AutoSize        =   0   'False
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Pause animaton"
      Height          =   435
      Left            =   4575
      TabIndex        =   1
      Top             =   2730
      Width           =   1515
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start animation"
      Height          =   435
      Left            =   2955
      TabIndex        =   0
      Top             =   2730
      Width           =   1515
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Command1_Click()
  Progress1.StartAnimation
End Sub

Private Sub Command3_Click()
  Progress1.StopAnimation
End Sub

Private Sub Form_Load()
  Set Progress1.Picture = LoadResPicture(102, vbResBitmap)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  Progress1.EndAnimation
  Unload Me
  End
End Sub

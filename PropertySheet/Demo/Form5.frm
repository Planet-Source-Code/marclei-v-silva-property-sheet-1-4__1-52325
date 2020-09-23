VERSION 5.00
Begin VB.Form Form5 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Picture selection"
   ClientHeight    =   1230
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6225
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1230
   ScaleWidth      =   6225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Select a picture"
      Height          =   885
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4815
      Begin VB.Image Image1 
         Height          =   480
         Index           =   0
         Left            =   180
         Picture         =   "Form5.frx":0000
         Stretch         =   -1  'True
         Top             =   240
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   1
         Left            =   810
         Picture         =   "Form5.frx":0642
         Stretch         =   -1  'True
         Top             =   240
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   2
         Left            =   1455
         Picture         =   "Form5.frx":0C7C
         Stretch         =   -1  'True
         Top             =   240
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   3
         Left            =   2085
         Picture         =   "Form5.frx":216E
         Stretch         =   -1  'True
         Top             =   240
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   4
         Left            =   2715
         Picture         =   "Form5.frx":24B0
         Stretch         =   -1  'True
         Top             =   240
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   5
         Left            =   3360
         Picture         =   "Form5.frx":28F2
         Stretch         =   -1  'True
         Top             =   240
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   6
         Left            =   3990
         Picture         =   "Form5.frx":2BFC
         Stretch         =   -1  'True
         Top             =   240
         Width           =   480
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Default         =   -1  'True
      Height          =   495
      Left            =   5070
      TabIndex        =   0
      Top             =   330
      Width           =   1035
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_PictureIndex As Integer

Public Property Get PictureIndex() As Integer
    PictureIndex = m_PictureIndex
End Property

Private Sub Command1_Click()
    m_PictureIndex = -1
    Unload Me
End Sub

Private Sub Form_Load()
    m_PictureIndex = -1
End Sub

Private Sub Form_Resize()
'    Command1.Left = (ScaleWidth - Command1.Width) / 2
'    Command1.Top = (ScaleHeight - Command1.Height) / 2
End Sub

Private Sub Image1_Click(Index As Integer)
    m_PictureIndex = Index
    Unload Me
End Sub

Private Sub Image1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If m_PictureIndex <> -1 Then
        Image1(m_PictureIndex).BorderStyle = 0
    End If
    Image1(Index).BorderStyle = 1
    m_PictureIndex = Index
End Sub

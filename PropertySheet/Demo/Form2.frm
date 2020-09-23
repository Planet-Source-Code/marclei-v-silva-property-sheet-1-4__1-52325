VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add New Property"
   ClientHeight    =   2490
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5040
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2490
   ScaleWidth      =   5040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Default         =   -1  'True
      Height          =   375
      Left            =   2400
      TabIndex        =   9
      Top             =   2010
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3720
      TabIndex        =   10
      Top             =   2010
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Property definition"
      Height          =   1755
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4815
      Begin VB.TextBox txtValue 
         Height          =   288
         Left            =   2160
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   600
         Width           =   1812
      End
      Begin VB.ComboBox cmbCategory 
         Height          =   288
         Left            =   2160
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1200
         Width           =   1812
      End
      Begin VB.ComboBox cmbType 
         Height          =   288
         Left            =   210
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1200
         Width           =   1812
      End
      Begin VB.TextBox txtProperty 
         Height          =   288
         Left            =   210
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   600
         Width           =   1812
      End
      Begin VB.Label Label1 
         Caption         =   "Default value"
         Height          =   288
         Index           =   3
         Left            =   2160
         TabIndex        =   3
         Top             =   360
         Width           =   1812
      End
      Begin VB.Label Label1 
         Caption         =   "Category"
         Height          =   288
         Index           =   2
         Left            =   2160
         TabIndex        =   7
         Top             =   960
         Width           =   1812
      End
      Begin VB.Label Label1 
         Caption         =   "Type"
         Height          =   288
         Index           =   1
         Left            =   210
         TabIndex        =   5
         Top             =   960
         Width           =   1812
      End
      Begin VB.Label Label1 
         Caption         =   "Property"
         Height          =   288
         Index           =   0
         Left            =   210
         TabIndex        =   1
         Top             =   360
         Width           =   1812
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdd_Click()
    Dim ValueType As Long
    Dim Prop As TProperty
    
    On Error Resume Next
    ' check if property already exists
    Set Prop = fMain.ps2.Properties(txtProperty.Text)
    If Prop Is Nothing Then
        ' property does not exist so...add to the list
        ValueType = cmbType.ItemData(cmbType.ListIndex)
        With fMain.ps2
            With .Categories(cmbCategory.Text)
                .Properties.Add _
                    txtProperty.Text, _
                    txtValue.Text, _
                    ValueType
            End With
        End With
    Else
        ' property exists...inform...
        MsgBox "Property already exists", vbExclamation
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim i As Integer
    
    txtProperty.Text = ""
    txtValue.Text = ""
    cmbCategory.Clear
    For i = 1 To fMain.ps2.Categories.Count
        cmbCategory.AddItem fMain.ps2.Categories(i).Caption
    Next
    
    cmbType.Clear
    AddItem cmbType, "psNone", -1
    AddItem cmbType, "psCustom", 0
    AddItem cmbType, "psInteger", 2
    AddItem cmbType, "psLong", 3
    AddItem cmbType, "psSingle", 4
    AddItem cmbType, "psDouble", 5
    AddItem cmbType, "psCurrency", 6
    AddItem cmbType, "psDate", 7
    AddItem cmbType, "psString", 8
    AddItem cmbType, "psObject", 9
    AddItem cmbType, "psBoolean", 11
    AddItem cmbType, "psDecimal", 14
    AddItem cmbType, "psByte", 17
    AddItem cmbType, "psFont", 240
    AddItem cmbType, "psPicture", 241
    AddItem cmbType, "psFile", 242
    AddItem cmbType, "psColor", 243
    AddItem cmbType, "psDropDownList", 244
    AddItem cmbType, "psCombo", 245
    AddItem cmbType, "psTime", 246
    AddItem cmbType, "psLongText", 247
    AddItem cmbType, "psFolder", 248
    AddItem cmbType, "psDropDownCheckList", 249
    
End Sub

Private Sub AddItem(Ctl As ComboBox, Caption As String, Value As Long)
    Ctl.AddItem Caption
    Ctl.ItemData(Ctl.NewIndex) = Value
End Sub

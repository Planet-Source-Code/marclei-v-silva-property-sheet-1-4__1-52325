VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "*\A..\Source\pPropertySheet.vbp"
Begin VB.Form Form1 
   Caption         =   "PropertySheet Demo"
   ClientHeight    =   7815
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   8955
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   7.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7815
   ScaleWidth      =   8955
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picTop 
      Align           =   1  'Align Top
      BackColor       =   &H8000000C&
      BorderStyle     =   0  'None
      Height          =   525
      Left            =   0
      ScaleHeight     =   525
      ScaleWidth      =   8955
      TabIndex        =   3
      Top             =   0
      Width           =   8955
      Begin VB.Image Image1 
         Height          =   480
         Left            =   30
         Picture         =   "Form1.frx":0CCA
         Top             =   30
         Width           =   480
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "PropertySheet control"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000018&
         Height          =   375
         Left            =   510
         TabIndex        =   4
         Top             =   60
         Width           =   6705
      End
   End
   Begin PropertySheet.TPropertySheet ps1 
      Height          =   6735
      Left            =   4080
      TabIndex        =   2
      Top             =   960
      Width           =   4725
      _ExtentX        =   8334
      _ExtentY        =   11880
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty CatFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowToolbar     =   0   'False
   End
   Begin PropertySheet.TPropertySheet ps2 
      Height          =   6045
      Left            =   120
      TabIndex        =   1
      Top             =   570
      Width           =   3795
      _ExtentX        =   6694
      _ExtentY        =   10663
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AutoSelect      =   0   'False
      CatBackColor    =   14737632
      CatForeColor    =   0
      BeginProperty CatFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DescriptionHeight=   700
      ShowDescription =   0   'False
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   7185
      Left            =   3990
      TabIndex        =   0
      Top             =   570
      Width           =   4905
      _ExtentX        =   8652
      _ExtentY        =   12674
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Alphabetic"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Categorized"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2010
      Top             =   6720
      _ExtentX        =   794
      _ExtentY        =   794
      BackColor       =   16777215
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   25
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1994
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1AA6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1BB8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1CCA
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1DDC
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1EEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2000
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2112
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2224
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2336
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2448
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":255A
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":266C
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":277E
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2890
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":29A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2AB4
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2BC6
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2CD8
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2DEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2EFC
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":300E
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":3120
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":3232
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":3344
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileLoad 
         Caption         =   "&Load"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
      End
      Begin VB.Menu mnuFileClear 
         Caption         =   "&Clear"
      End
      Begin VB.Menu mnuFileRestore 
         Caption         =   "&Restore"
      End
      Begin VB.Menu mnuBar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAddProperty 
         Caption         =   "&New Property"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "&Exit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' *******************************************************
' Project      : DemoPropertySheet
' Written By   : Marclei V Silva (MVS)
' Programmer   : Marclei V Silva (MVS) [Spnorte Consultoria de Informática]
' Date Writen  : 06/16/2000 -- 08:52:54
' Description  : This project will demonstrate the use
'              : of PropertySheet control
'              :
'              :
' *******************************************************
Option Explicit

' Enumerate Property Sheet images
Enum psStandardImages
    psImgClosedFolder = 1
    psImgOpenFolder = 2
    psImgUser = 3
    psImgHome = 4
    psImgPhone = 5
    psImgFax = 6
    psImgWebPage = 7
    psImgHyperlink = 8
    psImgMail = 9
    psImgLock = 10
    psImgPaperClip = 11
    psImgNotes = 12
    psImgPicture1 = 13
    psImgCalendar1 = 14
    psImgCalendar2 = 15
    psImgClock = 16
    psImgFont = 17
    psImgPicture2 = 18
    psImgFile = 19
    psImgFontColor = 20
    psImgBackColor = 21
    psImgLineColor = 22
    psImgZoom = 23
    psImgWidth = 24
    psImgHeight = 25
End Enum

Private Sub Form_Load()
    ' select propertysheet style
    TabStrip1_Click
    ' fill propertysheet1 properties
    AddPropertiesPS1
    ' fill propertysheet1 properties
    AddPropertiesPS2
    ' init description pane
End Sub

Private Sub Form_Resize()
    Static Resizing As Boolean
    If Not Resizing Then
      
        Resizing = True
      
       On Error GoTo err_resize
        If Height < 2325 Then Height = 2325
        If Width < 4560 Then Width = 4560
   
        ps2.Move 30, ps2.Top, ScaleWidth / 2 - 45, ScaleHeight - Me.picTop.Height - 45
        TabStrip1.Move ScaleWidth / 2 + 30, TabStrip1.Top, ScaleWidth / 2 - 45, ScaleHeight - (735)
        ps1.Move ScaleWidth / 2 + 90, TabStrip1.ClientTop, TabStrip1.ClientWidth, TabStrip1.ClientHeight
err_resize:
        Resizing = False
      
    End If

End Sub

Sub AddPropertiesPS1()
    With ps1
'<Added by: Project Administrator at: 1/4/2004-19:59:14 on machine: ZEUS>
        .Redraw = False
'</Added by: Project Administrator at: 1/4/2004-19:59:14 on machine: ZEUS>
        .ImageList = ImageList1
        .ShowToolTips = True
        With .Categories
            With .Add("Appearance", , "@Use these properties to change the" & vbCrLf & "PropertyList2 appearance").Properties
                .Add "BackColor", ps2.BackColor, psColor, , psImgBackColor, , "Returns/sets the background color of the object"
                .Add "CatBackColor", ps2.CatBackColor, psColor, , psImgBackColor, , "Returns/sets catagory cell background color"
                .Add "SelBackColor", ps2.SelBackColor, psColor, , psImgBackColor, , "Returns/sets selection background color"
                .Add "CatForeColor", ps2.CatForeColor, psColor, , psImgFontColor, , "Returns/sets category foreground color used to display text and graphics of an object"
                .Add "ForeColor", ps2.ForeColor, psColor, , psImgFontColor, , "Returns/sets foreground color used to display text and graphics of an object"
                .Add "SelForeColor", ps2.SelForeColor, psColor, , psImgFontColor, , "Returns/sets selection foreground color used to display text and graphics of an object"
                .Add "GridColor", ps2.GridColor, psColor, , psImgLineColor, , "Returns/sets object grid color"
                With .Add("BorderStyle", 0, psDropDownList)
                    .ListValues.Add 0, "0 - psBorderNone"
                    .ListValues.Add 1, "1 - psBorderSingle"
                    .Value = ps2.BorderStyle
                    .Description = "Returns/sets the border style for the object"
                End With
                With .Add("Appearance", 0, psDropDownList)
                    .ListValues.Add 0, "0 - psFlat"
                    .ListValues.Add 1, "1 - ps3D"
                    .Value = ps2.Appearance
                End With
                .Add("CatFont", ps2.CatFont, psFont).Format = "n (c)"
                .Add "Font", ps2.Font, psFont
                .Add "NameWidth", ps2.NameWidth
                With .Add("ShowCategories", ps2.ShowCategories)
                    With .ListValues
                        .Item(1).Caption = "No"
                        .Item(2).Caption = "Yes"
                    End With
                End With
                With .Add("Tooltips", ps2.ShowToolTips, , , psImgPicture2)
                    With .ListValues
                        .Item(1).Caption = "Hide"
                        .Item(2).Caption = "Show"
                    End With
                End With
                .Add "ShowToolbar", True
                .Add "ShowDescription", False
'<Added by: Project Administrator at: 31/3/2004-21:16:10 on machine: ZEUS>
                With .Add("EffectStyle", psNormal, psDropDownList)
                    With .ListValues
                        .Add psNormal, "psNormal"
                        .Add psSmooth, "psSmooth"
                    End With
                End With
'</Added by: Project Administrator at: 31/3/2004-21:16:10 on machine: ZEUS>
                With .Add("DescriptionHeight", ps2.DescriptionHeight, psInteger)
                    .UpDownIncrement = 50
                End With
            End With
            With .Add("Behavior", , "@Set the control behavior").Properties
                .Add "AllowEmptyValues", ps2.AllowEmptyValues
                .Add "AutoSelect", ps2.AutoSelect
                .Add "Expandable Categories", ps2.ExpandableCategories
                .Add "RequiresEnter", ps2.RequiresEnter
                .Add "Visible", True
            End With
            With .Add("Misc", , "@Other properties")
                With .Properties
                    .Add("(About)", "", psCustom, , psImgNotes, "Click the button for information about this control", "Show propertysheet about box").ForeColor = vbBlue
                    .Add("(Revisions)", "", psCustom, , psImgWebPage, "Click the button for revision file", "Open the revision file for PropertySheet control").ForeColor = vbBlue
                    .Add("(Readme)", "", psCustom, , psImgFile, "Click the button for read-me file", "Open the readme file for PropertySheet control").ForeColor = vbBlue
                End With
            End With
            With .Add("Position", , "@Fields in red are read-only")
                With .Properties
                    With .Add("Left", ps2.Left, , True)
                        .ForeColor = vbRed
                    End With
                    .Add("Width", ps2.Width, , , psImgWidth).SetRange 2100, 2790
                    .Add("Height", ps2.Height, , , psImgHeight).SetRange 100, 3300
                    .Add("Top", ps2.Top, , True).ForeColor = vbRed
                End With
            End With
            With .Add("Formats").Properties
                With .Add("ColorFormat", "RRGGBB", psCombo).ListValues
                    .Add "&HeH&", "VB"
                    .Add "$e", "Delphi"
                    .Add "#m", "HTML"
                    .Add "r g b", "Red Green Blue"
                End With
                With .Add("Date Format", "dd-MMM-yyyy", psCombo).ListValues
                    .Add "Long Date"
                    .Add "Medium Date"
                    .Add "Short Date"
                    .Add """Today is"" dddd dd "", a really nice day.""", "Really Long Date"
                End With
                With .Add("Boolean Format", 0, psBoolean).ListValues
                    .Item(1).Caption = "Like combobox"
                    .Item(2).Caption = "Like checkbox"
                End With
            End With
        End With
'<Added by: Project Administrator at: 1/4/2004-19:59:24 on machine: ZEUS>
        .Redraw = True
'</Added by: Project Administrator at: 1/4/2004-19:59:24 on machine: ZEUS>
    End With
End Sub

Private Sub ps1_Browse(ByVal Left As Variant, ByVal Top As Variant, ByVal Width As Variant, ByVal Prop As PropertySheet.TProperty)
    Select Case Prop.Caption
        Case "(About)"
            Dim fAbout As New Form4
            
            fAbout.Show vbModal
            Unload fAbout
            Set fAbout = Nothing
            
        Case "(Revisions)"
            OpenFile App.Path & "\revisions.rtf", "Revisions"
        
        Case "(Readme)"
            OpenFile App.Path & "\readme.rtf", "Readme"
    End Select
End Sub

Private Sub ps1_EditError(ErrMessage As String)
    MsgBox ErrMessage
End Sub

Private Sub ps1_GetDisplayString(ByVal Prop As PropertySheet.TProperty, DisplayString As String, UseDefault As Boolean)
    Select Case Prop.Caption
        ' nothing
    End Select
End Sub

Private Sub ps1_PropertyChanged(ByVal Prop As TProperty, NewValue As Variant, Cancel As Boolean)
    '<EhHeader>
    On Error GoTo ps1_PropertyChanged_Err
    '</EhHeader>
    Dim Txt As String
    
    With ps2
        Select Case Prop.Caption
            Case "Boolean Format"
                .Categories("Other Types").Properties("Boolean").Format = IIf(NewValue, "checkbox", "")
                
            Case "Expandable Categories"
                .ExpandableCategories = NewValue
            
            Case "AllowEmptyValues"
                .AllowEmptyValues = NewValue
            
            Case "AutoSelect"
                .AutoSelect = NewValue
                
            Case "BackColor"
                .BackColor = NewValue
            
            Case "BorderStyle"
                .BorderStyle = NewValue
            
            Case "Appearance"
                .Appearance = NewValue
                
            Case "CatBackColor"
                .CatBackColor = NewValue
                
            Case "CatFont"
                Set .CatFont = NewValue
                
            Case "CatForeColor"
                .CatForeColor = NewValue
                
            Case "ColorFormat"
                .Categories("Other types").Properties("color").Format = NewValue
                
            Case "Font"
                Set .Font = NewValue
                
            Case "ForeColor"
                .ForeColor = NewValue
                '.ResetForeColor
                
            Case "GridColor"
                .GridColor = NewValue
                
            Case "NameWidth"
                .NameWidth = NewValue
                NewValue = .NameWidth
                
            Case "SelBackColor"
                .SelBackColor = NewValue
                
            Case "SelForeColor"
                .SelForeColor = NewValue
                
            Case "ShowCategories"
                .ShowCategories = NewValue
                
            Case "RequiresEnter"
                .RequiresEnter = NewValue
                
            Case "Visible"
                .Visible = NewValue
                
            Case "ShowToolbar"
                .ShowToolbar = NewValue
                
            Case "ShowDescription"
                .ShowDescription = NewValue
                                
            Case "Tooltips"
                .ShowToolTips = NewValue
                
            Case "Height"
                .Height = NewValue
            
            Case "Width"
                .Width = NewValue
            
            Case "Date Format"
                .Categories("Date/time types").Properties("date").Format = NewValue
            
            Case "Time Format"
                .Categories("Date/time types").Properties("time").Format = NewValue
        
            Case "DescriptionHeight"
                ps2.DescriptionHeight = NewValue
            
'<Added by: Project Administrator at: 31/3/2004-21:16:35 on machine: ZEUS>
            Case "EffectStyle"
                ps2.EffectStyle = NewValue
'</Added by: Project Administrator at: 31/3/2004-21:16:35 on machine: ZEUS>
        End Select
    End With
    '<EhFooter>
    Exit Sub

ps1_PropertyChanged_Err:
    MsgBox Err.Description & vbCrLf & _
           "in DemoPropertySheet.Form1.ps1_PropertyChanged " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub AddPropertiesPS2()

    With ps2
        .Redraw = False
        .ImageList = ImageList1
        .Categories.Clear
        '
        ' numeric types
        '
        With .Categories.Add("Numeric Types", psImgOpenFolder)
            With .Properties
                .Add "Byte", 128, psByte, , , , "A byte property where values range from 0 to 255"
                With .Add("Currency", 12300, psCurrency, , , "psCurrency properties have a default format of ""$ #,##0.00""")
                    .Description = "Currency value"
                    .UpDownIncrement = 0.05
                End With
                .Add "Integer", 1, psInteger, , , , "Integer values range from 0 to 32768"
                With .Add("Long", 200, psLong, , , "This property has maximum and minimum values and an UpDown control.")
                    .Description = "Long value range from 0 to ??????"
                    .SetRange 100, 1000
                    .UpDownIncrement = 10
                End With
                .Add "Decimal", 32312223.21, psDecimal
                .Add "Double", 1639043.324, psDouble
                .Add "Single", 123 / 3, psSingle
            End With
        End With
        '
        ' date/time types
        '
        With .Categories.Add("Date/Time Types", psImgOpenFolder)
            With .Properties
                .Add "Time", Now(), psTime, , psImgClock, "PropertySheet supports Time and Date properties."
                .Add "Date", #10/22/1932#, psDate, , psImgCalendar2, "Date properties shows a calendar control to select a valid date."
            End With
        End With
        '
        ' string types
        '
        With .Categories.Add("String Types", psImgOpenFolder).Properties
            .Add "String", "Hello World!", psString                         ' a simple string property
            With .Add("String * 8", "12345678", psString)                   ' a string property with MaxLength=8
                .SetRange , 8
            End With
            .Add("Password", "pwd", psString).Format = "Password"           ' a password string
            .Add "Memo", "Text", psLongText                                 ' a memo property
            .Add "Combo", "Text", psCombo                                   ' a combo property
            ' you can get a category property using the "Item" property
            With .Item("Combo").ListValues
                .Add "Sample Item"
                .Add "Combo item", "New Combo Type"
            End With
        End With
        '
        ' File & Folders
        With .Categories.Add("File & Folder", psImgOpenFolder)
            With .Properties
                .Add "Folder", "C:\WINDOWS", psFolder, , psImgClosedFolder      ' browse for folder property
                With .Add("File", "C:\AUTOEXEC.BAT", psFile, , psImgPaperClip)  ' open file property
                    .Format = "Batch Files *.BAT|*.BAT"
                End With
            End With
        End With
        '
        ' object types
        '
        With .Categories.Add("Object Types", psImgOpenFolder)
            With .Properties
                .Add("Font", Me.Font, psFont, , psImgFont).Format = "(c npts)"
                .Add "Object", Nothing
                .Add("Picture", Nothing, psPicture, , psImgPicture1).Format = "CustomDisplay"
            End With
        End With
        '
        ' other types
        '
        With .Categories.Add("Other Types", psImgOpenFolder)
            With .Properties
                .Add "Array", Array(1, 2, 3)
                With .Add("CheckedList", "", psDropDownCheckList)
                    .Format = "CustomDisplay"
                    With .ListValues
                        .Add "Check1"
                        .Add "Check2"
                        .Add "Check3"
                        .Add "Check4"
                        .Add "Check5"
                    End With
                End With
                .Add "Boolean", False, psBoolean
                .Add("Color", vbBlue, psColor, , , "This property uses the CustomDisplay format.").Format = "CustomDisplay"
                With .Add("DropDown List", 0, psDropDownList, , , _
                   "@This kind of property" & vbCrLf & _
                   "is not limited only to" & vbCrLf & _
                   "Long values. You can use" & vbCrLf & _
                   "anything that can be" & vbCrLf & _
                   "stored in a Variant.").ListValues
                    .Add 0, "Item A"
                    .Add 1, "Item B"
                    .Add 2, "Item C"
                    .Add 3, "Item D"
                    .Add 4, "Item E"
                    .Add 5, "Item F"
                End With
                .Add("Popup window", -1, psCustom, , , "Show a poup window", "Custom poup window").Format = "CustomDisplay"
            End With
        End With
        '
        ' empty
        '
        With .Categories.Add("Empty Category")
            With .Properties
                ' nothing
            End With
        End With
        .Redraw = True
    End With
End Sub

Private Sub ps2_Browse(ByVal Left As Variant, ByVal Top As Variant, ByVal Width As Variant, ByVal Prop As PropertySheet.TProperty)
    If Prop.Caption = "Popup window" Then
    Dim f As Form5
    
    Set f = New Form5
    Load f
    f.Move Left, Top
    f.Show vbModal
    Prop.Value = f.PictureIndex
    Unload f
    Set f = Nothing
    End If
End Sub

Private Sub ps2_BrowseForFile(ByVal Prop As PropertySheet.TProperty, Title As String, InitDir As String, Filter As String, FilterIndex As Integer, flags As Long)
    ' BrowseForFile allows to configure file dialog before opening
    If Prop.Caption = "File" Then
        Title = "Open"
        Filter = "Batch files (*.BAT)|*.BAT"
        FilterIndex = 1
    End If
End Sub

Private Sub ps2_EditError(ErrMessage As String)
    ' show error message
    MsgBox ErrMessage
End Sub

Private Sub ps2_GetDisplayString(ByVal Prop As TProperty, DisplayString As String, UseDefault As Boolean)
    ' you may customize the display string of the property
    ' just modify thr DisplayString variable accordingly your needs
    Select Case Prop.Caption
        Case "CheckedList"
            DisplayString = "(Flags)"
            
        Case "Color"
            Select Case Prop.Value
                ' Show the color name instead of RGB components
                Case vbRed
                    DisplayString = "Red"
            
                Case vbGreen
                    DisplayString = "Green"
               
                Case vbBlue
                    DisplayString = "Blue"
               
                Case vbYellow
                    DisplayString = "Yellow"
            
                Case Else
                    ' Use the default string
                    ' for any other color
                    UseDefault = True
            End Select
         
        Case "Picture"
            If TypeName(Prop.Value) = "Picture" Then
                If Prop.Value.Type = 1 Then
                    DisplayString = "BMP"
                ElseIf Prop.Value.Type = 2 Then
                    DisplayString = "WMF"
                ElseIf Prop.Value.Type = 3 Then
                    DisplayString = "ICO"
                End If
            Else
                DisplayString = "None"
            End If
            DisplayString = "(" & DisplayString & ")"
            
        Case "Object"
            DisplayString = "(Object)"
         
        Case "Array"
            DisplayString = "Array of " & (UBound(Prop.Value) - LBound(Prop.Value) + 1) & " elements"
        
        Case "Popup window"
            If Prop.Value = -1 Then
                DisplayString = "(None)"
            Else
                DisplayString = "(Image " & Prop.Value & ")"
            End If
        
    End Select
End Sub

Private Sub mnuAddProperty_Click()
    ' adds a new property to the propertysheet control
    Form2.Show vbModal
End Sub

Private Sub mnuFileClear_Click()
    ps2.Clear
End Sub

Private Sub mnuFileExit_Click()
    ' bye bye
    Unload Me
End Sub

Private Sub mnuFileLoad_Click()
    ps2.LoadFromFile App.Path & "\PropertySheet.ini", "Grid"
End Sub

Private Sub mnuFileRestore_Click()
    AddPropertiesPS2
End Sub

Private Sub mnuFileSave_Click()
    ps2.SaveToFile App.Path & "\PropertySheet.ini", "Grid"
End Sub

Private Sub TabStrip1_Click()
    On Error Resume Next
    If TabStrip1.SelectedItem.Index = 1 Then
        ps1.ShowCategories = False
    Else
        ps1.ShowCategories = True
    End If
    ps1.SetFocus
End Sub

Private Sub OpenFile(FileName As String, Title As String)
    Dim f As New Form3
    
    f.Execute FileName, Title
End Sub
'-- end code

VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2865
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10020
   LinkTopic       =   "Form1"
   ScaleHeight     =   2865
   ScaleWidth      =   10020
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   8580
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   300
      Width           =   1335
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2175
      Left            =   2100
      TabIndex        =   2
      Top             =   120
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   3836
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      TextBackground  =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList2"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8820
      Top             =   780
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.DirListBox Dir1 
      Height          =   1665
      Left            =   180
      TabIndex        =   1
      Top             =   600
      Width           =   1755
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   9300
      Top             =   1560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0454
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Right click on a file and then drag 'n' drop it onto a file-accepting window, e.g., Explorer, or the VB IDE's project window."
      Height          =   315
      Left            =   240
      TabIndex        =   4
      Top             =   2460
      Width           =   9615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
' File Drag and Drop from VB
' By Jolyon Bloomfield October '00
' Jolyon_B@Hotmail.Com
' ICQ: 11084041
'
' If you use this, please give me credit ;)
'

Private Sub Combo1_Click()
ListView1.View = Combo1.ItemData(Combo1.ListIndex)
End Sub

Private Sub Dir1_Change()
ListView1.ListItems.Clear
ListView1.ColumnHeaders.Clear
ListView1.ColumnHeaders.Add , , "Filename", 80 * Screen.TwipsPerPixelX
Dim Filename As String
With ListView1.ListItems
  Filename = Dir(fixed(Dir1.path) & "*.*", vbArchive Or vbHidden Or vbReadOnly Or vbSystem)
  Do While Filename <> ""
    .Add , fixed(Dir1.path) & Filename, Filename, 1, 1
    Filename = Dir()
  Loop
End With
End Sub

Private Sub Drive1_Change()
On Error Resume Next
Dir1.path = Drive1.Drive
If Err Then Drive1.Drive = Dir1.path
End Sub

Private Sub Form_Load()
Combo1.AddItem "0 - Icons"
Combo1.ItemData(Combo1.NewIndex) = 0
Combo1.AddItem "1 - Small icons"
Combo1.ItemData(Combo1.NewIndex) = 1
Combo1.AddItem "2 - List"
Combo1.ItemData(Combo1.NewIndex) = 2
Combo1.AddItem "3 - Report"
Combo1.ItemData(Combo1.NewIndex) = 3
Combo1.ListIndex = 0

Dir1.path = Drive1.Drive
End Sub

Private Function fixed(ByVal path As String) As String
fixed = path & IIf(Right(path, 1) = "\", "", "\")
End Function

Private Sub ListView1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbRightButton Then ListView1.OLEDrag
End Sub

Private Sub ListView1_OLEStartDrag(Data As MSComctlLib.DataObject, AllowedEffects As Long)
AllowedEffects = vbDropEffectCopy
Data.Clear
Data.Files.Add ListView1.SelectedItem.Key
Data.SetData , vbCFFiles
End Sub

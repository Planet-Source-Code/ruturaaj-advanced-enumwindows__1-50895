VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Advanced EnumWindows"
   ClientHeight    =   4875
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7770
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain'.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4875
   ScaleWidth      =   7770
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      ForeColor       =   &H00C00000&
      Height          =   2655
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   7575
      Begin MSComctlLib.ListView lstWinList 
         Height          =   2055
         Left            =   120
         TabIndex        =   10
         Top             =   480
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   3625
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "imgList"
         SmallIcons      =   "imgList"
         ColHdrIcons     =   "imgList"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Window Handle"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Window Text"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "IsVisible"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "IsEnabled"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   " Result"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1245
      End
   End
   Begin VB.Frame Frame2 
      ForeColor       =   &H00C00000&
      Height          =   2175
      Left            =   4800
      TabIndex        =   2
      Top             =   2640
      Width           =   2895
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1575
         Left            =   120
         ScaleHeight     =   1545
         ScaleWidth      =   2625
         TabIndex        =   6
         Top             =   480
         Width           =   2655
         Begin VB.CommandButton cmdRefresh 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Reset"
            Height          =   495
            Left            =   240
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   840
            Width           =   2175
         End
         Begin VB.CommandButton cmdExit 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "Re&fresh"
            Height          =   495
            Left            =   240
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   240
            Width           =   2175
         End
      End
      Begin VB.Label Label3 
         BackColor       =   &H00808080&
         Caption         =   "Actions"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      ForeColor       =   &H00C00000&
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   2640
      Width           =   4575
      Begin VB.ListBox lstCondition 
         Appearance      =   0  'Flat
         Height          =   1590
         ItemData        =   "frmMain'.frx":038B
         Left            =   120
         List            =   "frmMain'.frx":03A7
         TabIndex        =   1
         Top             =   480
         Width           =   4335
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         Caption         =   " Filter Options "
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1245
      End
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   -120
      Top             =   3240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain'.frx":04C1
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain'.frx":0A5C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain'.frx":0FF7
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()
    GetWinInfo
End Sub

Private Sub cmdRefresh_Click()
    lstCondition.Selected(0) = True
    EnumCondition = No_Filter
    GetWinInfo
End Sub

Private Sub Form_Load()
    lstWinList.ColumnHeaders(1).Width = lstWinList.Width * 0.2
    lstWinList.ColumnHeaders(2).Width = lstWinList.Width * 0.4
    lstWinList.ColumnHeaders(3).Width = lstWinList.Width * 0.2
    lstWinList.ColumnHeaders(4).Width = lstWinList.Width * 0.2
    
    lstCondition.Selected(0) = True
    EnumCondition = No_Filter
    GetWinInfo
End Sub

Private Sub lstCondition_Click()
    Select Case lstCondition.ListIndex
        
        'No Filter
        Case 0:
            EnumCondition = No_Filter
            GetWinInfo
            
        'Only Visible
        Case 1:
            EnumCondition = Only_Visible
            GetWinInfo
        
        'Only Enabled
        Case 2:
            EnumCondition = Only_Enabled
            GetWinInfo
        
        'Visible with WindowText
        Case 3:
            EnumCondition = Only_Visible_WinTextNotEmpty
            GetWinInfo
        
        'Enabled-Visible
        Case 4:
            EnumCondition = Only_Enabled_Visible
            GetWinInfo
        
        'Enabled-NonVisible
        Case 5:
            EnumCondition = Only_Enabled_NonVisible
            GetWinInfo
        
        'Disabled-Visible
        Case 6:
            EnumCondition = Only_Disabled_Visible
            GetWinInfo
        
        'Disabled-NonVisible
        Case 7:
            EnumCondition = Only_Disabled_NonVisible
            GetWinInfo
    End Select
End Sub


VERSION 5.00
Object = "{13709F57-6507-4C51-8323-B204803FD2D1}#4.0#0"; "LayerMan.ocx"
Begin VB.Form frmTbar 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   5  'Sizable ToolWindow
   ClientHeight    =   5955
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   3840
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5955
   ScaleWidth      =   3840
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frAbout 
      Caption         =   "About"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5655
      Left            =   -60
      TabIndex        =   0
      Top             =   315
      Width           =   3900
      Begin vbpLayerTest.LimitSizing LimitSizing1 
         Left            =   2955
         Top             =   336
         _ExtentX        =   609
         _ExtentY        =   582
         MinWidth        =   300
         MinHeight       =   265
         MaxWidth        =   510
         MaxHeight       =   450
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "here."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   2250
         MouseIcon       =   "frmTbar.frx":0000
         MousePointer    =   99  'Custom
         TabIndex        =   16
         Top             =   4665
         Width           =   390
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "• Steve McMahon - ssubtmr.dll is required for the size limit control. Visit his sweet ass VBAccelerator site by clicking"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   630
         Left            =   120
         TabIndex        =   15
         Top             =   4290
         Width           =   3375
      End
      Begin VB.Label Label6 
         Caption         =   " Ray For More Information."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1170
         TabIndex        =   7
         Top             =   5025
         Width           =   2445
      End
      Begin VB.Label lblemail 
         AutoSize        =   -1  'True
         Caption         =   "Email"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   765
         MouseIcon       =   "frmTbar.frx":08CA
         MousePointer    =   99  'Custom
         TabIndex        =   6
         Top             =   5025
         Width           =   360
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "• John Percival - His SizeLimit control found on VB Square was used with minor modifications. Thanks for the great control John!"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   630
         Left            =   165
         TabIndex        =   5
         Top             =   3615
         Width           =   3375
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "• Ray Hildenbrand - Created this sample application and all functionallity found in the layer control, layer manager."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Left            =   165
         TabIndex        =   3
         Top             =   2055
         Width           =   3375
      End
      Begin VB.Label Label2 
         Caption         =   "Code by these authors was used (in part) to create this example."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   120
         TabIndex        =   2
         Top             =   1620
         Width           =   3510
      End
      Begin VB.Label lblAbout 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmTbar.frx":1194
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1200
         Left            =   150
         TabIndex        =   1
         Top             =   300
         Width           =   3615
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "• Jim Williams - His JW_Cooltabs control found on PSC was used with minor modifications. Thanks for the great control Jim!"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Left            =   165
         TabIndex        =   4
         Top             =   2895
         Width           =   3375
      End
   End
   Begin vbpLayerItem.LayerContainer LayerContainer1 
      Height          =   5370
      Left            =   0
      TabIndex        =   21
      Top             =   630
      Width           =   3825
      _ExtentX        =   6747
      _ExtentY        =   9472
      highlightColor  =   13537847
   End
   Begin VB.PictureBox picControls 
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   15
      ScaleHeight     =   315
      ScaleWidth      =   3840
      TabIndex        =   18
      Top             =   345
      Width           =   3840
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         ItemData        =   "frmTbar.frx":12B8
         Left            =   1980
         List            =   "frmTbar.frx":12BF
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   30
         Width           =   1005
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         ItemData        =   "frmTbar.frx":12C9
         Left            =   135
         List            =   "frmTbar.frx":12D0
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   15
         Width           =   1005
      End
   End
   Begin vbpLayerTest.CoolTabs CoolTabs1 
      Height          =   690
      Left            =   -15
      TabIndex        =   17
      Top             =   30
      Width           =   3870
      _ExtentX        =   6826
      _ExtentY        =   1217
      CaptionTab1     =   "Layers"
      CaptionTab2     =   "About"
      CaptionTab3     =   "Usage"
      CaptionTab4     =   ""
      VisibleTabs     =   2
   End
   Begin VB.Frame frUsage 
      Caption         =   "Usage"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5655
      Left            =   0
      TabIndex        =   8
      Top             =   315
      Width           =   3855
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "• Double Click on a layer to change the caption"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   630
         Left            =   135
         TabIndex        =   14
         Top             =   4785
         Width           =   3660
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmTbar.frx":12DC
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Left            =   135
         TabIndex        =   13
         Top             =   255
         Width           =   3660
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "• The 'layers' can be toggled between visible and not visible by clicking on the eye icon "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   120
         TabIndex        =   12
         Top             =   1245
         Width           =   3645
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmTbar.frx":136D
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1200
         Left            =   120
         TabIndex        =   11
         Top             =   1755
         Width           =   3645
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmTbar.frx":149E
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1230
         Left            =   120
         TabIndex        =   10
         Top             =   3030
         Width           =   3645
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "• The layer manager arranges the controls' Zorder according to their order in the list top down."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   630
         Left            =   135
         TabIndex        =   9
         Top             =   4305
         Width           =   3660
      End
   End
End
Attribute VB_Name = "frmTbar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CoolTabs1_RoundBtnClick(ActiveTabIs As Integer)
    MsgBox "Menu Requested for Tab number " & ActiveTabIs
End Sub

Private Sub CoolTabs1_TabClick(Index As Integer)
    Select Case Index
        Case 0
            LayerContainer1.Visible = True
            frAbout.Visible = False
            frUsage.Visible = False
            LayerContainer1.ZOrder
        Case 1
            LayerContainer1.Visible = False
            frAbout.Visible = True
            frUsage.Visible = False
            frAbout.ZOrder
            Me.Height = 6500
        Case 2
            LayerContainer1.Visible = False
            frAbout.Visible = False
            frUsage.Visible = True
            frUsage.ZOrder
            Me.Height = 6500
    End Select
    
End Sub

Private Sub Form_Load()
   LimitSizing1.Formhwnd = Me.hwnd
   LimitSizing1.Enabled = True
   Combo1.ListIndex = 0
   Combo2.ListIndex = 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    LimitSizing1.Enabled = False
End Sub

Private Sub Form_Resize()
   'If Not Me.Width < 4050 Then
        CoolTabs1.Move 0, -1, Me.Width - 90, Me.Height - 60
        picControls.Width = Me.Width
        LayerContainer1.Move 0, picControls.Top + picControls.Height, CoolTabs1.Width - 60, CoolTabs1.Height - (picControls.Top + picControls.Height + 300)
        frAbout.Move -15, picControls.Top + picControls.Height + 30, CoolTabs1.Width, CoolTabs1.Height - (picControls.Top + picControls.Height + 300) - 30
        frUsage.Move -15, picControls.Top + picControls.Height + 30, CoolTabs1.Width, CoolTabs1.Height - (picControls.Top + picControls.Height + 300) - 30
   'Else
   ' Me.Width = 4050
   'End If
End Sub

Private Sub Label14_Click()
    AccessWeb
End Sub

Private Sub LayerContainer1_LayerDuplicated(ByRef objectToDuplicate As Object)
    MsgBox "Duplicate code not complete yet", vbInformation, App.Title

End Sub

Private Sub lblemail_Click()
    sendemail
End Sub

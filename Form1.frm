VERSION 5.00
Object = "*\AvbpLayerControl.vbp"
Begin VB.Form frmTbar 
   Caption         =   "Form1"
   ClientHeight    =   7185
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8580
   LinkTopic       =   "Form1"
   ScaleHeight     =   7185
   ScaleWidth      =   8580
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   285
      Left            =   1140
      TabIndex        =   1
      Top             =   3135
      Width           =   3990
   End
   Begin vbpLayerItem.LayerContainer LayerContainer1 
      Height          =   3435
      Left            =   1125
      TabIndex        =   0
      Top             =   3450
      Width           =   3960
      _ExtentX        =   6985
      _ExtentY        =   6059
   End
End
Attribute VB_Name = "frmTbar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Dim tmpObj As Control
    Set tmpObj = Picture1(LayerContainer1.TotalLayers)
    LayerContainer1.AddLayerItem tmpObj
End Sub


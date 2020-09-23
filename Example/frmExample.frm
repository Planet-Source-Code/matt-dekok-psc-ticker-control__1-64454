VERSION 5.00
Object = "{46D2BBF7-F93B-4480-9522-165661C2E939}#59.0#0"; "Ticker.ocx"
Begin VB.Form frmExample 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PSC Ticker"
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   1935
   Icon            =   "frmExample.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   1935
   StartUpPosition =   3  'Windows Default
   Begin PSCTicker.ctlTicker ctlTicker1 
      Height          =   2445
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   4154
   End
End
Attribute VB_Name = "frmExample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Me.Height = ctlTicker1.Height + 375
End Sub

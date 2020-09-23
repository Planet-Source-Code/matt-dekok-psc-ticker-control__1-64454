VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.dll"
Begin VB.UserControl ctlTicker 
   ClientHeight    =   3360
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3765
   ScaleHeight     =   3360
   ScaleMode       =   0  'User
   ScaleWidth      =   3765
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2445
      Left            =   0
      ScaleHeight     =   2314.793
      ScaleMode       =   0  'User
      ScaleWidth      =   1935
      TabIndex        =   0
      Top             =   -90
      Width           =   1935
      Begin SHDocVwCtl.WebBrowser WebBrowser1 
         Height          =   2775
         Left            =   -120
         TabIndex        =   1
         Top             =   0
         Width           =   2655
         ExtentX         =   4683
         ExtentY         =   4895
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   0
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   "http:///"
      End
   End
End
Attribute VB_Name = "ctlTicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Default Property Values:
Const m_def_PSCWorld = 1
'Property Variables:
Dim m_PSCWorld As Integer
'Event Declarations:
Event Resize() 'MappingInfo=UserControl,UserControl,-1,Resize
Attribute Resize.VB_Description = "Occurs when a form is first displayed or the size of an object changes."

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub UserControl_Initialize()
    On Error GoTo ph1
    Open App.Path & "\ticker.htm" For Input As #1
    Close #1
    GoTo ph2
ph1:
    Open App.Path & "\ticker.htm" For Output As #1
        Print #1, Replace(LoadResString(101), "http://www.Planet-Source-Code.com/vb/linktous/ScrollingCode.asp?lngWId=1", "http://www.Planet-Source-Code.com/vb/linktous/ScrollingCode.asp?lngWId=1")
    Close #1
ph2:
    WebBrowser1.Navigate App.Path & "\ticker.htm", 2
End Sub

Private Sub UserControl_Resize()
    'Do not let control be resized
    UserControl.Width = Picture1.Width
    UserControl.Height = Picture1.Height + Picture1.Top
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,1,0,1
Public Property Get PSCWorld() As Integer
Attribute PSCWorld.VB_Description = "Sets/Returns a world at Planet-Source-Code.com."
    PSCWorld = m_PSCWorld
End Property

Public Property Let PSCWorld(ByVal New_PSCWorld As Integer)
    If Ambient.UserMode Then Err.Raise 382
    m_PSCWorld = New_PSCWorld
    Open App.Path & "\ticker.htm" For Output As #1
        Print #1, Replace(LoadResString(101), "http://www.Planet-Source-Code.com/vb/linktous/ScrollingCode.asp?lngWId=1", "http://www.Planet-Source-Code.com/vb/linktous/ScrollingCode.asp?lngWId=" & PSCWorld)
    Close #1
    WebBrowser1.Navigate App.Path & "\ticker.htm", 2
    PropertyChanged "PSCWorld"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_PSCWorld = m_def_PSCWorld
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_PSCWorld = PropBag.ReadProperty("PSCWorld", m_def_PSCWorld)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("PSCWorld", m_PSCWorld, m_def_PSCWorld)
End Sub

Private Sub WebBrowser1_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
Dim n As Integer
n = StrComp(URL, "http://www.Planet-Source-Code.com/vb/linktous/ScrollingCode.asp?lngWId=" & m_def_PSCWorld, 1)
n = n + StrComp(URL, "http://www.Planet-Source-Code.com/vb/linktous/ScrollingCode.asp?lngWId=-" & m_def_PSCWorld, 1)
n = n + StrComp(URL, "http://www.Planet-Source-Code.com/vb/linktous/ScrollingCode.asp?lngWId=" & m_PSCWorld, 1)
n = n + StrComp(URL, "http://www.Planet-Source-Code.com/vb/linktous/ScrollingCode.asp?lngWId=-" & m_PSCWorld, 1)
If Left$(URL, 57) = "http://www.planet-source-code.com/vb/scripts/ShowCode.asp" Then MsgBox "fg"
If n = 5 Then
    ShellExecute 0&, vbNullString, URL, vbNullString, vbNullString, vbNormalFocus
    Cancel = True
End If
End Sub

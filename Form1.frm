VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form Form1 
   Caption         =   "TAS Geocoding Simple Weather"
   ClientHeight    =   5400
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12810
   LinkTopic       =   "Form1"
   ScaleHeight     =   5400
   ScaleWidth      =   12810
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   800
      Left            =   3915
      Top             =   3000
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3120
      Top             =   3045
   End
   Begin VB.Frame Frame1 
      Caption         =   "Your Zip Code"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   3075
      TabIndex        =   1
      Top             =   645
      Width           =   1860
      Begin VB.CommandButton Command1 
         Caption         =   "Go"
         Height          =   300
         Left            =   1305
         TabIndex        =   3
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   60
         TabIndex        =   2
         Text            =   "85219"
         Top             =   255
         Width           =   1095
      End
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   1230
      Left            =   1005
      TabIndex        =   0
      Top             =   2580
      Width           =   1425
      ExtentX         =   2514
      ExtentY         =   2170
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
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
      Location        =   ""
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   5805
      Top             =   3630
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Here is your forecast from Intellicast. May the tightwad weisles at Intellicast roll over in their graves. LOL
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const BS_FLAT = &H8000&
Private Const GWL_STYLE = (-16)
Private Const WS_CHILD = &H40000000
Private ErrorCode As Boolean
Private WeatherURL As String
Private LocCode As String
Private Sub GetGeo()
    Dim i As Long
    Dim Filehedder As String
    Dim InetData As String
    WriteToBrowser "Downloading geo data for " & Text1.Text & " area code."
    InetData = Inet1.OpenURL("http://www.intellicast.com/IcastRSS/FcstRSS.aspx?loc=" & Text1.Text, icString)
    For i = 1 To Len(InetData)
        If Mid(InetData, i, Len("error page")) = "error page" Then
            WriteToBrowser "Comunication Error !"
            ErrorCode = True
            Exit Sub
        End If
    Next i
    
    InetData = SimpleHTMLGet(InetData, "<link>", "</link>")
    LocCode = GetLocCode(Trim(InetData))
    Debug.Print LocCode
    InetData = Replace(InetData, "&amp;", "&")
    GetWeather InetData
End Sub
Public Function GetLocCode(dURL As String) As String
    Dim m As Long
    Dim GetChr0 As String
    Dim GetChr1 As String
    GetLocCode = ""
    For m = 1 To Len(dURL)
        GetChr0 = Right(dURL, m)
        GetChr1 = Left(GetChr0, 1)
        If GetChr1 = "=" Then
            GetLocCode = Right(GetChr0, m - 1): Exit Function
        End If
    Next m
End Function
Private Function GetWeather(Optional dURL As String)
    Dim i As Long
    Dim Filehedder As String
    Dim InetData As String
    If dURL = "" Then dURL = WeatherURL
    WriteToBrowser "Downloading weather data for " & Text1.Text & " area code."
    InetData = Inet1.OpenURL(dURL, icString)
    Filehedder = Inet1.GetHeader()
    For i = 1 To Len(Filehedder)
        If Mid(Filehedder, i, Len("404")) = "404" Then
            WriteToBrowser "Comunication Error !"
            ErrorCode = True
            Exit Function
        End If
    Next i
    WeatherURL = dURL
    InetData = Trim(InetData)
    InetData = SimpleHTMLGet(InetData, "<!-- Observation -->", "<!-- Area Guid -->")
    InetData = Replace(InetData, Chr(34), Chr(39))
    InetData = SimpleHTMLRemove(InetData, "<div class='Container ChromeLight'>", "</div>")
    InetData = Replace(InetData, "Send this Forecast Daily to your E-mail Free &raquo;", "")
    InetData = Replace(InetData, "<em>", "")
    InetData = Replace(InetData, "</em>", "")
    Call BuildDisplay(InetData)
    ErrorCode = False
    Timer1.Enabled = True
End Function
Public Function BuildDisplay(DispData As String)
    WebBrowser1.Navigate2 "about:blank"
    Call AreWeReady
    Call WebBrowser1.Document.Script.Document.Clear
    Call WebBrowser1.Document.Script.Document.write("<html><head>")
    'Customize the page styling.
    Call WebBrowser1.Document.Script.Document.write("<style type='text/css'>")
    Call WebBrowser1.Document.Script.Document.write("a{text-decoration:none}")
    Call WebBrowser1.Document.Script.Document.write(".rss-box{margin:1em;background-color:#EEEEEE; border: 1px solid #000000;width:95%;height:95%; }")
    Call WebBrowser1.Document.Script.Document.write(".Field{font-family:times;font-size:12px;font-weight:bold;text-align:center;}")
    Call WebBrowser1.Document.Script.Document.write(".Value{font-family:times;font-size:12px;font-weight:normal;text-align:left;}")
    Call WebBrowser1.Document.Script.Document.write(".Fine{font-family:times;font-size:12px;font-weight:normal;text-align:center;}")
    Call WebBrowser1.Document.Script.Document.write(".Forecast{font-family:times;width:40px;height:40px;border:solid 1px #FFF;text-align:center;}")
    Call WebBrowser1.Document.Script.Document.write(".ForecastBox{font-family:times;width:100px;height:100px;text-align:center;}")
    Call WebBrowser1.Document.Script.Document.write("</style>")
    '****************************
    Call WebBrowser1.Document.Script.Document.write("</head><body bgcolor='#FFFFFF'>")
    Call WebBrowser1.Document.Script.Document.write("<center><div class='rss-box'><b>")
    Call WebBrowser1.Document.Script.Document.write(DispData)
    Call WebBrowser1.Document.Script.Document.write("<br><hr><center><a href='http://www.tas-independent-programming.com' target='_Blank'><b>Written By T.A.S. Independent Programming</b></a></center>")
    Call WebBrowser1.Document.Script.Document.write("</div></center>")
    Call WebBrowser1.Document.Script.Document.write("</body></html>")
    DoEvents
    Call WebBrowser1.Refresh
End Function
Private Function WriteToBrowser(dMessage As String)
    WebBrowser1.Navigate2 "about:blank"
    Call AreWeReady
    Call WebBrowser1.Document.Script.Document.Clear
    Call WebBrowser1.Document.Script.Document.write("<html><head>")
    Call WebBrowser1.Document.Script.Document.write("</head>")
    Call WebBrowser1.Document.Script.Document.write("<body bgcolor='#808080' link='#00FF00' vlink='#00FF00' alink='#00FF00'>")
    Call WebBrowser1.Document.Script.Document.write("<table border='0' width='100%' height='100%' cellspacing='0' cellpadding='0'>")
    Call WebBrowser1.Document.Script.Document.write("<tr><td width='100%' height='88' valign='middle' align='center'>")
    Call WebBrowser1.Document.Script.Document.write("<table border='0' width='100%' cellspacing='0' cellpadding='0'><tr><td>")
    Call WebBrowser1.Document.Script.Document.write("<hr><p align='center'><b><font face='Arial' size='6'>" & dMessage & "</font></b></p><hr>")
    Call WebBrowser1.Document.Script.Document.write("</td></tr></table></td></tr></table></center></div>")
    Call WebBrowser1.Document.Script.Document.write("</body></html>")
    DoEvents
    Call WebBrowser1.Refresh
End Function
Private Sub Command1_Click()
    SaveSetting "Simple Weather", "Settings", "ZIP", Text1.Text
    GetGeo
End Sub
Private Sub Form_Load()
    If App.PrevInstance Then End
    btnFlat Command1
    Text1.Text = GetSetting("Simple Weather", "Settings", "ZIP", "Enter Zip")
    WriteToBrowser "Please enter your zip code above and click 'Go'."
    If Text1.Text <> "Enter Zip" Then
        Timer2.Enabled = True
    End If
End Sub
Private Sub Form_Resize()
    Frame1.Top = Me.ScaleTop
    Frame1.Left = Me.ScaleLeft + 40
    WebBrowser1.Top = Me.ScaleTop + Frame1.Height + 40
    WebBrowser1.Left = Me.ScaleLeft + 40
    WebBrowser1.Height = (Me.ScaleHeight - Frame1.Height) - 100
    WebBrowser1.Width = (Me.ScaleWidth - 100)
End Sub
'Without end tags
Private Function SimpleHTMLGet(SearchString As String, FirstTag As String, SecondTag As String) As String
    Dim MyPos1 As String ' Changed to string because integer was causing overflow error.
    Dim MyPos2 As String ' Changed to string because integer was causing overflow error.
    MyPos1 = InStr(1, UCase(SearchString), UCase(FirstTag), 1)
    MyPos2 = InStr(MyPos1, UCase(SearchString), UCase(SecondTag), 1)
    SimpleHTMLGet = Mid(SearchString, MyPos1 + Len(FirstTag), MyPos2 - MyPos1 - Len(FirstTag))
End Function
Private Function SimpleHTMLRemove(SearchString As String, FirstTag As String, SecondTag As String) As String
    Dim MyPos1 As String ' Changed to string because integer was causing overflow error.
    Dim MyPos2 As String ' Changed to string because integer was causing overflow error.
    Dim Combined As String
    MyPos1 = InStr(1, UCase(SearchString), UCase(FirstTag), 1)
    MyPos2 = InStr(MyPos1, UCase(SearchString), UCase(SecondTag), 1)
    SimpleHTMLRemove = Mid(SearchString, "1", (MyPos1 - Len(FirstTag)) + 1) + Mid(SearchString, MyPos2, Len(SearchString) - (MyPos2 + 1))
End Function
Private Sub AreWeReady()
    Do While WebBrowser1.ReadyState <> READYSTATE_COMPLETE
        Sleep 100
        DoEvents
    Loop
End Sub
Public Function btnFlat(Button As CommandButton)
    SetWindowLong Button.hwnd, GWL_STYLE, WS_CHILD Or BS_FLAT
    Button.Visible = True 'Make the button visible (its automaticly hidden when the SetWindowLong call is executed because we reset the button's Attributes)
End Function
Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then GetGeo
End Sub
Private Sub Timer1_Timer()
    If ErrorCode = False And Minute(Now) = "0" And Second(Now) = "0" Then GetWeather
End Sub
Private Sub Timer2_Timer()
    Timer2.Enabled = False
    GetGeo
End Sub

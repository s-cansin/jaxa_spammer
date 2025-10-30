VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form mailoptimize 
   Caption         =   "Email Listesi Kombine Optimizasyon Aracý"
   ClientHeight    =   4140
   ClientLeft      =   6285
   ClientTop       =   5850
   ClientWidth     =   8325
   LinkTopic       =   "Form1"
   ScaleHeight     =   4140
   ScaleWidth      =   8325
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text2 
      Height          =   1815
      Left            =   5520
      MultiLine       =   -1  'True
      TabIndex        =   16
      Top             =   4800
      Visible         =   0   'False
      Width           =   2055
   End
   Begin MSComDlg.CommonDialog CommonDialog2 
      Left            =   7680
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Optimize edilmiþ listeyi farklý kaydet"
      Height          =   1335
      Left            =   6960
      TabIndex        =   15
      Top             =   1440
      Width           =   1215
   End
   Begin VB.ListBox List2 
      Height          =   2985
      ItemData        =   "Form3.frx":0000
      Left            =   3600
      List            =   "Form3.frx":0002
      TabIndex        =   14
      Top             =   600
      Width           =   2895
   End
   Begin VB.ListBox other 
      Height          =   1815
      ItemData        =   "Form3.frx":0004
      Left            =   5280
      List            =   "Form3.frx":0006
      TabIndex        =   13
      Top             =   4800
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.ListBox inbox 
      Height          =   1815
      ItemData        =   "Form3.frx":0008
      Left            =   5040
      List            =   "Form3.frx":000A
      TabIndex        =   10
      Top             =   4800
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.ListBox yorktrade 
      Height          =   1815
      ItemData        =   "Form3.frx":000C
      Left            =   4800
      List            =   "Form3.frx":000E
      TabIndex        =   9
      Top             =   4800
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.ListBox mynet 
      Height          =   1815
      ItemData        =   "Form3.frx":0010
      Left            =   4560
      List            =   "Form3.frx":0012
      TabIndex        =   8
      Top             =   4800
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.ListBox aol 
      Height          =   1815
      ItemData        =   "Form3.frx":0014
      Left            =   4320
      List            =   "Form3.frx":0016
      TabIndex        =   7
      Top             =   4800
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.ListBox microsoft 
      Height          =   1815
      ItemData        =   "Form3.frx":0018
      Left            =   4080
      List            =   "Form3.frx":001A
      TabIndex        =   6
      Top             =   4800
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.ListBox google 
      Height          =   1815
      ItemData        =   "Form3.frx":001C
      Left            =   3840
      List            =   "Form3.frx":001E
      TabIndex        =   5
      Top             =   4800
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.ListBox yahoo 
      Height          =   1815
      ItemData        =   "Form3.frx":0020
      Left            =   3600
      List            =   "Form3.frx":0022
      TabIndex        =   4
      Top             =   4800
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Email Listesini Optimize Et"
      Height          =   375
      Left            =   3600
      TabIndex        =   3
      Top             =   120
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      ForeColor       =   &H8000000C&
      Height          =   375
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Göster"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.ListBox List1 
      Height          =   2985
      ItemData        =   "Form3.frx":0024
      Left            =   240
      List            =   "Form3.frx":0026
      TabIndex        =   0
      Top             =   600
      Width           =   2895
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7080
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Image Image1 
      Height          =   1095
      Left            =   6600
      Picture         =   "Form3.frx":0028
      Top             =   1560
      Width           =   300
   End
   Begin VB.Label total2 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   7200
      TabIndex        =   18
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Optimize Edilmiþ / Toplam E-Mail: :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   17
      Top             =   3720
      Width           =   3495
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000C&
      X1              =   3360
      X2              =   3360
      Y1              =   120
      Y2              =   3600
   End
   Begin VB.Label total 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   2520
      TabIndex        =   12
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Orjinal / Toplam E-Mail:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   3720
      Width           =   2415
   End
End
Attribute VB_Name = "mailoptimize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()

CommonDialog1.DialogTitle = "Open Email List DB File (*.txt)"
CommonDialog1.Filter = "Text DB Files(*.txt)|*.txt;"
CommonDialog1.ShowOpen

If CommonDialog1.FileName <> "" Then
total = "0"
total2 = "0"


List1.Clear
List2.Clear

yahoo.Clear
google.Clear
microsoft.Clear
aol.Clear
mynet.Clear
yorktrade.Clear
inbox.Clear
other.Clear

Text2.Text = ""

Text1.Text = CommonDialog1.FileName


Open CommonDialog1.FileName For Input As #1

Do While Not EOF(1)
Input #1, Read

total.Caption = total.Caption + 1

List1.AddItem Read



yahoo_str = InStr(1, Read, "@yahoo") + InStr(1, Read, "@yahoogroups")
google_str = InStr(1, Read, "@gmail")
microsoft_str = InStr(1, Read, "@hotmail") + InStr(1, Read, "@msn") + InStr(1, Read, "@live")
aol_str = InStr(1, Read, "@aim") + InStr(1, Read, "@aol")
mynet_str = InStr(1, Read, "@mynet")
yorktrade_str = InStr(1, Read, "@yorktrade")
inbox_str = InStr(1, Read, "@inbox")
other_str = yahoo_str + google_str + microsoft_str + aol_str + mynet_str + yorktrade_str + inbox_str


If yahoo_str > 0 Then yahoo.AddItem Read
If google_str > 0 Then google.AddItem Read
If microsoft_str > 0 Then microsoft.AddItem Read
If aol_str > 0 Then aol.AddItem Read
If mynet_str > 0 Then mynet.AddItem Read
If yorktrade_str > 0 Then yorktrade.AddItem Read
If inbox_str > 0 Then inbox.AddItem Read
If other_str = 0 Then other.AddItem Read



Loop

Close #1
End If

End Sub

Private Sub Command2_Click()
On Error Resume Next

For i = 1 To List1.ListCount
X = i - 1

yahoo_addr = yahoo.List(X)
google_addr = google.List(X)
microsoft_addr = microsoft.List(X)
aol_addr = aol.List(X)
mynet_addr = mynet.List(X)
yorktrade_addr = yorktrade.List(X)
inbox_addr = inbox.List(X)
other_addr = other.List(X)

If yahoo_addr <> "" Then List2.AddItem yahoo.List(X): Text2.Text = Text2.Text & yahoo.List(X) & vbCrLf
If google_addr <> "" Then List2.AddItem google.List(X): Text2.Text = Text2.Text & google.List(X) & vbCrLf
If microsoft_addr <> "" Then List2.AddItem microsoft.List(X): Text2.Text = Text2.Text & microsoft.List(X) & vbCrLf
If aol_addr <> "" Then List2.AddItem aol.List(X): Text2.Text = Text2.Text & aol.List(X) & vbCrLf
If mynet_addr <> "" Then List2.AddItem mynet.List(X): Text2.Text = Text2.Text & mynet.List(X) & vbCrLf
If yorktrade_addr <> "" Then List2.AddItem yorktrade.List(X): Text2.Text = Text2.Text & yorktrade.List(X) & vbCrLf
If inbox_addr <> "" Then List2.AddItem inbox.List(X): Text2.Text = Text2.Text & inbox.List(X) & vbCrLf
If other_addr <> "" Then List2.AddItem other.List(X): Text2.Text = Text2.Text & other.List(X) & vbCrLf

total2.Caption = total2.Caption + 1
Next
End Sub

Private Sub Command3_Click()
CommonDialog2.DialogTitle = "Open Email List DB File (*.txt)"
CommonDialog2.Filter = "Text DB Files(*.txt)|*.txt;"
CommonDialog2.ShowSave
If CommonDialog2.FileName <> "" Then
Open CommonDialog2.FileName For Output As #1
Print #1, Text2.Text
Close #1
If total.Caption = total2.Caption Then
MsgBox ("Liste optimizasyon iþlemi baþarýyla gerçekleþtirilmiþtir!!!")
Else

If total.Caption < total2.Caption Then
missingemails = total.Caption - total2.Caption
Else
missingemails = total2.Caption - total.Caption
End If


MsgBox ("Optimize edilmiþ liste oluþturuldu fakat yinede bir sorun var, " & missingemails & " adet email adresi yeni oluþturulan listede yok!")
End If
End If
End Sub

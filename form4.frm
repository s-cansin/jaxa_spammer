VERSION 5.00
Begin VB.Form ipdegistirici 
   Caption         =   "IP Deðiþtirme Yöneticisi"
   ClientHeight    =   7365
   ClientLeft      =   7305
   ClientTop       =   3990
   ClientWidth     =   5970
   LinkTopic       =   "Form1"
   ScaleHeight     =   7365
   ScaleWidth      =   5970
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   4200
      TabIndex        =   19
      Top             =   2355
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "-"
      Height          =   375
      Left            =   3120
      TabIndex        =   18
      Top             =   2760
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "+"
      Height          =   375
      Left            =   3120
      TabIndex        =   17
      Top             =   2310
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Tamam"
      Height          =   375
      Left            =   2040
      TabIndex        =   16
      Top             =   6720
      Width           =   2295
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   2400
      TabIndex        =   13
      Text            =   "admin"
      Top             =   6000
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   2400
      TabIndex        =   12
      Text            =   "admin"
      Top             =   5520
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2400
      TabIndex        =   8
      Top             =   5010
      Width           =   975
   End
   Begin VB.ListBox List1 
      Height          =   1230
      ItemData        =   "form4.frx":0000
      Left            =   840
      List            =   "form4.frx":0002
      TabIndex        =   5
      Top             =   2160
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1440
      TabIndex        =   2
      Top             =   1290
      Width           =   975
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Gönderim belli periyotlarla durdurulsun ve Modem IP si yenilensin"
      Height          =   255
      Left            =   600
      TabIndex        =   1
      Top             =   4440
      Width           =   4935
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Gönderim belli periyotlarla durdurulsun ve local IP deðiþtirilsin"
      Height          =   255
      Left            =   600
      TabIndex        =   0
      Top             =   840
      Width           =   4575
   End
   Begin VB.Label Label10 
      Caption         =   "IP Ekle:"
      Height          =   255
      Left            =   3600
      TabIndex        =   20
      Top             =   2400
      Width           =   615
   End
   Begin VB.Label Label9 
      Caption         =   "WAN IP (INTERNET IP SÝ) DEÐÝÞÝM AYARLARI"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   600
      TabIndex        =   15
      Top             =   3840
      Width           =   4695
   End
   Begin VB.Label Label8 
      Caption         =   "LOCAL IP (PC IP SÝ) DEÐÝÞÝM AYARLARI"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   600
      TabIndex        =   14
      Top             =   360
      Width           =   3615
   End
   Begin VB.Label Label7 
      Caption         =   "Modem Þifre:"
      Height          =   255
      Left            =   1320
      TabIndex        =   11
      Top             =   6030
      Width           =   975
   End
   Begin VB.Label Label6 
      Caption         =   "Modem Kullanýcý Adý:"
      Height          =   255
      Left            =   720
      TabIndex        =   10
      Top             =   5565
      Width           =   1575
   End
   Begin VB.Label Label5 
      Caption         =   "sn"
      Height          =   255
      Left            =   3480
      TabIndex        =   9
      Top             =   5040
      Width           =   255
   End
   Begin VB.Label Label4 
      Caption         =   "Periyot:"
      Height          =   255
      Left            =   1800
      TabIndex        =   7
      Top             =   5040
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "Local IP Havuzu"
      Height          =   255
      Left            =   1200
      TabIndex        =   6
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "sn"
      Height          =   255
      Left            =   2520
      TabIndex        =   4
      Top             =   1320
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "Periyot:"
      Height          =   255
      Left            =   840
      TabIndex        =   3
      Top             =   1320
      Width           =   615
   End
End
Attribute VB_Name = "ipdegistirici"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()

If Check1.Value = 1 Then

Text1.Enabled = False
List1.Enabled = False
Else
Text1.Enabled = True
List1.Enabled = True
End If

End Sub

Private Sub Check2_Click()

If Check2.Value = 1 Then

Text2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False

Else
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
End If
End Sub

Private Sub Command1_Click()
Unload Me
End Sub


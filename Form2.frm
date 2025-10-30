VERSION 5.00
Begin VB.Form mailhesaplayici 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "E-Mail Gönderim Hesaplama Aracý"
   ClientHeight    =   1335
   ClientLeft      =   8130
   ClientTop       =   5835
   ClientWidth     =   3885
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1335
   ScaleWidth      =   3885
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2880
      TabIndex        =   3
      Text            =   "500"
      Top             =   195
      Width           =   495
   End
   Begin VB.Frame Frame1 
      Height          =   495
      Left            =   1440
      TabIndex        =   1
      Top             =   600
      Width           =   2295
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   175
         Width           =   2175
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Hesapla"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   690
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Günde Kaç Email gönderilecek?"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   240
      Width           =   2415
   End
End
Attribute VB_Name = "mailhesaplayici"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Label1.Caption = Fix(86400 / Text1.Text) & " sn"
End Sub


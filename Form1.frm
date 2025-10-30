VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form mailgonderici 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "JAXA Geliþmiþ Toplu Mail Gönderim Motoru"
   ClientHeight    =   6630
   ClientLeft      =   6945
   ClientTop       =   5835
   ClientWidth     =   6885
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6630
   ScaleWidth      =   6885
   Begin VB.Frame Frame1 
      Caption         =   "TALÝMATLAR"
      Height          =   6375
      Left            =   120
      TabIndex        =   26
      Top             =   120
      Width           =   6615
      Begin VB.CommandButton Command4 
         Caption         =   "ÝLERÝ >"
         Height          =   615
         Left            =   5280
         TabIndex        =   28
         Top             =   5640
         Width           =   1215
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Gönderim Yöneticisi"
         Height          =   615
         Left            =   2760
         TabIndex        =   54
         Top             =   4800
         Width           =   2415
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Mail Listesi Optimizasyon Aracý"
         Height          =   615
         Left            =   240
         TabIndex        =   38
         Top             =   5520
         Width           =   2415
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Gönderim Süresi Hesaplayýcý"
         Height          =   615
         Left            =   240
         TabIndex        =   33
         Top             =   4800
         Width           =   2415
      End
      Begin VB.Label Label14 
         Caption         =   $"Form1.frx":15162
         Height          =   4455
         Left            =   120
         TabIndex        =   27
         Top             =   360
         Width           =   6375
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "MAÝL ÝÇERÝÐÝ"
      Height          =   6375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6615
      Begin VB.CommandButton Command7 
         Caption         =   "< GERÝ"
         Height          =   615
         Left            =   3840
         TabIndex        =   30
         Top             =   5400
         Width           =   855
      End
      Begin VB.CommandButton Command6 
         Caption         =   "ÝLERÝ >"
         Height          =   615
         Left            =   4920
         TabIndex        =   29
         Top             =   5400
         Width           =   855
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   1560
         TabIndex        =   23
         Text            =   "10"
         Top             =   5595
         Width           =   495
      End
      Begin VB.Timer Timer1 
         Left            =   3240
         Top             =   5520
      End
      Begin VB.TextBox Text3 
         Height          =   3975
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   1320
         Width           =   6255
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Mail Gövde Dosyasý (Template)"
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   480
         Width           =   2655
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         ForeColor       =   &H8000000C&
         Height          =   375
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   480
         Width           =   3375
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   720
         TabIndex        =   1
         Top             =   960
         Width           =   5655
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   2640
         Top             =   5520
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label13 
         Caption         =   "Gönderim Aralýðý:"
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   5640
         Width           =   1455
      End
      Begin VB.Label Label8 
         Caption         =   "sn"
         Height          =   255
         Left            =   2160
         TabIndex        =   24
         Top             =   5640
         Width           =   255
      End
      Begin VB.Label Label11 
         Caption         =   "Konu:"
         Height          =   255
         Left            =   225
         TabIndex        =   5
         Top             =   1005
         Width           =   495
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "ALICILAR"
      Height          =   6375
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   6615
      Begin VB.CommandButton Command9 
         Caption         =   "< GERÝ"
         Height          =   495
         Left            =   120
         TabIndex        =   31
         Top             =   5640
         Width           =   2895
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Bekle"
         Height          =   495
         Left            =   1680
         TabIndex        =   32
         Top             =   5160
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Göndermeye Baþla"
         Height          =   495
         Left            =   120
         TabIndex        =   22
         Top             =   5160
         Width           =   1575
      End
      Begin VB.TextBox sec 
         Height          =   285
         Left            =   3000
         TabIndex        =   37
         Text            =   "0"
         Top             =   2880
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Yarým Kalan Gönderim"
         Height          =   255
         Left            =   4440
         TabIndex        =   35
         Top             =   550
         Width           =   1935
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Yeni Gönderim"
         Height          =   255
         Left            =   4440
         TabIndex        =   34
         Top             =   275
         Value           =   -1  'True
         Width           =   1815
      End
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
         ForeColor       =   &H8000000C&
         Height          =   375
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   350
         Width           =   2415
      End
      Begin VB.ListBox List3 
         Height          =   1620
         ItemData        =   "Form1.frx":154CE
         Left            =   120
         List            =   "Form1.frx":154D0
         TabIndex        =   10
         Top             =   3480
         Width           =   2895
      End
      Begin VB.ListBox List2 
         Height          =   1425
         ItemData        =   "Form1.frx":154D2
         Left            =   120
         List            =   "Form1.frx":154D4
         TabIndex        =   9
         Top             =   1320
         Width           =   2895
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Alýcý Listesi"
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   350
         Width           =   1455
      End
      Begin VB.ListBox List1 
         Height          =   3765
         ItemData        =   "Form1.frx":154D6
         Left            =   3240
         List            =   "Form1.frx":154D8
         TabIndex        =   7
         Top             =   1500
         Width           =   3225
      End
      Begin VB.Label Label16 
         Caption         =   "Henüz Gönderilmemiþ Adresler"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3750
         TabIndex        =   36
         Top             =   1240
         Width           =   2175
      End
      Begin VB.Label Label9 
         Caption         =   "Kullanýlacak Hesap:"
         Height          =   255
         Left            =   3240
         TabIndex        =   21
         Top             =   5400
         Width           =   1455
      End
      Begin VB.Label Label10 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   4800
         TabIndex        =   20
         Top             =   5400
         Width           =   1575
      End
      Begin VB.Label Label12 
         Caption         =   "Toplam Ýþlem:"
         Height          =   255
         Left            =   3600
         TabIndex        =   19
         Top             =   5805
         Width           =   1095
      End
      Begin VB.Label Label6 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   4680
         TabIndex        =   18
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label Label5 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   2040
         TabIndex        =   17
         Top             =   3120
         Width           =   975
      End
      Begin VB.Label Label4 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   1680
         TabIndex        =   16
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Bütün Adresler:"
         Height          =   255
         Left            =   3360
         TabIndex        =   15
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Gönderilemeyen Adresler:"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   3120
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Gönderilen Adresler:"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   375
         Left            =   4800
         TabIndex        =   12
         Top             =   5725
         Width           =   1695
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "GÖNDERÝM YÖNETÝCÝ"
      Height          =   6375
      Left            =   120
      TabIndex        =   39
      Top             =   120
      Visible         =   0   'False
      Width           =   6615
      Begin VB.CheckBox Check1 
         Caption         =   "Yapýlsýn (Gönderim aralýðý 30 sn den az ise seçilmemelidir)"
         Enabled         =   0   'False
         Height          =   195
         Left            =   360
         TabIndex        =   53
         Top             =   5280
         Width           =   5775
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   4440
         TabIndex        =   51
         Text            =   "0"
         Top             =   2280
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Timer Timer2 
         Left            =   3960
         Top             =   2160
      End
      Begin VB.CommandButton Command1k 
         Caption         =   "Tamam"
         Height          =   375
         Left            =   2880
         TabIndex        =   50
         Top             =   5760
         Width           =   1695
      End
      Begin VB.CheckBox Check2k 
         Caption         =   "Bilgisayar kapatýlsýn"
         Height          =   375
         Left            =   360
         TabIndex        =   49
         Top             =   3000
         Width           =   1815
      End
      Begin VB.TextBox Text2k 
         Enabled         =   0   'False
         Height          =   285
         Left            =   840
         TabIndex        =   47
         Text            =   "samedcansin@hotmail.com"
         Top             =   4035
         Width           =   2055
      End
      Begin VB.CheckBox Check3k 
         Caption         =   "Ýþlem istatistikleri þu mail adresine postalansýn"
         Height          =   255
         Left            =   360
         TabIndex        =   46
         Top             =   3600
         Width           =   3495
      End
      Begin VB.CheckBox Check1k 
         Caption         =   "Program kapatýlsýn"
         Height          =   255
         Left            =   360
         TabIndex        =   45
         Top             =   2520
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.TextBox Text1k 
         Enabled         =   0   'False
         Height          =   285
         Left            =   4560
         TabIndex        =   43
         Text            =   "1000"
         Top             =   1425
         Width           =   975
      End
      Begin VB.OptionButton Option2k 
         Caption         =   "Mail listesindeki þu sayýda adrese gönderim yapýlýnca"
         Height          =   255
         Left            =   360
         TabIndex        =   42
         Top             =   1440
         Width           =   4095
      End
      Begin VB.OptionButton Option1k 
         Caption         =   "Mail Listesindeki tüm adreslere gönderim yapýlýnca"
         Height          =   315
         Left            =   360
         TabIndex        =   41
         Top             =   840
         Value           =   -1  'True
         Width           =   4095
      End
      Begin VB.Label Label15 
         Caption         =   "Her gönderimde email doðrulama yapýlsýn mý?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   52
         Top             =   4800
         Width           =   3855
      End
      Begin VB.Label Label3k 
         Caption         =   "Mail:"
         Height          =   255
         Left            =   360
         TabIndex        =   48
         Top             =   4080
         Width           =   495
      End
      Begin VB.Label Label2k 
         Caption         =   "Gönderim iþlemi bitince ne yapýlsýn?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   44
         Top             =   2040
         Width           =   2535
      End
      Begin VB.Label Label1k 
         Caption         =   "Gönderim iþlemi ne zaman bitsin?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   40
         Top             =   360
         Width           =   2535
      End
   End
End
Attribute VB_Name = "mailgonderici"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Check1k_Click()
If Check1k.Value = 1 Then
Check2k.Value = 0
End If
End Sub

Private Sub Check2k_Click()
If Check2k.Value = 1 Then
Check1k.Value = 0
End If
End Sub

Private Sub Check3k_Click()
If Text2k.Enabled = True Then
Text2k.Enabled = False
Else
Text2k.Enabled = True
End If
End Sub

Private Sub Command1_Click()
If Text4.Text = 0 Then
zeromessage = MsgBox("Gönderim aralýðýný 0 olarak belirtmiþsiniz. Geri dönün ve düzeltin!" & vbCrLf & "Gönderim aralýðýný virgüllü belirtirseniz sistem otomatik olarak 0 deðerini verir...")
Else
Command1.Enabled = False
Command3.Enabled = False
Command11.Enabled = False
Command5.Enabled = True
Text4.Enabled = False

Timer1.Interval = 1000
Option1.Enabled = False
Option2.Enabled = False
Text6.Text = 0
End If
End Sub



Private Sub Command10_Click()
mailoptimize.Show
End Sub

Private Sub Command11_Click()
Frame1.Visible = False
Frame2.Visible = False
Frame3.Visible = False
Frame4.Visible = True
End Sub


Private Sub Command12_Click()
ipdegistirici.Show
End Sub

Private Sub Command1k_Click()
Frame1.Visible = True
Frame2.Visible = False
Frame3.Visible = False
Frame4.Visible = False
End Sub

Private Sub Command2_Click()

CommonDialog1.DialogTitle = "Open Access DB File"
CommonDialog1.Filter = "Template Files, Text Files(*.htm, *.html, *.txt, *.asp )|*.htm; *.html; *.txt; *.asp"
CommonDialog1.ShowOpen
If CommonDialog1.FileName <> "" Then

Text1.Text = CommonDialog1.FileName



Text3.Text = ""

Open CommonDialog1.FileName For Input As #1

Do While Not EOF(1)
Input #1, Read
Text3.Text = Text3.Text & Read
Loop

Close #1



End If


End Sub

Private Sub Command3_Click()

CommonDialog1.DialogTitle = "Open HTML Template File"
CommonDialog1.Filter = "Text Files of Email List(*.txt)|*.txt"
CommonDialog1.ShowOpen


List1.Clear
List2.Clear
List3.Clear

Label4.Caption = "0"
Label5.Caption = "0"
Label6.Caption = "0"
Label7.Caption = "0"

Timer2.Interval = 0

On Error Resume Next

If CommonDialog1.FileName <> "" Then

Text2.Text = CommonDialog1.FileName

If Option1.Value = True Then


On Error Resume Next

Open CommonDialog1.FileName For Input As #1



Do While Not EOF(1)
Input #1, Read
If Read <> "" And Read <> " " Then
List1.AddItem Read
Label6.Caption = Label6.Caption + 1
End If
Loop

Close #1


Else

dosyaadi = CommonDialog1.FileName

dosyaadi2 = Left(dosyaadi, Len(dosyaadi) - 4)



Err.Number = 0

On Error Resume Next

Open dosyaadi2 & "_gonderilen.txt" For Input As #1
If Err.Number = 0 Then
Do While Not EOF(1)
Input #1, Read
If Read <> "" And Read <> " " Then
List2.AddItem Read
Label4.Caption = Label4.Caption + 1
Label7.Caption = Label7.Caption + 1
End If
Loop

Close #1

End If





Err.Number = 0

Open dosyaadi2 & "_gitmeyen.txt" For Input As #1
If Err.Number = 0 Then

Do While Not EOF(1)
Input #1, Read
If Read <> "" And Read <> " " Then
List3.AddItem Read
Label5.Caption = Label5.Caption + 1
Label7.Caption = Label7.Caption + 1
End If
Loop

Close #1

End If




Err.Number = 0
Open dosyaadi2 & "_gonderilmeyen.txt" For Input As #1
If Err.Number = 0 Then
Do While Not EOF(1)
Input #1, Read
If Read <> "" And Read <> " " Then
List1.AddItem Read
End If
Loop

Close #1
End If



Err.Number = 0
Open CommonDialog1.FileName For Input As #1
If Err.Number = 0 Then
total = 0

Do While Not EOF(1)
Input #1, Read
If Read <> "" And Read <> " " Then
total = total + 1
End If
Loop

Close #1
End If



End If

End If
End Sub





Private Sub Command4_Click()
Frame1.Visible = False
Frame2.Visible = True
Frame3.Visible = False
Frame4.Visible = False
End Sub

Private Sub Command5_Click()
If Timer1.Interval = 0 Then

Timer1.Interval = 1000
Command5.Caption = "Bekle"
Command1.Enabled = False
Command11.Enabled = False

Command3.Enabled = False
Option1.Enabled = False
Option2.Enabled = False

Else

Timer1.Interval = 0
Command5.Caption = "Devam"

Command3.Enabled = True
Option1.Enabled = True
Option2.Enabled = True
Command11.Enabled = True

End If
End Sub

Private Sub Command6_Click()
Frame1.Visible = False
Frame2.Visible = False
Frame3.Visible = True
Frame4.Visible = False
End Sub

Private Sub Command7_Click()
Frame1.Visible = True
Frame2.Visible = False
Frame3.Visible = False
Frame4.Visible = False
End Sub



Private Sub Command8_Click()
mailhesaplayici.Show
End Sub

Private Sub Command9_Click()
Frame1.Visible = False
Frame2.Visible = True
Frame3.Visible = False
Frame4.Visible = False
End Sub

Private Sub Form_Load()
Frame1.Visible = True
Frame2.Visible = False
Frame3.Visible = False
Command5.Enabled = False


Dim ol As Outlook.Application
    Dim ns As NameSpace
    Dim oRec As Recipient
    Set ol = New Outlook.Application
    Set ns = ol.GetNamespace("MAPI")
    Call ns.Logon(, , , False)
    Set oRec = ns.CurrentUser
    
    On Error Resume Next
    Label10.Caption = oRec.Name
    
    If oRec.Name = "Unknown" Then
    MsgBox ("Outlook ta hesabýnýz görünmüyor! Eðer mevcut hesabýnýz varsa ve gönderen kiþi adý 'Unknown' olarak isimlendirilmiþse lütfen deðiþtiriniz.")
    End
    End If
    
    
    If Err.Number = 287 Then
    MsgBox ("Lütfen programý tekrar çalýþtýrýn ve açýlan Outlook güvenlik teyitlerini onaylayýn!")
    End
    End If


End Sub




Private Sub Option1k_Click()
Text1k.Enabled = False
End Sub

Private Sub Option2k_Click()
Text1k.Enabled = True
End Sub



Private Sub Text4_LostFocus()
Text4.Text = Replace(Text4.Text, ".", ",")
Text4.Text = Fix(Text4.Text)
End Sub

Private Sub Timer1_Timer()
DoEvents


sec.Text = sec.Text + 1

If sec.Text = Text4.Text Then
sec.Text = 0

varforend = List1.ListCount


If Option2k.Value = True And Label7.Caption = Text1k.Text Then
varforend = 0
End If






If varforend = 0 Then
Timer1.Interval = 0
Command1.Enabled = True
Command3.Enabled = True
Command11.Enabled = True
Option1.Enabled = True
Option2.Enabled = True
Timer2.Interval = 10000
Text4.Enabled = True
Else

List1.ListIndex = 0
Label7.Caption = Label7.Caption + 1

emailaddress = List1.Text

List1.RemoveItem (List1.ListIndex)




If List1.ListCount > 0 Then
gonderilmeyen = ""
For i = 0 To List1.ListCount - 1
gonderilmeyen = gonderilmeyen & List1.List(i) & vbCrLf
Next

Open Left(Text2.Text, Len(Text2.Text) - 4) & "_gonderilmeyen" & ".txt" For Output As #1
Print #1, gonderilmeyen
Close #1
End If




If List1.ListCount > 0 Then
List1.ListIndex = 0
End If





If Check1.Value = 1 Then ' KONTROL BAÞLANGICI

result = ""
On Error Resume Next
URLString = emailaddress
    mURL = Trim("http://www.webservicex.net/ValidateEmail.asmx/IsValidEmail?Email=" & URLString)
        Set objXMLHTTP = New MSXML2.XMLHTTP
 objXMLHTTP.Open "GET", mURL, False
    
   objXMLHTTP.Send
    
    result = objXMLHTTP.responseText
    
    result = Replace(result, "<?xml version=""1.0"" encoding=""utf-8""?>", "")
result = Replace(result, "<boolean xmlns=""http://www.webservicex.net"">", "")
result = Replace(result, "</boolean>", "")
result = Replace(result, vbCrLf, "")




  If result = "false" Then
  
Open Left(Text2.Text, Len(Text2.Text) - 4) & "_yanlisadresler" & ".txt" For Output As #1
Print #1, emailaddress
Close #1


  
  
  GoTo hata

End If

If Err.Number <> 0 Then Err.Number = 0






End If ''''KONTROL BÝTÝÞÝ

On Error GoTo hata


Dim uygulama As Outlook.Application
Dim program As Outlook.MailItem
Dim yazi As String

Set uygulama = CreateObject("Outlook.Application")
Set program = uygulama.CreateItem(olMailItem)

program.Subject = Text5.Text
Set myRecipients = program.Recipients
myRecipients.Add emailaddress


program.HTMLBody = Text3.Text




program.Send
List2.AddItem emailaddress
Label4.Caption = Label4.Caption + 1

Open Left(Text2.Text, Len(Text2.Text) - 4) & "_gonderilen" & ".txt" For Append As #1
Print #1, emailaddress
Close #1



With List2
If .ListCount > 0 Then
.ListIndex = (.ListCount - 1)
End If
End With

Exit Sub

hata:
List3.AddItem emailaddress
Label5.Caption = Label5.Caption + 1
Err.Number = 0
Open Left(Text2.Text, Len(Text2.Text) - 4) & "_gitmeyen" & ".txt" For Append As #1
Print #1, emailaddress
Close #1


With List2
If .ListCount > 0 Then
.ListIndex = (.ListCount - 1)
End If
End With


End If
End If
End Sub

Private Sub Timer2_Timer()
Text6.Text = Text6.Text + 1

Select Case Text6.Text

Case 1
If Check3k.Value = 1 Then
''ÝSTATÝSTÝK POSTALA



On Error Resume Next
Dim uygulama As Outlook.Application
Dim program As Outlook.MailItem
Dim yazi As String

Set uygulama = CreateObject("Outlook.Application")
Set program = uygulama.CreateItem(olMailItem)

program.Subject = "toplu mail gönderimi Ýistatistikleri"
Set myRecipients = program.Recipients
myRecipients.Add Text2k.Text


program.HTMLBody = "Toplu mail gönderiminiz bitmiþtir... Ýstatistikler Aþaðýda verilmiþtir:" & vbclrf & "<li>Toplam Mail :" & Label6.Caption & "<li>Yapýlan Deneme:" & Label7.Caption & "<li>Henüz Yapýlmayan Deneme:" & Label6.Caption - Label7.Caption & "<li>Gönderilen:" & Label4.Caption & "<li>Gitmeyen:" & Label5.Caption & "<br><br>Gönderim aþaðýdaki gibidir<br>" & Text5.Text & "<br>" & Text3.Text


program.Send





End If


Case 2
If Check1k.Value = 1 Then
''PROGRAMI KAPAT

Unload mailgonderici
Unload mailhesaplayici
Unload mailoptimize

End If

Case 3
If Check2k.Value = 1 Then
''BÝLGÝSAYARI KAPAT
Shell ("shutdown -s -t 1")
End If

Timer2.Interval = 0


End Select

End Sub

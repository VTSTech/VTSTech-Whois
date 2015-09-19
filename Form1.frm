VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3050F1C5-98B5-11CF-BB82-00AA00BDCE0B}#4.0#0"; "mshtml.tlb"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00404040&
   Caption         =   "VTSTech-Whois v0.0.1"
   ClientHeight    =   7500
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   7050
   FillColor       =   &H00808080&
   FillStyle       =   0  'Solid
   ForeColor       =   &H00808080&
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7500
   ScaleWidth      =   7050
   StartUpPosition =   3  'Windows Default
   Begin MSHTMLCtl.Scriptlet Scriptlet1 
      Height          =   915
      Left            =   0
      TabIndex        =   18
      Top             =   6600
      Width           =   7250
      Scrollbar       =   0   'False
      URL             =   "http://coinurl.com/get.php?id=35561"
   End
   Begin VB.OptionButton Option11 
      BackColor       =   &H00404040&
      Caption         =   "DOMAIN"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Left            =   3240
      TabIndex        =   17
      Top             =   120
      Width           =   1095
   End
   Begin VB.OptionButton Option10 
      BackColor       =   &H00404040&
      Caption         =   "IANA"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Left            =   2040
      TabIndex        =   14
      Top             =   600
      Width           =   735
   End
   Begin VB.OptionButton Option9 
      BackColor       =   &H00404040&
      Caption         =   "AFRINIC"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Left            =   120
      TabIndex        =   12
      Top             =   600
      Width           =   1095
   End
   Begin VB.OptionButton Option8 
      BackColor       =   &H00404040&
      Caption         =   "LACNIC"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Left            =   120
      TabIndex        =   11
      Top             =   360
      Width           =   975
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   4935
      Left            =   105
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   10
      Text            =   "Form1.frx":000C
      Top             =   840
      Width           =   6840
   End
   Begin VB.OptionButton Option7 
      BackColor       =   &H00404040&
      Caption         =   "JPNIC"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Left            =   1200
      TabIndex        =   9
      Top             =   360
      Width           =   855
   End
   Begin VB.OptionButton Option6 
      BackColor       =   &H00404040&
      Caption         =   "KRNIC"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Left            =   2040
      TabIndex        =   8
      Top             =   360
      Width           =   975
   End
   Begin VB.OptionButton Option5 
      BackColor       =   &H00404040&
      Caption         =   "TRACERT"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Left            =   3240
      TabIndex        =   7
      Top             =   600
      Width           =   1215
   End
   Begin VB.OptionButton Option4 
      BackColor       =   &H00404040&
      Caption         =   "PING"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Left            =   3240
      TabIndex        =   6
      Top             =   360
      Width           =   855
   End
   Begin VB.OptionButton Option3 
      BackColor       =   &H00404040&
      Caption         =   "APNIC"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Left            =   2040
      TabIndex        =   5
      Top             =   120
      Width           =   855
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00404040&
      Caption         =   "RIPE"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Left            =   1200
      TabIndex        =   4
      Top             =   120
      Width           =   735
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00404040&
      Caption         =   "ARIN"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "WHOIS"
      Height          =   285
      Left            =   5280
      TabIndex        =   1
      Top             =   480
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   285
      Left            =   4680
      TabIndex        =   0
      Text            =   "888.888.888.888"
      Top             =   75
      Width           =   2055
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   120
      Top             =   5880
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      Caption         =   "[pgp]"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   195
      Left            =   4193
      TabIndex        =   16
      Top             =   6120
      Width           =   465
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      Caption         =   "veritas@vts-tech.org"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   195
      Left            =   2393
      TabIndex        =   15
      Top             =   6120
      Width           =   1770
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      Caption         =   "1VTSgzD24bjkSGdD7kvauxkxHZ4yiwhdU"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   195
      Left            =   1748
      TabIndex        =   13
      Top             =   6360
      Width           =   3555
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      Caption         =   "www.VTS-Tech.org"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   195
      Left            =   2708
      TabIndex        =   2
      Top             =   5880
      Width           =   1635
   End
   Begin VB.Menu File 
      Caption         =   "File"
      Begin VB.Menu Save 
         Caption         =   "Save"
      End
   End
   Begin VB.Menu Quit 
      Caption         =   "Quit"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Banner
Dim dns
Private Const WSADescription_Len As Long = 256
Private Const WSASYS_Status_Len As Long = 128
Private Const WS_VERSION_REQD As Long = &H101
Private Const IP_SUCCESS As Long = 0
Private Const SOCKET_ERROR As Long = -1
Private Const AF_INET As Long = 2

Private Type WSADATA
  wVersion As Integer
  wHighVersion As Integer
  szDescription(0 To WSADescription_Len) As Byte
  szSystemStatus(0 To WSASYS_Status_Len) As Byte
  imaxsockets As Integer
  imaxudp As Integer
  lpszvenderinfo As Long
End Type

Private Declare Function WSAStartup Lib "wsock32" _
  (ByVal VersionReq As Long, _
   WSADataReturn As WSADATA) As Long
  
Private Declare Function WSACleanup Lib "wsock32" () As Long

Private Declare Function inet_addr Lib "wsock32" _
  (ByVal s As String) As Long

Private Declare Function gethostbyaddr Lib "wsock32" _
  (haddr As Long, _
   ByVal hnlen As Long, _
   ByVal addrtype As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" _
   Alias "RtlMoveMemory" _
  (xDest As Any, _
   xSource As Any, _
   ByVal nbytes As Long)
   
Private Declare Function lstrlen Lib "kernel32" _
   Alias "lstrlenA" _
  (lpString As Any) As Long
  Private Function GetHostNameFromIP(ByVal sAddress As String) As String

   Dim ptrHosent As Long
   Dim hAddress As Long
   Dim nbytes As Long
   
   If SocketsInitialize() Then

     'convert string address to long
      hAddress = inet_addr(sAddress)
      
      If hAddress <> SOCKET_ERROR Then
         
        'obtain a pointer to the HOSTENT structure
        'that contains the name and address
        'corresponding to the given network address.
         ptrHosent = gethostbyaddr(hAddress, 4, AF_INET)
   
         If ptrHosent <> 0 Then
         
           'convert address and
           'get resolved hostname
            CopyMemory ptrHosent, ByVal ptrHosent, 4
            nbytes = lstrlen(ByVal ptrHosent)
         
            If nbytes > 0 Then
               sAddress = Space$(nbytes)
               CopyMemory ByVal sAddress, ByVal ptrHosent, nbytes
               GetHostNameFromIP = sAddress
            End If
         
         Else
            'MsgBox "Call to gethostbyaddr failed."
         End If 'If ptrHosent
      
      SocketsCleanup
      
      Else
         'MsgBox "String passed is an invalid IP."
      End If 'If hAddress
   
   Else
      MsgBox "Sockets failed to initialize."
   End If  'If SocketsInitialize
      
End Function

Private Function SocketsInitialize() As Boolean

   Dim WSAD As WSADATA
   
   SocketsInitialize = WSAStartup(WS_VERSION_REQD, WSAD) = IP_SUCCESS
    
End Function


Private Sub SocketsCleanup()
   
   If WSACleanup() <> 0 Then
       MsgBox "Windows Sockets error occurred in Cleanup.", vbExclamation
   End If
    
End Sub
'Private WithEvents objDOS As DOSOutputs
Private Sub Command1_Click()
dns = ""
dns = GetHostNameFromIP(Text1.Text)
Text2.Text = ""
Banner = 0
If Option1.Value = True Then
Winsock1.Close
Winsock1.RemoteHost = "whois.arin.net"
Winsock1.RemotePort = 43
Winsock1.Connect
ElseIf Option2.Value = True Then
Winsock1.Close
Winsock1.RemoteHost = "whois.ripe.net"
Winsock1.RemotePort = 43
Winsock1.Connect
ElseIf Option3.Value = True Then
Winsock1.Close
Winsock1.RemoteHost = "whois.apnic.net"
Winsock1.RemotePort = 43
Winsock1.Connect
ElseIf Option4.Value = True Then
Shell ("cmd.exe /K ping.exe " & Text1.Text), vbNormalFocus
ElseIf Option5.Value = True Then
Shell ("cmd.exe /K tracert.exe " & Text1.Text), vbNormalFocus
ElseIf Option6.Value = True Then
Winsock1.Close
Winsock1.RemoteHost = "whois.nic.or.kr"
Winsock1.RemotePort = 43
Winsock1.Connect
ElseIf Option7.Value = True Then
Winsock1.Close
Winsock1.RemoteHost = "whois.nic.ad.jp"
Winsock1.RemotePort = 43
Winsock1.Connect
ElseIf Option8.Value = True Then
Winsock1.Close
Winsock1.RemoteHost = "whois.lacnic.net"
Winsock1.RemotePort = 43
Winsock1.Connect
ElseIf Option9.Value = True Then
Winsock1.Close
Winsock1.RemoteHost = "whois.afrinic.net"
Winsock1.RemotePort = 43
Winsock1.Connect
ElseIf Option10.Value = True Then
Winsock1.Close
Winsock1.RemoteHost = "whois.iana.org"
Winsock1.RemotePort = 43
Winsock1.Connect
ElseIf Option11.Value = True Then
Winsock1.Close
Winsock1.RemoteHost = "whois.uwhois.net"
Winsock1.RemotePort = 43
Winsock1.Connect
Else
MsgBox "Select a WHOIS Database first!"
End If

End Sub

Private Sub Command2_Click()
Unload Form1
End Sub

Private Sub Form_Load()
On Error Resume Next
'Option1 = ARIN
'Option2 = RIPE
'Option3 = APNIC
'Option4 = PING
'Option5 = TRACERT
'Option6 = KRNIC
Option1.Enabled = True
Option2.Enabled = True
Option3.Enabled = True
Option4.Enabled = True
Option5.Enabled = True
Option6.Enabled = True
Option7.Enabled = True
Option8.Enabled = True
Option9.Enabled = True
Option10.Enabled = True
'http://coinurl.com/get.php?id=35561
Scriptlet1.Width = 7045
Scriptlet1.Height = 905
Text1.Text = Winsock1.LocalIP
Text2.Text = "Written by VTSTech (Veritas Technical Solutions)" & vbCrLf & "Web: www.VTS-Tech.org" & vbCrLf & "E-Mail: veritas@vts-tech.org" & vbCrLf & "XMPP: veritas@creep.im" & vbCrLf & "BTC 1VTSgzD24bjkSGdD7kvauxkxHZ4yiwhdU"
Text2.Text = Text2.Text & vbCrLf & vbCrLf
Text2.Text = Text2.Text & "-----BEGIN PGP PUBLIC KEY BLOCK-----" & vbCrLf
Text2.Text = Text2.Text & "" & vbCrLf
Text2.Text = Text2.Text & "mQGiBFPHmXYRBADt/EkhMXIeEvcI7+N8d9G+c7Bq/SqYDRXzu5Ewv9MNWUqDyyGp" & vbCrLf
Text2.Text = Text2.Text & "Z5FzgtSeSPLcCA8elv0h71uVDI+PnwLY2Bn0pcJN81KP2snH5X36RjPuRGhz4rvJ" & vbCrLf
Text2.Text = Text2.Text & "IN87/F9aWYrcXHhQcWC3FHiB9tEUQFlCfNWb/zP+yPH9KEdt4hhsNi10mwCgg8YA" & vbCrLf
Text2.Text = Text2.Text & "xnlC2O9UiFq+r1gJElWknHED/1QL+9o3XQ5mNRB/aQTpp9Mhgvn8gXHomLRLoA0h" & vbCrLf
Text2.Text = Text2.Text & "ejsLVrEHRPg2birQC/ry4NP8QKXisR/WeNIhRhMtaAT+cWGdnxlnTwtlDtquoJCG" & vbCrLf
Text2.Text = Text2.Text & "jMYaC6k5iVbQsGBiWrpzG6ey3Lp0oATYcfjBTqvbyrFU3Je0s1xtTUTL8gH0NoAb" & vbCrLf
Text2.Text = Text2.Text & "SA3zBACPJd2VQlYEiflVQ3FE5GkiD1iUIEdkYTrZ8FTp5lxxa1mIOebvynYB5xLI" & vbCrLf
Text2.Text = Text2.Text & "mwnY2xaMsuN5IsWFyXPFv2FqwM7Kf8Amn5N5E+oyVt4USPl3/l2u3iFev9oGeH2o" & vbCrLf
Text2.Text = Text2.Text & "ANB8QhhRtq1BmP38eDcPJmQpRFaA+BS9SJm+mmDIJkyKyQ3xOrQoVmVyaXRhcyAo" & vbCrLf
Text2.Text = Text2.Text & "VlRTVGVjaCkgPHZlcml0YXNAdnRzLXRlY2gub3JnPohbBBMRAgAbBQJTx5l2BgsJ" & vbCrLf
Text2.Text = Text2.Text & "CAcDAgMVAgMDFgIBAh4BAheAAAoJEMJ0uz0W6NU0lvYAn00fR/DBHa1atUm4DUKV" & vbCrLf
Text2.Text = Text2.Text & "qUH94rD+AJ9S7utJ0phinxgY7ifxfX1+g74KwrkCDQRTx5l4EAgAvFSy5qqNKypt" & vbCrLf
Text2.Text = Text2.Text & "rI0PUGA1IkEgk9E21bfNmAVzizm0gEtJSmsMPD4qfFirfiuVpFbUn//PyJFGF5ep" & vbCrLf
Text2.Text = Text2.Text & "OCuhXXR4pU4dM3P6dfDpUimPz5nm5VuuCtIEHqMVq9h4BZObV+03uOu2L26YGXeY" & vbCrLf
Text2.Text = Text2.Text & "hufam0XAuPGJ0ohwABdOnumUbvVS7cy3aMmigbvubdfQnXgfKyeW+6TR9M1xLgo3" & vbCrLf
Text2.Text = Text2.Text & "X1F3bS1GDuj6suJkYFBCQSD4CmCgWkdC64Qp4aqUdruRUyU9IwgAbUiyerfs5oy9" & vbCrLf
Text2.Text = Text2.Text & "y8I6Ht+yplF1Lyq9BQoRguHw/ZjM5TnvNYTpQQRqw+YqcTL5YrUJrYvaVFaqNNyw" & vbCrLf
Text2.Text = Text2.Text & "LP2wyxBVxwADBQf/X32aRvWfUcBuru+3ZZ8PDPIc2N43YcpThr6dCkEx/I4dcn6t" & vbCrLf
Text2.Text = Text2.Text & "60TmO3xtWiGTfyZrafKYctI0smYyz5eDNcmJfwJrei3aECaRv9yZM26m1fOzgVQW" & vbCrLf
Text2.Text = Text2.Text & "4halibOcdpL3I4kR5nTfiGKddGSGPpcpG4MOtwtnFIRCC51BBngE3SsWIwRNnQas" & vbCrLf
Text2.Text = Text2.Text & "Yr7Hl4k1tikwNC/v/zzWiAt7ryhiWKEPT/fiZflf54oh9sWq+LhdzweTKxTxX9+m" & vbCrLf
Text2.Text = Text2.Text & "kFLjU7oBq0J9w8yNii4XMyj+Myf75kXPr9XyVK3fu9cvMVK0p3N/PLLp/uF10axK" & vbCrLf
Text2.Text = Text2.Text & "vho28zvQx6Tyvq7z8+Ep5HERfJkM+ew4z0OEf4hGBBgRAgAGBQJTx5l4AAoJEMJ0" & vbCrLf
Text2.Text = Text2.Text & "uz0W6NU0vZkAnjyBXJKdNzauwZNcc1GnlAnKbh/EAJkBYw8+jPBlX0vC4KV4t2A1" & vbCrLf
Text2.Text = Text2.Text & "OVNYbg==                                                        " & vbCrLf
Text2.Text = Text2.Text & "=SyO+                                                           " & vbCrLf
Text2.Text = Text2.Text & "-----END PGP PUBLIC KEY BLOCK-----                              " & vbCrLf
Text2.Text = Text2.Text & "                                                                " & vbCrLf
Banner = 0
End Sub

Private Sub Label1_Click()
Shell ("cmd.exe /c start http://www.vts-tech.org"), vbHide
End Sub

Private Sub Label3_Click()
Clipboard.Clear
Clipboard.SetText "1VTSgzD24bjkSGdD7kvauxkxHZ4yiwhdU"
MsgBox "BitCoin Address copied to clipboard"
End Sub

Private Sub Label4_Click()
Shell ("cmd.exe /c start http://vts-tech.org/veritas.key.pub.asc"), vbHide
End Sub

Private Sub Option1_Click()
Text2.Text = ""
End Sub
Private Sub Option2_Click()
Text2.Text = ""
End Sub
Private Sub Option3_Click()
Text2.Text = ""
End Sub
Private Sub Option6_Click()
Text2.Text = ""
End Sub
Private Sub Option7_Click()
Text2.Text = ""
End Sub
Private Sub Option8_Click()
Text2.Text = ""
End Sub
Private Sub Option9_Click()
Text2.Text = ""
End Sub
Private Sub Option10_Click()
Text2.Text = ""
End Sub

Private Sub Quit_Click()
Unload Form1
End
End Sub

Private Sub Save_Click()
Open VB.App.Path & "\VTSTech-Whois.txt" For Output As #1
Print #1, Text2.Text
Close #1
MsgBox "Output written to " & VB.App.Path & "\VTSTech-Whois.txt"
End Sub

Private Sub Winsock1_Connect()
On Error Resume Next
If Option1.Value = True Then
Winsock1.SendData "n + " & Text1.Text & vbCrLf
ElseIf Option2.Value = True Then
Winsock1.SendData "-B " & Text1.Text & vbCrLf
ElseIf Option3.Value = True Then
Winsock1.SendData Text1.Text & vbCrLf
ElseIf Option6.Value = True Then
Winsock1.SendData Text1.Text & vbCrLf
ElseIf Option7.Value = True Then
Winsock1.SendData Text1.Text & vbCrLf
ElseIf Option8.Value = True Then
Winsock1.SendData Text1.Text & vbCrLf
ElseIf Option9.Value = True Then
Winsock1.SendData "-B " & Text1.Text & vbCrLf
ElseIf Option10.Value = True Then
Winsock1.SendData Text1.Text & vbCrLf
ElseIf Option11.Value = True Then
Winsock1.SendData Text1.Text & vbCrLf
End If
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next
Winsock1.GetData SockData, vbString
For x = 1 To Len(SockData)
SockSrch = Replace(SockData, Chr(10), vbCrLf)
Next x

If Len(dns) > 1 And Len(SockSrch) > 10 And Banner = 0 Then
Text2.Text = "Hostname: " & dns & vbCrLf & SockSrch
Banner = 1
ElseIf Banner = 0 And Len(dns) = 0 Then
Text2.Text = "IP: " & Text1.Text & vbCrLf & SockSrch
Banner = 1
ElseIf Banner = 1 Then
Text2.Text = Text2.Text & SockSrch
End If

End Sub

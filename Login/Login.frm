VERSION 5.00
Begin VB.Form Login 
   BackColor       =   &H80000009&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3915
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6015
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3915
   ScaleWidth      =   6015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cmbUsrType 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2400
      TabIndex        =   7
      Top             =   240
      Width           =   2295
   End
   Begin VB.Timer LoginLock 
      Interval        =   8000
      Left            =   0
      Top             =   2400
   End
   Begin VB.Timer tmrBlink 
      Interval        =   200
      Left            =   0
      Top             =   2880
   End
   Begin VB.ComboBox cmbUser 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3120
      TabIndex        =   5
      Top             =   1680
      Width           =   2295
   End
   Begin VB.TextBox txtPass 
      Height          =   285
      Left            =   3120
      TabIndex        =   4
      Top             =   2520
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Quit"
      Height          =   375
      Left            =   4560
      TabIndex        =   1
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Login"
      Height          =   375
      Left            =   3240
      TabIndex        =   0
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label lblLock 
      BackColor       =   &H8000000E&
      Caption         =   "You have attemped to login three (3) times. You are temporarily Lock-out!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   14
      Top             =   840
      Width           =   5535
   End
   Begin VB.Image Image2 
      Height          =   225
      Left            =   1080
      Picture         =   "Login.frx":0000
      Top             =   240
      Width           =   225
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter your Password"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000011&
      Height          =   255
      Left            =   3240
      TabIndex        =   13
      Top             =   2280
      Width           =   2175
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Select your Username"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000011&
      Height          =   255
      Left            =   3240
      TabIndex        =   12
      Top             =   1440
      Width           =   2175
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "[2]"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000A&
      Height          =   375
      Left            =   2880
      TabIndex        =   11
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "[1]"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000A&
      Height          =   375
      Left            =   2880
      TabIndex        =   10
      Top             =   1200
      Width           =   375
   End
   Begin VB.Label UserLogin 
      BackStyle       =   0  'Transparent
      Caption         =   "User's Login"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   3360
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   1680
      Left            =   0
      Picture         =   "Login.frx":0312
      Top             =   2400
      Width           =   1740
   End
   Begin VB.Label Label3 
      Caption         =   "User Type:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1320
      TabIndex        =   8
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label4 
      Height          =   735
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   6015
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1440
      TabIndex        =   3
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Username:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1440
      TabIndex        =   2
      Top             =   1680
      Width           =   1575
   End
End
Attribute VB_Name = "Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ctr As Byte
Dim UserType As Integer
Dim UserID As Integer

Private Sub cmbUser_Click()
UserID = Me.cmbUser.ItemData(Me.cmbUser.ListIndex)
txtPass.SetFocus
End Sub

Private Sub cmbUsrType_Click()
UserType = Me.cmbUsrType.ItemData(Me.cmbUsrType.ListIndex)
If cmbUsrType.Text = "Administrator" Then
    'Place your code here
End If

End Sub

Private Sub Command1_Click()
If txtPass.Text = "" Then
    MsgBox "Password is empty! Please fill to continue...", vbExclamation + vbOKOnly, "Empty"
    txtPass.SetFocus
    Exit Sub
End If

Set RstRecord = New ADODB.Recordset

    With RstRecord
        .Open "Select * from Users where Username ='" & cmbUser.Text & "' and Password ='" & txtPass & "'", ConConnection, adOpenStatic, adLockOptimistic
        
            If .EOF Then
                MsgBox "Invalid Login!", vbExclamation
                txtPass.SetFocus
                ctr = ctr + 1
            Else
                Unload Me
                Sample.Show
                MsgBox "Successfully Login!", vbInformation
            End If
            
                If ctr = 3 Then
                    objLock (True)
                    UserLogin.Caption = "Lock!"
                End If
                
                If ctr = 6 Then
                    MsgBox "You hae attempted to login six (6) times. this Program will now close!", vbExclamation, "Close"
                    End
                End If
    End With


    
End Sub

Private Sub Command2_Click()
Dim answer As Integer
    answer = MsgBox("Thank you for downloading! Program by Froilan C. Alejandro, 4th year BSIT student of Isabela State University, not yet a programmer but has knowledge in programming!", vbOKOnly, "Thank you!")
     If vbOK Then
        End
    End If
End Sub

Private Sub Form_Load()

Call DBConnection

'CENTER FORM
Call CenterForm(Me)

Call FillAll

objLock (False)

txtPass.PasswordChar = "*"
End Sub
'ADD DATA TO COMBOBOX
Sub FillAll()
Dim rsShow As New ADODB.Recordset
    With rsShow
    
        .Open "Select User_Type from UserType", ConConnection, adOpenKeyset, adLockPessimistic
        If Not .BOF Then
          .MoveFirst
        End If
          Me.cmbUsrType.Clear
          Do While Not .EOF
              Me.cmbUsrType.AddItem !User_Type
              'Me.cmbGender.ItemData(Me.cmbGender.NewIndex) = !Username
              .MoveNext
          Loop
        .Close
        
        .Open "Select Username from Users", ConConnection, adOpenKeyset, adLockPessimistic
        If Not .BOF Then
          .MoveFirst
        End If
          Me.cmbUser.Clear
          Do While Not .EOF
              Me.cmbUser.AddItem !Username
              'Me.cmbGender.ItemData(Me.cmbGender.NewIndex) = !Username
              .MoveNext
          Loop
        .Close
    End With
End Sub
'LOCK OBJECTS
Sub objLock(x As Boolean)
cmbUser.Enabled = Not x
txtPass.Enabled = Not x
Command1.Enabled = Not x
Command2.Enabled = Not x
lblLock.Visible = x
 LoginLock.Enabled = True
End Sub

Private Sub LoginLock_Timer()
If LoginLock.Interval = 8000 Then
    objLock (False)
    UserLogin.Caption = "User's Login"
End If
LoginLock.Enabled = False
End Sub

Private Sub tmrBlink_Timer()
If UserLogin.Visible = False Then
    UserLogin.Visible = True
Else
    UserLogin.Visible = False
End If
End Sub

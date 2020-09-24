VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form Sample 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sample List"
   ClientHeight    =   4155
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6045
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4155
   ScaleWidth      =   6045
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "About"
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   3600
      Width           =   855
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2640
      Top             =   1800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483624
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sample.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3015
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   5318
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ColHdrIcons     =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483624
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "UserID"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Username"
         Object.Width           =   2647
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Password"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "Sample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command2_Click()
About.Show
End Sub

Private Sub Form_Load()
'CENTER FORM
Call CenterForm(Me)

LIST
End Sub

Sub LIST()
On Error Resume Next

ListView1.ListItems.Clear

Dim criteria As String

Set RstRecord = New ADODB.Recordset

    With RstRecord
    
        criteria = "Select * from Users order by Username asc"
        
            .Open criteria, ConConnection, 3, 3
                
            Do While Not .EOF
            
            ListView1.ListItems.Add , , !UserID, 1, 1
            
            ListView1.ListItems(ListView1.ListItems.Count).SubItems(1) = !Username
            
            ListView1.ListItems(ListView1.ListItems.Count).SubItems(2) = !Password
                                           
            .MoveNext
            
            Loop
                       
        .Close
        
    End With
    
    Set RstRecord = Nothing
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim answer As Integer
    answer = MsgBox("Thank you for downloading! Program by Froilan C. Alejandro, 4th year BSIT student of Isabela State University, not yet a programmer but has knowledge in programming!", vbOKOnly, "Thank you!")
    If vbOK Then
        End
    End If
End Sub

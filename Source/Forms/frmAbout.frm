VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About RandArray"
   ClientHeight    =   10155
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14550
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10155
   ScaleWidth      =   14550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdDone 
      Caption         =   "Done"
      Height          =   375
      Left            =   12720
      TabIndex        =   1
      Top             =   9600
      Width           =   1455
   End
   Begin VB.TextBox txtAbout 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9015
      Left            =   360
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   360
      Width           =   13935
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdDone_Click()
    Unload Me
    
End Sub

Private Sub Form_Load()
    Dim iFile As String
    Dim iHand As Integer
    Dim iData As String
    iData = String(10000, " ")
    
    iFile = App.Path & "\RandArray.txt"
    If Len(Dir(iFile)) > 0 Then
        iHand = FreeFile()
        Open iFile For Binary Access Read As #iHand
        Get #iHand, , iData
        Me.txtAbout.Text = iData
        Close #iHand
    End If
    
        
    
End Sub

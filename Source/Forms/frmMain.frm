VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "RandArray"
   ClientHeight    =   7065
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8925
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7065
   ScaleWidth      =   8925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdDone 
      Caption         =   "Done"
      Height          =   375
      Left            =   6840
      TabIndex        =   7
      Top             =   6360
      Width           =   1575
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "About"
      Height          =   375
      Left            =   6840
      TabIndex        =   6
      Top             =   2160
      Width           =   1575
   End
   Begin VB.TextBox txtUDTCount 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   5040
      TabIndex        =   4
      Text            =   "0"
      Top             =   6480
      Width           =   1215
   End
   Begin VB.CommandButton cmdAlpha 
      Caption         =   "Alpha"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6840
      TabIndex        =   3
      ToolTipText     =   "Click this to sort the elements alphabetically."
      Top             =   1560
      Width           =   1575
   End
   Begin VB.CommandButton cmdRandom 
      Caption         =   "Randomize"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6840
      TabIndex        =   2
      ToolTipText     =   "Click this to randomize the elements."
      Top             =   960
      Width           =   1575
   End
   Begin VB.CommandButton cmdBuildUDT 
      Caption         =   "Build UDTData"
      Height          =   375
      Left            =   6840
      TabIndex        =   1
      ToolTipText     =   "Click this to enter the number of elements to create."
      Top             =   360
      Width           =   1575
   End
   Begin VB.ListBox lstUDTData 
      Height          =   6105
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   6015
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Number of UDT Items:"
      Height          =   255
      Left            =   2760
      TabIndex        =   5
      Top             =   6480
      Width           =   2175
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAbout_Click()
    frmAbout.Show vbModal, Me
    
End Sub

Private Sub cmdAlpha_Click()
    quickSortLast udtData, 0, numUDTData - 1
    showUDTData

End Sub

Private Sub cmdBuildUDT_Click()
    Dim iVal As Long
    iVal = InputBox("Enter number of items to create:", "UDT Items")
    If Len(iVal) > 0 Then
        numUDTData = Val(iVal)
        
        udtData = createUDTArray(numUDTData)
        quickSortLast udtData, 0, numUDTData - 1
        showUDTData
        Me.cmdRandom.Enabled = True
        Me.cmdAlpha.Enabled = True
        
    End If
    Me.txtUDTCount.Text = Format(numUDTData)
    
    
End Sub
Private Sub showUDTData()
    Dim ctr As Long
    Me.lstUDTData.Clear
    For ctr = 0 To numUDTData - 1
        Me.lstUDTData.AddItem udtData(ctr).lastName & ", " & udtData(ctr).firstName & " : " & Format(udtData(ctr).rndValue, "0.00000")
    Next
    
End Sub

Private Sub cmdDone_Click()
    Unload Me
    End
    
End Sub

Private Sub cmdRandom_Click()
    randomizeArray
    
    showUDTData
    
End Sub

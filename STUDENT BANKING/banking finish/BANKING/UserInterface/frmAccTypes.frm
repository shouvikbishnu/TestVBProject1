VERSION 5.00
Begin VB.Form frmAccTypes 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Account Types"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5535
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   5535
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   0
      TabIndex        =   10
      Top             =   2160
      Width           =   5535
      Begin VB.CommandButton cmdQuit 
         Caption         =   "&Quit"
         Height          =   375
         Left            =   4320
         TabIndex        =   19
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   375
         Left            =   3240
         TabIndex        =   18
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   2160
         TabIndex        =   17
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   375
         Left            =   1200
         TabIndex        =   16
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add New"
         Height          =   375
         Left            =   240
         TabIndex        =   15
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton cmdFirst 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   240
         Picture         =   "frmAccTypes.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   600
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdPrevious 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   600
         Picture         =   "frmAccTypes.frx":0342
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   600
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdNext 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   4560
         Picture         =   "frmAccTypes.frx":0684
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   600
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdLast 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   4920
         Picture         =   "frmAccTypes.frx":09C6
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   600
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.Label lblStatus 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   960
         TabIndex        =   20
         Top             =   600
         Width           =   3600
      End
   End
   Begin VB.TextBox txtMinBalance 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1350
      TabIndex        =   9
      Top             =   1770
      Width           =   1935
   End
   Begin VB.TextBox txtInterestRate 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1350
      TabIndex        =   7
      Top             =   1395
      Width           =   1455
   End
   Begin VB.TextBox txtDescription 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1350
      TabIndex        =   5
      Top             =   1020
      Width           =   3375
   End
   Begin VB.TextBox txtAccountName 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1350
      TabIndex        =   3
      Top             =   630
      Width           =   3375
   End
   Begin VB.TextBox txtAccountID 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1350
      TabIndex        =   1
      Top             =   255
      Width           =   1335
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "MinBalance:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   4
      Left            =   435
      TabIndex        =   8
      Top             =   1815
      Width           =   885
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "InterestRate:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   3
      Left            =   405
      TabIndex        =   6
      Top             =   1440
      Width           =   915
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Description:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   2
      Left            =   480
      TabIndex        =   4
      Top             =   1065
      Width           =   840
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "AccountName:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   255
      TabIndex        =   2
      Top             =   675
      Width           =   1065
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "AccountID:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   510
      TabIndex        =   0
      Top             =   300
      Width           =   810
   End
End
Attribute VB_Name = "frmAccTypes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAdd_Click()
NewRecord = True

cmdAdd.Enabled = False
cmdSave.Enabled = True

cmdCancel.Enabled = True
cmdEdit.Enabled = False
cmdQuit.Enabled = False
Call UnLock_Form_Controls(Me)
Call clear_Form_Controls(Me)
Call GenerateNewAccountCode
txtAccountID.Locked = True
txtAccountName.SetFocus
End Sub

Private Sub cmdCancel_Click()
cmdAdd.Enabled = True
cmdSave.Enabled = False

cmdCancel.Enabled = False
cmdEdit.Enabled = True
cmdQuit.Enabled = True
With rsAccTypes
    If NewRecord = True Then
        .CancelUpdate
        NewRecord = False
    Else
        .CancelUpdate
    End If
    Call DisplayaccTypes(rsAccTypes)
End With
Call Lock_Form_Controls(Me)
End Sub

Private Sub cmdEdit_Click()
NewRecord = False
cmdAdd.Enabled = False
cmdSave.Enabled = True

cmdCancel.Enabled = False
cmdEdit.Enabled = False
cmdQuit.Enabled = False
Call UnLock_Form_Controls(Me)
End Sub

Private Sub cmdFirst_Click()
Call MoveToFirst(rsAccTypes)
Call DisplayaccTypes(rsAccTypes)
lblStatus.Caption = CStr("Record :" & rsAccTypes.AbsolutePosition & " of " & rsAccTypes.RecordCount)
End Sub

Private Sub cmdLast_Click()
Call MoveToLast(rsAccTypes)
Call DisplayaccTypes(rsAccTypes)
lblStatus.Caption = CStr("Record :" & rsAccTypes.AbsolutePosition & " of " & rsAccTypes.RecordCount)
End Sub

Private Sub cmdNext_Click()
Call MoveToNext(rsAccTypes)
Call DisplayaccTypes(rsAccTypes)
lblStatus.Caption = CStr("Record :" & rsAccTypes.AbsolutePosition & " of " & rsAccTypes.RecordCount)

End Sub

Private Sub cmdPrevious_Click()
Call MoveToPrev(rsAccTypes)
Call DisplayaccTypes(rsAccTypes)
lblStatus.Caption = CStr("Record :" & rsAccTypes.AbsolutePosition & " of " & rsAccTypes.RecordCount)
End Sub

Private Sub cmdQuit_Click()
Unload Me
End Sub

Public Sub GenerateNewAccountCode()
    Dim lastnumber As Long, newnumber As Long
    'Check if there are records in the file
    With rsAccTypes
    If .BOF = True And .EOF = True Then
        lastnumber = 1000
    Else
        .MoveLast
        lastnumber = !AccountID
    End If
    'Generate New Number
    newnumber = lastnumber + 1
    txtAccountID.Text = newnumber
    End With
End Sub

Private Sub cmdSave_Click()
With rsAccTypes
If NewRecord = True Then .AddNew
!AccountID = txtAccountID.Text
!AccountName = txtAccountName.Text
!Description = txtDescription.Text
!IntrestRate = txtInterestRate.Text
!MinBalance = txtMinBalance.Text
.Update
End With
End Sub

Private Sub Form_Load()
Call connectDatabase
cmdAdd.Enabled = True
cmdSave.Enabled = False
cmdCancel.Enabled = False
cmdEdit.Enabled = False
cmdQuit.Enabled = True
Call Lock_Form_Controls(Me)
Call DisplayaccTypes(rsAccTypes)
lblStatus.Caption = CStr("Record :" & rsAccTypes.AbsolutePosition & " of " & rsAccTypes.RecordCount)
Call DisplayaccTypes(rsAccTypes)
End Sub

Public Sub DisplayaccTypes(myRs As Recordset)
With myRs
If .BOF = True And .EOF = True Then Exit Sub

    txtAccountID.Text = !AccountID
    txtAccountName.Text = !AccountName
    txtDescription.Text = !Description
    txtInterestRate.Text = !InterestRate
    txtMinBalance.Text = !MinBalance
End With
End Sub

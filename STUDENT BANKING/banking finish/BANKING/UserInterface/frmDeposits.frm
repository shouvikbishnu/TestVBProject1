VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDeposits 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Deposits"
   ClientHeight    =   4365
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6375
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   6375
   Begin MSComCtl2.DTPicker txtDated 
      Height          =   375
      Left            =   4200
      TabIndex        =   25
      Top             =   240
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      Format          =   22740993
      CurrentDate     =   38293
   End
   Begin VB.TextBox txtTransactionID 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1575
      TabIndex        =   15
      Top             =   240
      Width           =   1095
   End
   Begin VB.TextBox txtAccountNo 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1575
      TabIndex        =   14
      Top             =   1005
      Width           =   2295
   End
   Begin VB.TextBox txtNarration 
      Appearance      =   0  'Flat
      Height          =   645
      Left            =   1575
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   13
      Top             =   1380
      Width           =   3375
   End
   Begin VB.TextBox txtAmountDeposited 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1575
      TabIndex        =   12
      Top             =   2115
      Width           =   1575
   End
   Begin VB.TextBox txtMode 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1590
      TabIndex        =   11
      Top             =   2790
      Width           =   2400
   End
   Begin VB.TextBox txtCheckNo 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1575
      TabIndex        =   10
      Top             =   3240
      Width           =   3375
   End
   Begin VB.OptionButton optCash 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Cash"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1590
      TabIndex        =   9
      Top             =   2550
      Width           =   975
   End
   Begin VB.OptionButton optCheque 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Cheque"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2670
      TabIndex        =   8
      Top             =   2550
      Width           =   1095
   End
   Begin VB.ComboBox cboCustomerNo 
      Height          =   315
      Left            =   1590
      TabIndex        =   7
      Top             =   630
      Width           =   2055
   End
   Begin VB.OptionButton optOthers 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Other..Specify"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4110
      TabIndex        =   6
      Top             =   2520
      Width           =   1815
   End
   Begin VB.TextBox txtOther 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4110
      TabIndex        =   5
      Top             =   2790
      Width           =   1815
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   375
      Left            =   1320
      TabIndex        =   4
      Top             =   3840
      Width           =   975
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "&Quit"
      Height          =   375
      Left            =   5040
      TabIndex        =   3
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   375
      Left            =   3720
      TabIndex        =   2
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2400
      TabIndex        =   1
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton cmdDeposit 
      Caption         =   "&Deposit"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   3840
      Width           =   1095
   End
   Begin VB.Label lblBalance 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   4320
      TabIndex        =   24
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TransactionID:"
      Height          =   195
      Index           =   0
      Left            =   495
      TabIndex        =   23
      Top             =   285
      Width           =   1050
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CustomerNo:"
      Height          =   195
      Index           =   1
      Left            =   630
      TabIndex        =   22
      Top             =   660
      Width           =   915
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "AccountNo:"
      Height          =   195
      Index           =   2
      Left            =   690
      TabIndex        =   21
      Top             =   1050
      Width           =   855
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Narration:"
      Height          =   195
      Index           =   3
      Left            =   855
      TabIndex        =   20
      Top             =   1425
      Width           =   690
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "AmountDeposited:"
      Height          =   195
      Index           =   4
      Left            =   240
      TabIndex        =   19
      Top             =   2160
      Width           =   1305
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mode:"
      Height          =   195
      Index           =   5
      Left            =   1095
      TabIndex        =   18
      Top             =   2550
      Width           =   450
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CheckNo:"
      Height          =   195
      Index           =   6
      Left            =   825
      TabIndex        =   17
      Top             =   3285
      Width           =   720
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Dated:"
      Height          =   195
      Index           =   7
      Left            =   3585
      TabIndex        =   16
      Top             =   300
      Width           =   480
   End
End
Attribute VB_Name = "frmDeposits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim currBalance As Currency

Private Sub cboCustomerNo_Click()

Set rsTemp = New ADODB.Recordset
rsTemp.Open "Select * FROM tblCustomers WHERE CustomerID='" & cboCustomerNo.Text & "'", cnBank, adOpenKeyset, adLockOptimistic
With rsTemp

If .RecordCount > 0 Then
    txtAccountNo.Text = !AccountNo
    
    txtNarration.SetFocus
Else
MsgBox "Invalid Customer Code", vbInformation

txtAccountNo.Text = ""
Exit Sub
End If
.Close
End With

Set rsTemp = New ADODB.Recordset
rsTemp.Open "Select * FROM tblBalances WHERE CustomerID='" & cboCustomerNo.Text & "'", cnBank, adOpenKeyset, adLockOptimistic
With rsTemp
If .RecordCount > 0 Then
lblBalance.Caption = !Balance
Else
Exit Sub
End If
.Close
End With

End Sub

Private Sub cboCustomerNo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Call cboCustomerNo_Click
End Sub

Private Sub cmdDeposit_Click()
NewRecord = True
Call clear_Form_Controls(Me)
Call GenerateNewTransactCode
cboCustomerNo.SetFocus
End Sub

Private Sub cmdQuit_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()
If txtTransactionID.Text = "" Then
MsgBox "Please Enter the Transaction ID.", vbInformation
txtTransactionID.SetFocus
Exit Sub
End If

If cboCustomerNo.Text = "" Then
MsgBox "Please Enter the Customer ID", vbInformation
cboCustomerNo.SetFocus
Exit Sub
End If

If txtAccountNo.Text = "" Then
MsgBox "Please Enter the Account No.", vbInformation
txtAccountNo.SetFocus
Exit Sub
End If

If txtNarration.Text = "" Then
MsgBox "Please Enter the Narration.", vbInformation
txtNarration.SetFocus
Exit Sub
End If

 If txtAmountDeposited.Text = "" Then
MsgBox "Please Enter the Amount to Deposit.", vbInformation
txtAmountDeposited.SetFocus
Exit Sub
End If

If txtMode.Text = "" Then
MsgBox "Please select Transaction Mode.", vbInformation
txtMode.SetFocus
Exit Sub
End If

If txtCheckNo.Text = "" Then
MsgBox "Please Enter the Check No.", vbInformation
txtCheckNo.SetFocus
Exit Sub
End If


With rsDeposit
If NewRecord = True Then .AddNew
!TransactionID = txtTransactionID.Text
!CustomerID = cboCustomerNo.Text
!AccountNo = txtAccountNo.Text
!Narration = txtNarration.Text
!AmountDeposited = txtAmountDeposited.Text
!Mode = txtMode.Text
!CheckNO = txtCheckNo.Text
!Dated = txtDated.Value
.Update
End With

currBalance = (CCur(lblBalance.Caption) + CCur(txtAmountDeposited.Text))

With rsTransactions
.AddNew
!CustomerID = cboCustomerNo.Text
!AccountNo = txtAccountNo.Text
!Narration = txtNarration.Text
!CheckNO = txtCheckNo.Text
!Dated = txtDated.Value
!Debit = txtAmountDeposited.Text
!Mode = txtMode.Text
!Credit = "00"
!Balance = currBalance
.Update
End With

Set rsTemp = New ADODB.Recordset
rsTemp.Open "Select * FROM tblBalances WHERE CustomerID='" & cboCustomerNo.Text & "'", cnBank, adOpenKeyset, adLockOptimistic
With rsTemp
!Balance = currBalance
.Update
.Requery
.Close
End With

'rsBalances.Open "UPDATE tblBalances SET Balance ='" & currBalance & "' WHERE CustomerID='" & cboCustomerNo.Text & "'", cnBank, adOpenKeyset, adLockOptimistic

End Sub

Private Sub Form_Load()
Call connectDatabase

With rsCustomers
For X = 1 To .RecordCount
cboCustomerNo.AddItem !CustomerID
.MoveNext
Next X
End With
txtDated.Value = Date
End Sub
Public Sub GenerateNewTransactCode()
    Dim lastnumber As Long, newnumber As Long
    'Check if there are records in the file
    With rsDeposit
    If .BOF = True And .EOF = True Then
        lastnumber = 1000
    Else
        .MoveLast
        lastnumber = !TransactionID
    End If
    'Generate New Number
    newnumber = lastnumber + 1
    txtTransactionID.Text = newnumber
    End With
End Sub

Private Sub optCash_Click()
txtCheckNo.Text = "N/A"
txtCheckNo.Locked = True
txtMode.Text = "CASH"

End Sub

Private Sub optCheque_Click()
txtMode.Text = "CHEQUE"
txtCheckNo.Text = ""
txtCheckNo.Locked = False
txtCheckNo.SetFocus
End Sub

Private Sub optOthers_Click()
txtOther.Text = ""
txtMode.Text = ""
txtCheckNo.Text = "N/A"
txtCheckNo.Locked = True
txtOther.SetFocus
End Sub

Private Sub txtAccountNo_Change()
txtNarration.SetFocus
End Sub


Private Sub txtAmountDeposited_KeyPress(KeyAscii As Integer)
Call ValidNumeric(KeyAscii)
If KeyAscii = 13 Then optCash.SetFocus
End Sub

Private Sub txtCheckNo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdSave.SetFocus
End Sub

Private Sub txtMode_LostFocus()
If txtMode.Text = "" Then
MsgBox "Select the Mode of Transaction", vbInformation
End If
End Sub

Private Sub txtNarration_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
If KeyAscii = 13 Then txtAmountDeposited.SetFocus
End Sub

Private Sub txtOther_LostFocus()
txtMode.Text = txtOther.Text
End Sub








VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmWithdrawal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Withdrawals"
   ClientHeight    =   3285
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6450
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3285
   ScaleWidth      =   6450
   Begin MSComCtl2.DTPicker txtDated 
      Height          =   375
      Left            =   4200
      TabIndex        =   17
      Top             =   120
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      _Version        =   393216
      Format          =   22675457
      CurrentDate     =   38293
   End
   Begin VB.CommandButton cmdWithdraw 
      Caption         =   "&Withdraw"
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2520
      TabIndex        =   8
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   375
      Left            =   3840
      TabIndex        =   7
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "&Quit"
      Height          =   375
      Left            =   5160
      TabIndex        =   6
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   375
      Left            =   1440
      TabIndex        =   5
      Top             =   2760
      Width           =   975
   End
   Begin VB.ComboBox cboCustomerNo 
      Height          =   315
      Left            =   1710
      TabIndex        =   4
      Top             =   510
      Width           =   2055
   End
   Begin VB.TextBox txtAmountWithdrawn 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1695
      TabIndex        =   3
      Top             =   2235
      Width           =   1920
   End
   Begin VB.TextBox txtNarration 
      Appearance      =   0  'Flat
      Height          =   765
      Left            =   1695
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   1260
      Width           =   3375
   End
   Begin VB.TextBox txtAccountNo 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1695
      TabIndex        =   1
      Top             =   885
      Width           =   2295
   End
   Begin VB.TextBox txtTransactionID 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1695
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Dated:"
      Height          =   195
      Index           =   7
      Left            =   3705
      TabIndex        =   16
      Top             =   180
      Width           =   480
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "AmountWithdrawn:"
      Height          =   195
      Index           =   4
      Left            =   315
      TabIndex        =   15
      Top             =   2280
      Width           =   1350
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Narration:"
      Height          =   195
      Index           =   3
      Left            =   975
      TabIndex        =   14
      Top             =   1305
      Width           =   690
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "AccountNo:"
      Height          =   195
      Index           =   2
      Left            =   810
      TabIndex        =   13
      Top             =   930
      Width           =   855
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "CustomerNo:"
      Height          =   195
      Index           =   1
      Left            =   750
      TabIndex        =   12
      Top             =   540
      Width           =   915
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "TransactionID:"
      Height          =   195
      Index           =   0
      Left            =   615
      TabIndex        =   11
      Top             =   165
      Width           =   1050
   End
   Begin VB.Label lblBalance 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4440
      TabIndex        =   10
      Top             =   720
      Width           =   1815
   End
End
Attribute VB_Name = "frmWithdrawal"
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

 If txtAmountWithdrawn.Text = "" Then
MsgBox "Please Enter the Amount to Deposit.", vbInformation
txtAmountWithdrawn.SetFocus
Exit Sub
End If



With rsWithdrawal
If NewRecord = True Then .AddNew
!TransactionID = txtTransactionID.Text
!CustomerID = cboCustomerNo.Text
!AccountNo = txtAccountNo.Text
!Narration = txtNarration.Text
!AmountWithdrawn = txtAmountWithdrawn.Text


!Dated = txtDated.Value
.Update
End With

currBalance = (CCur(lblBalance.Caption) - CCur(txtAmountWithdrawn.Text))

With rsTransactions
.AddNew
!CustomerID = cboCustomerNo.Text
!AccountNo = txtAccountNo.Text
!Narration = txtNarration.Text
!CheckNO = "N/A"
!Dated = txtDated.Value
!Debit = "00"
!Mode = "N/A"
!Credit = txtAmountWithdrawn.Text
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

Private Sub cmdWithdraw_Click()
NewRecord = True
Call clear_Form_Controls(Me)
Call GenerateNewTransactCode
cboCustomerNo.SetFocus
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
    With rsWithdrawal
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

Private Sub txtAccountNo_Change()
txtNarration.SetFocus
End Sub


Private Sub txtAmountwithdrawn_KeyPress(KeyAscii As Integer)
Call ValidNumeric(KeyAscii)

End Sub

Private Sub txtNarration_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
If KeyAscii = 13 Then txtAmountWithdrawn.SetFocus
End Sub











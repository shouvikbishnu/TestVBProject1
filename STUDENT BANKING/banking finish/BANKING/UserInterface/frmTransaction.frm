VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmTransaction 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edit Transaction"
   ClientHeight    =   4635
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin MSComCtl2.DTPicker txtDated 
      Height          =   375
      Left            =   1080
      TabIndex        =   21
      Top             =   2280
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      Format          =   22675457
      CurrentDate     =   38293
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   2280
      TabIndex        =   18
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   375
      Left            =   960
      TabIndex        =   17
      Top             =   4200
      Width           =   1215
   End
   Begin VB.TextBox txtMode 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1065
      TabIndex        =   16
      Top             =   3810
      Width           =   2055
   End
   Begin VB.TextBox txtBalance 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1065
      TabIndex        =   14
      Top             =   3420
      Width           =   1695
   End
   Begin VB.TextBox txtCredit 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1065
      TabIndex        =   12
      Top             =   3045
      Width           =   1695
   End
   Begin VB.TextBox txtDebit 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1065
      TabIndex        =   10
      Top             =   2670
      Width           =   1695
   End
   Begin VB.TextBox txtCheckNo 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1065
      TabIndex        =   7
      Top             =   1905
      Width           =   3375
   End
   Begin VB.TextBox txtNarration 
      Appearance      =   0  'Flat
      Height          =   645
      Left            =   1080
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   1170
      Width           =   3375
   End
   Begin VB.TextBox txtAccountNo 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1065
      TabIndex        =   3
      Top             =   780
      Width           =   3375
   End
   Begin VB.TextBox txtCustomerID 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1065
      TabIndex        =   1
      Top             =   405
      Width           =   3375
   End
   Begin VB.TextBox txtCode 
      Appearance      =   0  'Flat
      Height          =   225
      Left            =   3840
      TabIndex        =   19
      Top             =   450
      Width           =   495
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Code:"
      Height          =   195
      Index           =   9
      Left            =   3330
      TabIndex        =   20
      Top             =   480
      Width           =   420
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Mode:"
      Height          =   195
      Index           =   8
      Left            =   585
      TabIndex        =   15
      Top             =   3855
      Width           =   450
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Balance:"
      Height          =   195
      Index           =   7
      Left            =   405
      TabIndex        =   13
      Top             =   3465
      Width           =   630
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Credit:"
      Height          =   195
      Index           =   6
      Left            =   585
      TabIndex        =   11
      Top             =   3090
      Width           =   450
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Debit:"
      Height          =   195
      Index           =   5
      Left            =   615
      TabIndex        =   9
      Top             =   2715
      Width           =   420
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Dated:"
      Height          =   195
      Index           =   4
      Left            =   555
      TabIndex        =   8
      Top             =   2325
      Width           =   480
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "CheckNo:"
      Height          =   195
      Index           =   3
      Left            =   315
      TabIndex        =   6
      Top             =   1950
      Width           =   720
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Narration:"
      Height          =   195
      Index           =   2
      Left            =   345
      TabIndex        =   4
      Top             =   1215
      Width           =   690
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "AccountNo:"
      Height          =   195
      Index           =   1
      Left            =   180
      TabIndex        =   2
      Top             =   825
      Width           =   855
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "CustomerID:"
      Height          =   195
      Index           =   0
      Left            =   165
      TabIndex        =   0
      Top             =   450
      Width           =   870
   End
End
Attribute VB_Name = "frmTransaction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()
Set rsTemp = New ADODB.Recordset
rsTemp.Open "Select * FROM tblTransactions WHERE Code=" & txtCode.Text & " ", cnBank, adOpenKeyset, adLockOptimistic
With rsTemp
!CustomerID = txtCustomerID.Text
!AccountNo = txtAccountNo.Text
!Narration = txtNarration.Text
!CheckNO = txtCheckNo.Text
!Dated = txtDated.Value
!Debit = txtDebit.Text
!Mode = txtMode.Text
!Credit = txtCredit.Text
!Balance = txtBalance.Text
.Update
End With
rsTemp.Close
rsTransactions.Requery
frmTransactions.lvwTransactions.Refresh
Unload Me
End Sub

Private Sub Form_Load()
Call connectDatabase
End Sub

Public Sub DisplayTransact(myRs As Recordset)

With myRs
    cboCustomerNo.Text = !CustomerID
    txtAccountNo.Text = !AccountNo
    txtNarration.Text = !Narration
    txtCheckNo.Text = !CheckNO
    txtDated.Value = !Dated
    txtAmountDeposited.Text = !Debit
    txtMode.Text = !Mode
    txtCredit.Text = !Credit
    txtBalance.Text = !Balance
End With

End Sub


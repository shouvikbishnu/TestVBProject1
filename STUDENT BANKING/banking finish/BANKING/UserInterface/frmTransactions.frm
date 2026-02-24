VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmTransactions 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transactions"
   ClientHeight    =   6435
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10410
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6435
   ScaleWidth      =   10410
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPrintAll 
      Caption         =   "Print All"
      Height          =   255
      Left            =   3720
      TabIndex        =   21
      Top             =   6120
      Width           =   1815
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print Statement"
      Height          =   255
      Left            =   1800
      TabIndex        =   20
      Top             =   6120
      Width           =   1815
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit"
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   6120
      Width           =   1575
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "&Quit"
      Height          =   375
      Left            =   9480
      TabIndex        =   18
      Top             =   840
      Width           =   735
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Choose the View Mode"
      Height          =   615
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Width           =   5775
      Begin VB.OptionButton Option2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "View All"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2040
         TabIndex        =   17
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Custom"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   360
         TabIndex        =   16
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   10215
      Begin VB.ComboBox cboAccNo 
         Height          =   315
         Left            =   3960
         TabIndex        =   11
         Text            =   "Select..."
         Top             =   960
         Width           =   1815
      End
      Begin VB.ComboBox cboFirst 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1680
         TabIndex        =   10
         Text            =   "Select..."
         Top             =   960
         Width           =   2175
      End
      Begin VB.ComboBox cboCustomerID 
         Height          =   315
         Left            =   120
         TabIndex        =   9
         Text            =   "Select..."
         Top             =   960
         Width           =   1455
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Dated"
         Height          =   1215
         Left            =   5880
         TabIndex        =   3
         Top             =   120
         Width           =   3375
         Begin VB.CommandButton cmdOk 
            Caption         =   "Proceed"
            Height          =   315
            Left            =   1320
            TabIndex        =   4
            Top             =   840
            Width           =   855
         End
         Begin MSComCtl2.DTPicker dtTo 
            Height          =   255
            Left            =   1800
            TabIndex        =   5
            Top             =   480
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   450
            _Version        =   393216
            Format          =   58523649
            CurrentDate     =   38311
         End
         Begin MSComCtl2.DTPicker dtFrom 
            Height          =   255
            Left            =   120
            TabIndex        =   6
            Top             =   480
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   450
            _Version        =   393216
            Format          =   58523649
            CurrentDate     =   38311
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "From:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   240
            Width           =   735
         End
         Begin VB.Label dtToj 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "To:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1800
            TabIndex        =   7
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "&Refresh"
         Height          =   375
         Left            =   9360
         TabIndex        =   2
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Customer ID"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "First Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1680
         TabIndex        =   13
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Account Number"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3960
         TabIndex        =   12
         Top             =   720
         Width           =   1815
      End
   End
   Begin MSComctlLib.ListView lvwTransactions 
      Height          =   4335
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   7646
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   14737632
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   10
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Customer ID"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Account No:"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Narration"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Dated"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Debit"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Credit"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Mode Of Payment"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Cheque No."
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Balance"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Code"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmTransactions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lstItem As ListItem
Private Sub cboAccNo_Click()
Set rsTemp = New ADODB.Recordset
rsTemp.Open "Select * from tblCustomers Where AccountNo='" & cboAccNo.Text & "'", cnBank, adOpenKeyset, adLockOptimistic
With rsTemp
If .RecordCount > 0 Then
cboAccNo = !AccountNo
cboCustomerID = !CustomerID
cboFirst = !FirstName
Else
MsgBox "Invalid customer ID/Name/Account NO. Please Try Again", vbInformation
Exit Sub
End If
.Close
End With


Set rsTemp = New ADODB.Recordset
rsTemp.Open "Select * from tblTransactions Where AccountNo='" & cboAccNo.Text & "'", cnBank, adOpenKeyset, adLockOptimistic
With rsTemp
If .RecordCount > 0 Then
lvwTransactions.ListItems.Clear
Call LoadListView(rsTemp)
'cboAccNo = !AccountNo
'cboCustomerID = !CustomerID
'cboFirst = !FirstName
Else
MsgBox "Invalid customer ID/Name/Account NO. Please Try Again", vbInformation
Exit Sub
End If
.Close
End With


End Sub

Private Sub cboCustomerID_Click()

Set rsTemp = New ADODB.Recordset
rsTemp.Open "Select * from tblCustomers Where customerID='" & cboCustomerID.Text & "'", cnBank, adOpenKeyset, adLockOptimistic
With rsTemp
If .RecordCount > 0 Then
cboAccNo = !AccountNo
cboCustomerID = !CustomerID
cboFirst = !FirstName
Else
MsgBox "Invalid customer ID/Name/Account NO. Please Try Again", vbInformation
Exit Sub
End If
.Close
End With

Set rsTemp = New ADODB.Recordset
rsTemp.Open "Select * from tblTransactions Where customerID='" & cboCustomerID.Text & "'", cnBank, adOpenKeyset, adLockOptimistic
With rsTemp
If .RecordCount > 0 Then
lvwTransactions.ListItems.Clear
Call LoadListView(rsTemp)
Else
MsgBox "No Transactions bearing this customer ID. Please Try Again", vbInformation
Exit Sub
End If
.Close
End With
End Sub

Private Sub cboFirst_Click()
Set rsTemp = New ADODB.Recordset
rsTemp.Open "Select * from tblCustomers Where FirstName='" & cboFirst.Text & "'", cnBank, adOpenKeyset, adLockOptimistic
With rsTemp
If .RecordCount > 0 Then
cboAccNo = !AccountNo
cboCustomerID = !CustomerID
cboFirst = !FirstName
Else
MsgBox "Invalid customer ID/Name/Account NO. Please Try Again", vbInformation
Exit Sub
End If
.Close
End With


Set rsTemp = New ADODB.Recordset
rsTemp.Open "Select * from tblTransactions Where CustomerID='" & cboCustomerID.Text & "'", cnBank, adOpenKeyset, adLockOptimistic
With rsTemp
If .RecordCount > 0 Then
lvwTransactions.ListItems.Clear
Call LoadListView(rsTemp)
Else
MsgBox "No Transactions bearing this customers' first name. Please Try Again", vbInformation
Exit Sub
End If
.Close
End With
End Sub

Private Sub cmdEdit_Click()
With rsTransactions
.MoveFirst
While Not .EOF

If lvwTransactions.SelectedItem.ListSubItems(9) = !Code Then
    frmTransaction.txtCustomerID.Text = !CustomerID
    frmTransaction.txtAccountNo.Text = !AccountNo
    frmTransaction.txtNarration.Text = !Narration
    frmTransaction.txtCheckNo.Text = !CheckNO
    frmTransaction.txtDated.Value = !Dated
    frmTransaction.txtDebit.Text = !Debit
    frmTransaction.txtMode.Text = !Mode
    frmTransaction.txtCredit.Text = !Credit
    frmTransaction.txtBalance.Text = !Balance
    frmTransaction.txtCode.Text = !Code
    .MoveLast
    .MoveNext
    Else
    .MoveNext
    End If
Wend
frmTransaction.Show
End With
End Sub

Private Sub cmdOk_Click()
Set rsTemp = New ADODB.Recordset
rsTemp.Open "Select * from tblTransactions Where Dated BETWEEN #" & dtFrom.Value & "# AND #" & dtTo.Value & "#", cnBank, adOpenKeyset, adLockOptimistic
With rsTemp
If .RecordCount > 0 Then
lvwTransactions.ListItems.Clear
Call LoadListView(rsTemp)
Else
MsgBox "No Transactions Were carried out between these Dates. Please Try Again", vbInformation
Exit Sub
End If
.Close
End With
End Sub

Private Sub cmdPrint_Click()
With deBank
If .rscmdStatement_Grouping.State = adStateOpen Then .rscmdStatement_Grouping.Close
   .cmdStatement_Grouping Val(lvwTransactions.SelectedItem.Text)
   rptStatement.Show
End With
''Dim strSql As String
'''strSql = "SELECT tblCustomers.FirstName,tblCustomers.LastName,tblCustomers.Address,tblCustomers.PostalCode,tblCustomers.Location,tblCustomers.OpeningBalance,tblCustomers.CustomerID,tblTransactions.AccountNo,tblTransactions.Debit,tblTransactions.Credit,tblTransactions.Dated,tblTransactions.Mode,tblTransactions.CheckNo,tblTransactions.Code from tblCustomers,tblTransactions where tblCustomers.CustomerID=tblTransactions.CustomerID AND tblTransactions.Code=4'" ''& lvwTransactions.SelectedItem.ListSubItems(9).Text & "'"
''strSql = " SELECT tblCustomers.AccountNo, tblCustomers.Address,"
''strSql = strSql & "tblCustomers.LastName, tblCustomers.FirstName,"
''strSql = strSql & "tblCustomers.CustomerID,"
''strSql = strSql & "tblCustomers.Location, tblCustomers.OpeningBalance, "
''strSql = strSql & "tblCustomers.PostalCode, tblTransactions.Code, "
''strSql = strSql & "tblTransactions.CheckNo, tblTransactions.Credit, "
''strSql = strSql & "tblTransactions.Debit, tblTransactions.Mode "
''
''strSql = strSql & "From tblCustomers, tblTransactions "
''strSql = strSql & "Where tblCustomers.CustomerID = tblTransactions.CustomerID"
''
''Set rsTemp = New ADODB.Recordset
''rsTemp.Open strSql, cnBank, adOpenKeyset, adLockOptimistic
''With deBank
'''.Commands("cmdStatement").Parameters = lvwTransactions.SelectedItem.Text
'''.cmdStatement_Grouping( ,lvwTransactions.SelectedItem.ListSubItems(9))=
''.Commands("cmdStatement_Grouping").Parameters("tblCustomersCode") = "1"
''End With

'With rsTemp
'If .RecordCount > 0 Then
'Set rptStatement.DataSource = Nothing
'Set rptStatement.DataSource = rsTemp
'rptStatement.Show
'Else
'MsgBox "few"
'End If
'End With

End Sub

Private Sub cmdPrintAll_Click()
rptStatement.Show
End Sub

Private Sub cmdQuit_Click()
Unload Me
End Sub


Private Sub DTPicker1_Click()

End Sub

Private Sub cmdRefresh_Click()
lvwTransactions.Refresh
End Sub


Private Sub Command1_Click()

End Sub

'Text1.Text = lvwTransactions.SelectedItem.Text


Private Sub Form_Load()

Call connectDatabase
Call LoadListView(rsTransactions)
With rsCustomers
.MoveFirst
For X = 1 To .RecordCount
    cboCustomerID.AddItem !CustomerID
    cboFirst.AddItem !FirstName
    cboAccNo.AddItem !AccountNo
    .MoveNext
Next X
End With
Frame1.Enabled = False
End Sub


Private Sub lvwTransactions_ColumnClick(ByVal ColumnHeader As _
    MSComctlLib.ColumnHeader)
    ' Sort according to data in this column.
    If lvwTransactions.Sorted And _
        ColumnHeader.Index - 1 = lvwTransactions.SortKey Then
        ' Already sorted on this column, just invert the sort order.
        lvwTransactions.SortOrder = 1 - lvwTransactions.SortOrder
    Else
        lvwTransactions.SortOrder = lvwAscending
        lvwTransactions.SortKey = ColumnHeader.Index - 1
    End If
    lvwTransactions.Sorted = True
    
End Sub

Public Sub LoadListView(myRs As Recordset)
With myRs
While Not .EOF
'lvwTransactions.ListItems.Add , , !CustomerID & " " & !AccountNo ' & " " & !Narration & " " & !Dated & " " & !Debit & " " & !Credit
Set lstItem = lvwTransactions.ListItems.Add(, , !CustomerID)
lstItem.SubItems(1) = !AccountNo
lstItem.SubItems(2) = !Narration
lstItem.SubItems(3) = !Dated
lstItem.SubItems(4) = Format(!Debit, "#,###,##00.00")
lstItem.SubItems(5) = Format(!Credit, "#,###,##00.00")
lstItem.SubItems(6) = !Mode
lstItem.SubItems(7) = !CheckNO
lstItem.SubItems(8) = !Balance
lstItem.SubItems(9) = !Code
.MoveNext
Wend
End With
End Sub

Private Sub lvwTransactions_DblClick()
Call cmdEdit_Click
End Sub

Private Sub Option1_Click()
Frame1.Enabled = True
End Sub

Private Sub Option2_Click()
Frame1.Enabled = False
Set rsTemp = New ADODB.Recordset
rsTemp.Open "Select * from tblTransactions", cnBank, adOpenKeyset, adLockOptimistic
With rsTemp
If .RecordCount > 0 Then
lvwTransactions.ListItems.Clear
Call LoadListView(rsTemp)
Else
MsgBox "Database Empty..", vbInformation
Exit Sub
End If
.Close
End With
End Sub


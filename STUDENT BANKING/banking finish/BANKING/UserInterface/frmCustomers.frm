VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCustomers 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Customers"
   ClientHeight    =   5550
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7155
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5550
   ScaleWidth      =   7155
   Begin VB.Frame Frame5 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      TabIndex        =   33
      Top             =   3840
      Width           =   7095
      Begin VB.TextBox txtEmail 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   285
         Left            =   4080
         TabIndex        =   19
         Top             =   360
         Width           =   2895
      End
      Begin MSMask.MaskEdBox txtMobileNo 
         Height          =   285
         Left            =   2280
         TabIndex        =   18
         Top             =   360
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         ForeColor       =   64
         MaxLength       =   15
         Mask            =   "(####)-(######)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtPhoneNo 
         Height          =   285
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   503
         _Version        =   393216
         ClipMode        =   1
         Appearance      =   0
         MaxLength       =   17
         Mask            =   "(###)-###########"
         PromptChar      =   " "
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Email Address:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Index           =   11
         Left            =   4080
         TabIndex        =   36
         Top             =   120
         Width           =   1245
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MobileNo:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Index           =   10
         Left            =   2280
         TabIndex        =   35
         Top             =   120
         Width           =   810
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PhoneNo(Trunk):"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Index           =   9
         Left            =   120
         TabIndex        =   34
         Top             =   120
         Width           =   1395
      End
   End
   Begin VB.Frame Frame4 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      TabIndex        =   31
      Top             =   2040
      Width           =   7095
      Begin VB.ComboBox cboAccType 
         ForeColor       =   &H00000040&
         Height          =   315
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   2295
      End
      Begin VB.TextBox txtOpeningBal 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000040&
         Height          =   285
         Left            =   4560
         TabIndex        =   13
         Top             =   360
         Width           =   1695
      End
      Begin VB.TextBox txtAccountNo 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   285
         Left            =   2520
         TabIndex        =   12
         Top             =   360
         Width           =   1800
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Opening Balance"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Index           =   14
         Left            =   4560
         TabIndex        =   42
         Top             =   120
         Width           =   1365
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Account Type:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Index           =   13
         Left            =   120
         TabIndex        =   40
         Top             =   120
         Width           =   1170
      End
      Begin VB.Label lblMin 
         Caption         =   "0"
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Account No:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Index           =   5
         Left            =   2520
         TabIndex        =   32
         Top             =   120
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   0
      TabIndex        =   4
      Top             =   4560
      Width           =   7095
      Begin VB.CommandButton cmdLast 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   6600
         Picture         =   "frmCustomers.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   600
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdNext 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   6240
         Picture         =   "frmCustomers.frx":0342
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   600
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdPrevious 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   480
         Picture         =   "frmCustomers.frx":0684
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   600
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdFirst 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   120
         Picture         =   "frmCustomers.frx":09C6
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   600
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add New"
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   375
         Left            =   1320
         TabIndex        =   25
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   2400
         TabIndex        =   24
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   375
         Left            =   3600
         TabIndex        =   23
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   375
         Left            =   4800
         TabIndex        =   22
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmdQuit 
         Caption         =   "&Quit"
         Height          =   375
         Left            =   6000
         TabIndex        =   21
         Top             =   120
         Width           =   975
      End
      Begin VB.Label lblStatus 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   960
         TabIndex        =   47
         Top             =   600
         Width           =   5160
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   3000
      Width           =   7095
      Begin VB.TextBox txtAddress 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   285
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   1680
      End
      Begin VB.TextBox txtPostalCode 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   285
         Left            =   1920
         TabIndex        =   15
         Top             =   360
         Width           =   1560
      End
      Begin VB.TextBox txtLocation 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   285
         Left            =   4200
         TabIndex        =   16
         Top             =   360
         Width           =   2775
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Index           =   6
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   765
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PostalCode:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Index           =   7
         Left            =   1920
         TabIndex        =   2
         Top             =   120
         Width           =   990
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Location / Town:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Index           =   8
         Left            =   4200
         TabIndex        =   1
         Top             =   120
         Width           =   1350
      End
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   0
      TabIndex        =   26
      Top             =   480
      Width           =   7095
      Begin MSComCtl2.DTPicker txtDateJoined 
         Height          =   375
         Left            =   4440
         TabIndex        =   10
         Top             =   960
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         Format          =   19726337
         CurrentDate     =   38293
      End
      Begin VB.TextBox txtIDNO 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   285
         Left            =   2040
         TabIndex        =   9
         Top             =   960
         Width           =   2040
      End
      Begin VB.ComboBox cboContactTitle 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   345
         Left            =   120
         TabIndex        =   8
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox txtCustomerID 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   360
         Width           =   1440
      End
      Begin VB.TextBox txtFirstName 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   285
         Left            =   2040
         TabIndex        =   6
         Top             =   360
         Width           =   2295
      End
      Begin VB.TextBox txtLastName 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   285
         Left            =   4440
         TabIndex        =   7
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DateJoined:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Index           =   12
         Left            =   4440
         TabIndex        =   41
         Top             =   720
         Width           =   960
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ContactTitle:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Index           =   3
         Left            =   120
         TabIndex        =   39
         Top             =   720
         Width           =   1035
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "National ID NO:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Index           =   4
         Left            =   2040
         TabIndex        =   38
         Top             =   720
         Width           =   1155
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer ID:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Index           =   0
         Left            =   75
         TabIndex        =   29
         Top             =   120
         Width           =   1080
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "First Name:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Index           =   1
         Left            =   2025
         TabIndex        =   28
         Top             =   120
         Width           =   930
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Last Name:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Index           =   2
         Left            =   4455
         TabIndex        =   27
         Top             =   120
         Width           =   915
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   0
      TabIndex        =   30
      Top             =   0
      Width           =   7095
   End
End
Attribute VB_Name = "frmCustomers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cboAccType_Click()
Set rsTemp = New ADODB.Recordset
rsTemp.Open "Select * FROM tblAccTypes WHERE AccountName='" & cboAccType.Text & "'", cnBank, adOpenKeyset, adLockOptimistic
With rsTemp
If .RecordCount > 0 Then
    lblMin = !MinBalance
    Else
        Exit Sub
End If
End With
End Sub

Private Sub cboAccType_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtAccountNo.SetFocus
End Sub

Private Sub cboContactTitle_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
If KeyAscii = 13 Then txtIDNO.SetFocus
End Sub

Private Sub cmdAdd_Click()
NewRecord = True

cmdAdd.Enabled = False
cmdSave.Enabled = True
cmdDelete.Enabled = False
cmdCancel.Enabled = True
cmdEdit.Enabled = False
cmdQuit.Enabled = False
Call UnLock_Form_Controls(Me)
Call clear_Form_Controls(Me)
Call GenerateNewCustomerCode
txtCustomerID.Locked = True
txtFirstName.SetFocus
End Sub

Private Sub cmdCancel_Click()
cmdAdd.Enabled = True
cmdSave.Enabled = False
cmdDelete.Enabled = True
cmdCancel.Enabled = False
cmdEdit.Enabled = True
cmdQuit.Enabled = True
With rsCustomers
    If NewRecord = True Then
        .CancelUpdate
        NewRecord = False
    Else
        .CancelUpdate
    End If
    Call DisplayCustomers(rsCustomers)
End With
Call Lock_Form_Controls(Me)
End Sub

Private Sub cmdDelete_Click()
cmdAdd.Enabled = True
cmdSave.Enabled = False
cmdDelete.Enabled = True
cmdCancel.Enabled = False
cmdEdit.Enabled = False
cmdQuit.Enabled = True
With rsCustomers
If .BOF = True And .EOF = True Then
MsgBox "Nothing  to delete", vbInformation
cmdDelete.Enabled = False
Exit Sub
End If
.Delete
Call clear_Form_Controls(Me)
.MoveFirst
If .BOF = True Or .EOF = True Then
cmdDelete.Enabled = False
Exit Sub
End If
Call DisplayCustomers(rsCustomers)
End With
End Sub

Private Sub cmdEdit_Click()
NewRecord = False
cmdAdd.Enabled = False
cmdSave.Enabled = True
cmdDelete.Enabled = False
cmdCancel.Enabled = False
cmdEdit.Enabled = False
cmdQuit.Enabled = False
Call UnLock_Form_Controls(Me)
End Sub

Private Sub cmdFirst_Click()
Call MoveToFirst(rsCustomers)
Call DisplayCustomers(rsCustomers)
lblStatus.Caption = CStr("Record :" & rsCustomers.AbsolutePosition & " of " & rsCustomers.RecordCount)
End Sub

Private Sub cmdLast_Click()
Call MoveToLast(rsCustomers)
Call DisplayCustomers(rsCustomers)
lblStatus.Caption = CStr("Record :" & rsCustomers.AbsolutePosition & " of " & rsCustomers.RecordCount)
End Sub

Private Sub cmdNext_Click()
Call MoveToNext(rsCustomers)
Call DisplayCustomers(rsCustomers)
lblStatus.Caption = CStr("Record :" & rsCustomers.AbsolutePosition & " of " & rsCustomers.RecordCount)

End Sub

Private Sub cmdPrevious_Click()
Call MoveToPrev(rsCustomers)
Call DisplayCustomers(rsCustomers)
lblStatus.Caption = CStr("Record :" & rsCustomers.AbsolutePosition & " of " & rsCustomers.RecordCount)
End Sub

Private Sub cmdQuit_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()


        If txtFirstName.Text = "" Then
        Call Messager
        txtFirstName.SetFocus
        Exit Sub
        End If
        
        If txtLastName.Text = "" Then
        Call Messager
        txtLastName.SetFocus
        Exit Sub
        End If
        
        If cboContactTitle.Text = "" Then
        Call Messager
        cboContactTitle.SetFocus
        Exit Sub
        End If
        
        If txtIDNO.Text = "" Then
        Call Messager
        txtIDNO.SetFocus
        Exit Sub
        End If
        
         If txtAccountNo.Text = "" Then
         Call Messager
         txtAccountNo.SetFocus
        Exit Sub
        End If
        
        If txtAddress.Text = "" Then
        Call Messager
         txtAddress.SetFocus
        Exit Sub
        End If
        
        If txtPostalCode.Text = "" Then
        Call Messager
        txtPostalCode.SetFocus
        Exit Sub
        End If
                 
        If txtLocation.Text = "" Then
        Call Messager
        txtLocation.SetFocus
        Exit Sub
        End If
         
        If txtPhoneNo.Text = "" Then
        Call Messager
        txtPhoneNo.SetFocus
        Exit Sub
        End If
                 
        If txtMobileNo.Text = "" Then
        Call Messager
         txtMobileNo.SetFocus
        Exit Sub
        End If
        
        If txtEmail.Text = "" Then
        txtEmail.Text = "N/A"
        txtEmail.SetFocus
        Exit Sub
        End If
        
        If txtOpeningBal.Text = "" Then
        Call Messager
        txtOpeningBal.SetFocus
        Exit Sub
        End If
        
        If cboAccType.Text = "" Then
        Call Messager
        cboAccType.SetFocus
        Exit Sub
        End If
        
        
If CCur(txtOpeningBal.Text) < CCur(lblMin.Caption) Then
MsgBox "Opening balance should be atleast " & lblMin.Caption & " for this type of Account", vbInformation
Call selectTextControl(txtOpeningBal)
Exit Sub
End If

With rsCustomers
    If NewRecord = True Then .AddNew
        !CustomerID = txtCustomerID.Text
        !FirstName = txtFirstName.Text
        !LastName = txtLastName.Text
        !ContactTitle = cboContactTitle.Text
        !IDNO = txtIDNO.Text
        !AccountNo = txtAccountNo.Text
        !Address = txtAddress.Text
        !PostalCode = txtPostalCode.Text
        !Location = txtLocation.Text
        !PhoneNo = txtPhoneNo.Text
        !MobileNo = txtMobileNo.Text
        !Email = txtEmail.Text
        !DateJoined = txtDateJoined.Value
        !OpeningBalance = txtOpeningBal.Text
        !AccountType = cboAccType.Text
    .Update
    .Requery
End With

With rsBalances
    If NewRecord = True Then .AddNew
        !CustomerID = txtCustomerID.Text
        !AccountNo = txtAccountNo.Text
        !Balance = txtOpeningBal.Text
    .Update
End With


cmdAdd.Enabled = True
cmdSave.Enabled = False
cmdDelete.Enabled = True
cmdCancel.Enabled = False
cmdEdit.Enabled = True
cmdQuit.Enabled = True
Call Lock_Form_Controls(Me)
End Sub



Private Sub Form_Load()
'Customer Details
Call connectDatabase
cmdAdd.Enabled = True
cmdSave.Enabled = False
cmdDelete.Enabled = True
cmdCancel.Enabled = False
cmdEdit.Enabled = True
cmdQuit.Enabled = True

With cboContactTitle
.AddItem "MR."
.AddItem "MRS."
.AddItem "MISS."
.AddItem "DR."
.AddItem "PROFF."
.AddItem "SIR."
.AddItem "REV."
.AddItem "FR."
End With

With rsAccTypes
For X = 1 To .RecordCount
cboAccType.AddItem !AccountName
.MoveNext
Next X
End With

Call DisplayCustomers(rsCustomers)
Call Lock_Form_Controls(Me)
lblStatus.Caption = CStr("Record :" & rsCustomers.AbsolutePosition & " of " & rsCustomers.RecordCount)
End Sub
Public Sub GenerateNewCustomerCode()
    Dim lastnumber As Long, newnumber As Long
    'Check if there are records in the file
    With rsCustomers
    If .BOF = True And .EOF = True Then
        lastnumber = 2004000
    Else
        .MoveLast
        lastnumber = !CustomerID
    End If
    'Generate New Number
    newnumber = lastnumber + 1
    txtCustomerID.Text = newnumber
    End With
End Sub

Public Sub DisplayCustomers(myRs As Recordset)

With myRs
If .BOF = True And .EOF = True Then Exit Sub
On Error Resume Next
txtCustomerID.Text = !CustomerID
txtFirstName.Text = !FirstName
txtLastName.Text = !LastName
cboContactTitle.Text = !ContactTitle
txtIDNO.Text = !IDNO
txtAccountNo.Text = !AccountNo
txtAddress.Text = !Address
txtPostalCode.Text = !PostalCode
txtLocation.Text = !Location
txtPhoneNo.Text = !PhoneNo
txtMobileNo.Text = !MobileNo
txtEmail.Text = !Email
'txtDateJoined.Mask = ""
txtDateJoined.Value = !DateJoined
txtOpeningBal.Text = !OpeningBalance
cboAccType.Text = !AccountType
End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
Call disconnectDatabase
End Sub

Private Sub txtAccountNo_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
If KeyAscii = 13 Then txtOpeningBal.SetFocus
End Sub

Private Sub txtAddress_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtPostalCode.SetFocus
End Sub

Private Sub txtDateJoined_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtAddress.SetFocus
End Sub

Private Sub txtEmail_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(LCase$(Chr$(KeyAscii)))
End Sub

Private Sub txtFirstName_KeyPress(KeyAscii As Integer)

KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
Select Case KeyAscii
 Case Asc(" ")
 Case 65 To 90
 Case 97 To 122
 Case 32
 Case 13
 Case 8
 Case 127
 Case Else
  MsgBox "Invalid Input", vbOKOnly + vbExclamation
  KeyAscii = 0
End Select

If KeyAscii = 13 Then txtLastName.SetFocus
End Sub

Private Sub txtIDNO_Change()
If Len((txtIDNO.Text)) > 8 Then
MsgBox "Limited to 8 Numerics", vbInformation
Call selectTextControl(txtIDNO)
Exit Sub
End If
End Sub

Private Sub txtIDNO_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cboAccType.SetFocus
End Sub

Private Sub txtLastName_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
Select Case KeyAscii
 Case Asc(" ")
 Case 65 To 90
 Case 97 To 122
 Case 32
 Case 13
 Case 8
 Case 127
 Case Else
  MsgBox "Invalid Input", vbOKOnly + vbExclamation
  KeyAscii = 0
End Select

If KeyAscii = 13 Then cboContactTitle.SetFocus
End Sub

Private Sub txtLocation_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
Select Case KeyAscii
 Case Asc(" ")
 Case 65 To 90
 Case 97 To 122
 Case 32
 Case 13
 Case 8
 Case 127
 Case Else
  MsgBox "Invalid Input", vbOKOnly + vbExclamation
  KeyAscii = 0
End Select
If KeyAscii = 13 Then txtPhoneNo.SetFocus
End Sub

Private Sub txtMobileNo_GotFocus()
Call selectMaskControl(txtMobileNo)
End Sub

Private Sub txtMobileNo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtEmail.SetFocus
End Sub

Private Sub txtOpeningBal_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtAddress.SetFocus
End Sub

Private Sub txtPhoneNo_GotFocus()
Call selectMaskControl(txtPhoneNo)
End Sub

Private Sub txtPhoneNo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtMobileNo.SetFocus
End Sub

Private Sub txtPostalCode_Change()
If Len((txtPostalCode.Text)) > 5 Then
MsgBox "Limited to 5 Numerics", vbInformation
Call selectTextControl(txtPostalCode)
Exit Sub
End If
End Sub

Private Sub txtPostalCode_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtLocation.SetFocus
End Sub

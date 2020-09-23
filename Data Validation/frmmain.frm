VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmmain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Data Validation Program"
   ClientHeight    =   9555
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7860
   Icon            =   "frmmain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9555
   ScaleWidth      =   7860
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog dlgcommon 
      Left            =   7200
      Top             =   8400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   550
      Left            =   840
      TabIndex        =   29
      Top             =   8280
      Width           =   1815
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear"
      Enabled         =   0   'False
      Height          =   550
      Left            =   5280
      TabIndex        =   5
      Top             =   8280
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      Caption         =   "Enter Your Details"
      Height          =   7695
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   7455
      Begin VB.CommandButton cmdCheck 
         Caption         =   "&Check the validity of the card for its Printout"
         Enabled         =   0   'False
         Height          =   735
         Left            =   5520
         TabIndex        =   26
         Top             =   6600
         Width           =   1695
      End
      Begin VB.TextBox txtCreditCardNo 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2160
         TabIndex        =   25
         Text            =   "Optional"
         Top             =   6840
         Width           =   3255
      End
      Begin VB.TextBox txtEmailId 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1080
         TabIndex        =   23
         Top             =   6000
         Width           =   3855
      End
      Begin VB.ComboBox cboBlood 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmmain.frx":628A
         Left            =   1560
         List            =   "frmmain.frx":62AC
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   5160
         Width           =   1335
      End
      Begin VB.ComboBox cboYYYY 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmmain.frx":62E0
         Left            =   3480
         List            =   "frmmain.frx":6422
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   4440
         Width           =   735
      End
      Begin VB.ComboBox cboMM 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmmain.frx":66A2
         Left            =   2640
         List            =   "frmmain.frx":66CA
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   4440
         Width           =   615
      End
      Begin VB.ComboBox cboDD 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmmain.frx":66F5
         Left            =   1800
         List            =   "frmmain.frx":6756
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   4440
         Width           =   615
      End
      Begin VB.TextBox txtTelMobile 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2160
         TabIndex        =   11
         Text            =   "Optional"
         Top             =   3600
         Width           =   2055
      End
      Begin VB.TextBox txtTelResi 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2520
         TabIndex        =   9
         Top             =   2760
         Width           =   1575
      End
      Begin VB.TextBox txtAddress 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         TabIndex        =   7
         Top             =   2040
         Width           =   5055
      End
      Begin VB.TextBox txtRoll 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         TabIndex        =   4
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   1320
         TabIndex        =   2
         Top             =   480
         Width           =   2895
      End
      Begin VB.Label Label11 
         Height          =   255
         Left            =   4800
         TabIndex        =   32
         Top             =   480
         Width           =   735
      End
      Begin VB.Label lblCompanyOfNo 
         Alignment       =   2  'Center
         Height          =   255
         Left            =   4440
         TabIndex        =   30
         Top             =   2760
         Width           =   2415
      End
      Begin VB.Image picamerican 
         Height          =   975
         Left            =   5520
         Picture         =   "frmmain.frx":67CD
         Stretch         =   -1  'True
         Top             =   5520
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Image picmastercard 
         Height          =   975
         Left            =   5520
         Picture         =   "frmmain.frx":7383
         Stretch         =   -1  'True
         Top             =   5520
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Image picvisa 
         Height          =   975
         Left            =   5520
         Picture         =   "frmmain.frx":8709
         Stretch         =   -1  'True
         Top             =   5520
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label lblCheck 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2160
         TabIndex        =   28
         Top             =   7320
         Width           =   1095
      End
      Begin VB.Label lblType 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3720
         TabIndex        =   27
         Top             =   7320
         Width           =   1695
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Credit / Debit Card No :"
         Height          =   195
         Left            =   240
         TabIndex        =   24
         Top             =   6840
         Width           =   1665
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Email  :"
         Height          =   195
         Left            =   240
         TabIndex        =   22
         Top             =   6045
         Width           =   510
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Blood Group  :"
         Height          =   195
         Left            =   240
         TabIndex        =   20
         Top             =   5250
         Width           =   1020
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Height          =   255
         Left            =   4560
         TabIndex        =   19
         Top             =   4440
         Width           =   1575
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "(YYYY)"
         Height          =   195
         Left            =   3600
         TabIndex        =   18
         Top             =   4800
         Width           =   510
      End
      Begin VB.Label Label8 
         Caption         =   "(MM)"
         Height          =   255
         Left            =   2760
         TabIndex        =   16
         Top             =   4800
         Width           =   375
      End
      Begin VB.Label Label7 
         Caption         =   "(DD)"
         Height          =   255
         Left            =   1920
         TabIndex        =   14
         Top             =   4800
         Width           =   375
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Date Of Birth  :"
         Height          =   195
         Left            =   240
         TabIndex        =   12
         Top             =   4455
         Width           =   1050
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Telephone (Mobile)  :"
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   3660
         Width           =   1500
      End
      Begin VB.Label label4 
         AutoSize        =   -1  'True
         Caption         =   "Telephone (Residence/Co.)  :"
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   2865
         Width           =   2115
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Address  :"
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   2070
         Width           =   705
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Roll No.  :"
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   1275
         Width           =   705
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Name  :"
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   555
      End
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Programmer  :  Daljit Singh Kalsi"
      Height          =   195
      Left            =   5520
      TabIndex        =   31
      Top             =   9000
      Width           =   2235
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim inttemp2 As Integer                'Check for even & odd months ("integer" type value)
Dim sngtemp2 As Single                 'Check for even & odd months ("single" type value)
Dim inttemp4 As Integer                'Check for leap year ("integer" type Value)
Dim sngtemp4 As Single                 'Check for leap year ("single" type Value)
Dim Symbols, ch, SymbolsInName As String
Dim flagat As Boolean                  'Check for existence of '@' character in emailID textbox.
Dim flagdot As Boolean                 'Check for existence of '.' character in emailID textbox.
Dim flagalpha As Boolean               ' Flag for alphabetic character. Will be true if typed character is an alphabet.
Dim CC As String                       ' Trimmed string
Dim CheckSum As Integer                ' Running sum
Dim Dbl As Integer                     ' Doubling flag
Dim Idx As Integer                     ' String position
Dim Digit As Integer                   ' Extracted digit
Dim value As Integer

Private Sub cboBlood_Click()

If cboBlood.ListIndex <> -1 Then
    txtEmailId.Enabled = True
Else
    txtEmailId.Enabled = False
End If

End Sub

Private Sub cboDD_Click()
If cboDD.ListIndex <> -1 Then
    cboMM.Enabled = True
Else
    cboMM.Enabled = False
End If

If Val(cboYYYY.List(cboYYYY.ListIndex)) = 1937 And Val(cboMM.List(cboMM.ListIndex)) = 9 And Val(cboDD.List(cboDD.ListIndex)) > 7 Then
   MsgBox "Dates after 7th of september 1937 till Next month are not valid dates.", vbInformation, "Invalid Date Entered"
   Label13.Caption = ""
   cboDD.ListIndex = -1
   cboMM.ListIndex = -1
   cboYYYY.ListIndex = -1
   cboDD.SetFocus
End If

If Val(cboMM.List(cboMM.ListIndex)) = 2 And Val(cboDD.List(cboDD.ListIndex)) > 29 Then
    MsgBox "Invalid Date Entered. February cannot have " & cboDD.List(cboDD.ListIndex) & " days.", vbInformation, "Invalid Date"
    Label13.Caption = ""
    cboBlood.Enabled = False
    cboDD.ListIndex = -1
    cboMM.ListIndex = -1
    cboYYYY.ListIndex = -1
    cboDD.SetFocus
End If

If Val(cboYYYY.List(cboYYYY.ListIndex)) <> -1 Then
    If sngtemp4 <> inttemp4 And Val(cboMM.List(cboMM.ListIndex)) = 2 And Val(cboDD.List(cboDD.ListIndex)) > 28 Then
        MsgBox "Invalid Date Entered. Non Leap year cannot have " & cboDD.List(cboDD.ListIndex) & " days.", vbInformation, "Invalid Date"
        Label13.Caption = ""
        cboBlood.Enabled = False
        cboDD.ListIndex = -1
        cboMM.ListIndex = -1
        cboYYYY.ListIndex = -1
        cboDD.SetFocus
    End If
End If

If inttemp2 = sngtemp2 And Val(cboMM.List(cboMM.ListIndex)) < 8 And Val(cboMM.List(cboMM.ListIndex)) > 0 Then  'Means even months less than 8th month
    If Val(cboDD.List(cboDD.ListIndex)) = 31 Then
       MsgBox "Invalid Date Entered. This month cannot have " & cboDD.List(cboDD.ListIndex) & " days.", vbInformation, "Invalid Date"
       Label13.Caption = ""
       cboBlood.Enabled = False
       cboDD.ListIndex = -1
       cboMM.ListIndex = -1
       cboYYYY.ListIndex = -1
       cboDD.SetFocus
    End If
End If

If inttemp2 <> sngtemp2 And Val(cboMM.List(cboMM.ListIndex)) > 7 Then 'Means even months greater than 8th month
    If Val(cboDD.List(cboDD.ListIndex)) = 31 Then
       MsgBox "Invalid Date Entered. This month cannot have " & cboDD.List(cboDD.ListIndex) & " days.", vbInformation, "Invalid Date"
       Label13.Caption = ""
       cboBlood.Enabled = False
       cboDD.ListIndex = -1
       cboMM.ListIndex = -1
       cboYYYY.ListIndex = -1
       cboDD.SetFocus
    End If
End If

End Sub

Private Sub cboMM_Click()

If cboMM.ListIndex <> -1 Then
    cboYYYY.Enabled = True
Else
    cboYYYY.Enabled = False
End If

If Val(cboYYYY.List(cboYYYY.ListIndex)) = 1937 And Val(cboMM.List(cboMM.ListIndex)) = 9 And Val(cboDD.List(cboDD.ListIndex)) > 7 Then
   MsgBox "Dates after 7th of september 1937 till Next month are not valid dates.", vbInformation, "Invalid Date Entered"
   Label13.Caption = ""
   cboDD.ListIndex = -1
   cboMM.ListIndex = -1
   cboYYYY.ListIndex = -1
   cboDD.SetFocus
End If

If sngtemp4 <> inttemp4 And Val(cboMM.List(cboMM.ListIndex)) = 2 And Val(cboDD.List(cboDD.ListIndex)) > 28 Then
    MsgBox "Invalid Date Entered. February cannot have " & cboDD.List(cboDD.ListIndex) & " days.", vbInformation, "Invalid Date"
    Label13.Caption = ""
    cboBlood.Enabled = False
    cboDD.ListIndex = -1
    cboMM.ListIndex = -1
    cboYYYY.ListIndex = -1
    cboDD.SetFocus
End If

If Val(cboMM.List(cboMM.ListIndex)) = 2 And Val(cboDD.List(cboDD.ListIndex)) > 28 Then
    MsgBox "Invalid Date Entered. February cannot have " & cboDD.List(cboDD.ListIndex) & " days.", vbInformation, "Invalid Date"
    Label13.Caption = ""
    cboDD.ListIndex = -1
    cboMM.ListIndex = -1
    cboYYYY.ListIndex = -1
    cboDD.SetFocus
End If

sngtemp2 = Val(cboMM.List(cboMM.ListIndex)) / 2
inttemp2 = Val(cboMM.List(cboMM.ListIndex)) / 2

If inttemp2 = sngtemp2 And Val(cboMM.List(cboMM.ListIndex)) < 8 And Val(cboMM.List(cboMM.ListIndex)) > 0 Then  'Means even months less than 8th month
    If Val(cboDD.List(cboDD.ListIndex)) = 31 Then
       MsgBox "Invalid Date Entered. This month cannot have " & cboDD.List(cboDD.ListIndex) & " days.", vbInformation, "Invalid Date"
       Label13.Caption = ""
       cboBlood.Enabled = False
       cboDD.ListIndex = -1
       cboMM.ListIndex = -1
       cboYYYY.ListIndex = -1
       cboDD.SetFocus
    End If
End If

If inttemp2 <> sngtemp2 And Val(cboMM.List(cboMM.ListIndex)) > 7 Then 'Means even months greater than 8th month
    If Val(cboDD.List(cboDD.ListIndex)) = 31 Then
       MsgBox "Invalid Date Entered. This month cannot have " & cboDD.List(cboDD.ListIndex) & " days.", vbInformation, "Invalid Date"
       Label13.Caption = ""
       cboBlood.Enabled = False
       cboDD.ListIndex = -1
       cboMM.ListIndex = -1
       cboYYYY.ListIndex = -1
       cboDD.SetFocus
    End If
End If

End Sub


Private Sub cboYYYY_Click()

If cboYYYY.ListIndex <> -1 Then
    cboBlood.Enabled = True
Else
    cboBlood.Enabled = False
End If

If Val(cboYYYY.List(cboYYYY.ListIndex)) = 1937 And Val(cboMM.List(cboMM.ListIndex)) = 9 And Val(cboDD.List(cboDD.ListIndex)) > 7 Then
   MsgBox "Dates after 7th of september 1937 till Next month are not valid dates.", vbInformation, "Invalid Date Entered"
   Label13.Caption = ""
   cboDD.ListIndex = -1
   cboMM.ListIndex = -1
   cboYYYY.ListIndex = -1
   cboDD.SetFocus
End If

sngtemp4 = Val(cboYYYY.List(cboYYYY.ListIndex)) / 4
inttemp4 = Val(cboYYYY.List(cboYYYY.ListIndex)) / 4

If inttemp4 = sngtemp4 And Val(cboMM.List(cboMM.ListIndex)) = 2 And Val(cboDD.List(cboDD.ListIndex)) > 29 Then
        MsgBox "Invalid Date Entered. A Leap Year cannot have " & cboDD.List(cboDD.ListIndex) & " days.", vbInformation, "Invalid Date"
        cboBlood.Enabled = False
        cboDD.ListIndex = -1
        cboMM.ListIndex = -1
        cboYYYY.ListIndex = -1
        Label13.Caption = ""
        cboDD.SetFocus
ElseIf sngtemp4 = inttemp4 Then
    Label13.Caption = "Leap Year"
Else
    Label13.Caption = "Not a Leap Year"
    If sngtemp4 <> inttemp4 And Val(cboMM.List(cboMM.ListIndex)) = 2 And Val(cboDD.List(cboDD.ListIndex)) > 28 Then
        MsgBox "Invalid Date Entered. A Non Leap Year cannot have " & cboDD.List(cboDD.ListIndex) & " days.", vbInformation, "Invalid Date"
        Label13.Caption = ""
        cboBlood.Enabled = False
        cboDD.ListIndex = -1
        cboMM.ListIndex = -1
        cboYYYY.ListIndex = -1
        cboDD.SetFocus
     End If
End If

End Sub

Private Sub cmdCheck_Click()
CC$ = Trim$(txtCreditCardNo.Text)     ' Trim extra blanks
CheckSum = 0                           ' Start with 0 CheckSum
Dbl = 0                                ' Start with a non -doubling
For Idx = Len(CC$) To 1 Step -1        ' Working backwards
   Digit = Asc(Mid$(CC$, Idx, 1))      ' Isolate character
   If Digit > 47 Then                  ' Skip if not a digit
      If Digit < 58 Then
         Digit = Digit - 48            ' Remove ASCII bias
         If Dbl Then                   ' If in the "double-add" phase
            Digit = Digit + Digit      '   then double first
            If Digit > 9 Then
               Digit = Digit - 9       ' Cast nines
            End If
         End If
         Dbl = Not Dbl                 ' Flip doubling flag
         CheckSum = CheckSum + Digit   ' Add to running sum
         If CheckSum > 9 Then          ' Cast tens
            CheckSum = CheckSum - 10   ' (same as MOD 10 but faster)
         End If
      End If
   End If
Next

validcard = (CheckSum = 0)   ' Must sum to 0

If CheckSum = 0 And txtCreditCardNo.Text = "" Then
  lblCheck.Caption = ""
ElseIf CheckSum = 0 Then
    lblCheck.Caption = "Valid"
    If Left(txtCreditCardNo.Text, 1) = "4" Then
       lblType.Caption = "Visa Card"
       picvisa.Visible = True
       picamerican.Visible = False
       picmastercard.Visible = False
    ElseIf Left(txtCreditCardNo.Text, 2) = "37" Then
       lblType.Caption = "American Express Card"
       picvisa.Visible = False
       picamerican.Visible = True
       picmastercard.Visible = False
    ElseIf Left(txtCreditCardNo.Text, 1) = "5" Then
       lblType.Caption = "Master Card"
       picvisa.Visible = False
       picamerican.Visible = False
       picmastercard.Visible = True
    ElseIf Left(txtCreditCardNo.Text, 1) = "6" Then
       lblType.Caption = "Discover Card"
    End If
Else
    lblCheck.Caption = "InValid"
    picvisa.Visible = False
    picamerican.Visible = False
    picmastercard.Visible = False
    lblType.Caption = ""
End If

txtCreditCardNo.SetFocus
End Sub

Private Sub cmdClear_Click()
txtName.Text = ""
txtRoll.Text = ""
txtAddress.Text = ""
txtTelResi.Text = ""
txtTelMobile.Text = ""
Label13.Caption = ""
lblCompanyOfNo.Caption = ""
cboDD.ListIndex = -1
cboMM.ListIndex = -1
cboYYYY.ListIndex = -1
cboBlood.ListIndex = -1
txtEmailId.Text = ""
txtCreditCardNo.Text = ""
lblCheck.Caption = ""
lblType.Caption = ""

txtTelMobile.Text = "Optional"
txtCreditCardNo.Text = "Optional"

txtRoll.Enabled = False
txtAddress.Enabled = False
txtTelResi.Enabled = False
txtTelMobile.Enabled = False
cboDD.Enabled = False
cboMM.Enabled = False
cboYYYY.Enabled = False
cboBlood.Enabled = False
txtEmailId.Enabled = False
txtCreditCardNo.Enabled = False
cmdCheck.Enabled = False
cmdClear.Enabled = False
cmdPrint.Enabled = False

picvisa.Visible = False
picamerican.Visible = False
picmastercard.Visible = False

txtName.SetFocus

End Sub


Private Sub cmdPrint_Click()
dlgcommon.ShowPrinter
Printer.FontName = "Times New Roman"
Printer.CurrentX = 4500
Printer.CurrentY = 500
Printer.FontSize = 28
Printer.Print "Students Details"
Printer.Print "   ------------------------------------------------------------"
Printer.FontSize = 20
Printer.Print ""

Printer.CurrentX = 1000
Printer.Print " Name               :    "; txtName.Text
Printer.Print ""

Printer.CurrentX = 1000
Printer.Print " Roll No            :    "; txtRoll.Text
Printer.Print ""

Printer.CurrentX = 1000
Printer.Print " Address            :    "; txtAddress.Text
Printer.Print ""

Printer.CurrentX = 1000
Printer.Print " Tel (Resi)         :    "; txtTelResi.Text
Printer.Print ""

Printer.CurrentX = 1000
If txtTelMobile.Text <> "" Then
    Printer.Print " Tel (Mob)         :    "; txtTelMobile.Text
    Printer.Print ""
End If

Printer.CurrentX = 1000
Printer.Print " Date Of Birth    :    "; cboDD.List(cboDD.ListIndex); " / "; cboMM.List(cboMM.ListIndex); " / "; cboYYYY.List(cboYYYY.ListIndex)
Printer.Print ""

Printer.CurrentX = 1000
Printer.Print " Blood Group      :    "; cboBlood.List(cboBlood.ListIndex)
Printer.Print

Printer.CurrentX = 1000
Printer.Print " Email Id             :    "; txtEmailId.Text
Printer.Print ""

Printer.CurrentX = 1000
If lblCheck.Caption = "Valid" Then
    Printer.Print " Credit Card No  :    "; txtCreditCardNo.Text;
End If

Printer.EndDoc
End Sub


Private Sub txtAddress_Change()
txtAddress.MaxLength = 60
flagalpha = False
If txtAddress.Text <> "" And Len(txtAddress.Text) > 4 Then
    txtTelResi.Enabled = True
ElseIf Len(txtAddress.Text) < 5 Then
    txtTelResi.Enabled = False
End If
End Sub

Private Sub txtCreditCardNo_Change()
If IsNumeric(txtCreditCardNo.Text) Then
   'Enter Data
   cmdCheck.Enabled = True
   txtCreditCardNo.MaxLength = 20
Else
   txtCreditCardNo.Text = ""
   cmdCheck.Enabled = False
   lblCheck.Caption = ""
End If
End Sub

Private Sub txtCreditCardNo_GotFocus()
If txtCreditCardNo.Text = "Optional" Then
    txtCreditCardNo.Text = ""
End If
End Sub

Private Sub txtCreditCardNo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
        txtCreditCardNo.Enabled = False
        MsgBox "Right Clicking is disabled.", vbInformation, "Sorry No Right Click Here."
        txtCreditCardNo.Enabled = True
End If
End Sub

Private Sub txtEmailId_Change()
If txtEmailId.Text = "" Then
   txtCreditCardNo.Enabled = False
   cmdClear.Enabled = False
   cmdPrint.Enabled = False
End If
End Sub

Private Sub txtEmailId_KeyPress(KeyAscii As Integer)

Symbols = "`~' !@#$%^&*()_-+/.{}[]:;<>,.?|\0123456789"""
' We do not want these symbols as first character in the textbox


SymbolsInName = "`~' !#$%^&*()-+/{}[]:;<>,?|\"""
' We do not want these symbols as characters in the emailid

   ch = Chr$(KeyAscii)
   
   If InStr(SymbolsInName, ch) Then
         KeyAscii = 0
    End If
   
   If InStr(Symbols, ch) And txtEmailId.Text = "" Then
         KeyAscii = 0
   ElseIf InStr(txtEmailId.Text, "@") Then
         flagat = True
         
     Select Case ch
          Case """", " '", " ", "!", "@", "#", "$", "%", "^", "&", "*", "(", ")", "_", "+", "/", "{", "}", "[", "]", ":", ";", "<", ">", ",", ".", "?", "|", "\"
            If flagat = True Then
              If flagdot = False Then
               If ch = "." Then
                    txtEmailId.SelText = ".co"
                    flagdot = True
                    txtCreditCardNo.Enabled = True
                    cmdPrint.Enabled = True
                    cmdClear.Enabled = True
                  Else
                    KeyAscii = 0
               End If
              End If
            End If
      End Select
     
   ElseIf Not InStr(txtEmailId.Text, "@") Then
         flagat = False
   End If

End Sub

Private Sub txtEmailId_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
        txtEmailId.Enabled = False
        MsgBox "Right Clicking is disabled.", vbInformation, "Sorry No Right Click Here."
        txtEmailId.Enabled = True
End If
End Sub

Private Sub txtName_Change()
If txtName.Text <> "" Then
value = Asc(txtName.Text)

If value > 64 And value < 91 Or value > 96 And value < 123 Then

If txtName.Text <> "" Then
    txtName.MaxLength = 25
    If Len(txtName.Text) > 2 Then
        txtRoll.Enabled = True
    Else
        txtRoll.Enabled = False
    End If
End If

Else
 txtName.Text = ""
 End If
End If
End Sub


Private Sub txtName_KeyPress(KeyAscii As Integer)

'Initially flagalpha = false

If txtName.Text = "" Then
     flagalpha = False
End If

' With this code, when we clear the text box by using 'Backspace' then we need to do flagalpha to be = false
' If not, our first character will accept space if pressed.


If KeyAscii = 22 Then
    MsgBox "No Copy & Paste Allowed", vbInformation, "Pasting Not Permitted here."
    KeyAscii = 0
End If


If flagalpha = False Then           ' Means typed character is one of the symbol
    Symbols = " `~'!@#$%^&*()_-+/.{}[]:;<>,.?|\0123456789=""" ' We do not want these characters in our textbox as first character.
Else
    Symbols = "`~'!@#$%^&*()_-+/.{}[]:;<>,.?|\0123456789="""  ' We do not want these characters in our textbox
End If
    
    ch = Chr$(KeyAscii) 'Returns a String containing the character
                        'associated with the specified ascii code'Saves (OR Say returns) the character in "ch"
                        'by user's keystroke.
                        'Here only 1 character is stored at a time.
                        'If user enters more than 1 character i.e.
                        'presses the keys more than once (Which is obvious)
                        'then that new character is replaced by the older one.
                        
    
If InStr(Symbols, ch) Then      ' Means typed character is one of the symbol
       KeyAscii = 0
ElseIf Not InStr(Symbols, ch) Then  'Means a typed character is an alphabet
       flagalpha = True
End If
    ' "Instr" Returns the position of the first occurrence of one string
    ' within another. Here when the user presses any key, a keypress event
    ' occures and that character is saved in "ch". Thus it forms the first
    ' occurence of one of the string. Another string is "Symbols". Thus "ch"
    ' is compared with each of the character in "symbols". If the match is
    ' found, keyascii = 0 & we know that keyascii is the only parameter taken
    ' by this procedure (i.e. keyPress event). As the value = 0 which returns
    ' nothing (not even a blank space), this is reflected in the actual program
    ' while typing. Thus every keystroke is a first occurence with respect to
    ' the previous one.
   
If flagalpha = True And ch = " " Then
    KeyAscii = 32
    flagalpha = False
End If


End Sub

Private Sub txtName_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
        txtName.Enabled = False
        MsgBox "Right Clicking is disabled.", vbInformation, "Sorry No Right Click Here."
        txtName.Enabled = True
    End If
End Sub

Private Sub txtRoll_Change()
If IsNumeric(txtRoll.Text) Then
      'Enter Roll No.
      txtRoll.MaxLength = 3
      txtAddress.Enabled = True
   Else
      txtRoll.Text = ""
      txtAddress.Enabled = False
   End If

End Sub


Private Sub txtRoll_LostFocus()
If Val(txtRoll.Text) < 300 Or Val(txtRoll.Text) > 364 Then
    MsgBox "Roll Nos. should be in the range of 300 & 364.", vbInformation, "Invalid Roll No."
    txtRoll.Text = ""
    txtRoll.SetFocus
    txtAddress.Enabled = False
End If
End Sub

Private Sub txtRoll_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
        txtRoll.Enabled = False
        MsgBox "Right Clicking is disabled.", vbInformation, "Sorry No Right Click Here."
        txtRoll.Enabled = True
End If
End Sub

Private Sub txtTelMobile_Change()
If IsNumeric(txtTelMobile.Text) Then
   txtTelMobile.MaxLength = 10
   If Len(txtTelMobile.Text) = 10 And Val(txtTelMobile.Text) > 9000000000# Then
       cboDD.Enabled = True
   ElseIf txtTelMobile.Text <> "" And Val(txtTelMobile.Text) < 9000000000# Then
       cboDD.Enabled = False
       cmdPrint.Enabled = False
   End If
  
Else
    If txtTelMobile.Text = "" Then
       cboDD.Enabled = True
       If txtEmailId.Text <> "" Then
              cmdPrint.Enabled = True
       End If
    End If
    txtTelMobile.Text = ""
End If
End Sub

Private Sub txtTelMobile_GotFocus()
txtTelMobile.Text = ""
cboDD.Enabled = True
End Sub

Private Sub txtTelResi_Change()
If IsNumeric(txtTelResi.Text) Then
   'Enter Tel.
    txtTelResi.MaxLength = 8
    If Len(txtTelResi.Text) = 8 And Val(txtTelResi) > 20000000 Then
       txtTelMobile.Enabled = True
       If txtEmailId.Text <> "" Then
              cmdPrint.Enabled = True
       End If
       If Left(txtTelResi.Text, 1) = "2" Then
          lblCompanyOfNo.Caption = "MTNL Phone."
       ElseIf Left(txtTelResi.Text, 1) = "3" Then
          lblCompanyOfNo.Caption = "Reliance Phone."
       ElseIf Left(txtTelResi.Text, 1) = "5" Then
          lblCompanyOfNo.Caption = "Tata Phone (Old No)."
       ElseIf Left(txtTelResi.Text, 1) = "6" Then
          lblCompanyOfNo.Caption = "Tata Phone (New No)."
       Else
          lblCompanyOfNo.Caption = "Unknown Company."
          txtTelMobile.Enabled = False
       End If
       
    Else
       txtTelMobile.Enabled = False
       cmdPrint.Enabled = False
       lblCompanyOfNo.Caption = ""
    End If
Else
    txtTelMobile.Enabled = False
    cmdPrint.Enabled = False
    txtTelResi.Text = ""
    lblCompanyOfNo.Caption = ""
End If

End Sub

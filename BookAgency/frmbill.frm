VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bill"
   ClientHeight    =   6945
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7590
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6945
   ScaleWidth      =   7590
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   495
      Index           =   2
      Left            =   4440
      TabIndex        =   24
      Top             =   6120
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Print"
      Height          =   495
      Index           =   1
      Left            =   3120
      TabIndex        =   23
      Top             =   6120
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      Height          =   495
      Index           =   0
      Left            =   1800
      TabIndex        =   22
      Top             =   6120
      Width           =   1095
   End
   Begin VB.VScrollBar VScroll1 
      Enabled         =   0   'False
      Height          =   1935
      Left            =   6960
      Min             =   1
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   3600
      Value           =   1
      Width           =   255
   End
   Begin VB.TextBox Text9 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   5880
      Locked          =   -1  'True
      TabIndex        =   19
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   5640
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Height          =   2055
      Left            =   720
      TabIndex        =   6
      Top             =   3480
      Width           =   6495
      Begin VB.TextBox Text8 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   1
         Left            =   4920
         Locked          =   -1  'True
         TabIndex        =   18
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox Text7 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   1
         Left            =   3840
         TabIndex        =   5
         Text            =   "0"
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox Text6 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   1
         Left            =   2760
         TabIndex        =   4
         Text            =   "0"
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox Text5 
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   2535
      End
      Begin VB.Label Label6 
         Caption         =   "Book Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   960
         TabIndex        =   17
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "Quantity"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2880
         TabIndex        =   16
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label8 
         Caption         =   "Rate"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4080
         TabIndex        =   15
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label9 
         Caption         =   "Amount"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5040
         TabIndex        =   14
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.TextBox Text4 
      DataField       =   "Date"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   5040
      Locked          =   -1  'True
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   1080
      Width           =   1695
   End
   Begin VB.TextBox Text3 
      DataField       =   "BilNo"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1080
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      DataField       =   "Add"
      DataSource      =   "Data1"
      Height          =   855
      Left            =   2760
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Top             =   2520
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      DataField       =   "Cname"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   2760
      TabIndex        =   1
      Top             =   1920
      Width           =   2655
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   7215
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Dilip Book Agency"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1680
         TabIndex        =   12
         Top             =   240
         Width           =   3975
      End
   End
   Begin Crystal.CrystalReport rptbill 
      Left            =   3480
      Top             =   5640
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "bill.rpt"
   End
   Begin VB.Label Label10 
      Caption         =   "Total Amount"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4440
      TabIndex        =   20
      Top             =   5640
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   10
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Customer Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   9
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   8
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Bill No"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   7
      Top             =   1080
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'************************************************
'A Simple Database Programming Example         **
'Complete with reports and bill                **
'Program By Sachin Palewar                     **
'E-Mail palewar@hotmail.com                    **
'URL :- http://members.tripod.com/compuwhizkid **
'************************************************
Dim db As Database
Dim rec As Recordset
Dim sav As Boolean
Private Sub Command1_Click(Index As Integer)
Select Case Index
Case 0
If Command1(0).Caption = "New" Then
newrecord 'clear form
Else
saverecord 'save filled form
End If
Case 1
prnrecord 'print bill
Case 2
Unload Me
End Select
End Sub
Private Sub Form_Load()
Set db = OpenDatabase("jaggu.mdb")
newrecord
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set rec = Nothing
Set db = Nothing
End Sub

Private Sub Text6_GotFocus(Index As Integer)
Text6(Index).SelLength = 1
End Sub
Private Sub Text7_GotFocus(Index As Integer)
Text7(Index).SelLength = 1
End Sub

Private Sub Text6_LostFocus(Index As Integer)
If IsNumeric(Text6(Index)) = False Then ' Allows only numeric text
Text6(Index).Text = 0
Text6(Index).SetFocus
Exit Sub
End If
Text8(Index) = Val(Text6(Index)) * Val(Text7(Index)) 'Calculates Amount
End Sub
Private Sub Text7_KeyPress(Index As Integer, KeyAscii As Integer)

If Index = Text7.UBound Then
If KeyAscii = 9 Or KeyAscii = 13 Then   'If Enter key is pressed new set
                                      'of text boxes ate displayed
addtbox ' Sub which displays new set of txtboxes
End If
End If
End Sub
Sub addtbox() 'Sub which loads new set of txtboxes at runtime
If Text8.UBound = 4 Then
VScroll1.Enabled = True
VScroll1.Min = 1
VScroll1.LargeChange = 3
End If
Load Text5(Text5.UBound + 1)
Text5(Text5.UBound).Visible = True
Text5(Text5.UBound).Text = ""
Text5(Text5.UBound).Move Text5(1).Left, Text5(Text5.UBound - 1).Top + 360, Text5(1).Width, Text5(1).Height
Load Text6(Text6.UBound + 1)
Text6(Text6.UBound).Visible = True
Text6(Text6.UBound).Move Text6(1).Left, Text6(Text6.UBound - 1).Top + 360, Text6(1).Width, Text6(1).Height
Text6(Text6.UBound) = 0
Load Text7(Text7.UBound + 1)
Text7(Text7.UBound).Visible = True
Text7(Text7.UBound).Move Text7(1).Left, Text7(Text7.UBound - 1).Top + 360, Text7(1).Width, Text7(1).Height
Text7(Text7.UBound) = 0
Load Text8(Text8.UBound + 1)
Text8(Text8.UBound).Visible = True
Text8(Text8.UBound).Move Text8(1).Left, Text8(Text8.UBound - 1).Top + 360, Text8(1).Width, Text8(1).Height
Text8(Text8.UBound) = 0
Text5(Text5.UBound).SetFocus
VScroll1.Max = CInt(Text8.UBound) - 3
VScroll1.Value = VScroll1.Max
End Sub
Private Sub Text7_lostfocus(Index As Integer)
If IsNumeric(Text7(Index)) = False Then 'Only numbers are allowd to be enteed
Text7(Index) = 0
Text7(Index).SetFocus
Exit Sub
End If
Text8(Index) = Val(Text6(Index)) * Val(Text7(Index))
End Sub

Private Sub Text8_Change(Index As Integer) ' Calulates Final Amount
Text9 = 0
Dim i  As Integer
For i = 1 To Text8.UBound
Text9 = Val(Text9) + Val(Text8(i))
Next
End Sub
Private Sub VScroll1_Change()
If VScroll1.Value >= 1 Then
scrol VScroll1.Value
End If
End Sub

Private Sub VScroll1_Scroll()
If VScroll1.Value >= 1 Then
scrol VScroll1.Value
End If
End Sub
Sub scrol(n As Integer) 'sub for scrolling txtboxes when value of scrollbar changes
Text5(n).Top = 600
Text5(n).ZOrder 0
For i = n + 1 To n + 3
Text5(i).Top = Text5(i - 1).Top + 360
Text5(i).ZOrder 0
Next
Text6(n).Top = 600
Text6(n).ZOrder 0
For i = n + 1 To n + 3
Text6(i).Top = Text5(i - 1).Top + 360
Text6(i).ZOrder 0
Next
Text7(n).Top = 600
Text7(n).ZOrder 0
For i = n + 1 To n + 3
Text7(i).Top = Text5(i - 1).Top + 360
Text7(i).ZOrder 0
Next
Text8(n).Top = 600
Text8(n).ZOrder 0
For i = n + 1 To n + 3
Text8(i).Top = Text5(i - 1).Top + 360
Text8(i).ZOrder 0
Next
End Sub
Sub saverecord() 'sub for saving filled form data
If Text1 = "" Then 'Validation Begins Here
Beep
MsgBox "Please Enter Customer Name", vbCritical, "Data Entry Error"
Text1.SetFocus
Exit Sub
End If
For i = 1 To Text5.UBound
If Text5(i) = "" Then
MsgBox "Please Enter A Name of Book", vbCritical, "Data Entry Error"
Text5(i).SetFocus
Exit Sub
End If
If Text6(i) = 0 Then
MsgBox "Please Enter the Quantity of Book", vbCritical, "Data Entry Error"
Text6(i).SetFocus
Exit Sub
End If
If Text7(i) = 0 Then
MsgBox "Please Enter the Rate of Book", vbCritical, "Data Entry Error"
Text7(i).SetFocus
Exit Sub
End If
Next 'Validation Ends Here
Set rec = db.OpenRecordset("select * from bill")
rec.AddNew
rec!bilno = Text3
rec!cname = Text1
If Text2 <> "" Then
rec!Add = Text2
End If
rec!Date = Text4
rec.Update
Set rec = db.OpenRecordset("select * from bill_detail")
For i = 1 To Text8.UBound
rec.AddNew
rec!billno = Text3
If Text5(i) <> "" Then
rec!bname = Text5(i)
End If
If Text6(i) <> "" Then
rec!qty = Text6(i)
End If
If Text7(i) <> "" Then
rec!price = Text7(i)
End If
rec.Update
Next
Command1(0).Caption = "New"
sav = True
End Sub
Sub newrecord() 'Sub for clearing form
Dim i As Byte
Text4 = Format(Now, "dd-mmm-yyyy") 'Displays Today's Date
Set rec = db.OpenRecordset("select bilno from bill order by bilno")
rec.MoveLast
Text3 = rec!bilno + 1
Text1 = ""
Text2 = ""
i = Text5.UBound
While i <> 1
Unload Text5(i)
Unload Text6(i)
Unload Text7(i)
Unload Text8(i)
i = i - 1
Wend
Text5(1) = ""
Text6(1) = 0
Text7(1) = 0
Text8(1) = 0
Text9 = 0
sav = False
Command1(0).Caption = "Save"
End Sub
Sub prnrecord() 'Sub for printing bill
Dim a As Byte
a = Text3.Text
If sav = False Then
saverecord
End If
rptbill.DataFiles(0) = "jaggu.mdb"
rptbill.SelectionFormula = "{BILL.BILNO}=" & a
rptbill.PrintReport
End Sub


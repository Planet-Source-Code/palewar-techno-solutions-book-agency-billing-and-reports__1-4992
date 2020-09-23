VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "Dilip Book Agency"
   ClientHeight    =   7545
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   9210
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin Crystal.CrystalReport CrystalReport3 
      Left            =   240
      Top             =   3480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "annual.rpt"
   End
   Begin Crystal.CrystalReport CrystalReport2 
      Left            =   240
      Top             =   3000
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "monthly.rpt"
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   240
      Top             =   2520
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "daily.rpt"
   End
   Begin VB.Menu bill 
      Caption         =   "&Bill"
   End
   Begin VB.Menu rep 
      Caption         =   "&Reports"
      Begin VB.Menu tod 
         Caption         =   "&Today's Report"
      End
      Begin VB.Menu mon 
         Caption         =   "&Monthly Report"
      End
      Begin VB.Menu ann 
         Caption         =   "&Annual Report"
      End
   End
End
Attribute VB_Name = "MDIForm1"
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
Private Sub ann_Click() 'Prints annual report
CrystalReport3.DataFiles(0) = "Jaggu.mdb"
CrystalReport3.PrintReport

End Sub

Private Sub bill_Click()
Form1.Show

End Sub

Private Sub mon_Click() 'Prints monthly report
CrystalReport2.DataFiles(0) = "Jaggu.mdb"
CrystalReport2.PrintReport

End Sub

Private Sub tod_Click() 'Prints daily report
CrystalReport1.DataFiles(0) = "Jaggu.mdb"
CrystalReport1.PrintReport

End Sub

VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form FrmCaftrea 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ÇáßÇÝÊÑíÇ"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5880
   Icon            =   "A.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   4590
   ScaleWidth      =   5880
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   3840
      Width           =   5655
      Begin VB.CommandButton Command4 
         Caption         =   "ÊÚÏíá ÇáÕäÝ"
         Height          =   375
         Left            =   2880
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   160
         Width           =   1335
      End
      Begin VB.CommandButton Command3 
         Cancel          =   -1  'True
         Caption         =   "ÅÛáÇÞ"
         Height          =   375
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   160
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "ÍÐÝ ÇáÕäÝ"
         Height          =   375
         Left            =   1560
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   160
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "ÕäÝ ÌÏíÏ"
         Height          =   375
         Left            =   4320
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   160
         Width           =   1215
      End
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      DataField       =   "pamount"
      DataSource      =   "Data1"
      Height          =   495
      Left            =   2400
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Text            =   "Text3"
      Top             =   2.45745e5
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      DataField       =   "pofonep"
      DataSource      =   "Data1"
      Height          =   495
      Left            =   2400
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   2.45745e5
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      DataField       =   "nofp"
      DataSource      =   "Data1"
      Height          =   495
      Left            =   2400
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   2.45745e5
      Width           =   1215
   End
   Begin VB.PictureBox Skin1 
      Height          =   480
      Left            =   2760
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   9
      Top             =   2040
      Width           =   1200
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "E:\almokhtar network\Server\mokdatabase.dll"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   2400
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "coffee"
      RightToLeft     =   -1  'True
      Top             =   2160
      Visible         =   0   'False
      Width           =   1140
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "A.frx":29C12
      Height          =   3615
      Left            =   120
      OleObjectBlob   =   "A.frx":29C26
      TabIndex        =   0
      Top             =   120
      Width           =   5655
   End
End
Attribute VB_Name = "FrmCaftrea"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************
'* Al-Mokhtar For Networks Server       *
'*  By : Mokhtar saied saleh            *
'*      Syria - Abokamal                *
'*      WWW.ABOKAMAL.COM                *
'*  MOKHTAR_SS@HOTMAIL.COM              *
'*       0096394467547                  *
'****************************************
Private Sub Command1_Click()
usrcommand2 = "Add"
cafadd.Show
cafadd.Caption = "ÇÖÇÝÉ ÕäÝ"
StayOnTop cafadd

End Sub

Private Sub Command2_Click()
Dim mok
mok = MsgBox(" åá ÃäÊ ãÊÃßÏ Ãäß ÊÑíÏ ÍÐÝ ÇáÕäÝ ÇáãÓãøì " & Text1.Text, 64 + vbYesNo, "äÙÇã ÇáãÎÊÇÑ")
If mok = vbYes Then
On Error Resume Next
Data1.Recordset.Delete
Data1.Refresh
DBGrid1.ReBind
DBGrid1.Refresh
If Text1.Text = "" Then
Command4.Enabled = False
Command2.Enabled = False
Else
Command4.Enabled = True
Command2.Enabled = True
End If
Else
Exit Sub
End If

End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Command4_Click()
usrcommand2 = "Edit"
cafadd.Show
cafadd.Caption = "ÊÚÏíá ÈíÇäÇÊ ÇáÕäÝ"
StayOnTop cafadd

End Sub

Private Sub DBGrid1_Click()
If Text1.Text = "" Then
Command4.Enabled = False
Command2.Enabled = False
Else
Command4.Enabled = True
Command2.Enabled = True
End If

End Sub

Private Sub Form_Load()
Data1.DatabaseName = App.Path & ("\mokdatabase.dll")
Data1.RecordSource = "select * from coffee"
Data1.Refresh
Skin1.LoadSkin App.Path & ("\green.skn")
Skin1.ApplySkin Me.hwnd
If Text1.Text = "" Then
Command4.Enabled = False
Command2.Enabled = False
Else
Command4.Enabled = True
Command2.Enabled = True
End If

End Sub

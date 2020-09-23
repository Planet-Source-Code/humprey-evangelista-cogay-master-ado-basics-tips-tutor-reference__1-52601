VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmMain 
   Caption         =   "Comgen System's Using ADO"
   ClientHeight    =   5775
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7485
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5775
   ScaleWidth      =   7485
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cboID 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3720
      TabIndex        =   7
      Top             =   240
      Width           =   2295
   End
   Begin VB.CommandButton cmdFirst 
      Height          =   540
      Left            =   120
      Picture         =   "frmMain.frx":27C92
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5160
      UseMaskColor    =   -1  'True
      Width           =   705
   End
   Begin VB.CommandButton cmdPrevious 
      Height          =   540
      Left            =   840
      Picture         =   "frmMain.frx":27FD4
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5160
      UseMaskColor    =   -1  'True
      Width           =   705
   End
   Begin VB.CommandButton cmdNext 
      Height          =   540
      Left            =   1560
      Picture         =   "frmMain.frx":28316
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5160
      UseMaskColor    =   -1  'True
      Width           =   705
   End
   Begin VB.CommandButton cmdLast 
      Height          =   540
      Left            =   2280
      Picture         =   "frmMain.frx":28658
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5160
      UseMaskColor    =   -1  'True
      Width           =   705
   End
   Begin MSDataGridLib.DataGrid dgrRecordset 
      Height          =   4215
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   7435
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   0   'False
      Appearance      =   0
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdCloseCon 
      Caption         =   "&Kill Records and Close Connection "
      Height          =   615
      Left            =   1920
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "&Connect to Database"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
   Begin ADO_Tutor.General General1 
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   240
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------
'   Just dot forget to comment,
'   Vote and acknowlege.
'   Let us Encourage others to
'   Upload for Learning Purpose
'-----------------------------------

Option Explicit
Dim strDBname As String

Private Sub cboID_Click()
'Make sure that you start searching from the first record
    objADORs.MoveFirst
    objADORs.Find "CustomerID = '" & cboID.Text & "'", , adSearchForward
End Sub

Private Sub cmdCloseCon_Click()
  
    objADORs.Close
    Set objADORs = Nothing
    objADOConn.Close
    Set objADOConn = Nothing
    cboID.Clear
    DisableAllNavCmd
    
    cmdConnect.Enabled = True
    cmdCloseCon.Enabled = False
    
End Sub

Private Sub cmdConnect_Click()
'Set ADO Objects
    Set objADOConn = New Connection
    Set objADORs = New Recordset
    Set objADOCmd = New Command
    
'Set Properties of the Connection object
    With objADOConn
        .CursorLocation = adUseClient
        'open oledb provider for Microsoft Access
        .Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\database\" & strDBname & ";"
    End With
    
'Set Properties of the Command Object
    With objADOCmd
        .CommandType = adCmdText
        .Name = "GetCustomers"
        .ActiveConnection = objADOConn
        .CommandText = "select * from customers"
    End With
       
'Execute the command object and pass the results unto the recordset object
    objADOConn.GetCustomers objADORs
    
'Load Data to Combo before setting the grids
'data source for speed optimization

'Load Data (Customer ID's) onto the combobox
    cboID.Clear
    cboID.Text = objADORs.Fields(0)
    While Not objADORs.EOF
        cboID.AddItem objADORs.Fields(0)
        objADORs.MoveNext
    Wend
    
'Set Datasource of the Database Grid Object
    Set dgrRecordset.DataSource = objADORs

            
'Set controls to their proper states
    cmdConnect.Enabled = False
    cmdCloseCon.Enabled = True
    EnableAllNavCmd
End Sub

Private Sub Form_Load()
'name of my database
    strDBname = "Northwind.mdb"
'Set Controls into their proper states
    cmdConnect.Enabled = True
    cmdCloseCon.Enabled = False
    DisableAllNavCmd
    Beep
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
'trap the close button on your Forms Control Area
    If MsgBox("Are you sure you want to exit?", vbQuestion + vbYesNo, "Comgen Systems") = vbYes Then
    'Close the objects before closing your Application to save resources
        If objADORs.State = adStateOpen Then
            objADORs.Close
            Set objADORs = Nothing
        End If
        
        If objADOConn.State = adStateOpen Then
            objADOConn.Close
            Set objADOConn = Nothing
        End If
        End
    Else
        Cancel = 1
    End If
End Sub

Private Sub cmdFirst_Click()
On Error GoTo GoFirstError
    
    EnableAllNavCmd
    
    objADORs.MoveFirst
    cmdFirst.Enabled = False
    cmdPrevious.Enabled = False
    Exit Sub
GoFirstError:
    MsgBox Err.Description
End Sub

Private Sub cmdLast_Click()
On Error GoTo GoLastError
    
    EnableAllNavCmd
    
    objADORs.MoveLast
    cmdLast.Enabled = False
    cmdNext.Enabled = False
    Exit Sub
GoLastError:
    MsgBox Err.Description
End Sub

Private Sub cmdNext_Click()
On Error GoTo GoNextError
    
    EnableAllNavCmd
    
    If Not objADORs.EOF Then objADORs.MoveNext
    
    If objADORs.EOF And objADORs.RecordCount > 0 Then
        cmdLast.Enabled = False
        cmdNext.Enabled = False
        MsgBox "Last Record Reached", vbOKOnly, "Comgen Systems"
        
        'moved off the end so go back
        objADORs.MoveLast
    End If
    Exit Sub
GoNextError:
    MsgBox Err.Description
End Sub

Private Sub cmdPrevious_Click()
    On Error GoTo GoPrevError
    
    EnableAllNavCmd
    
    If Not objADORs.BOF Then objADORs.MovePrevious
    
    If objADORs.BOF And objADORs.RecordCount > 0 Then
        cmdFirst.Enabled = False
        cmdPrevious.Enabled = False
        MsgBox "First Record Reached", vbOKOnly, "Comgen Systems"
        
        'moved off the end so go back
        objADORs.MoveFirst
    End If
    Exit Sub
GoPrevError:
    MsgBox Err.Description
End Sub

Private Sub EnableAllNavCmd()
    cmdFirst.Enabled = True
    cmdPrevious.Enabled = True
    cmdNext.Enabled = True
    cmdLast.Enabled = True
End Sub

Private Sub DisableAllNavCmd()
    cmdFirst.Enabled = False
    cmdPrevious.Enabled = False
    cmdNext.Enabled = False
    cmdLast.Enabled = False
End Sub


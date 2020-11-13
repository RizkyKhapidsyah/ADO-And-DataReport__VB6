VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00008080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Learn Add new,Update, delete and Print"
   ClientHeight    =   3924
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   4572
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3924
   ScaleWidth      =   4572
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Txtname 
      Height          =   288
      Left            =   864
      TabIndex        =   1
      Top             =   432
      Width           =   1884
   End
   Begin VB.CommandButton cmdaddnew 
      Caption         =   "&Add New"
      Height          =   324
      Left            =   3048
      TabIndex        =   4
      Top             =   72
      Width           =   1380
   End
   Begin VB.CommandButton cmddelete 
      Caption         =   "&Delete"
      Height          =   324
      Left            =   3048
      TabIndex        =   6
      Top             =   816
      Width           =   1380
   End
   Begin VB.CommandButton cmdupdate 
      Caption         =   "&Update"
      Height          =   324
      Left            =   3048
      TabIndex        =   5
      Top             =   444
      Width           =   1380
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Print"
      Height          =   324
      Left            =   3048
      TabIndex        =   7
      Top             =   1188
      Width           =   1380
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   2316
      Left            =   48
      TabIndex        =   8
      ToolTipText     =   "Click this flex grid to put the records to the boxes"
      Top             =   1560
      Width           =   4476
      _ExtentX        =   7895
      _ExtentY        =   4085
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      BackColor       =   16776960
      BackColorFixed  =   65280
      BackColorBkg    =   32896
      GridColor       =   255
      GridLinesFixed  =   1
      PictureType     =   1
      BorderStyle     =   0
      Appearance      =   0
      FormatString    =   "ID    | Name                          | Age        |               Sex"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox Txtid 
      Height          =   288
      Left            =   864
      TabIndex        =   0
      Top             =   60
      Width           =   1884
   End
   Begin VB.TextBox Txtsex 
      Height          =   288
      Left            =   864
      TabIndex        =   3
      Top             =   1200
      Width           =   1884
   End
   Begin VB.TextBox Txtage 
      Height          =   288
      Left            =   864
      TabIndex        =   2
      Top             =   828
      Width           =   1884
   End
   Begin VB.Label Label4 
      BackColor       =   &H000040C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Sex :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   300
      Left            =   120
      TabIndex        =   12
      Top             =   1188
      Width           =   660
   End
   Begin VB.Label Label3 
      BackColor       =   &H000040C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Age :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   300
      Left            =   120
      TabIndex        =   11
      Top             =   780
      Width           =   660
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H000040C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Name :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   240
      Left            =   96
      TabIndex        =   10
      Top             =   420
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H000040C0&
      BackStyle       =   0  'Transparent
      Caption         =   "ID No :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   240
      Left            =   96
      TabIndex        =   9
      Top             =   72
      Width           =   708
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents Con As ADODB.Connection
Attribute Con.VB_VarHelpID = -1
Dim WithEvents rst As ADODB.Recordset
Attribute rst.VB_VarHelpID = -1
Dim cmd As ADODB.Command
Private Sub cmdaddnew_Click()

chec:
On Error GoTo errh

Set rst = New ADODB.Recordset 'specifying attributes to this recordset

With rst

    .ActiveConnection = Con
    .CursorLocation = adUseClient
    .CursorType = adOpenDynamic
    .LockType = adLockOptimistic

.Open "tab1" 'opening tab1 table

End With


'adding records from textbox to recordset

With rst

    .AddNew
    
    .Fields!id = StrConv(Txtid, vbProperCase)
    .Fields!Name = StrConv(Txtname, vbProperCase)
    .Fields!age = StrConv(Txtage, vbProperCase)
    .Fields!sex = StrConv(Txtsex, vbProperCase)
    
    .Update
    
End With

' clearing the text boxes

Txtname = ""
Txtid = ""
Txtage = ""
Txtsex = ""

' closing the recordset
rst.Close
Set rst = Nothing

Call dload ' calling private procedure to fill the flexgrid

errh:                     'in case of error, informing the user

If Err.Description <> vbNullString Then
    MsgBox Err.Description
End If
    

End Sub

Private Sub cmddelete_Click()

Set cmd = New ADODB.Command ' using command object to execute sql commands

With cmd

    .ActiveConnection = Con
    .CommandType = adCmdText
    .CommandText = "delete from tab1 where id = '" & Txtid & "'"
    .Execute

End With

Set cmd = Nothing

' clearing all the text boxes

Txtname = ""
Txtid = ""
Txtage = ""
Txtsex = ""

Call dload ' calling procedure to fill flexgrid


End Sub

Private Sub cmdupdate_Click()

On Error GoTo errhan

Set rst = New ADODB.Recordset

With rst

    .CursorLocation = adUseClient
    .ActiveConnection = Con
    .CursorType = adOpenDynamic
    .LockType = adLockPessimistic
    
    .Open "select * from tab1 where id='" & Txtid.Text & "'" 'opening the recordset
    
    .Fields!Name = StrConv(Txtname, vbProperCase)
    .Fields!sex = StrConv(Txtsex, vbProperCase)
    .Fields!age = StrConv(Txtage, vbProperCase)
    
    .Update ' updating the recordset

End With

Set rst = Nothing

Call dload

Txtname = ""
Txtid = ""
Txtage = ""
Txtsex = ""

errhan:

If Err.Description <> vbNullString Then
    MsgBox Err.Description
End If

End Sub
Public Sub connect()

Set Con = New ADODB.Connection

Con.CursorLocation = adUseClient

' use this code to connect to the database using universal data link

'Con.Open "File Name=" & App.Path & "\test.udl"

Con.Open "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=" & App.Path & "\test.mdb"

If Con.Provider = "SQLOLEDB.1" Then
    
    DataEnvironment1.Connections(2).Open Con

Else

    DataEnvironment1.Connections(1).Open Con
    
End If

Call dload

End Sub
Private Sub dload()

MSFlexGrid1.Rows = 1

Set rst = New ADODB.Recordset

    rst.ActiveConnection = Con
    rst.CursorLocation = adUseClient
    rst.CursorType = adOpenDynamic
    rst.LockType = adLockOptimistic
    rst.Source = "tab1"
    rst.Open

While Not rst.EOF() ' checking end of file

    MSFlexGrid1.AddItem rst!id & Chr(9) & rst!Name & Chr(9) & rst!age & Chr(9) & rst!sex 'adding records to flexgrid
    
    rst.MoveNext

Wend

Set rst = Nothing

End Sub

Private Sub Command1_Click()

With DataEnvironment1

    If Con.Provider = "SQLOLEDB.1" Then
            
            
            .Commands(2).CommandType = adCmdText
            .Commands(2).CommandText = "SELECT * FROM tab1 where id = '" & Txtid.Text & "'"
            .Commands(2).Execute
            
        DataReport2.Show
            
            
          If .rsCommand2.State = 1 Then
          
            .rsCommand2.Close
          
          End If
        
     
    Else
    
     
            .Commands(1).CommandType = adCmdText
            .Commands(1).CommandText = "SELECT * FROM tab1 where id = '" & Txtid.Text & "'"
            .Commands(1).Execute
        
        DataReport1.Show
        
        If .rsCommand1.State = 1 Then
          
            .rsCommand1.Close
          
        End If
     
    
    End If

End With

End Sub

Private Sub Form_Load()

Call connect
 
End Sub

Private Sub Form_Unload(Cancel As Integer)

Con.Close
Set Con = Nothing

End Sub

Private Sub MSFlexGrid1_Click()

With MSFlexGrid1 ' populating the text boxes when user clicks the flexgrid

    .Col = 0
        Txtid.Text = .Text
    .Col = 1
        Txtname.Text = .Text
    .Col = 2
        Txtage.Text = .Text
    .Col = 3
        Txtsex.Text = .Text
        
End With

End Sub



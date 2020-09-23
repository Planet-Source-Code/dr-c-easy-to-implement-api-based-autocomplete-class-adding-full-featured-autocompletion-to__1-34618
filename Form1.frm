VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   1320
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   0  'Table
      RecordSource    =   ""
      Top             =   2520
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CheckBox Check4 
      Caption         =   "ForceListUsage"
      Height          =   255
      Left            =   2160
      TabIndex        =   7
      Top             =   1200
      Width           =   1575
   End
   Begin VB.CheckBox Check3 
      Caption         =   "AutoUpdateList"
      Height          =   255
      Left            =   2160
      TabIndex        =   5
      Top             =   840
      Width           =   1455
   End
   Begin VB.CheckBox Check2 
      Caption         =   "LimitToList"
      Height          =   255
      Left            =   2160
      TabIndex        =   4
      Top             =   480
      Width           =   1455
   End
   Begin VB.CheckBox Check1 
      Caption         =   "AutoDropdown"
      Height          =   255
      Left            =   2160
      TabIndex        =   2
      Top             =   120
      Value           =   1  'Checked
      Width           =   1455
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   3360
      TabIndex        =   1
      Top             =   2520
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "clsAutoComplete error messages are displayed here..."
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   2040
      Width           =   4575
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "www.dtdn.com/dev"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   1200
      MouseIcon       =   "Form1.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   2880
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents clsAC As clsAutoComplete 'define a class instance
Attribute clsAC.VB_VarHelpID = -1

Private Sub Check1_Click()                  'turn autodropdown on/off
    If Check1.Value Then                    'if checked
        clsAC.AutoDropdown = True           'turn on
    Else                                    'if not checked
        clsAC.AutoDropdown = False          'turn off
    End If
End Sub

Private Sub Check2_Click()                  'turn limittolist on/off
    If Check2.Value Then                    'if checked
        clsAC.LimitToList = True            'turn on
    Else                                    'if not checked
        clsAC.LimitToList = False           'turn off
    End If
End Sub

Private Sub Check3_Click()                  'turn autoupdatelist on/off
    If Check3.Value Then                    'if checked
        clsAC.AutoUpdateList = True         'turn on
    Else                                    'if not checked
        clsAC.AutoUpdateList = False        'turn off
    End If
End Sub

Private Sub Check4_Click()                  'turn forcelistusage on/off
    If Check4.Value Then                    'if checked
        clsAC.ForceListUsage = True         'turn on
    Else                                    'if not checked
        clsAC.ForceListUsage = False        'turn off
    End If
End Sub

Private Sub clsAC_Message(ByVal strMessage As String)
    Label1.Caption = strMessage             'set caption to status message from class
End Sub

Private Sub cmdClose_Click()
    Set clsAC = Nothing                     'destroy the class
    Unload Me                               'unload the form
End Sub

Private Sub Form_Load()
    Data1.Connect = "Access"                'set data control properties
    Data1.DatabaseName = App.Path & "\AutoComplete.mdb" 'db in current dir
    Data1.Exclusive = False                 'do not lock db
    Data1.ReadOnly = True                   'not in update mode either
    Data1.RecordsetType = 0                 'table
    Data1.RecordSource = "tblValues"        'target table in db
    Set clsAC = New clsAutoComplete         'initialize mem for class
    Set clsAC.LinkedComboBox = Combo1       'assign the combo to the class
'    Combo1.AddItem "aabbcc"                 'add sample data to the combo
'    Combo1.AddItem "aaccbb"
'    Combo1.AddItem "bbaacc"
'    Combo1.AddItem "bbccaa"
'    Combo1.AddItem "ccaabb"
'    Combo1.AddItem "ccbbaa"
    'sequence of assigning RowSourceField and RowSource does NOT matter
    clsAC.RowSourceField = "Value"          'set row source field
    Set clsAC.RowSource = Data1             'set rowsource to the data control
End Sub

Private Sub Label10_Click()                 'redirect to Source Code Utopia
  Dim q As Variant
  q = "http://www.dtdn.com/dev"
  q = ShellExecute(0&, vbNullString, q, vbNullString, vbNullString, vbNormalFocus)
End Sub

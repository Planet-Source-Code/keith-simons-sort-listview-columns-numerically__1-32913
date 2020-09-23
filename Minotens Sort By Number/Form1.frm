VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form Form1 
   BackColor       =   &H80000013&
   Caption         =   "Minoten's Sort By Number Module Example"
   ClientHeight    =   8475
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8955
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8475
   ScaleWidth      =   8955
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      Caption         =   "Sort without minoten's module."
      Height          =   375
      Left            =   4920
      TabIndex        =   5
      Top             =   7080
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   3240
      TabIndex        =   3
      Text            =   "200"
      Top             =   7080
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Generate List"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   6960
      Width           =   2175
   End
   Begin MSComctlLib.ListView ListView3 
      Height          =   375
      Left            =   7920
      TabIndex        =   1
      Top             =   7080
      Visible         =   0   'False
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   6735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   11880
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Column 1"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Column 2"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Column 3"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Column 4"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Column 5"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Coded by Minoten - Minoten@hotmail.com"
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   8160
      Width           =   9015
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "CAUTION: Because of the looping involved in the module, slower computers may freeze for a moment."
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   7680
      Width           =   8895
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      Height          =   15
      Left            =   5640
      TabIndex        =   6
      Top             =   7200
      Width           =   15
   End
   Begin VB.Label Label1 
      Caption         =   "Rows:"
      Height          =   255
      Left            =   2760
      TabIndex        =   4
      Top             =   7200
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strCLicked As Boolean


Private Sub Command1_Click()
MsgBox ListView1.ListItems.Count
End Sub

Private Sub Command2_Click()
On Error GoTo oops:

'All this does is insert numbers into listview randomizing bold/checked/forecolor for module testing.
ListView1.ListItems.Clear

    Dim strCount, strCount2, strCount3, strCount4 As Integer
    strCount = 1
    
    Do Until strCount = Text1.Text + 1
    strCount2 = Int(1000 * Rnd)
    
    
    ListView1.ListItems.Add , , strCount2
    strCount3 = ListView1.ListItems.Count
    ListView1.ListItems(strCount3).ListSubItems.Add , , strCount2
    ListView1.ListItems(strCount3).ListSubItems.Add , , strCount2
    ListView1.ListItems(strCount3).ListSubItems.Add , , strCount2
    ListView1.ListItems(strCount3).ListSubItems.Add , , strCount2 & " w/ letters"
    
     If Int(2 * Rnd) = 1 Then
     ListView1.Checkboxes = True
     ListView1.ListItems(strCount3).Checked = True
     End If
     
     If Int(2 * Rnd) = 1 Then
     ListView1.ListItems(strCount3).Bold = True
     End If
     
     If Int(2 * Rnd) = 1 Then
     ListView1.ListItems(strCount3).ForeColor = &HC0&
     End If
    
     strCount4 = 1
     Do Until strCount4 = 5
        If Int(2 * Rnd) = 1 Then
        ListView1.ListItems(strCount3).ListSubItems(strCount4).ForeColor = &HC0&
        End If
        If Int(2 * Rnd) = 1 Then
        ListView1.ListItems(strCount3).ListSubItems(strCount4).Bold = True
        End If
     strCount4 = strCount4 + 1
     Loop
    
     
    strCount = strCount + 1
    strCount2 = strCount2 - 1
    Loop
    
Exit Sub

oops:
MsgBox "error creating listview numbers"

End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo oops:

    If Check1.Value = 1 Then
        If strCLicked = True Then
        ListView1.SortOrder = lvwAscending
        ListView1.Sorted = True
        strCLicked = False
        Else
        ListView1.SortOrder = lvwDescending
        ListView1.Sorted = True
        strCLicked = True
        End If
    Else
        If strCLicked = True Then
        '### An Example on how to call on the code ###
        sortbynum ListView1, ListView3, ColumnHeader.Index - 1, False
        strCLicked = False
        Else
        '### An Example on how to call on the code ###
        sortbynum ListView1, ListView3, ColumnHeader.Index - 1, True
        strCLicked = True
        End If
    End If

oops:
End Sub

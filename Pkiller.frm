VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00800000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ProcessKiller by Joe Wong (ICQ:7601450)"
   ClientHeight    =   3690
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5415
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   7.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Pkiller.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   5415
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin ComctlLib.ListView lvw 
      Height          =   2895
      Left            =   67
      TabIndex        =   0
      Top             =   120
      Width           =   5280
      _ExtentX        =   9313
      _ExtentY        =   5106
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   4194304
      BackColor       =   12648447
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00800000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   900
      TabIndex        =   1
      Top             =   3000
      Width           =   3615
      Begin VB.CommandButton CmdEnd 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "End Process"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1920
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton CmdView 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "List Process"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   1455
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdView_Click()
Dim i As Long
Dim proc As PROCESSENTRY32
Dim snap As Long
Dim exename As String
lvw.ListItems.Clear 'clear listview contents
snap = CreateToolhelpSnapshot(TH32CS_SNAPall, 0) 'get snapshot handle
proc.dwSize = Len(proc)
theloop = ProcessFirst(snap, proc)       'first process and return value
i = 0
While theloop <> 0      'next process
exename = proc.szExeFile
ret = lvw.ListItems.Add(, "first" & CStr(i), exename)   'add process name to listview
lvw.ListItems("first" & CStr(i)).SubItems(1) = proc.th32ProcessID   'add process ID to listview
i = i + 1
theloop = ProcessNext(snap, proc)
Wend
CloseHandle snap       'close snapshot handle
End Sub
Private Sub CmdEnd_Click()
Dim i As Long
hand = OpenProcess(process_terminate, True, CLng(lvw.SelectedItem.SubItems(1))) 'get process handle
TerminateProcess hand, 0      'end define process
Call CmdView_Click
End Sub

Private Sub Form_Load()
Dim header As ColumnHeader
lvw.View = lvwReport
lvw.ColumnHeaders.Clear
Set header = lvw.ColumnHeaders.Add(, "first", "Process", 4200)  'set listview width
Set header = lvw.ColumnHeaders.Add(, "second", "ID", 1200)
lvw.Refresh
End Sub


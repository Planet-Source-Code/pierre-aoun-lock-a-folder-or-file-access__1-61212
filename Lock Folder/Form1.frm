VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1515
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4080
   LinkTopic       =   "Form1"
   ScaleHeight     =   1515
   ScaleWidth      =   4080
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdunlock 
      Caption         =   "unlock"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2040
      TabIndex        =   2
      Top             =   720
      Width           =   1575
   End
   Begin VB.CommandButton cmdlock 
      Caption         =   "lock"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Text            =   "c:\temp"
      Top             =   120
      Width           =   3735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Author: Pierre AOUN
'email: pierre_aoun@hotmail.com
Option Explicit
Private Const FILE_LIST_DIRECTORY = &H1
Private Const FILE_SHARE_READ = &H1&
Private Const FILE_SHARE_DELETE = &H4&
Private Const OPEN_EXISTING = 3
Private Const FILE_FLAG_BACKUP_SEMANTICS = &H2000000
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal PassZero As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal PassZero As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Dim File_Share_Flag As Long
Dim hDir As Long
Private Sub cmdlock_Click()
    Dim PathDir As String
    PathDir = Text1.Text
    hDir = CreateFile(PathDir, FILE_LIST_DIRECTORY, File_Share_Flag, _
                      ByVal 0&, OPEN_EXISTING, FILE_FLAG_BACKUP_SEMANTICS, ByVal 0&)
    cmdlock.Enabled = False
    cmdunlock.Enabled = True
End Sub
Private Sub cmdunlock_Click()
    CloseHandle hDir
    cmdlock.Enabled = True
    cmdunlock.Enabled = False
End Sub
Private Sub Form_Load()
    File_Share_Flag = 0 'if =FILE_SHARE_READ then read only (for example)
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Call cmdunlock_Click
End Sub

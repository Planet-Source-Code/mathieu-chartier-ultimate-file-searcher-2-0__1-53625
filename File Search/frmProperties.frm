VERSION 5.00
Begin VB.Form frmProperties 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Properties"
   ClientHeight    =   2625
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   5070
   Icon            =   "frmProperties.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   175
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   338
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   255
      Left            =   2520
      TabIndex        =   6
      Top             =   2280
      Width           =   2295
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2280
      Width           =   2295
   End
   Begin VB.CheckBox chkHidden 
      Caption         =   "&Hidden"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      TabIndex        =   4
      Top             =   1440
      Width           =   2295
   End
   Begin VB.CheckBox chkArchive 
      Caption         =   "&Archive"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      TabIndex        =   3
      Top             =   1800
      Width           =   2295
   End
   Begin VB.CheckBox chkReadOnly 
      Caption         =   "&Read Only"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   2415
   End
   Begin VB.CheckBox chkSystem 
      Caption         =   "&System"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1800
      Width           =   2415
   End
   Begin VB.Label lblPath 
      Caption         =   "C:\CRAP"
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4815
   End
End
Attribute VB_Name = "frmProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim Attr As VbFileAttribute
    If chkReadOnly Then Attr = Attr Or vbReadOnly
    If chkHidden Then Attr = Attr Or vbHidden
    If chkSystem Then Attr = Attr Or vbSystem
    If chkArchive Then Attr = Attr Or vbArchive
    Call SetAttr(lblPath, Attr)
    Unload Me
End Sub

Private Sub Form_Load()
    Show
    frmMain.Enabled = False
End Sub

Sub SetFileName(FileName As String)
    Dim Attr As Long
    lblPath = FileName
    Attr = GetAttr(FileName)
    chkReadOnly = -((Attr And vbReadOnly) <> 0)
    chkHidden = -((Attr And vbHidden) <> 0)
    chkSystem = -((Attr And vbSystem) <> 0)
    chkArchive = -((Attr And vbArchive) <> 0)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmMain.Enabled = True
End Sub

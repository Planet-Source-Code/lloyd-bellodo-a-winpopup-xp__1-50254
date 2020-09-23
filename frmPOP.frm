VERSION 5.00
Begin VB.Form frmPOP 
   BackColor       =   &H0000FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Winpopup XP"
   ClientHeight    =   3150
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4635
   Icon            =   "frmPOP.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   4635
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cboTo 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   345
      Left            =   420
      Sorted          =   -1  'True
      TabIndex        =   4
      Top             =   60
      Width           =   2775
   End
   Begin VB.CommandButton cmdSend 
      BackColor       =   &H0000FFFF&
      Caption         =   "Send"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3330
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   60
      UseMaskColor    =   -1  'True
      Width           =   1245
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      Caption         =   "Message:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2640
      Left            =   60
      TabIndex        =   1
      Top             =   435
      Width           =   4515
      Begin VB.TextBox txtMsg 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   2280
         Left            =   150
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   3
         Top             =   240
         Width           =   4230
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "To:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   75
      TabIndex        =   0
      Top             =   120
      Width           =   300
   End
End
Attribute VB_Name = "frmPOP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSend_Click()
    ' HEY! TRY TO READ THIS CODE.
    ' IT'S VERY SIMPLE!
    ' COMMENTS ARE FOR YOU
    ' TO UNDERSTAND IT EASILY.

    ' send message through NET.EXE with syntax:
    ' "NET<SPACE>SEND<SPACE><RECEIVERNAME><SPACE><MESSAGE>"
    Shell "net send " & Trim(cboTo.Text) & " " & Trim(txtMsg.Text)
    
    DoEvents
    
    ' NOTE:
    ' The previous ONE-LINE CODE is enough to send a message,
    ' the following codes are help for users.
    ' it clears the message and sets the focus to the message box,
    ' and saves the every new receiver name without duplicating
    
    txtMsg.Text = ""  ' erase msg previous message
    txtMsg.SetFocus   ' setsfocus on the message box
    
    Dim EXISTS As Boolean  ' prepare a variable that holds a boolean value
                           ' whether receiver name already exists or not
    EXISTS = False         ' sets it to false
    
    ' compare receiver name to each and every list on the combo box
    ' if found sets variable to TRUE
    For Counter = 0 To cboTo.ListCount - 1
        If cboTo.List(ctr) = cboTo.Text Then EXISTS = True
    Next
    
    ' add receiver name if such name don't exists
    If EXISTS = False Then cboTo.AddItem cboTo.Text
    
    ' YOU DON'T HAVE TO VOTE FOR ME, FELLOW PROGRAMMERS.
    ' JUST VISIT "www.angdatingdaan.org"
    ' IT'S FOR YOU...
    ' GOD LOVES YOU!
End Sub

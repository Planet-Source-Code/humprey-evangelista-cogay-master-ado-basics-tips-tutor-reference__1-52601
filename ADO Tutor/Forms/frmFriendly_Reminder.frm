VERSION 5.00
Begin VB.Form frmFriendly_Reminder 
   Caption         =   "Friendly Reminder"
   ClientHeight    =   3735
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5355
   Icon            =   "frmFriendly_Reminder.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3735
   ScaleWidth      =   5355
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtReminder 
      Height          =   3015
      Left            =   360
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "frmFriendly_Reminder.frx":27C92
      Top             =   240
      Width           =   4335
   End
End
Attribute VB_Name = "frmFriendly_Reminder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
     txtReminder.Text = "Copy Right Comgen Systems 2004" & vbCrLf & _
                "'-----------------------------------" & vbCrLf & _
                "'   Just dont forget to comment," & vbCrLf & _
                "'   Vote and acknowlege." & vbCrLf & _
                "'   LET US ENCOURAGE OTHERS " & vbCrLf & _
                "'   Upload for Learning Purpose" & vbCrLf & _
                "'-----------------------------------"
End Sub

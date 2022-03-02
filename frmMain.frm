VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMain 
   Caption         =   "Tags Mercado Livre"
   ClientHeight    =   5640
   ClientLeft      =   18225
   ClientTop       =   6465
   ClientWidth     =   3300
   OleObjectBlob   =   "frmMain.frx":0000
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub bt_novo_arquivo_Click()
    redimdoc
End Sub

Private Sub bt_tag_1_Click()
    copy_tag_1
End Sub

Private Sub bt_tag_2_Click()
    copy_tag_2
End Sub

Private Sub bt_tag_3_Click()
    copy_tag_3
End Sub

Private Sub CheckBox1_Click()

End Sub

Private Sub op_jadlog_Click()

End Sub

Private Sub OptionButton1_Click()

End Sub

Private Sub UserForm_Activate()
   'frmMain.ct_active_doc
End Sub


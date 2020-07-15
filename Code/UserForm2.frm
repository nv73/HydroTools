VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "About"
   ClientHeight    =   5685
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3765
   OleObjectBlob   =   "UserForm2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_initialize()

Dim quote As Variant
Dim randNum As Integer

ReDim quote(7)

quote = Array("The secret to doing anything is believing that you can do it. Anything that you believe you can do strong enough, you can do. Anything. As long as you believe.", "I remember when my Dad told me as a kid, ‘If you want to catch a rabbit, stand behind a tree and make a noise like a carrot. Then when the rabbit comes by you grab him. Works pretty good until you try to figure out what kind of noise a carrot makes…", "Just gotta beat the devil out of it.", "This is your world, you're the creator, find freedom on the canvas.", "If you believe that you can do it, then you can do it.", "Talk to the tree, make friends with it.", "We don't make mistakes, we just have happy accidents.")

randNum = Int((7 - 1 + 1) * Rnd + 1)

randQuote.Caption = quote(randNum)

End Sub



Private Sub CommandButton1_Click()
Unload Me

End Sub




VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmReviewFinalize 
   Caption         =   "FP-GEN: Review & Finalize"
   ClientHeight    =   8085
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15270
   OleObjectBlob   =   "frmReviewFinalize.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmReviewFinalize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub lblTitle_Click()

End Sub

Private Sub UserForm_Initialize()
    modReviewFinalize.Form_Initialize Me
End Sub

Private Sub ListSections_Click()
    modReviewFinalize.Form_SectionSelect Me
End Sub

Private Sub cmdSaveDraft_Click()
    modReviewFinalize.Form_SaveDraft Me
End Sub

Private Sub cmdValidate_Click()
    modReviewFinalize.Form_Validate Me
End Sub

Private Sub cmdGenerateFP_Click()
    modReviewFinalize.Form_GenerateFP Me
End Sub

Private Sub cmdRegenerate_Click()
    modReviewFinalize.Form_Regenerate Me
End Sub

Private Sub cmdDiff_Click()
    modReviewFinalize.Form_ShowDiff Me
End Sub

Private Sub cmdApprove_Click()
    modReviewFinalize.Form_Approve Me
End Sub

Private Sub cmdLock_Click()
    modReviewFinalize.Form_Lock Me
End Sub

Private Sub cmdRecomputeFee_Click()
    modReviewFinalize.Form_RecomputeFee Me
End Sub

Private Sub cmdAttachCV_Click()
    modReviewFinalize.Form_AttachCV Me
End Sub

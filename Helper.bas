Public WithEvents App As Application

Private Sub App_DocumentChange()
    Main.HandyRef_UpdateRibbonState
End Sub

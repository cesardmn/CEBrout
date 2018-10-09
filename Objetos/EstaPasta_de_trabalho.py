Private Sub Workbook_Open()
    
    Application.ScreenUpdating = False
    
        'bloqueia celulas '
        Worksheets("home").Unprotect Password:="adminadmin"
        Worksheets("home").Protect Password:="adminadmin", UserInterfaceOnly:=True
        
        Worksheets("entrada").Unprotect Password:="adminadmin"
        Worksheets("entrada").Protect Password:="adminadmin", UserInterfaceOnly:=True
        
        Worksheets("transfer").Unprotect Password:="adminadmin"
        Worksheets("transfer").Protect Password:="adminadmin", UserInterfaceOnly:=True
        
        Worksheets("relot").Unprotect Password:="adminadmin"
        Worksheets("relot").Protect Password:="adminadmin", UserInterfaceOnly:=True
        
        Worksheets("saida").Unprotect Password:="adminadmin"
        Worksheets("saida").Protect Password:="adminadmin", UserInterfaceOnly:=True
        
        'Habilita tela cheia'
        Application.DisplayFullScreen = True
        ActiveWindow.DisplayGridlines = False
        ActiveWindow.DisplayHeadings = False
        ActiveWindow.DisplayWorkbookTabs = False
        ActiveWindow.DisplayHorizontalScrollBar = False
        ActiveWindow.DisplayVerticalScrollBar = False
        Application.DisplayFormulaBar = False
        
       
    Sheets("home").Select
        
    Application.ScreenUpdating = True

End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Application.ScreenUpdating = False

        'Desabilita tela cheia ao sair'
        Application.DisplayFullScreen = False
        ActiveWindow.DisplayGridlines = True
        ActiveWindow.DisplayHeadings = True
        ActiveWindow.DisplayWorkbookTabs = True
        ActiveWindow.DisplayHorizontalScrollBar = True
        ActiveWindow.DisplayVerticalScrollBar = True
        Application.DisplayFormulaBar = True
        
        
    Application.ScreenUpdating = True

End Sub




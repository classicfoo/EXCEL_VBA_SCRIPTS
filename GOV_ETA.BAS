Sub gov_eta()

'Autofit, Autofilter and show only certain columns

    Columns("A:FP").Select
    Columns("A:FP").EntireColumn.AutoFit
    Selection.AutoFilter
    Selection.EntireColumn.Hidden = True

'a = order num
    Columns("A:A").Select
    Selection.EntireColumn.Hidden = False

'j = customer name
    Columns("J:J").Select
    Selection.EntireColumn.Hidden = False

'ed = initial edd
    Columns("ED:ED").Select
    Selection.EntireColumn.Hidden = False
    
'es = revised edd
    Columns("ES:ES").Select
    Selection.EntireColumn.Hidden = False

'ap = fdd y/n
    Columns("AP:AP").Select
    Selection.EntireColumn.Hidden = False

'eb = fdd
    Columns("EB:EB").Select
    Selection.EntireColumn.Hidden = False
    
'o = po number
    Columns("O:O").Select
    Selection.EntireColumn.Hidden = False
    
'cr = Product
    Columns("CR:CR").Select
    Selection.EntireColumn.Hidden = False
    
'bn = qty
    Columns("BN:BN").Select
    Selection.EntireColumn.Hidden = False
    
End Sub

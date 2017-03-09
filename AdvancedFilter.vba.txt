Sub FilterCopyToOtherSheetGHMNE()
With Sheets
    .Add().Name = "Cape Unit"
    .Add().Name = "Majors"
    .Add().Name = "Mariner"
    .Add().Name = "Metro"
    .Add().Name = "North"
    .Add().Name = "NorthWest"
    .Add().Name = "TeleCenter"
    .Add().Name = "West"
    .Add().Name = "ENM"
    .Add().Name = "Fallriver"
End With
    
    Sheets("Total").Range("A:A:I:I").AdvancedFilter _
        Action:=xlFilterCopy, _
        CriteriaRange:=Sheets("Total").Range("L1:L2"), _
        CopyToRange:=Sheets("Cape Unit").Range("A1"), _
        Unique:=False
    Sheets("Total").Range("A:A:I:I").AdvancedFilter _
        Action:=xlFilterCopy, _
        CriteriaRange:=Sheets("Total").Range("M1:M2"), _
        CopyToRange:=Sheets("Majors").Range("A1"), _
        Unique:=False
    Sheets("Total").Range("A:A:I:I").AdvancedFilter _
        Action:=xlFilterCopy, _
        CriteriaRange:=Sheets("Total").Range("N1:N2"), _
        CopyToRange:=Sheets("Mariner").Range("A1"), _
        Unique:=False
    Sheets("Total").Range("A:A:I:I").AdvancedFilter _
        Action:=xlFilterCopy, _
        CriteriaRange:=Sheets("Total").Range("O1:O2"), _
        CopyToRange:=Sheets("Metro").Range("A1"), _
        Unique:=False
    Sheets("Total").Range("A:A:I:I").AdvancedFilter _
        Action:=xlFilterCopy, _
        CriteriaRange:=Sheets("Total").Range("P1:P2"), _
        CopyToRange:=Sheets("North").Range("A1"), _
        Unique:=False
    Sheets("Total").Range("A:A:I:I").AdvancedFilter _
        Action:=xlFilterCopy, _
        CriteriaRange:=Sheets("Total").Range("Q1:Q2"), _
        CopyToRange:=Sheets("NorthWest").Range("A1"), _
        Unique:=False
    Sheets("Total").Range("A:A:I:I").AdvancedFilter _
        Action:=xlFilterCopy, _
        CriteriaRange:=Sheets("Total").Range("R1:R2"), _
        CopyToRange:=Sheets("TeleCenter").Range("A1"), _
        Unique:=False
    Sheets("Total").Range("A:A:I:I").AdvancedFilter _
        Action:=xlFilterCopy, _
        CriteriaRange:=Sheets("Total").Range("S1:S2"), _
        CopyToRange:=Sheets("West").Range("A1"), _
        Unique:=False
    Sheets("Total").Range("A:A:I:I").AdvancedFilter _
        Action:=xlFilterCopy, _
        CriteriaRange:=Sheets("Total").Range("T1:T2"), _
        CopyToRange:=Sheets("ENM").Range("A1"), _
        Unique:=False
    Sheets("Total").Range("A:A:I:I").AdvancedFilter _
        Action:=xlFilterCopy, _
        CriteriaRange:=Sheets("Total").Range("U1:U2"), _
        CopyToRange:=Sheets("Fallriver").Range("A1"), _
        Unique:=False
End Sub

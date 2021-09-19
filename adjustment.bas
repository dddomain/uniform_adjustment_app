Option Explicit

Sub Planning()

    'テーブルを変数に格納
    Dim UniformList As ListObject: Set UniformList = Sheet1.ListObjects(1)
    
    'Dim uni As Uniform: Set uni = New Uniform
    'uni.Initialize UniformList.ListRows(1).Range
    
    With UniformList
        Dim a As integer
        Dim b As integer
        Dim c As integer

        For a = 2 To .ListRows.Count + 1
            If .Range(a, 6).Value = "不足" And .Range(a, 11) = Empty Then

                For b = 2 To .ListRows.Count + 1
                    '早期離脱しつつ検索
                    If _
                    .Range(b, 11).Value = Empty And _
                    .Range(b, 5).Value <> .Range(a, 5).Value And _
                    .Range(b, 2).Value = .Range(a, 2).Value And _
                    .Range(b, 3).Value = .Range(a, 3).Value And _
                    .Range(b, 4).Value = .Range(a, 4).Value _
                    Then
                    
                        .Range(a, 11).Value = .Range(b, 5).Value + "から"
                        .Range(b, 11).Value = .Range(a, 5).Value + "へ"
                        .Range(a, 12).Value = a
                        .Range(b, 12).Value = a
                        Exit For
                    End If
                Next b
            End If
        Next a
        
        
        For c = 2 To .ListRows.Count + 1
        
            If .Range(c, 6).Value = "不足" And .Range(c, 11) = Empty Then
                .Range(c, 11).Value = "購入"
            ElseIf .Range(c, 6).Value = "余剰" And .Range(c, 11) = Empty Then
                .Range(c, 11).Value = "保留"
            End If
        Next c
        
    End With

End Sub



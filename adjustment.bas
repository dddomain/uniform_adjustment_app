Option Explicit



Sub MySub()

    Dim uni As Uniform: Set uni = New Uniform
    uni.Initialize Sheet1.ListObjects(1).ListRows(1).Range


End Sub

Sub search()

    'テーブルを変数に格納
    Dim UniformList As ListObject: Set UniformList = Sheet1.ListObjects(1)
    
    'Dim uni As Uniform: Set uni = New Uniform
    'uni.Initialize UniformList.ListRows(1).Range
    
    With UniformList
        Dim a As Long
        Dim b As Long
        Dim c As Long

        For a = 2 To .ListRows.Count + 1
            If .Range(a, 6).Value = "不足" And .Range(a, 11) = Empty Then
                'MsgBox .Range(a, 5) + "は不足しています。"

                For b = 2 To .ListRows.Count + 1
                    '早期離脱しつつ検索
                    If _
                        .Range(b, 5).Value <> .Range(a, 5).Value And _
                        .Range(b, 2).Value = .Range(a, 2).Value And _
                        .Range(b, 3).Value = .Range(a, 3).Value And _
                        .Range(b, 4).Value = .Range(a, 4).Value _
                    Then
                    
                    'MsgBox .Range(a, 5) + "には" + .Range(b, 5) + "から移動します。"
                    .Range(a, 11).Value = .Range(b, 5).Value + "から"
                    .Range(b, 11).Value = .Range(a, 5).Value + "へ"
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
    '一行目から順に、EODが不足 AND PassがNull の行を探す
    

End Sub

Sub auto_open()  'ブックの起動と同時に動作

If Time > TimeValue("17:57:00") And Time < TimeValue("18:03:00") Then '5時57分～6時3分の間に起動するとマクロが動作


For I = 2 To 1 Step -1  'シート1、2を走査するための繰り返し文

ThisWorkbook.Worksheets(I).Activate  'シートIをアクティブに

Call 試作1号  'プロシージャ試作1号を呼び出し

Next I


End If

End Sub


///////////////////////////////////////////////////////////////////////////////////////////


Sub auto_open()


If Time > TimeValue("17:57:00") And Time < TimeValue("18:03:00") And _
   Time > TimeValue("11:57:00") And Time < TimeValue("12:03:00") And _
   Time > TimeValue("14:57:00") And Time < TimeValue("15:03:00") Then '5時57分～6時3分の間に起動するとマクロが動作


For I = 2 To 1 Step -1  'シート1、2を走査するための繰り返し文

ThisWorkbook.Worksheets(I).Activate  'シートIをアクティブに

Call main処理  'プロシージャ試作1号を呼び出し

Next I


End If


End Sub

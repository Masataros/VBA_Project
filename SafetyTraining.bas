Sub main処理()


Dim today As Date
Dim I As Integer
Dim cnt As Integer
 
 
 today = Date    '変数todayに今日の日付を代入


For I = 2 To 100 Step 1  'シートを走査するための繰り返し文

 Dim oApp As New Outlook.Application      'outlookを起動
 Dim oItem As Outlook.MailItem            'メールのオプションを作成
 Set oItem = oApp.CreateItem(olMailItem)  'オブジェクト変数にoutlookと互換性を持たせる


 If DateDiff("d", Weekday(today), vbSunday) = 0 And DateDiff("d", today, Cells(I, 5)) = 2 And Cells(I, 8) = 0 Then  '土曜日の時の処理(2日前を走査)

    oItem.To = Cells(I, 3)                 'アドレス
    oItem.Subject = "安全講習について"     '件名
    oItem.Importance = olImportanceNormal  'メールの重要度設定
    oItem.Body = ""
    
    oItem.Send  '送信する場合
    
   ' oItem.Display  '送信せずに画面を表示する場合
    
    Cells(I, 6).Value = 1  'メールを送ったデータに確認印をする
    
   Set oItem = Nothing     'メールオプションを消す
   Set oApp = Nothing      'outlookを閉じる
    
  End If
    
 
 If DateDiff("d", today, Cells(I, 5)) = 1 And Cells(I, 8) = 0 Then '平日の時の走査(1日前を走査)

    oItem.To = Cells(I, 3)                 'アドレス
    oItem.Subject = "安全講習について"     '件名
    oItem.Importance = olImportanceNormal  'メールの重要度
    oItem.Body = ""
    
    oItem.Send  '送信する場合
    
   ' oItem.Display  '送信せずに画面を表示する場合
    
    Cells(I, 8).Value = 1  'メールを送ったデータに確認印をする
    
   Set oItem = Nothing     'メールオプションを消す
   Set oApp = Nothing      'outlookを閉じる

  End If

Next I

ActiveWorkbook.Save  'シートを保存


End Sub

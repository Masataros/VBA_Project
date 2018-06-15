Sub メール送信()
    
    Dim I As Integer
    Dim c As Integer
    
    c = 0
  
For I = 7 To 98 Step 1 'シートを走査するための繰り返し文

 Dim oApp As New Outlook.Application      'outlookを起動
 Dim oItem As Outlook.MailItem            'メールのオプションを作成
 Set oItem = oApp.CreateItem(olMailItem)  'オブジェクト変数にoutlookと互換性を持たせる

    
    
    If Cells(I, 16) = 0 Then
    
        If Cells(I, 1) = 2 Then
        
            oItem.To = Cells(I, 10)
            oItem.Subject = "ものづくりクリポ抽選結果のお知らせ"
    
            oItem.Body = Cells(I, 5) & "様　　" & "当選番号　" & Cells(I, 2) & vbNewLine & _
            Cells(I + 1, 5) & "様　　" & "当選番号　" & Cells(I + 1, 2) & vbNewLine & _
             "この度申込みいただいた「ものづくりクリポ2016」につきまして、抽選の結果、ご当選となりましたのでお知らせいたします。" _
            & vbNewLine & "こちらのメールは当日まで大切に保存してください。" & vbNewLine & "「ものづくりクリポ2016」" & vbNewLine & "■日時：11月6日（日）13時～15時" & vbNewLine & "■会場：大阪産業大学16号館6階16606教室" _
            & vbNewLine & "・当日は12時30分より受付を開始しますので、イベント開始10分前までに受付を完了してください。" & vbNewLine & "・受付の際に当選番号と名前をスタッフに申し出て、名札を貰ってください。" _
            & vbNewLine & "・万が一、参加出来ない場合はお手数ですが下記連絡先までご連絡下さい。" & vbNewLine & "スタッフ一同、心よりお待ちしています。"
    
            ' oItem.Send  '送信する場合
    
            oItem.Display  '送信せずに画面を表示する場合は、
    
            Set oItem = Nothing
            Set oApp = Nothing
            
            Cells(I, 16).Value = 1  'メールを送ったデータに確認印をする
            Cells(I + 1, 16).Value = 1 'メールを送ったデータに確認印をする
            
            I = I + 1
            

            
            
         ElseIf Cells(I, 1) = 3 Then
        
            oItem.To = Cells(I, 10)
            oItem.Subject = "ものづくりクリポ抽選結果のお知らせ"
    
            oItem.Body = Cells(I, 5) & "様　　" & "当選番号　" & Cells(I, 2) & vbNewLine & _
            Cells(I + 1, 5) & "様　　" & "当選番号　" & Cells(I + 1, 2) & vbNewLine & _
            Cells(I + 2, 5) & "様　　" & "当選番号　" & Cells(I + 2, 2) & vbNewLine & _
             "この度申込みいただいた「ものづくりクリポ2016」につきまして、抽選の結果、ご当選となりましたのでお知らせいたします。" _
            & vbNewLine & "こちらのメールは当日まで大切に保存してください。" & vbNewLine & "「ものづくりクリポ2016」" & vbNewLine & "■日時：11月6日（日）13時～15時" & vbNewLine & "■会場：大阪産業大学16号館6階16606教室" _
            & vbNewLine & "・当日は12時30分より受付を開始しますので、イベント開始10分前までに受付を完了してください。" & vbNewLine & "・受付の際に当選番号と名前をスタッフに申し出て、名札を貰ってください。" _
            & vbNewLine & "・万が一、参加出来ない場合はお手数ですが下記連絡先までご連絡下さい。" & vbNewLine & "スタッフ一同、心よりお待ちしています。"
    
            ' oItem.Send  '送信する場合
    
            oItem.Display  '送信せずに画面を表示する場合は、
    
            Set oItem = Nothing
            Set oApp = Nothing
            
            Cells(I, 16).Value = 1  'メールを送ったデータに確認印をする
            Cells(I + 1, 16).Value = 1 'メールを送ったデータに確認印をする
            Cells(I + 2, 16).Value = 1 'メールを送ったデータに確認印をする
            
            I = I + 2
            
        
           
           
           ElseIf Cells(I, 1) = 30 Then
        
            oItem.To = Cells(I, 10)
            oItem.Subject = "ものづくりクリポ抽選結果のお知らせ"
    
            oItem.Body = Cells(I, 5) & "様　　" & "当選番号　" & Cells(I, 2) & vbNewLine & _
            Cells(I + 1, 5) & "様　　" & "当選番号　" & Cells(I + 1, 2) & vbNewLine & _
            Cells(95, 5) & "様　　" & "当選番号　" & Cells(95, 2) & vbNewLine & _
             "この度申込みいただいた「ものづくりクリポ2016」につきまして、抽選の結果、ご当選となりましたのでお知らせいたします。" _
            & vbNewLine & "こちらのメールは当日まで大切に保存してください。" & vbNewLine & "「ものづくりクリポ2016」" & vbNewLine & "■日時：11月6日（日）13時～15時" & vbNewLine & "■会場：大阪産業大学16号館6階16606教室" _
            & vbNewLine & "・当日は12時30分より受付を開始しますので、イベント開始10分前までに受付を完了してください。" & vbNewLine & "・受付の際に当選番号と名前をスタッフに申し出て、名札を貰ってください。" _
            & vbNewLine & "・万が一、参加出来ない場合はお手数ですが下記連絡先までご連絡下さい。" & vbNewLine & "スタッフ一同、心よりお待ちしています。"
    
            ' oItem.Send  '送信する場合
    
            oItem.Display  '送信せずに画面を表示する場合は、
    
            Set oItem = Nothing
            Set oApp = Nothing
            
            Cells(I, 16).Value = 1  'メールを送ったデータに確認印をする
            Cells(I + 1, 16).Value = 1 'メールを送ったデータに確認印をする
            Cells(95, 16).Value = 1 'メールを送ったデータに確認印をする
            
            I = I + 2
            
           
           
           
           ElseIf Cells(I, 1) = 0 Then
        
            oItem.To = Cells(I, 10)
            oItem.Subject = "ものづくりクリポ抽選結果のお知らせ"
    
            oItem.Body = Cells(I, 5) & "様　　" & "当選番号　" & Cells(I, 2) & vbNewLine & _
             "この度申込みいただいた「ものづくりクリポ2016」につきまして、抽選の結果、ご当選となりましたのでお知らせいたします。" _
            & vbNewLine & "こちらのメールは当日まで大切に保存してください。" & vbNewLine & "「ものづくりクリポ2016」" & vbNewLine & "■日時：11月6日（日）13時～15時" & vbNewLine & "■会場：大阪産業大学16号館6階16606教室" _
            & vbNewLine & "・当日は12時30分より受付を開始しますので、イベント開始10分前までに受付を完了してください。" & vbNewLine & "・受付の際に当選番号と名前をスタッフに申し出て、名札を貰ってください。" _
            & vbNewLine & "・万が一、参加出来ない場合はお手数ですが下記連絡先までご連絡下さい。" & vbNewLine & "スタッフ一同、心よりお待ちしています。"
    
            ' oItem.Send  '送信する場合
    
            oItem.Display  '送信せずに画面を表示する場合は、
    
            Set oItem = Nothing
            Set oApp = Nothing
            
            Cells(I, 16).Value = 1  'メールを送ったデータに確認印をする
            
           
     
     End If
     
     If c > 8 Then
        Exit For
     End If
     
     c = c + 1
     
     
     
     End If
     
     
     
        
Next I


End Sub

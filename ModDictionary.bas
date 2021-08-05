Attribute VB_Name = "ModDictionary"
Option Explicit

'連想配列関連モジュール

Private Sub TestMakeDictFromArray()
    
    'テスト用の配列の定義
    Dim KaisoArray2D, ItemArray1D
    KaisoArray2D = Application.Transpose(Application.Transpose( _
                    Array(Array("愛媛県", "県庁所在地"), _
                    Array("愛媛県", "特産品"), _
                    Array("愛媛県", "特産品"), _
                    Array("愛媛県", "特産品"), _
                    Array("愛媛県", "ゆるキャラ"), _
                    Array("愛媛県", "ゆるキャラ"), _
                    Array("香川県", "県庁所在地"), _
                    Array("香川県", "特産品"), _
                    Array("香川県", "特産品"), _
                    Array("香川県", "特産品"), _
                    Array("香川県", "特産品"), _
                    Array("香川県", "ゆるキャラ"), _
                    Array("徳島県", "県庁所在地"), _
                    Array("徳島県", "特産品"), _
                    Array("徳島県", "ゆるキャラ"), _
                    Array("徳島県", "ゆるキャラ"), _
                    Array("高知県", "県庁所在地"), _
                    Array("高知県", "特産品"), _
                    Array("高知県", "特産品"), _
                    Array("高知県", "ゆるキャラ")) _
                    ))
    ItemArray1D = Application.Transpose(Application.Transpose( _
            Array("松山市", "みかん", "タオル", "真珠", "バリィさん", "みきゃん", "高松市", "うどん", "醤油", "オリーブ", "素麺", "うどん脳", "徳島市", "すだち", "ししゃも猫", "すだちくん", "高松市", "かつお", "酒", "しんじょうくん") _
            ))
    
    '配列の中身を表示確認
    Call DPH(KaisoArray2D, , "KaisoArray2D")
    Call DPH(ItemArray1D, , "ItemArray1D")
    
    '階層型連想配列作成
    Dim OutputDict As Object
    Set OutputDict = MakeDictFromArray(KaisoArray2D, ItemArray1D)
    
    '出力テスト
    Debug.Print OutputDict("愛媛県")("特産品")(2) '→タオル
    Debug.Print OutputDict("愛媛県")("ゆるキャラ")(1) '→バリィさん
    Debug.Print OutputDict("徳島県")("ゆるキャラ")(1) '→ししゃも猫
    Debug.Print OutputDict("香川県")("県庁所在地")(1) '→高松市

End Sub

Function MakeDictFromArrayWithItem(KaisoArray2D, KeyArray1D, ItemArray2D)
    
    '引数チェック
    Call CheckArray2D(KaisoArray2D, "KaisoArray2D") '2次元配列かチェック
    Call CheckArray2DStart1(KaisoArray2D, "KaisoArray2D") '要素の開始番号が1かチェック
    Call CheckArray1D(KeyArray1D, "KeyArray1D") '1次元配列かチェック
    Call CheckArray1DStart1(KeyArray1D, "KeyArray1D") '要素の開始番号が1かチェック
    Call CheckArray2D(ItemArray2D, "ItemArray2D") '2次元配列かチェック
    Call CheckArray2DStart1(ItemArray2D, "ItemArray2D") '要素の開始番号が1かチェック
    
    If UBound(KaisoArray2D, 1) <> UBound(ItemArray2D, 1) Then
        MsgBox ("「KaisoArray2D」と「ItemArray2D」の縦要素数を一致させてください")
        Stop
        End
    End If
    
    If UBound(KeyArray1D, 1) <> UBound(ItemArray2D, 2) Then
        MsgBox ("「KeyArray1D」の要素数と「ItemArray2D」の横要素数を一致させてください")
        Stop
        End
    End If
    
    '計算処理
    Dim DictArray
    DictArray = 配列から連想配列リスト作成(ItemArray2D, KeyArray1D)
    Dim Output As Object
    Set Output = 配列から階層型連想配列作成(KaisoArray2D, DictArray)
    
    Set MakeDictFromArray = Output
    
End Function

Function MakeDictFromArray(KaisoArray2D, ItemArray1D)

    '引数チェック
    Call CheckArray2D(KaisoArray2D, "KaisoArray2D") '2次元配列かチェック
    Call CheckArray2DStart1(KaisoArray2D, "KaisoArray2D") '要素の開始番号が1かチェック
    Call CheckArray1D(ItemArray1D, "ItemArray1D") '1次元配列かチェック
    Call CheckArray1DStart1(ItemArray1D, "ItemArray1D") '要素の開始番号が1かチェック
    If UBound(KaisoArray2D, 1) <> UBound(ItemArray1D, 1) Then
        MsgBox ("「KaisoArray2D」と「ItemArray1D」の縦要素数を一致させてください")
        Stop
        End
    End If
    
    Dim I&, J&, K&, M&, N& '数え上げ用(Long型)
    N = UBound(KaisoArray2D, 1)
    M = UBound(KaisoArray2D, 2)
    
    Dim Output As Object
    Set Output = CreateObject("Scripting.Dictionary")
    
    Dim TmpUnique, TmpValue, TmpArray, TmpItiArray, TmpDict As Object, TmpItemArray1D
    Dim Dummy
    TmpUnique = 配列の指定列のユニーク値リスト取得(KaisoArray2D, 1)
    For Each TmpValue In TmpUnique
        Dummy = 配列の先頭列値に一致するもののみ配列抽出(KaisoArray2D, TmpValue)
        TmpArray = Dummy(1)
        TmpItiArray = Dummy(2)
        TmpItemArray1D = 一次元配列の指定範囲取得(ItemArray1D, TmpItiArray)
        
        If M = 1 Then
            Output.Add TmpValue, TmpItemArray1D
        Else
            Set TmpDict = 配列から階層型連想配列作成(TmpArray, TmpItemArray1D)
            Output.Add TmpValue, TmpDict
        End If
    Next
        
    Set MakeDictFromArray = Output
        
End Function

Private Function 配列から連想配列リスト作成(ItemArray2D, KeyArray1D)
    
    Dim TmpDict As Object
    Dim I&, J&, K&, M&, N& '数え上げ用(Long型)
    N = UBound(ItemArray2D, 1)
    M = UBound(ItemArray2D, 2)
    Dim Output
    ReDim Output(1 To N)
    For I = 1 To N
        Set TmpDict = CreateObject("Scripting.Dictionary")
        For J = 1 To M
            TmpDict.Add KeyArray1D(J), ItemArray2D(I, J)
        Next J
        
        Set Output(I) = TmpDict
    Next I
    
    配列から連想配列リスト作成 = Output

End Function

Private Function 配列から階層型連想配列作成(KaisoArray2D, ItemArray1D)
    
    Dim I&, J&, K&, M&, N& '数え上げ用(Long型)
    N = UBound(KaisoArray2D, 1)
    M = UBound(KaisoArray2D, 2)
    
    Dim Output As Object
    Set Output = CreateObject("Scripting.Dictionary")
    
    Dim TmpUnique, TmpValue, TmpArray, TmpItiArray, TmpDict As Object, TmpItemArray1D
    Dim Dummy
    TmpUnique = 配列の指定列のユニーク値リスト取得(KaisoArray2D, 1)
    For Each TmpValue In TmpUnique
        Dummy = 配列の先頭列値に一致するもののみ配列抽出(KaisoArray2D, TmpValue)
        TmpArray = Dummy(1)
        TmpItiArray = Dummy(2)
        TmpItemArray1D = 一次元配列の指定範囲取得(ItemArray1D, TmpItiArray)
        
        If M = 1 Then
            Output.Add TmpValue, TmpItemArray1D
        Else
            Set TmpDict = 配列から階層型連想配列作成(TmpArray, TmpItemArray1D)
            Output.Add TmpValue, TmpDict
        End If
    Next
        
    Set 配列から階層型連想配列作成 = Output
        
End Function

Private Function 配列の先頭列値に一致するもののみ配列抽出(InputArray2D, InputValue)
    
    Dim I&, J&, K&, M&, N& '数え上げ用(Long型)
    N = UBound(InputArray2D, 1)
    M = UBound(InputArray2D, 2)
    Dim Count&
    K = 0
    For I = 1 To N
        If InputArray2D(I, 1) = InputValue Then
            K = K + 1
        End If
    Next I
    
    Count = K
    Dim OutputArray, ItiArray
    
    If M > 1 Then
        ReDim OutputArray(1 To Count, 1 To M - 1)
        K = 0
        For I = 1 To N
            If InputArray2D(I, 1) = InputValue Then
                K = K + 1
                For J = 1 To M - 1
                    OutputArray(K, J) = InputArray2D(I, J + 1)
                Next J
            End If
        Next I
    End If
    
    ReDim ItiArray(1 To Count)
    K = 0
    For I = 1 To N
        If InputArray2D(I, 1) = InputValue Then
            K = K + 1
            ItiArray(K) = I
        End If
    Next I
        
    
    Dim Output(1 To 2)
    Output(1) = OutputArray
    Output(2) = ItiArray
    
    配列の先頭列値に一致するもののみ配列抽出 = Output
    
End Function

Private Function 配列の指定列のユニーク値リスト取得(InputArray, Col%)
    
    Dim TmpDict As Object
    Set TmpDict = CreateObject("Scripting.Dictionary")
    Dim I&, J&, K&, M&, N& '数え上げ用(Long型)
    N = UBound(InputArray, 1)
    
    For I = 1 To N
        If TmpDict.Exists(InputArray(I, Col)) = False Then
            TmpDict.Add InputArray(I, Col), ""
        End If
    Next I
    
    Dim Output
    Output = TmpDict.Keys
    Output = Application.Transpose(Application.Transpose(Output))
    
    配列の指定列のユニーク値リスト取得 = Output

End Function

Private Function 一次元配列の指定範囲取得(Array1D, ItiArray1D)
    
    Dim I&, J&, K&, M&, N& '数え上げ用(Long型)
    N = UBound(ItiArray1D, 1)
    Dim TmpIti&
    
    Dim Output
    ReDim Output(1 To N)
    For I = 1 To N
        TmpIti = ItiArray1D(I)
        If IsObject(Array1D(TmpIti)) = True Then
            Set Output(I) = Array1D(TmpIti)
        Else
            Output(I) = Array1D(TmpIti)
        End If
    Next I
    
    一次元配列の指定範囲取得 = Output
    
End Function


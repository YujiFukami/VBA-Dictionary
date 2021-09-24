Attribute VB_Name = "ModMakeDictFromArrayWithItem"
Option Explicit

'MakeDictFromArrayWithItem               ・・・元場所：FukamiAddins3.ModDictionary
'MakeDictFromArray1D                     ・・・元場所：FukamiAddins3.ModDictionary
'CheckArray1D                            ・・・元場所：FukamiAddins3.ModDictionary
'CheckArray1DStart1                      ・・・元場所：FukamiAddins3.ModDictionary
'配列から連想配列リスト作成              ・・・元場所：FukamiAddins3.ModDictionary
'配列から階層型連想配列作成              ・・・元場所：FukamiAddins3.ModDictionary
'配列の先頭列値に一致するもののみ配列抽出・・・元場所：FukamiAddins3.ModDictionary
'配列の指定列のユニーク値リスト取得      ・・・元場所：FukamiAddins3.ModDictionary
'一次元配列の指定範囲取得                ・・・元場所：FukamiAddins3.ModDictionary
'CheckArray2D                            ・・・元場所：FukamiAddins3.ModDictionary
'CheckArray2DStart1                      ・・・元場所：FukamiAddins3.ModDictionary

'------------------------------


'連想配列関連モジュール
'------------------------------


Function MakeDictFromArrayWithItem(VerticalKeyArray2D, HorizontalKeyArray1D, ItemArray2D)
'二次元配列から連想配列を作成する。
'連想配列の末端要素がさらに連想配列（要素連想配列）となっていて、もととなるキー配列、アイテム配列を入力する。
'各配列の要素の開始番号は1とすること
'20210806作成

'VerticalKeyArray2D   :Keyとなる縦二次元配列
'HorizontalKeyArray1D :要素連想配列のキー 一次元配列
'ItemArray2D          :要素連想配列のアイテム 二次元配列


    '引数チェック
    Call CheckArray2D(VerticalKeyArray2D, "VerticalKeyArray2D") '2次元配列かチェック
    Call CheckArray2DStart1(VerticalKeyArray2D, "VerticalKeyArray2D") '要素の開始番号が1かチェック
    Call CheckArray1D(HorizontalKeyArray1D, "HorizontalKeyArray1D") '1次元配列かチェック
    Call CheckArray1DStart1(HorizontalKeyArray1D, "HorizontalKeyArray1D") '要素の開始番号が1かチェック
    Call CheckArray2D(ItemArray2D, "ItemArray2D") '2次元配列かチェック
    Call CheckArray2DStart1(ItemArray2D, "ItemArray2D") '要素の開始番号が1かチェック
    
    If UBound(VerticalKeyArray2D, 1) <> UBound(ItemArray2D, 1) Then
        MsgBox ("「VerticalKeyArray2D」と「ItemArray2D」の縦要素数を一致させてください")
        Stop
        End
    End If
    
    If UBound(HorizontalKeyArray1D, 1) <> UBound(ItemArray2D, 2) Then
        MsgBox ("「HorizontalKeyArray1D」の要素数と「ItemArray2D」の横要素数を一致させてください")
        Stop
        End
    End If
    
    '計算処理
    Dim DictArray
    DictArray = 配列から連想配列リスト作成(ItemArray2D, HorizontalKeyArray1D)
    Dim Output As Object
    
    Dim VerticalKeyArray1D
    If UBound(VerticalKeyArray2D, 2) = 1 Then
        '縦二次元配列の二次元要素数が1だったら一次元配列に変換して処理
        VerticalKeyArray1D = Application.Transpose(VerticalKeyArray2D)
        Set Output = MakeDictFromArray1D(VerticalKeyArray1D, DictArray)
    Else
        Set Output = 配列から階層型連想配列作成(VerticalKeyArray2D, DictArray)
    End If
    
    Set MakeDictFromArrayWithItem = Output
    
End Function

Private Function MakeDictFromArray1D(KeyArray1D, ItemArray1D)
'配列から連想配列を作成する
'各配列の要素の開始番号は1とすること
'20210806作成

'KeyArray1D   ：Keyが入った一次元配列
'ItemArray1D  ：Itemが入った一次元配列

    '引数チェック
    Call CheckArray1D(KeyArray1D, "KeyArray1D") '2次元配列かチェック
    Call CheckArray1DStart1(KeyArray1D, "KeyArray1D") '要素の開始番号が1かチェック
    Call CheckArray1D(ItemArray1D, "ItemArray1D") '1次元配列かチェック
    Call CheckArray1DStart1(ItemArray1D, "ItemArray1D") '要素の開始番号が1かチェック
    If UBound(KeyArray1D, 1) <> UBound(ItemArray1D, 1) Then
        MsgBox ("「KeyArray1D」と「ItemArray1D」の縦要素数を一致させてください")
        Stop
        End
    End If
    
    Dim I&, J&, K&, M&, N& '数え上げ用(Long型)
    N = UBound(KeyArray1D, 1)
    
    Dim Output As Object
    Set Output = CreateObject("Scripting.Dictionary")
    
    Dim TmpKey$
    
    For I = 1 To N
        TmpKey = KeyArray1D(I)
        If Output.Exists(TmpKey) = False Then
            Output.Add TmpKey, ItemArray1D(I)
        End If
    Next I
    
    Set MakeDictFromArray1D = Output
        
End Function

Private Sub CheckArray1D(InputArray, Optional HairetuName$ = "配列")
'入力配列が1次元配列かどうかチェックする
'20210804

    Dim Dummy%
    On Error Resume Next
    Dummy = UBound(InputArray, 2)
    On Error GoTo 0
    If Dummy <> 0 Then
        MsgBox (HairetuName & "は1次元配列を入力してください")
        Stop
        Exit Sub '入力元のプロシージャを確認するために抜ける
    End If

End Sub

Private Sub CheckArray1DStart1(InputArray, Optional HairetuName$ = "配列")
'入力1次元配列の開始番号が1かどうかチェックする
'20210804

    If LBound(InputArray, 1) <> 1 Then
        MsgBox (HairetuName & "の開始要素番号は1にしてください")
        Stop
        Exit Sub '入力元のプロシージャを確認するために抜ける
    End If

End Sub

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
            If UBound(TmpItemArray1D, 1) = 1 And IsObject(TmpItemArray1D(1)) = True Then
                Output.Add TmpValue, TmpItemArray1D(1)
            Else
                Output.Add TmpValue, TmpItemArray1D
            End If
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

Private Sub CheckArray2D(InputArray, Optional HairetuName$ = "配列")
'入力配列が2次元配列かどうかチェックする
'20210804

    Dim Dummy2%, Dummy3%
    On Error Resume Next
    Dummy2 = UBound(InputArray, 2)
    Dummy3 = UBound(InputArray, 3)
    On Error GoTo 0
    If Dummy2 = 0 Or Dummy3 <> 0 Then
        MsgBox (HairetuName & "は2次元配列を入力してください")
        Stop
        Exit Sub '入力元のプロシージャを確認するために抜ける
    End If

End Sub

Private Sub CheckArray2DStart1(InputArray, Optional HairetuName$ = "配列")
'入力2次元配列の開始番号が1かどうかチェックする
'20210804

    If LBound(InputArray, 1) <> 1 Or LBound(InputArray, 2) <> 1 Then
        MsgBox (HairetuName & "の開始要素番号は1にしてください")
        Stop
        Exit Sub '入力元のプロシージャを確認するために抜ける
    End If

End Sub



Attribute VB_Name = "ModMakeDictFromArray1D"
Option Explicit

'MakeDictFromArray1D・・・元場所：FukamiAddins3.ModDictionary
'CheckArray1D       ・・・元場所：FukamiAddins3.ModDictionary
'CheckArray1DStart1 ・・・元場所：FukamiAddins3.ModDictionary

'------------------------------


'連想配列関連モジュール
'------------------------------


Function MakeDictFromArray1D(KeyArray1D, ItemArray1D)
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



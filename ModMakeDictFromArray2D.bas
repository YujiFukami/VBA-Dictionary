Attribute VB_Name = "ModMakeDictFromArray2D"
Option Explicit

'MakeDictFromArray2D                     ・・・元場所：FukamiAddins3.ModDictionary
'CheckArray2D                            ・・・元場所：FukamiAddins3.ModDictionary
'CheckArray2DStart1                      ・・・元場所：FukamiAddins3.ModDictionary
'CheckArray1D                            ・・・元場所：FukamiAddins3.ModDictionary
'CheckArray1DStart1                      ・・・元場所：FukamiAddins3.ModDictionary
'配列の指定列のユニーク値リスト取得      ・・・元場所：FukamiAddins3.ModDictionary
'配列の先頭列値に一致するもののみ配列抽出・・・元場所：FukamiAddins3.ModDictionary
'一次元配列の指定範囲取得                ・・・元場所：FukamiAddins3.ModDictionary
'配列から階層型連想配列作成              ・・・元場所：FukamiAddins3.ModDictionary



Public Function MakeDictFromArray2D(KaisoArray2D, ItemArray1D)
'配列から連想配列を作成する。
'二次元配列から階層状態を取得し、複数のKeyとして扱う
'各配列の要素の開始番号は1とすること
'20210806作成

'KaisoArray2D   ：二次元配列。階層状態になっていること。
'ItemArray1D    ：要素が入っている配列　一次元配列

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
    
    Dim M As Long
    Dim N As Long
    N = UBound(KaisoArray2D, 1)
    M = UBound(KaisoArray2D, 2)
    
    Dim Output As Object
    Set Output = CreateObject("Scripting.Dictionary")
    
    Dim TmpUnique
    Dim TmpValue
    Dim TmpArray
    Dim TmpItiArray
    Dim TmpDict       As Object
    Dim TmpItemArray1D
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
        
    Set MakeDictFromArray2D = Output
        
End Function

Private Sub CheckArray2D(InputArray, Optional HairetuName As String = "配列")
'入力配列が2次元配列かどうかチェックする
'20210804

    Dim Dummy2 As Integer
    Dim Dummy3 As Integer
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

Private Sub CheckArray2DStart1(InputArray, Optional HairetuName As String = "配列")
'入力2次元配列の開始番号が1かどうかチェックする
'20210804

    If LBound(InputArray, 1) <> 1 Or LBound(InputArray, 2) <> 1 Then
        MsgBox (HairetuName & "の開始要素番号は1にしてください")
        Stop
        Exit Sub '入力元のプロシージャを確認するために抜ける
    End If

End Sub

Private Sub CheckArray1D(InputArray, Optional HairetuName As String = "配列")
'入力配列が1次元配列かどうかチェックする
'20210804

    Dim Dummy As Integer
    On Error Resume Next
    Dummy = UBound(InputArray, 2)
    On Error GoTo 0
    If Dummy <> 0 Then
        MsgBox (HairetuName & "は1次元配列を入力してください")
        Stop
        Exit Sub '入力元のプロシージャを確認するために抜ける
    End If

End Sub

Private Sub CheckArray1DStart1(InputArray, Optional HairetuName As String = "配列")
'入力1次元配列の開始番号が1かどうかチェックする
'20210804

    If LBound(InputArray, 1) <> 1 Then
        MsgBox (HairetuName & "の開始要素番号は1にしてください")
        Stop
        Exit Sub '入力元のプロシージャを確認するために抜ける
    End If

End Sub

Private Function 配列の指定列のユニーク値リスト取得(InputArray, Col As Integer)
    
    Dim TmpDict As Object
    Dim I       As Long
    Dim N       As Long
    Set TmpDict = CreateObject("Scripting.Dictionary")
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

Private Function 配列の先頭列値に一致するもののみ配列抽出(InputArray2D, InputValue)
    
    Dim I     As Long
    Dim J     As Long
    Dim K     As Long
    Dim M     As Long
    Dim N     As Long
    Dim Count As Long
    N = UBound(InputArray2D, 1)
    M = UBound(InputArray2D, 2)
    K = 0
    For I = 1 To N
        If InputArray2D(I, 1) = InputValue Then
            K = K + 1
        End If
    Next I
    
    Count = K
    Dim OutputArray
    Dim ItiArray
    
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

Private Function 一次元配列の指定範囲取得(Array1D, ItiArray1D)
    
    Dim I      As Long
    Dim N      As Long
    Dim TmpIti As Long
    Dim Output
    N = UBound(ItiArray1D, 1)
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

Private Function 配列から階層型連想配列作成(KaisoArray2D, ItemArray1D)
    
    Dim I      As Long
    Dim J      As Long
    Dim K      As Long
    Dim M      As Long
    Dim N      As Long
    Dim Output As Object
    N = UBound(KaisoArray2D, 1)
    M = UBound(KaisoArray2D, 2)
    
    Set Output = CreateObject("Scripting.Dictionary")
    
    Dim TmpUnique
    Dim TmpValue
    Dim TmpArray
    Dim TmpItiArray
    Dim TmpDict       As Object
    Dim TmpItemArray1D
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



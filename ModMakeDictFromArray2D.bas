Attribute VB_Name = "ModMakeDictFromArray2D"
Option Explicit

'MakeDictFromArray2D                     �E�E�E���ꏊ�FFukamiAddins3.ModDictionary
'CheckArray2D                            �E�E�E���ꏊ�FFukamiAddins3.ModDictionary
'CheckArray2DStart1                      �E�E�E���ꏊ�FFukamiAddins3.ModDictionary
'CheckArray1D                            �E�E�E���ꏊ�FFukamiAddins3.ModDictionary
'CheckArray1DStart1                      �E�E�E���ꏊ�FFukamiAddins3.ModDictionary
'�z��̎w���̃��j�[�N�l���X�g�擾      �E�E�E���ꏊ�FFukamiAddins3.ModDictionary
'�z��̐擪��l�Ɉ�v������̂̂ݔz�񒊏o�E�E�E���ꏊ�FFukamiAddins3.ModDictionary
'�ꎟ���z��̎w��͈͎擾                �E�E�E���ꏊ�FFukamiAddins3.ModDictionary
'�z�񂩂�K�w�^�A�z�z��쐬              �E�E�E���ꏊ�FFukamiAddins3.ModDictionary



Public Function MakeDictFromArray2D(KaisoArray2D, ItemArray1D)
'�z�񂩂�A�z�z����쐬����B
'�񎟌��z�񂩂�K�w��Ԃ��擾���A������Key�Ƃ��Ĉ���
'�e�z��̗v�f�̊J�n�ԍ���1�Ƃ��邱��
'20210806�쐬

'KaisoArray2D   �F�񎟌��z��B�K�w��ԂɂȂ��Ă��邱�ƁB
'ItemArray1D    �F�v�f�������Ă���z��@�ꎟ���z��

    '�����`�F�b�N
    Call CheckArray2D(KaisoArray2D, "KaisoArray2D") '2�����z�񂩃`�F�b�N
    Call CheckArray2DStart1(KaisoArray2D, "KaisoArray2D") '�v�f�̊J�n�ԍ���1���`�F�b�N
    Call CheckArray1D(ItemArray1D, "ItemArray1D") '1�����z�񂩃`�F�b�N
    Call CheckArray1DStart1(ItemArray1D, "ItemArray1D") '�v�f�̊J�n�ԍ���1���`�F�b�N
    If UBound(KaisoArray2D, 1) <> UBound(ItemArray1D, 1) Then
        MsgBox ("�uKaisoArray2D�v�ƁuItemArray1D�v�̏c�v�f������v�����Ă�������")
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
    TmpUnique = �z��̎w���̃��j�[�N�l���X�g�擾(KaisoArray2D, 1)
    For Each TmpValue In TmpUnique
        Dummy = �z��̐擪��l�Ɉ�v������̂̂ݔz�񒊏o(KaisoArray2D, TmpValue)
        TmpArray = Dummy(1)
        TmpItiArray = Dummy(2)
        TmpItemArray1D = �ꎟ���z��̎w��͈͎擾(ItemArray1D, TmpItiArray)
        
        If M = 1 Then
            Output.Add TmpValue, TmpItemArray1D
        Else
            Set TmpDict = �z�񂩂�K�w�^�A�z�z��쐬(TmpArray, TmpItemArray1D)
            Output.Add TmpValue, TmpDict
        End If
    Next
        
    Set MakeDictFromArray2D = Output
        
End Function

Private Sub CheckArray2D(InputArray, Optional HairetuName As String = "�z��")
'���͔z��2�����z�񂩂ǂ����`�F�b�N����
'20210804

    Dim Dummy2 As Integer
    Dim Dummy3 As Integer
    On Error Resume Next
    Dummy2 = UBound(InputArray, 2)
    Dummy3 = UBound(InputArray, 3)
    On Error GoTo 0
    If Dummy2 = 0 Or Dummy3 <> 0 Then
        MsgBox (HairetuName & "��2�����z�����͂��Ă�������")
        Stop
        Exit Sub '���͌��̃v���V�[�W�����m�F���邽�߂ɔ�����
    End If

End Sub

Private Sub CheckArray2DStart1(InputArray, Optional HairetuName As String = "�z��")
'����2�����z��̊J�n�ԍ���1���ǂ����`�F�b�N����
'20210804

    If LBound(InputArray, 1) <> 1 Or LBound(InputArray, 2) <> 1 Then
        MsgBox (HairetuName & "�̊J�n�v�f�ԍ���1�ɂ��Ă�������")
        Stop
        Exit Sub '���͌��̃v���V�[�W�����m�F���邽�߂ɔ�����
    End If

End Sub

Private Sub CheckArray1D(InputArray, Optional HairetuName As String = "�z��")
'���͔z��1�����z�񂩂ǂ����`�F�b�N����
'20210804

    Dim Dummy As Integer
    On Error Resume Next
    Dummy = UBound(InputArray, 2)
    On Error GoTo 0
    If Dummy <> 0 Then
        MsgBox (HairetuName & "��1�����z�����͂��Ă�������")
        Stop
        Exit Sub '���͌��̃v���V�[�W�����m�F���邽�߂ɔ�����
    End If

End Sub

Private Sub CheckArray1DStart1(InputArray, Optional HairetuName As String = "�z��")
'����1�����z��̊J�n�ԍ���1���ǂ����`�F�b�N����
'20210804

    If LBound(InputArray, 1) <> 1 Then
        MsgBox (HairetuName & "�̊J�n�v�f�ԍ���1�ɂ��Ă�������")
        Stop
        Exit Sub '���͌��̃v���V�[�W�����m�F���邽�߂ɔ�����
    End If

End Sub

Private Function �z��̎w���̃��j�[�N�l���X�g�擾(InputArray, Col As Integer)
    
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
    
    �z��̎w���̃��j�[�N�l���X�g�擾 = Output

End Function

Private Function �z��̐擪��l�Ɉ�v������̂̂ݔz�񒊏o(InputArray2D, InputValue)
    
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
    
    �z��̐擪��l�Ɉ�v������̂̂ݔz�񒊏o = Output
    
End Function

Private Function �ꎟ���z��̎w��͈͎擾(Array1D, ItiArray1D)
    
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
    
    �ꎟ���z��̎w��͈͎擾 = Output
    
End Function

Private Function �z�񂩂�K�w�^�A�z�z��쐬(KaisoArray2D, ItemArray1D)
    
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
    TmpUnique = �z��̎w���̃��j�[�N�l���X�g�擾(KaisoArray2D, 1)
    For Each TmpValue In TmpUnique
        Dummy = �z��̐擪��l�Ɉ�v������̂̂ݔz�񒊏o(KaisoArray2D, TmpValue)
        TmpArray = Dummy(1)
        TmpItiArray = Dummy(2)
        TmpItemArray1D = �ꎟ���z��̎w��͈͎擾(ItemArray1D, TmpItiArray)
        
        If M = 1 Then
            If UBound(TmpItemArray1D, 1) = 1 And IsObject(TmpItemArray1D(1)) = True Then
                Output.Add TmpValue, TmpItemArray1D(1)
            Else
                Output.Add TmpValue, TmpItemArray1D
            End If
        Else
            Set TmpDict = �z�񂩂�K�w�^�A�z�z��쐬(TmpArray, TmpItemArray1D)
            Output.Add TmpValue, TmpDict
        End If
    Next
        
    Set �z�񂩂�K�w�^�A�z�z��쐬 = Output
        
End Function



Attribute VB_Name = "ModMakeDictFromArrayWithItem"
Option Explicit

'MakeDictFromArrayWithItem               �E�E�E���ꏊ�FFukamiAddins3.ModDictionary
'MakeDictFromArray1D                     �E�E�E���ꏊ�FFukamiAddins3.ModDictionary
'CheckArray1D                            �E�E�E���ꏊ�FFukamiAddins3.ModDictionary
'CheckArray1DStart1                      �E�E�E���ꏊ�FFukamiAddins3.ModDictionary
'�z�񂩂�A�z�z�񃊃X�g�쐬              �E�E�E���ꏊ�FFukamiAddins3.ModDictionary
'�z�񂩂�K�w�^�A�z�z��쐬              �E�E�E���ꏊ�FFukamiAddins3.ModDictionary
'�z��̐擪��l�Ɉ�v������̂̂ݔz�񒊏o�E�E�E���ꏊ�FFukamiAddins3.ModDictionary
'�z��̎w���̃��j�[�N�l���X�g�擾      �E�E�E���ꏊ�FFukamiAddins3.ModDictionary
'�ꎟ���z��̎w��͈͎擾                �E�E�E���ꏊ�FFukamiAddins3.ModDictionary
'CheckArray2D                            �E�E�E���ꏊ�FFukamiAddins3.ModDictionary
'CheckArray2DStart1                      �E�E�E���ꏊ�FFukamiAddins3.ModDictionary

'------------------------------


'�A�z�z��֘A���W���[��
'------------------------------


Function MakeDictFromArrayWithItem(VerticalKeyArray2D, HorizontalKeyArray1D, ItemArray2D)
'�񎟌��z�񂩂�A�z�z����쐬����B
'�A�z�z��̖��[�v�f������ɘA�z�z��i�v�f�A�z�z��j�ƂȂ��Ă��āA���ƂƂȂ�L�[�z��A�A�C�e���z�����͂���B
'�e�z��̗v�f�̊J�n�ԍ���1�Ƃ��邱��
'20210806�쐬

'VerticalKeyArray2D   :Key�ƂȂ�c�񎟌��z��
'HorizontalKeyArray1D :�v�f�A�z�z��̃L�[ �ꎟ���z��
'ItemArray2D          :�v�f�A�z�z��̃A�C�e�� �񎟌��z��


    '�����`�F�b�N
    Call CheckArray2D(VerticalKeyArray2D, "VerticalKeyArray2D") '2�����z�񂩃`�F�b�N
    Call CheckArray2DStart1(VerticalKeyArray2D, "VerticalKeyArray2D") '�v�f�̊J�n�ԍ���1���`�F�b�N
    Call CheckArray1D(HorizontalKeyArray1D, "HorizontalKeyArray1D") '1�����z�񂩃`�F�b�N
    Call CheckArray1DStart1(HorizontalKeyArray1D, "HorizontalKeyArray1D") '�v�f�̊J�n�ԍ���1���`�F�b�N
    Call CheckArray2D(ItemArray2D, "ItemArray2D") '2�����z�񂩃`�F�b�N
    Call CheckArray2DStart1(ItemArray2D, "ItemArray2D") '�v�f�̊J�n�ԍ���1���`�F�b�N
    
    If UBound(VerticalKeyArray2D, 1) <> UBound(ItemArray2D, 1) Then
        MsgBox ("�uVerticalKeyArray2D�v�ƁuItemArray2D�v�̏c�v�f������v�����Ă�������")
        Stop
        End
    End If
    
    If UBound(HorizontalKeyArray1D, 1) <> UBound(ItemArray2D, 2) Then
        MsgBox ("�uHorizontalKeyArray1D�v�̗v�f���ƁuItemArray2D�v�̉��v�f������v�����Ă�������")
        Stop
        End
    End If
    
    '�v�Z����
    Dim DictArray
    DictArray = �z�񂩂�A�z�z�񃊃X�g�쐬(ItemArray2D, HorizontalKeyArray1D)
    Dim Output As Object
    
    Dim VerticalKeyArray1D
    If UBound(VerticalKeyArray2D, 2) = 1 Then
        '�c�񎟌��z��̓񎟌��v�f����1��������ꎟ���z��ɕϊ����ď���
        VerticalKeyArray1D = Application.Transpose(VerticalKeyArray2D)
        Set Output = MakeDictFromArray1D(VerticalKeyArray1D, DictArray)
    Else
        Set Output = �z�񂩂�K�w�^�A�z�z��쐬(VerticalKeyArray2D, DictArray)
    End If
    
    Set MakeDictFromArrayWithItem = Output
    
End Function

Private Function MakeDictFromArray1D(KeyArray1D, ItemArray1D)
'�z�񂩂�A�z�z����쐬����
'�e�z��̗v�f�̊J�n�ԍ���1�Ƃ��邱��
'20210806�쐬

'KeyArray1D   �FKey���������ꎟ���z��
'ItemArray1D  �FItem���������ꎟ���z��

    '�����`�F�b�N
    Call CheckArray1D(KeyArray1D, "KeyArray1D") '2�����z�񂩃`�F�b�N
    Call CheckArray1DStart1(KeyArray1D, "KeyArray1D") '�v�f�̊J�n�ԍ���1���`�F�b�N
    Call CheckArray1D(ItemArray1D, "ItemArray1D") '1�����z�񂩃`�F�b�N
    Call CheckArray1DStart1(ItemArray1D, "ItemArray1D") '�v�f�̊J�n�ԍ���1���`�F�b�N
    If UBound(KeyArray1D, 1) <> UBound(ItemArray1D, 1) Then
        MsgBox ("�uKeyArray1D�v�ƁuItemArray1D�v�̏c�v�f������v�����Ă�������")
        Stop
        End
    End If
    
    Dim I&, J&, K&, M&, N& '�����グ�p(Long�^)
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

Private Sub CheckArray1D(InputArray, Optional HairetuName$ = "�z��")
'���͔z��1�����z�񂩂ǂ����`�F�b�N����
'20210804

    Dim Dummy%
    On Error Resume Next
    Dummy = UBound(InputArray, 2)
    On Error GoTo 0
    If Dummy <> 0 Then
        MsgBox (HairetuName & "��1�����z�����͂��Ă�������")
        Stop
        Exit Sub '���͌��̃v���V�[�W�����m�F���邽�߂ɔ�����
    End If

End Sub

Private Sub CheckArray1DStart1(InputArray, Optional HairetuName$ = "�z��")
'����1�����z��̊J�n�ԍ���1���ǂ����`�F�b�N����
'20210804

    If LBound(InputArray, 1) <> 1 Then
        MsgBox (HairetuName & "�̊J�n�v�f�ԍ���1�ɂ��Ă�������")
        Stop
        Exit Sub '���͌��̃v���V�[�W�����m�F���邽�߂ɔ�����
    End If

End Sub

Private Function �z�񂩂�A�z�z�񃊃X�g�쐬(ItemArray2D, KeyArray1D)
    
    Dim TmpDict As Object
    Dim I&, J&, K&, M&, N& '�����グ�p(Long�^)
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
    
    �z�񂩂�A�z�z�񃊃X�g�쐬 = Output

End Function

Private Function �z�񂩂�K�w�^�A�z�z��쐬(KaisoArray2D, ItemArray1D)
    
    Dim I&, J&, K&, M&, N& '�����グ�p(Long�^)
    N = UBound(KaisoArray2D, 1)
    M = UBound(KaisoArray2D, 2)
    
    Dim Output As Object
    Set Output = CreateObject("Scripting.Dictionary")
    
    Dim TmpUnique, TmpValue, TmpArray, TmpItiArray, TmpDict As Object, TmpItemArray1D
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

Private Function �z��̐擪��l�Ɉ�v������̂̂ݔz�񒊏o(InputArray2D, InputValue)
    
    Dim I&, J&, K&, M&, N& '�����グ�p(Long�^)
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
    
    �z��̐擪��l�Ɉ�v������̂̂ݔz�񒊏o = Output
    
End Function

Private Function �z��̎w���̃��j�[�N�l���X�g�擾(InputArray, Col%)
    
    Dim TmpDict As Object
    Set TmpDict = CreateObject("Scripting.Dictionary")
    Dim I&, J&, K&, M&, N& '�����グ�p(Long�^)
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

Private Function �ꎟ���z��̎w��͈͎擾(Array1D, ItiArray1D)
    
    Dim I&, J&, K&, M&, N& '�����グ�p(Long�^)
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
    
    �ꎟ���z��̎w��͈͎擾 = Output
    
End Function

Private Sub CheckArray2D(InputArray, Optional HairetuName$ = "�z��")
'���͔z��2�����z�񂩂ǂ����`�F�b�N����
'20210804

    Dim Dummy2%, Dummy3%
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

Private Sub CheckArray2DStart1(InputArray, Optional HairetuName$ = "�z��")
'����2�����z��̊J�n�ԍ���1���ǂ����`�F�b�N����
'20210804

    If LBound(InputArray, 1) <> 1 Or LBound(InputArray, 2) <> 1 Then
        MsgBox (HairetuName & "�̊J�n�v�f�ԍ���1�ɂ��Ă�������")
        Stop
        Exit Sub '���͌��̃v���V�[�W�����m�F���邽�߂ɔ�����
    End If

End Sub



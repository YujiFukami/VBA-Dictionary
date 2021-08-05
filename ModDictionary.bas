Attribute VB_Name = "ModDictionary"
Option Explicit

'�A�z�z��֘A���W���[��

Private Sub TestMakeDictFromArray()
    
    '�e�X�g�p�̔z��̒�`
    Dim KaisoArray2D, ItemArray1D
    KaisoArray2D = Application.Transpose(Application.Transpose( _
                    Array(Array("���Q��", "�������ݒn"), _
                    Array("���Q��", "���Y�i"), _
                    Array("���Q��", "���Y�i"), _
                    Array("���Q��", "���Y�i"), _
                    Array("���Q��", "���L����"), _
                    Array("���Q��", "���L����"), _
                    Array("���쌧", "�������ݒn"), _
                    Array("���쌧", "���Y�i"), _
                    Array("���쌧", "���Y�i"), _
                    Array("���쌧", "���Y�i"), _
                    Array("���쌧", "���Y�i"), _
                    Array("���쌧", "���L����"), _
                    Array("������", "�������ݒn"), _
                    Array("������", "���Y�i"), _
                    Array("������", "���L����"), _
                    Array("������", "���L����"), _
                    Array("���m��", "�������ݒn"), _
                    Array("���m��", "���Y�i"), _
                    Array("���m��", "���Y�i"), _
                    Array("���m��", "���L����")) _
                    ))
    ItemArray1D = Application.Transpose(Application.Transpose( _
            Array("���R�s", "�݂���", "�^�I��", "�^��", "�o���B����", "�݂����", "�����s", "���ǂ�", "�ݖ�", "�I���[�u", "�f��", "���ǂ�]", "�����s", "������", "��������L", "����������", "�����s", "����", "��", "���񂶂傤����") _
            ))
    
    '�z��̒��g��\���m�F
    Call DPH(KaisoArray2D, , "KaisoArray2D")
    Call DPH(ItemArray1D, , "ItemArray1D")
    
    '�K�w�^�A�z�z��쐬
    Dim OutputDict As Object
    Set OutputDict = MakeDictFromArray(KaisoArray2D, ItemArray1D)
    
    '�o�̓e�X�g
    Debug.Print OutputDict("���Q��")("���Y�i")(2) '���^�I��
    Debug.Print OutputDict("���Q��")("���L����")(1) '���o���B����
    Debug.Print OutputDict("������")("���L����")(1) '����������L
    Debug.Print OutputDict("���쌧")("�������ݒn")(1) '�������s

End Sub

Function MakeDictFromArrayWithItem(KaisoArray2D, KeyArray1D, ItemArray2D)
    
    '�����`�F�b�N
    Call CheckArray2D(KaisoArray2D, "KaisoArray2D") '2�����z�񂩃`�F�b�N
    Call CheckArray2DStart1(KaisoArray2D, "KaisoArray2D") '�v�f�̊J�n�ԍ���1���`�F�b�N
    Call CheckArray1D(KeyArray1D, "KeyArray1D") '1�����z�񂩃`�F�b�N
    Call CheckArray1DStart1(KeyArray1D, "KeyArray1D") '�v�f�̊J�n�ԍ���1���`�F�b�N
    Call CheckArray2D(ItemArray2D, "ItemArray2D") '2�����z�񂩃`�F�b�N
    Call CheckArray2DStart1(ItemArray2D, "ItemArray2D") '�v�f�̊J�n�ԍ���1���`�F�b�N
    
    If UBound(KaisoArray2D, 1) <> UBound(ItemArray2D, 1) Then
        MsgBox ("�uKaisoArray2D�v�ƁuItemArray2D�v�̏c�v�f������v�����Ă�������")
        Stop
        End
    End If
    
    If UBound(KeyArray1D, 1) <> UBound(ItemArray2D, 2) Then
        MsgBox ("�uKeyArray1D�v�̗v�f���ƁuItemArray2D�v�̉��v�f������v�����Ă�������")
        Stop
        End
    End If
    
    '�v�Z����
    Dim DictArray
    DictArray = �z�񂩂�A�z�z�񃊃X�g�쐬(ItemArray2D, KeyArray1D)
    Dim Output As Object
    Set Output = �z�񂩂�K�w�^�A�z�z��쐬(KaisoArray2D, DictArray)
    
    Set MakeDictFromArray = Output
    
End Function

Function MakeDictFromArray(KaisoArray2D, ItemArray1D)

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
            Output.Add TmpValue, TmpItemArray1D
        Else
            Set TmpDict = �z�񂩂�K�w�^�A�z�z��쐬(TmpArray, TmpItemArray1D)
            Output.Add TmpValue, TmpDict
        End If
    Next
        
    Set MakeDictFromArray = Output
        
End Function

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
            Output.Add TmpValue, TmpItemArray1D
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


Attribute VB_Name = "ModMakeDictFromArray1D"
Option Explicit

'MakeDictFromArray1D�E�E�E���ꏊ�FFukamiAddins3.ModDictionary
'CheckArray1D       �E�E�E���ꏊ�FFukamiAddins3.ModDictionary
'CheckArray1DStart1 �E�E�E���ꏊ�FFukamiAddins3.ModDictionary

'------------------------------


'�A�z�z��֘A���W���[��
'------------------------------


Function MakeDictFromArray1D(KeyArray1D, ItemArray1D)
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



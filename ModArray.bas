Attribute VB_Name = "ModArray"
Option Explicit

'�z��̏����֌W�̃v���V�[�W��

Function SortArrayByNetFramework(InputArray, Optional InputOrder As OrderType = xlAscending)
'�ꎟ���z���Net.Framework���g���ď����ɂ���
'20210726

    Dim DataList, I&, Msg$
    Set DataList = CreateObject("System.Collections.ArrayList")
    
    For I = LBound(InputArray, 1) To UBound(InputArray, 1)
        DataList.Add InputArray(I)
    Next I
    
    Dim Output
    ReDim Output(LBound(InputArray, 1) To UBound(InputArray, 1))
    
    If InputOrder = xlAscending Then
        DataList.Sort
    Else
        DataList.Reverse
    End If
    
    For I = 0 To DataList.Count - 1
        Output(I + LBound(InputArray, 1)) = DataList(I)
    Next I
    Set DataList = Nothing
    
    SortArrayByNetFramework = Output
    
End Function

Sub TestSortArray2D()
    Dim TmpList
    TmpList = Range("B20").CurrentRegion.Value
    Dim SortList
    
    SortList = SortArray2D(TmpList, 2)
    Call DPH(SortList)
    
End Sub

Function SortArray2D(InputArray2D, Optional SortCol%, Optional InputOrder As OrderType = xlAscending)
'�w���2�����z����A�w������ɕ��ёւ���
'�z��͕�������܂�ł��Ă��悢
'20210726

'InputArray2D�E�E�E���ёւ��Ώۂ�2�����z��
'SortCol�E�E�E���ёւ��̊�Ŏw�肷���ԍ�
'InputOrder�E�E�ExlAscending������, xlDescending���~��

    '�w����1�����z��Œ��o
    Dim KijunArray1D
    Dim MinRow&, MaxRow&
    MinRow = LBound(InputArray2D, 1)
    MaxRow = UBound(InputArray2D, 1)
    ReDim KijunArray1D(MinRow To MaxRow)
    Dim I&, J&, K&, M&, N& '�����グ�p(Long�^)
    For I = MinRow To MaxRow
        KijunArray1D(I) = InputArray2D(I, SortCol)
    Next I
    
    '���ёւ�
    Dim Output
    Output = SortArray2Dby1D(InputArray2D, KijunArray1D, InputOrder)
    SortArray2D = Output
    
End Function

Function SortArray2Dby1D(InputArray2D, ByVal KijunArray1D, Optional InputOrder As OrderType = xlAscending)
'�w���2�����z����A�w��1�����z�����ɕ��ёւ���
'�z��͕�������܂�ł��Ă��悢
'20210726

'InputArray2D�E�E�E���ёւ��Ώۂ�2�����z��
'KijunArray1D�E�E�E���ёւ��̊�ƂȂ�z��
'InputOrder�E�E�ExlAscending������, xlDescending���~��

    '���͒l�̃`�F�b�N
    Dim Dummy%
    On Error Resume Next
    Dummy = UBound(KijunArray1D, 2)
    On Error GoTo 0
    If Dummy <> 0 Then
        MsgBox ("��z���1�����z�����͂��Ă�������")
        Stop
        End
    End If
    
    Dummy = 0
    On Error Resume Next
    Dummy = UBound(InputArray2D, 2)
    On Error GoTo 0
    If Dummy = 0 Then
        MsgBox ("���ёւ��Ώۂ̔z���2�����z�����͂��Ă�������")
        Stop
        End
    End If
    
    Dim MinRow&, MaxRow&, MinCol&, MaxCol&
    MinRow = LBound(InputArray2D, 1)
    MaxRow = UBound(InputArray2D, 1)
    MinCol = LBound(InputArray2D, 2)
    MaxCol = UBound(InputArray2D, 2)
    If MinRow <> LBound(KijunArray1D, 1) Or MaxRow <> UBound(KijunArray1D, 1) Then
        MsgBox ("���ёւ��Ώۂ̔z��ƁA��z��̊J�n�A�I���v�f�ԍ�����v�����Ă�������")
        Stop
        End
    End If
    
    '��z��ɕ����񂪊܂܂�Ă���ꍇ��ISO�R�[�h�ɕϊ�
    Dim StrAruNaraTrue As Boolean
    StrAruNaraTrue = False
    Dim I&, J&, K&, M&, N& '�����グ�p(Long�^)
    Dim Tmp, TmpStr$
    For I = MinRow To MaxRow
        Tmp = KijunArray1D(I)
        If VarType(Tmp) = vbString Then
            StrAruNaraTrue = True
            Exit For
        End If
    Next I
    
    If StrAruNaraTrue Then
        For I = MinRow To MaxRow
            TmpStr = KijunArray1D(I)
            KijunArray1D(I) = ConvStrToISO(TmpStr)
        Next I
    End If
    
    '��z��𐳋K�����āA(1�`�v�f��)�̊Ԃ̐��l�ɂ���
    Dim Count&, MinNum#, MaxNum#
    Count = MaxRow - MinRow + 1
    MinNum = WorksheetFunction.Min(KijunArray1D)
    MaxNum = WorksheetFunction.Max(KijunArray1D)
    For I = MinRow To MaxRow
        KijunArray1D(I) = (KijunArray1D(I) - MinNum) / (MaxNum - MinNum) '(0�`1)�̊ԂŐ��K��
        KijunArray1D(I) = (Count - 1) * KijunArray1D(I) + 1 '(1�`�v�f��)�̊�
    Next I
    
    '���ёւ�(1,2,3�̔z�������ăN�C�b�N�\�[�g�ŕ��ёւ��āA�Ώۂ̔z�����ёւ����1,2,3�œ���ւ���)
    Dim TmpArray, Array123, TmpNum&
    ReDim Array123(MinRow To MaxRow)
    For I = MinRow To MaxRow
        If InputOrder = xlAscending Then
            '����
            Array123(I) = I - MinRow + 1
        Else
            '�~��
            Array123(MaxRow - I + 1) = I - MinRow + 1
        End If
    Next I
    
    Call DPH(Array123)
    Call DPH(KijunArray1D)
    Call SortArrayQuick(KijunArray1D, Array123)
        
    ReDim TmpArray(MinRow To MaxRow, MinCol To MaxCol)
    For I = MinRow To MaxRow
        TmpNum = Array123(I)
        For J = MinCol To MaxCol
            TmpArray(I, J) = InputArray2D(TmpNum, J)
        Next J
    Next I
    
    '�o��
    SortArray2Dby1D = TmpArray

End Function

Sub SortArrayQuick(KijunArray, Array123, Optional StartNum%, Optional EndNum%)
'�N�C�b�N�\�[�g��1�����z�����ёւ���
'���ёւ���̏��Ԃ��o�͂��邽�߂ɔz��uArray123�v�𓯎��ɕ��ёւ���
'20210726

'KijunArray�E�E�E���ёւ��Ώۂ̔z��i1�����z��j
'Array123�E�E�E�u1,2,3�v�̒l��������1�����z��
'StartNum�E�E�E�ċA�p�̈���
'EndNum�E�E�E�ċA�p�̈���

    If StartNum = 0 Then
        StartNum = LBound(KijunArray, 1)
    End If
    
    If EndNum = 0 Then
        EndNum = UBound(KijunArray, 1)
    End If
    
    Dim Tmp#, Counter#, I&, J&
    Counter = KijunArray((StartNum + EndNum) \ 2)
    I = StartNum - 1
    J = EndNum + 1
    
    '���ёւ��Ώۂ̔z��̏���
    Dim Col&, MinCol&, MaxCol&
    Dim Tmp2
    
    Do
        Do
            I = I + 1
        Loop While KijunArray(I) < Counter
        
        Do
            J = J - 1
        Loop While KijunArray(J) > Counter
        
        If I >= J Then Exit Do
        Tmp = KijunArray(J)
        KijunArray(J) = KijunArray(I)
        KijunArray(I) = Tmp
        
        Tmp2 = Array123(J)
        Array123(J) = Array123(I)
        Array123(I) = Tmp2
    
    Loop
    If I - StartNum > 1 Then
        Call SortArrayQuick(KijunArray, Array123, StartNum, I - 1) '�ċA�Ăяo��
    End If
    If EndNum - J > 1 Then
        Call SortArrayQuick(KijunArray, Array123, J + 1, EndNum) '�ċA�Ăяo��
    End If
End Sub

Function ConvStrToISO(InputStr$)
'���������ёւ��p��ISO�R�[�h�ɕϊ�
'20210726

    Dim Mojiretu As String
    Dim I&, J&, K&, M&, N& '�����グ�p(Long�^)
    Dim UniCode
    
    Dim UniMax&
    UniMax = 65536
    
    Dim StartKeta%, Kurai#
    StartKeta = 20 '����������������������������������������������
    Kurai = Exp(1) '����������������������������������������������
    
    Dim Output#
    
    If InputStr = "" Then
        Output = 0
    Else
        N = Len(InputStr)
        ReDim UniCode(1 To N)
        
        Output = 0
        For I = 1 To N
            UniCode(I) = Abs(AscW(Mid(InputStr, I, 1)))
            Output = Output + ((Kurai ^ StartKeta) / (UniMax) ^ (I - 1)) * UniCode(I)
            
        Next I
    End If
    
    ConvStrToISO = Output

End Function

Sub CheckArray1D(InputArray, Optional HairetuName$ = "�z��")
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

Sub CheckArray2D(InputArray, Optional HairetuName$ = "�z��")
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

Sub CheckArray1DStart1(InputArray, Optional HairetuName$ = "�z��")
'����1�����z��̊J�n�ԍ���1���ǂ����`�F�b�N����
'20210804

    If LBound(InputArray, 1) <> 1 Then
        MsgBox (HairetuName & "�̊J�n�v�f�ԍ���1�ɂ��Ă�������")
        Stop
        Exit Sub '���͌��̃v���V�[�W�����m�F���邽�߂ɔ�����
    End If

End Sub

Sub CheckArray2DStart1(InputArray, Optional HairetuName$ = "�z��")
'����2�����z��̊J�n�ԍ���1���ǂ����`�F�b�N����
'20210804

    If LBound(InputArray, 1) <> 1 Or LBound(InputArray, 2) <> 1 Then
        MsgBox (HairetuName & "�̊J�n�v�f�ԍ���1�ɂ��Ă�������")
        Stop
        Exit Sub '���͌��̃v���V�[�W�����m�F���邽�߂ɔ�����
    End If

End Sub

Sub ClipCopyArray2D(Array2D)
'2�����z���ϐ��錾�p�̃e�L�X�g�f�[�^�ɕϊ����āA�N���b�v�{�[�h�ɃR�s�[����
'20210805
    
    '�����`�F�b�N
    Call CheckArray2D(Array2D, "Array2D")
    Call CheckArray2DStart1(Array2D, "Array2D")
    
    Dim I&, J&, K&, M&, N& '�����グ�p(Long�^)
    N = UBound(Array2D, 1)
    M = UBound(Array2D, 2)
    
    Dim TmpValue
    Dim Output$
    
    Output = ""
    For I = 1 To N
        If I = 1 Then
            Output = Output & String(3, Chr(9)) & "Array(Array("
        Else
            Output = Output & String(3, Chr(9)) & "Array("
        End If
        
        For J = 1 To M
            TmpValue = Array2D(I, J)
            If IsNumeric(TmpValue) Then
                Output = Output & TmpValue
            Else
                Output = Output & """" & TmpValue & """"
            End If
            
            If J < M Then
                Output = Output & ","
            Else
                Output = Output & ")"
            End If
        Next J
        
        If I < N Then
            Output = Output & ", _" & vbLf
        Else
            Output = Output & ")"
        End If
    Next I
    
    Output = "Application.Transpose(Application.Transpose( _" & vbLf & Output & " _" & vbLf & String(3, Chr(9)) & "))"
    
    Call ClipboardCopy(Output, True)
    
End Sub

Sub ClipCopyArray1D(Array1D)
'1�����z���ϐ��錾�p�̃e�L�X�g�f�[�^�ɕϊ����āA�N���b�v�{�[�h�ɃR�s�[����
'20210805
    
    '�����`�F�b�N
    Call CheckArray1D(Array1D, "Array1D")
    Call CheckArray1DStart1(Array1D, "Array1D")
    
    Dim I&, J&, K&, M&, N& '�����グ�p(Long�^)
    N = UBound(Array1D, 1)
    
    Dim TmpValue
    Dim Output$
    
    Output = String(3, Chr(9)) & "Array("
    For I = 1 To N
        
        TmpValue = Array1D(I)
        If IsNumeric(TmpValue) Then
            Output = Output & TmpValue
        Else
            Output = Output & """" & TmpValue & """"
        End If
        
        If I < N Then
            Output = Output & ","
        Else
            Output = Output & ")"
        End If
        
    Next I
    
    Output = "Application.Transpose(Application.Transpose( _" & vbLf & Output & " _" & vbLf & String(3, Chr(9)) & "))"
    
    Call ClipboardCopy(Output, True)
    
End Sub

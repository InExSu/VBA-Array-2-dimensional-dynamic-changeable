VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "A2Dynamic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 1

' dynamic array 2-dimensional changeable
' array one-dimensional of arrays one-dimensional - index with 1
' both rows and columns
' without Property Get, without Property Set

' based on Dragokas https:'www.cyberforum.ru/vba/thread643752.html#post3398522

' in a regular module you need to place:
'Type A2_Rows.
' a1_ColS() As Variant
'End Type

' Usage:
' A2_Dyn As New A2Dynamic ' declare type
' A2_Dyn.Create rowS, colS ' Create an empty of a given size
' A2_Dyn.Element String, Colu, sVal ' put in element sVal and if row_ or Colu is more than available dynamically expand
' d2 = A2_Dyn.A2 ' return array regular static
' A2_Dyn.ColSCountChange ' line count change
' A2_Dyn.RowSCountChange ' columns count change
' var = A2_Dyn.Element(row_, colu) ' return the contents of the array element
' ...

' ������������ ������ 2������ ����������
' ������ ���������� �������� ���������� - ������ � 1
' � �� ������� � �� ��������
' ��� Property Get, ��� Property Set

' �� ������ Dragokas https:'www.cyberforum.ru/vba/thread643752.html#post3398522

' � ������ ������� ����� ����������:
'Type A2_Rows
'    a1_ColS() As Variant
'End Type

' �������������:
' A2_Dyn As New A2Dynamic ' �������� ���
' A2_Dyn.Create rowS, colS  ' Create ������ ��������� �������
' A2_Dyn.Element ������, Colu, sVal ' ��������� � ������� sVal � ���� row_ ��� Colu ������,  ��� ������� ����������� ���������
' d2 = A2_Dyn.A2 ' ������� ������ ������� �����������
' A2_Dyn.ColSCountChange ' ������  ����� ��������
' A2_Dyn.RowSCountChange ' ������� ����� ��������
' var = A2_Dyn.Element(row_, colu) ' ������� ���������� �������� �������
' ...

Private A2_RowS() As A2_RowS    ' ������ ������� �������������
Private row_ As Long
Private colu As Long


Sub Create( _
   rowS As Long, _
   colS As Long)
   
   ' ����� ��� ������� ������ ���������� � 1
    
   ReDim A2_RowS(1 To rowS)

   For row_ = 1 To UBound(A2_RowS)
      
      ReDim _
         A2_RowS(row_). _
         a1_ColS(1 To colS): Next
         
End Sub


Sub Fill_From(a2() As Variant)

   ' ��������� �� �������� ������� ������������
   ' ������ ������� ������������� �������� � 1
   
   Create AX_Len(a2), A2Width(a2)

   Static _
      diff_row_ As Long, _
      diff_colu As Long

   diff_row_ = LBound(A2_RowS) - LBound(a2, 1)
   diff_colu = LBound(A2_RowS(1).a1_ColS) - LBound(a2, 2)

   For row_ = LBound(a2, 1) To UBound(a2, 1)
      For colu = LBound(a2, 2) To UBound(a2, 2)
                                                
         A2_RowS(row_ + diff_row_). _
            a1_ColS(colu + diff_colu) = a2(row_, colu): Next: Next
                                                                  
End Sub


Function Element( _
   row_ As Long, _
   colu As Long, _
   Optional vvar As Variant) _
   As Variant
   
   ' ��������� ������� �������
   ' ��� ������������� ����������� ����������� "������"
   ' ��� ���������� �������� �� ����������� ������ �������

   If Not IsMissing(vvar) Then
   
      If row_ > UBound(A2_RowS) Then _
         RowSCountChange row_

      If colu > UBound(A2_RowS(1).a1_ColS) Then _
         ColSCountChange colu

      A2_RowS(row_).a1_ColS(colu) = vvar
      
   Else
      
      Element = A2_RowS(row_).a1_ColS(colu): End If
      
End Function


Sub RowSCountChange(rowS As Long)
   
   ' ����� ����� ��������
   
   ReDim Preserve A2_RowS(1 To rowS)
   
   For row_ = 1 To rowS
   
      ReDim Preserve A2_RowS(row_).a1_ColS(1 To ColSCount): Next
      
End Sub


Sub ColSCountChange(colu As Long)
   
   ' ����� �������� ��������
   
   For row_ = 1 To UBound(A2_RowS)
      
      ReDim Preserve A2_RowS(row_).a1_ColS(1 To colu): Next

End Sub


Function A2Return() As Variant()
        
   ' ������� ������� a2 ������
   ' ���� ������� ��������� � �������������:
   ' A2_RowS(row_).a1_ColS(colu)

   Static a2() As Variant
   ReDim a2(1 To UBound(A2_RowS), _
      1 To UBound(A2_RowS(1).a1_ColS))

   For row_ = 1 To UBound(a2)
      For colu = 1 To UBound(a2, 2)
         
         a2(row_, colu) = A2_RowS(row_).a1_ColS(colu): Next: Next

   If IsArray(a2) Then _
      A2Return = a2

End Function


Function A2Cut( _
   rowS As Long, _
   colS As Long) _
   As Variant()

   ' ������� ������ ����������� ����������

   RowSCountChange rowS
   ColSCountChange colS

   A2Cut = A2Return

End Function


Function RowsCount() As Long
   
   ' ������� ����� �����

   RowsCount = UBound(A2_RowS)
    
End Function


Function ColSCount() As Long

   ' ������� ����� ��������

   ColSCount = UBound(A2_RowS(1).a1_ColS)
      
End Function
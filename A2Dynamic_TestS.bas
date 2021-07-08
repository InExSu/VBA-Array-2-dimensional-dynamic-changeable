Attribute VB_Name = "A2Dynamic_TestS"
Option Explicit
Option Private Module
'@IgnoreModule
'@TestModule
'@Folder("Tests")

Private Mock As New Mock
Private A2Dyn As New A2Dynamic

Private Assert As Object
Private Fakes As Object


'@ModuleInitialize
Private Sub ModuleInitialize()
   'this method runs once per module.
   Set Assert = CreateObject("Rubberduck.AssertClass")
   Set Fakes = CreateObject("Rubberduck.FakesProvider")
End Sub


'@ModuleCleanup
Private Sub ModuleCleanup()
   'this method runs once per module.
   Set Assert = Nothing
   Set Fakes = Nothing
End Sub


'@TestInitialize
Private Sub TestInitialize()
   'This method runs before every test in the module..
End Sub


'@TestCleanup
Private Sub TestCleanup()
   'this method runs after every test in the module.
End Sub


'@TestMethod
Sub Create_TestMethod()
   On Error GoTo TestFail
   Dim varReturn As Variant
   Dim rowS As Long
   Dim colS As Long
   rowS = Mock.G_Long
   colS = Mock.G_Long
   A2Dyn.Create rowS, colS
   Mock.wb.Close False
   Exit Sub
TestFail:
   Mock.wb.Close False
   Assert.Fail "Test error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Sub Element_TestMethod()
   
   On Error GoTo TestFail
   
   Dim varReturn As Variant
   Dim row_ As Long
   Dim colu As Long
   Dim vvar As Variant
   
   Create_TestMethod
   
   row_ = Mock.G_Long
   colu = Mock.G_Long
   vvar = Mock.G_Variant
   varReturn = A2Dyn.Element(row_, colu, vvar)
   If A2Dyn.RowsCount <> row_ Then Err.Raise 567, "Element(row_,colu,vVar)"
   If A2Dyn.ColSCount <> colu Then Err.Raise 567, "Element(row_,colu,vVar)"
   Mock.wb.Close False
   Exit Sub
TestFail:
   Mock.wb.Close False
   Assert.Fail "Test error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Sub RowSCountChange_TestMethod()
   On Error GoTo TestFail
   Dim varReturn As Variant
   Dim rowS As Long
   
   Create_TestMethod
   
   rowS = Mock.G_Long + 1
   A2Dyn.RowSCountChange rowS
   If A2Dyn.RowsCount <> rowS Then Err.Raise 567, "", ""
   
   Exit Sub
TestFail:
   Mock.wb.Close False
   Assert.Fail "Test error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Sub ColSCountChange_TestMethod()
   On Error GoTo TestFail
   Dim varReturn As Variant
   Dim colu As Long

   Create_TestMethod

   colu = Mock.G_Long
   A2Dyn.ColSCountChange colu
   If A2Dyn.ColSCount <> colu Then Err.Raise 567, "", ""
   
   Exit Sub
TestFail:
   Mock.wb.Close False
   Assert.Fail "Test error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Sub A2Retu_TestMethod()
   On Error GoTo TestFail
   Dim varReturn As Variant

   Create_TestMethod

   varReturn = A2Dyn.A2Return
   'if varReturn <> 0 Then Err.Raise 567, "A2Retu{}(As(Variant(Dim((As(Variant)"
   ' Mock.wb.Close False
   Exit Sub
TestFail:
   ' Mock.wb.Close False
   Assert.Fail "Test error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Sub Fill_4_Test_TestMethod()
   On Error GoTo TestFail
   Dim varReturn As Variant

   varReturn = A2Dyn.Fill_4_Test
   'if varReturn <> 0 Then Err.Raise 567, "Fill_4_Test(Dim((As(Variant)"
   ' Mock.wb.Close False
   Exit Sub
TestFail:
   ' Mock.wb.Close False
   Assert.Fail "Test error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Sub A2Cut_TestMethod()
   On Error GoTo TestFail
   Dim varReturn As Variant
   Dim rowS As Long
   Dim colS As Long
   rowS = Mock.G_Long
   colS = Mock.G_Long
   varReturn = A2Dyn.A2Cut(rowS, colS)
   'if varReturn <> 0 Then Err.Raise 567, "A2Cut(rowS,colS)"
   ' Mock.wb.Close False
   Exit Sub
TestFail:
   ' Mock.wb.Close False
   Assert.Fail "Test error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Sub RowSCount_TestMethod()
   On Error GoTo TestFail
   Dim varReturn As Variant

   A2Dyn.RowsCount
   ' Mock.wb.Close False
   Exit Sub
TestFail:
   ' Mock.wb.Close False
   Assert.Fail "Test error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Sub ColSCount_TestMethod()
   On Error GoTo TestFail
   Dim varReturn As Variant

   A2Dyn.ColSCount
   ' Mock.wb.Close False
   Exit Sub
TestFail:
   ' Mock.wb.Close False
   Assert.Fail "Test error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Sub Fill_From_TestMethod()
   
   On Error GoTo TestFail
   
   Dim a2_Sour() As Variant
   Dim a2_Dest() As Variant
   
   a2_Sour = Mock.G_a2
   A2Dyn.Fill_From a2_Sour
   a2_Dest = A2Dyn.A2Return
   
   If A2S_Equal_Check(a2_Sour, a2_Dest)(1) = False Then Err.Raise 567, "", ""
 
   ReDim a2_Sour(0 To 0, 0 To 0)
   A2Dyn.Fill_From a2_Sour
   a2_Dest = A2Dyn.A2Return
   
   If a2_Dest(1, 1) <> a2_Sour(0, 0) Then Err.Raise 567, "", ""
    Mock.wb.Close False
   Exit Sub
TestFail:
    Mock.wb.Close False
   Assert.Fail "Test error: #" & Err.Number & " - " & Err.Description
End Sub


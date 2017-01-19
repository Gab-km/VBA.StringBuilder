VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StringBuilderTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

''' [NOTICE]
''' To run this test class, use Assert.bas in Ariawase.
''' see: https://github.com/vbaidiot/Ariawase

Sub WhenInitializeThenTheInstanceHasAnEmptyValue_Test()
    Dim sut As New StringBuilder
        
    Assert.AreEqual 0, sut.Length
    Assert.AreEqual 16, sut.Capacity
    Assert.AreEqual "", sut.ToString()
    
    Set sut = Nothing
End Sub

Sub GivenAStringWhenAppendItThenReturnTheInstanceWithAppendedTheString_Test()
    Dim sut As New StringBuilder
    Dim other As StringBuilder
        
    Set other = sut.Append("Test")
    Assert.IsInstanceOfTypeName "StringBuilder", other
    Assert.AreEq sut, other
    
    Set sut = Nothing
End Sub

Sub WhenAppendSomeTimesThenAppendValuesSequencially_Test()
    Dim sut As New StringBuilder
    
    sut.Append ("This")
    sut.Append (" ")
    sut.Append ("is")
    sut.Append (" ")
    sut.Append ("a")
    sut.Append (" ")
    sut.Append ("pen")
    sut.Append (".")
    Assert.AreEqual 14, sut.Length
    Assert.AreEqual 16, sut.Capacity
    Assert.AreEqual "This is a pen.", sut.ToString()
        
    Set sut = Nothing
End Sub

Sub GivenAStringWhenInsertItWithIndexThenReturnTheInstanceWithInsertedTheValue_Test()
    Dim sut As New StringBuilder
    Dim other As StringBuilder
        
    sut.Append ("Test")
    Set other = sut.Insert(1, "he b")
    Assert.IsInstanceOfTypeName "StringBuilder", other
    Assert.AreEq sut, other
    Assert.AreEqual "The best", sut.ToString()
    
    Set sut = Nothing
End Sub

Sub WhenInsertAtTheFirstThenInsertTheValue_Test()
    Dim sut As New StringBuilder
        
    sut.Append ("Test")
    sut.Insert 0, "AB "
    Assert.AreEqual "AB Test", sut.ToString()
    
    Set sut = Nothing
End Sub

Sub WhenInsertAtTheLastThenInsertTheValue_Test()
    Dim sut As New StringBuilder
        
    sut.Append ("Test")
    sut.Insert 4, "ing"
    Assert.AreEqual "Testing", sut.ToString()
    
    Set sut = Nothing
End Sub

Sub WhenInsertBeyondTheFirstThenRaiseError_Test()
    Dim sut As New StringBuilder
        
    sut.Append ("Test")
    Assert.IsErrMethod 9, sut, "Insert", Array(-1, "AB ")
        
    Set sut = Nothing
End Sub

Sub WhenInsertBeyondTheLastThenRaiseError_Test()
    Dim sut As New StringBuilder
        
    sut.Append ("Test")
    Assert.IsErrMethod 9, sut, "Insert", Array(5, "ing")
        
    Set sut = Nothing
End Sub

Sub LengthShouldReturnsItsLength_Test()
    Dim sut As New StringBuilder
    
    sut.Append ("Test")
    Assert.AreEqual 4, sut.Length
    
    Set sut = Nothing
End Sub

Sub CapacityShouldReturnsItsCapacity_Test()
    Dim sut As New StringBuilder
    
    Assert.AreEqual 16, sut.Capacity
    
    Set sut = Nothing
End Sub

Sub GivenTheInstanceWithNonEmptyStringWhenClearThenTheValueIsEmpty_Test()
    Dim sut As New StringBuilder
    
    sut.Append ("Spam!Spam!Spam!")
    Call sut.Clear
    Assert.AreEqual 0, sut.Length
    Assert.AreEqual "", sut.ToString()
    
    Set sut = Nothing
End Sub

Sub WhenTheInstanceAppend16CharactersThenCapacityDoesNotIncrease_Test()
    Dim sut As New StringBuilder
    
    Const value As String = "1234567890123456"
    sut.Append (value)
    Assert.AreEqual 16, sut.Length
    Assert.AreEqual 16, sut.Capacity
    Assert.AreEqual value, sut.ToString()
    
    Set sut = Nothing
End Sub

Sub WhenTheInstanceAppendOver16CharactersThenCapacityIncreases_Test()
    Dim sut As New StringBuilder
    
    Const value As String = "12345678901234567"
    sut.Append (value)
    Assert.AreEqual 17, sut.Length
    Assert.AreEqual 33, sut.Capacity
    Assert.AreEqual value, sut.ToString()
    
    Set sut = Nothing
End Sub

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StringBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''' ------------------------------------------------------------------------
''' Copyright 2017 Kazuhiro Matsushima
'''
''' Licensed under the Apache License, Version 2.0 (the "License");
''' you may not use this file except in compliance with the License.
''' You may obtain a copy of the License at
'''
'''     http://www.apache.org/licenses/LICENSE-2.0
'''
''' Unless required by applicable law or agreed to in writing, software
''' distributed under the License is distributed on an "AS IS" BASIS,
''' WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
''' See the License for the specific language governing permissions and
''' limitations under the License.
''' ------------------------------------------------------------------------
''' <summary>Represents a mutable string of characters.</summary>

Option Compare Binary
Option Explicit

Private Const DefaultBlockSize As Integer = 16

Private m_buffer As String
Private m_length As Integer

''' <summary>
''' Initializes a new instance of the StringBuilder class.
''' </summary>
Private Sub Class_Initialize()
    m_buffer = VBA.String(DefaultBlockSize, " ")
    m_length = 0
End Sub

''' <summary>
''' Finalizes this instance.
''' </summary>
Private Sub Class_Terminate()
    Call Clear
End Sub

''' <summary>
''' Gets the length of the current StringBuilder object.
''' </summary>
''' <returns>The length of this instance.</returns>
Public Property Get Length() As Integer
    Length = m_length
End Property

''' <summary>
''' Gets the maximum number of characters that can be contained in the memory allocated by the current instance.
''' </summary>
''' <returns>
''' The maximum number of characters that can be contained in the memory allocated by the current instance.
''' Its value can be equal to or greater than Length.
''' </returns>
Public Property Get Capacity() As Integer
    Capacity = Len(m_buffer)
End Property

''' <summary>
''' Appends a copy of the specified string to this instance.
''' </summary>
''' <param name="value">The string to append.</param>
''' <returns>A reference to this instance after the append operation has completed.</returns>
Public Function Append(ByVal value As String) As StringBuilder
    Dim valueLength As Integer
    valueLength = Len(value)
    ExpandIfNeeded (valueLength)
    
    Mid(m_buffer, m_length + 1) = value
    m_length = m_length + valueLength
    Set Append = Me
End Function

''' <summary>
''' Inserts a string into this instance at the specified character position.
''' </summary>
''' <param name="index">The position in this instance where insertion begins.</param>
''' <param name="value">The string to insert.</param>
''' <returns>A reference to this instance after the insert operation has completed.</returns>
Public Function Insert(ByVal index As Integer, ByVal value As String) As StringBuilder
    If index < 0 Then
        Err.Raise 9
    ElseIf m_length < index Then
        Err.Raise 9
    End If
    
    Dim valueLength As Integer
    valueLength = Len(value)
    ExpandIfNeeded (valueLength)
    
    Dim rest As String
    rest = Mid(m_buffer, index + 1)
    Mid(m_buffer, index + 1) = value
    Mid(m_buffer, index + valueLength + 1) = rest
    m_length = m_length + valueLength
    Set Insert = Me
End Function

''' <summary>
''' Expands a buffer of this instance if the buffer size will not be enough to append or
''' insert new value.
''' </summary>
''' <param name="increaseSize">A size to increase.</param>
Private Sub ExpandIfNeeded(ByVal increaseSize)
    Dim nextLength As Integer
    nextLength = m_length + increaseSize
    If Len(m_buffer) < nextLength Then
        Dim blockSize As Integer
        blockSize = Max(increaseSize, m_length)
        m_buffer = m_buffer + VBA.String(blockSize, " ")
    End If
End Sub

''' <summary>
''' Returns the larger of two integers.
''' </summary>
''' <param name="val1">The first of two integers to compare.</param>
''' <param name="val2">The second of two integers to compare.</param>
''' <returns>Parameter val1 or val2, whichever is larger.</returns>
Private Function Max(ByVal val1 As Integer, ByVal val2 As Integer) As Integer
    Max = IIf(val1 > val2, val1, val2)
End Function

''' <summary>
''' Removes all characters from the current StringBuilder instance.
''' </summary>
Public Sub Clear()
    m_buffer = ""
    m_length = 0
End Sub

''' <summary>
''' Returns a value indicating whether this instance is equal to a specified object.
''' </summary>
''' <returns>
''' True if this instance and sb have equal string and Capacity values; otherwise, False.
''' </returns>
Public Function Equals(ByVal sb As StringBuilder) As Boolean
    If sb.ToString() <> Me.ToString() Then
        Equals = False
    ElseIf sb.Capacity <> Me.Capacity Then
        Equals = False
    Else
        Equals = True
    End If
End Function

''' <summary>
''' Converts the value of this instance to a String.
''' </summary>
''' <returns>A string whose value is the same as this instance.</returns>
Public Function ToString() As String
    ToString = Left(m_buffer, m_length)
End Function

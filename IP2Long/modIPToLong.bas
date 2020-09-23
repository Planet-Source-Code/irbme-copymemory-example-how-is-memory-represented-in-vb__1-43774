Attribute VB_Name = "modIPToLong"
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)

'This structure is a representation of the long variable
'but the difference is we can access each byte seperately.

'It looks like this in memory:

'[][][][][][][][] [][][][][][][][] [][][][][][][][] [][][][][][][][]
'<----Part1-----> <----Part2-----> <----Part3-----> <----Part4---->
'<----8 bits----> <----8 bits----> <----8 bits----> <----8 bits--->
'<----1 byte----> <----1 byte----> <----1 byte----> <----1 byte--->

'Alternately we could use an array of 4 bytes like this:
'Private IPLong(1 to 4) As Byte
'It would look the exact same in memory. If we examine how a "long" looks in memory, we find it
'also looks the exact same as what is shown above.

Private Type IPLong
    Part1 As Byte
    Part2 As Byte
    Part3 As Byte
    Part4 As Byte
End Type


Private IPParts As IPLong
Public Const ERR_INVALIDIP As Long = -1


Public Function IPToLong(strIP As String) As Long

  Dim strIPParts() As String, i As Integer, lngIP As Long
    
    IPToLong = ERR_INVALIDIP
    
    If Not IsValidIP(strIP) Then Exit Function
    
    'First we split the IP up into its 4 seperate numbers
    strIPParts = Split(strIP, ".")
      
    'Now we convert each one to a byte and store it in each part of the IPLong structure.
    With IPParts:
        .Part1 = CByte(strIPParts(0))
        .Part2 = CByte(strIPParts(1))
        .Part3 = CByte(strIPParts(2))
        .Part4 = CByte(strIPParts(3))
    End With
    
    'Imagine we passed "127.0.0.1", the memory would now look like:
    
    '[0][1][1][1][1][1][1][1] [0][0][0][0][0][0][0][0] [0][0][0][0][0][0][0][0] [0][0][0][0][0][0][0][1]
    '<--------Part1---------> <--------Part2---------> <--------Part3---------> <--------Part4--------->
    '<---------127----------> <----------0-----------> <----------0-----------> <----------1----------->
    '<--------------------------------------------16777343--------------------------------------------->
    
    'Now all we simply do is copy this structure to a long variable.
    'The memory is exactly the same - a copy.
    'However we have just put it into a different data type so we can read it as a long
    CopyMemory ByVal VarPtr(lngIP), ByVal VarPtr(IPParts.Part1), 4
    IPToLong = lngIP
        
End Function


Private Function IsValidIP(strIP As String) As Boolean

  On Error Resume Next

  Dim strIPParts() As String, i As Integer, IPNumber As Byte
    
    'It has to have at least 1 dot in it.
    If InStr(1, strIP, ".") = 0 Then Exit Function
    
    strIPParts = Split(strIP, ".")
    
    'There must be 4 dots to be precise
    If UBound(strIPParts) <> 3 Then Exit Function
        
    For i = 0 To 3
        'The current part must be a number between 0 and 255
        IPNumber = CByte(strIPParts(i))
        If Err Then Exit Function
    Next i
    
    IsValidIP = True
        
End Function

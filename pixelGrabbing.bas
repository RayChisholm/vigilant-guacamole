Attribute VB_Name = "pixelGrabbing"
#If VBA7 Then
    Private Declare PtrSafe Function GetPixel Lib "gdi32" (ByVal hdc As LongPtr, ByVal x As Long, ByVal y As Long) As Long
    Private Declare PtrSafe Function GetCursorPos Lib "user32" (ByRef lpPoint As POINT) As LongPtr
    Private Declare PtrSafe Function GetWindowDC Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
#Else
    Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long,     ByVal y As Long) As Long
    Private Declare Function GetCursorPos Lib "user32" (ByRef lpPoint As POINT) As Long
    Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
#End If

Public Type POINT
    x As Long
    y As Long
End Type

Sub checkForMatch(col As Long, x As Long, y As Long)
Attribute checkForMatch.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim pLocation As POINT
    Dim lColour As Long
    
    Dim lDC As Variant
    lDC = GetWindowDC(0)
    
    pLocation.x = x
    pLocation.y = y
    
    lColour = GetPixel(lDC, pLocation.x, pLocation.y)
    
    If (lColour = col) Then
        Debug.Print ("Match")
    Else
        Debug.Print ("No Match")
    End If
    
    'Range("a1").Interior.color = lColour
    
End Sub

Sub returnMouseCoords()
Attribute returnMouseCoords.VB_ProcData.VB_Invoke_Func = "j\n14"
    Dim coords As POINT
    Call GetCursorPos(coords)
    Debug.Print ("x:" & coords.x & " y:" & coords.y)
End Sub

Sub returnMouseColor()
    Dim pLocation As POINT
    Dim lColour As Long
    
    Dim lDC As Variant
    lDC = GetWindowDC(0)
    
    lColour = GetPixel(lDC, pLocation.x, pLocation.y)
    Debug.Print (lColour)
End Sub

Sub returnMouseBoth()
    Dim pLocation As POINT
    Dim lColour As Long
    
    Dim lDC As Variant
    lDC = GetWindowDC(0)
    
    Call GetCursorPos(pLocation)
    
    lColour = GetPixel(lDC, pLocation.x, pLocation.y)
    Debug.Print ("Copied {C:" & lColour & " x:" & pLocation.x & " y:" & pLocation.y & "}")
    Sheets("Image").Range("z1").Value = ("C:" & lColour & " x:" & pLocation.x & " y:" & pLocation.y)
    Sheets("Image").Range("z1").Copy
    Sheets("Image").Range("z1").ClearContents
End Sub

Attribute VB_Name = "windowSnapper"
Option Explicit
Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare PtrSafe Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
 
Sub openNotepad()
  Shell "Notepad.exe", vbNormalFocus
End Sub
 
Sub moveNotepad()
  'Move Notepad to 0,0 and set its Width and height to 200,250 (pixels)
  Dim hNotepad As Long
  hNotepad = FindWindow("Notepad", vbNullString)
  MoveWindow hNotepad, 0, 0, 200, 250, 1
End Sub

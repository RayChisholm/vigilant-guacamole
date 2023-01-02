Attribute VB_Name = "imageGrabbing"
'--------------------------------------------------------------------------------------
'Image Grabbing, Moving, and Deleting
'

'   CopyScreen - Uses PrintScreen key to capture and paste a screenshot to a predetermined cell
Sub CopyScreen()
    Dim pasteArea As Range
    Set pasteArea = ThisWorkbook.Sheets("Image").Range("A1")
    
     Application.SendKeys "({1068})", True
         DoEvents
     ThisWorkbook.Sheets("Image").Activate
     ActiveSheet.Paste Destination:=pasteArea ' default target cell, where the topleft corner of our WHOLE screenshot is to be pasted
         Dim shp As Shape
         Dim h As Single, w As Single, l As Single, r As Single
     With ActiveSheet
         Set shp = .Shapes(.Shapes.Count)
     End With
     With shp
        h = -(635 - shp.Height)
        w = -(1225 - shp.Width)
        l = -(715 - shp.Height)
        r = -(2860 - shp.Width)
        .Height = 980
        .Width = 1080
        .LockAspectRatio = False
        With .PictureFormat
           '.CropRight = r
           '.CropLeft = w
           '.CropTop = h
           '.CropBottom = l
        End With
        'With .Line 'optional image borders
        '  .Weight = 1
        '  .DashStyle = msoLineSolid
        'End With
                ' Moving our cropped region to the target cell
        .Top = pasteArea.Top
        .Left = pasteArea.Left
    End With
End Sub

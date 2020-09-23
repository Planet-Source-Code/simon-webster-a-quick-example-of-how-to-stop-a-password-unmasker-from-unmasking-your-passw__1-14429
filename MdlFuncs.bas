Attribute VB_Name = "MdlFuncs"
'this function was originaly called when
'in mouse move or any mouse evens on the text
'feild you wanted but it doesnt work very well
'as it still lets the text to be unmasked when
'double clicked
'Public Function resettxt()
'
'   GetCursorPos pt2
'   Wnd = WindowFromPoint(pt2.X, pt2.Y)
'    If Wnd = frmmain.txt_pw.hwnd Then
'        SendMessage Wnd, EM_SETPASSWORDCHAR, Chr_Star_Ascii, 0
'        frmmain.txt_pw.Refresh
'   End If
'
'End Function

Public Function tmrreset()

        Wnd = frmmain.txt_pw.hwnd
        SendMessage Wnd, EM_SETPASSWORDCHAR, Chr_Star_Ascii, 0
        frmmain.txt_pw.Refresh

End Function

Function TickLoopreset()
   Etick = GetTickCount
     Do While GetTickCount - Etick < 5
        Wnd = frmmain.txt_pw.hwnd
        SendMessage Wnd, EM_SETPASSWORDCHAR, Chr_Star_Ascii, 0
        Etick = Etick + 5
        If bEnd = True Then Exit Do
        If bEnd = True Then Exit Function
        DoEvents
    Loop
End Function

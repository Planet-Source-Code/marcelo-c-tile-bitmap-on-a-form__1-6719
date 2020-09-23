Attribute VB_Name = "TileBitmapModule"

'// I am not getting rates or comments on my posts at planet
'// source code... so, if you are reading this and haven't
'// rated or commented, well, this will make you feel bad
'// enough to rate me next time :-)
'//
'// Cheers,
'// Marcelo
'//

'// I think the code is pretty much easy to follow,
'// its short, its fast, couldn't be faster i guess
'// after all we are using the famous bitblt api
'//
'// This code could easly be adapted to tile images on
'// pictureboxes.
'//
'// I made it specifically and restricted to a form for
'// posting it at planet for the newbies.



Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Const SRCCOPY = &HCC0020

Public Sub TileBitmap(Target As Form, Source As PictureBox)

BackupInformation_ScaleMode = Target.ScaleMode
BackupInformation_ScaleMode2 = Source.ScaleMode
Source.ScaleMode = 3
Target.ScaleMode = 3
Target.Cls
Target.AutoRedraw = True


For yDraw = 0 To Target.Height Step Source.ScaleHeight

    For Xdraw = 0 To Target.ScaleWidth Step Source.ScaleWidth

        BitBlt Target.hDC, Xdraw, yDraw, Source.ScaleWidth, Source.ScaleHeight, Source.hDC, 0, 0, SRCCOPY



    Next Xdraw

Next yDraw


Target.ScaleMode = BackupInformation_ScaleMode
Source.ScaleMode = BackupInformation_ScaleMode2

End Sub

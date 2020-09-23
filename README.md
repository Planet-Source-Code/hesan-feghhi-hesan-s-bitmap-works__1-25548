<div align="center">

## Hesan's Bitmap Works


</div>

### Description

This Program is used to add great effects to your images. Such as : Rotation, Blur, Noise, ... .
 
### More Info
 
Normal Pictures + Few Parameters

Great Pictures


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Hesan Feghhi](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/hesan-feghhi.md)
**Level**          |Intermediate
**User Rating**    |4.3 (13 globes from 3 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Graphics](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/graphics__1-46.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/hesan-feghhi-hesan-s-bitmap-works__1-25548/archive/master.zip)





### Source Code

```
Public Sub UnRGB(ByVal color As OLE_COLOR, ByRef R As Integer, ByRef G As Integer, ByRef B As Integer)
 B = color \ 65536
 G = (color \ 256) Mod 256
 R = color Mod 256
End Sub
'Destination Picture Cannont be the Source Picture itself
Public Sub bmp_Rotate(PicSrc As PictureBox, PicDest As PictureBox, degree As Single)
 yc = PicSrc.height / 30
 xc = PicSrc.width / 30
 R = Sqr(yc * yc)
 For x = 0 To (PicSrc.width - 80) / 5
 For y = 0 To (PicSrc.height - 80) / 5
 'Let the user see the result every 4000 pixels
 u = u + 1: If u = 8000 Then DoEvents: u = 0
 pix = PicSrc.Point(x * 5 - 5, y * 5 - 5)
 yy = y - PicSrc.height / 10
 xx = x - PicSrc.width / 10
 SinA = yy / R
 CosA = xx / R
 SinB = Sin(degree / 57.3248)
 CosB = Cos(degree / 57.3248)
 SinApB = SinA * CosB + CosA * SinB
 CosApB = CosA * CosB - SinA * SinB
 nx = R * CosApB
 ny = R * SinApB
 xorigin = PicDest.width / 2
 yorigin = PicDest.height / 2
 PicDest.PSet (xorigin + nx * 5, yorigin + ny * 5), pix
 Next y, x
End Sub
Public Sub bmp_Blur(PicSrc As PictureBox)
 For x = 1 To PicSrc.width / 15
 For y = 1 To PicSrc.height / 15
 'Let the user see the result every 4000 pixels
 u = u + 1: If u = 4000 Then DoEvents: u = 0
 p = PicSrc.Point(x * 15 - 15, y * 15 - 15)
 If p < 0 Then p = 0
 UnRGB p, R%, G%, B%
 p1 = PicSrc.Point((x - 1) * 15 - 15, y * 15 - 15)
 If p1 < 0 Then p1 = 0
 UnRGB p1, R1%, G1%, B1%
 p2 = PicSrc.Point((x + 1) * 15 - 15, y * 15 - 15)
 If p2 < 0 Then p2 = 0
 UnRGB p2, R2%, G2%, B2%
 p3 = PicSrc.Point(x * 15 - 15, (y - 1) * 15 - 15)
 UnRGB p3, R3%, G3%, B3%
 If p3 < 0 Then p3 = 0
 p4 = PicSrc.Point(x * 15 - 15, (y + 1) * 15 - 15)
 If p4 < 0 Then p4 = 0
 UnRGB p4, R4%, G4%, B4%
 R1 = (R1 + R) / 2
 R2 = (R2 + R) / 2
 R3 = (R3 + R) / 2
 R4 = (R4 + R) / 2
 G1 = (G1 + G) / 2
 G2 = (G2 + G) / 2
 G3 = (G3 + G) / 2
 G4 = (G4 + G) / 2
 B1 = (B1 + B) / 2
 B2 = (B2 + B) / 2
 B3 = (B3 + B) / 2
 B4 = (B4 + B) / 2
 PicSrc.PSet ((x - 1) * 15 - 15, y * 15 - 15), RGB(Fix(R1%), Fix(G1%), Fix(B1%))
 PicSrc.PSet ((x + 1) * 15 - 15, y * 15 - 15), RGB(Fix(R2%), Fix(G2%), Fix(B2%))
 PicSrc.PSet (x * 15 - 15, (y - 1) * 15 - 15), RGB(Fix(R3%), Fix(G3%), Fix(B3%))
 PicSrc.PSet (x * 15 - 15, (y + 1) * 15 - 15), RGB(Fix(R4%), Fix(G4%), Fix(B4%))
 Next y, x
End Sub
'Destination Cannont be the Source Picture itself
Public Sub bmp_FlipHorizontal(PicSrc As PictureBox, PicDest As PictureBox)
 For x = 0 To PicSrc.width / 15
 For y = 0 To PicSrc.height / 15
 'Let the user see the result every 4000 pixels
 u = u + 1: If u = 4000 Then DoEvents: u = 0
 pix = PicSrc.Point(x * 15, y * 15)
 PicDest.PSet (PicSrc.width - (x * 15) - 80, y * 15), pix
 Next y, x
End Sub
'Destination Cannont be the Source Picture itself
Public Sub bmp_FlipVertical(PicSrc As PictureBox, PicDest As PictureBox)
 For x = 0 To PicSrc.width / 15
 For y = 0 To PicSrc.height / 15
 'Let the user see the result every 4000 pixels
 u = u + 1: If u = 4000 Then DoEvents: u = 0
 pix = PicSrc.Point(x * 15, y * 15)
 PicDest.PSet (x * 15, PicSrc.height - (y * 15) - 80), pix
 Next y, x
End Sub
Public Sub bmp_GrayScale(PicSrc As PictureBox)
 For x = 0 To PicSrc.width / 15
 For y = 0 To PicSrc.height / 15
 'Let the user see the result every 4000 pixels
 u = u + 1: If u = 4000 Then DoEvents: u = 0
 pix = PicSrc.Point(x * 15, y * 15)
 UnRGB pix, R%, G%, B%
 GC = (R + G + B) / 3
 PicSrc.PSet (x * 15, y * 15), RGB(GC, GC, GC)
 Next y, x
End Sub
Public Sub bmp_AddNoise(PicSrc As PictureBox)
 For x = 1 To PicSrc.width / 50
 For y = 1 To PicSrc.height / 50
 'Let the user see the result every 4000 pixels
 u = u + 1: If u = 4000 Then DoEvents: u = 0
 UnRGB pix, R%, G%, B%
 GC = (R + G + B) / 3
 R = Fix(Rnd * 255)
 G = Fix(Rnd * 255)
 B = Fix(Rnd * 255)
 PicSrc.PSet (x * 50 + 45 * Rnd - 50, y * 50 + 45 * Rnd - 50), RGB(R, G, B)
 Next y, x
End Sub
```


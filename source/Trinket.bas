Attribute VB_Name = "Trinket"
'===============================================================================
' Макрос           : Trinket
' Версия           : 2022.01.26
' Сайты            : https://vk.com/elvin_macro
'                    https://github.com/elvin-nsk
' Автор            : elvin-nsk (me@elvin.nsk.ru, https://vk.com/elvin_macro)
'===============================================================================

Option Explicit

Const RELEASE As Boolean = True

'===============================================================================

Private Const TrimBoxWidth As Double = 27.4 'мм
Private Const TrimBoxHeight As Double = 36 'мм
Private Const FactorToFit As Double = 0.05
Private Const BackgroundEnlargeMult As Double = 1.4
Private Const TrimBoxBackgroundColor As String = "CMYK,USER,0,0,0,0"

'===============================================================================

Sub Start()

  If RELEASE Then On Error GoTo Catch
  
  If ActiveDocument Is Nothing Then
    VBA.MsgBox "Нет активного документа"
    Exit Sub
  End If
  ActiveDocument.Unit = cdrMillimeter
  
  Dim Bitmaps As ShapeRange
  Set Bitmaps = GetBitmaps
  If Bitmaps.Count = 0 Then Exit Sub
  
  Dim Stackables As New Collection
  
  lib_elvin.BoostStart "Trinket", RELEASE
  
  Dim Shape As Shape
  For Each Shape In Bitmaps
    Stackables.Add ProcessBitmap(Shape)
  Next Shape
  
  Dim StartingPoint As IPoint
  Set StartingPoint = FreePoint.Create(ActivePage.LeftX, ActivePage.TopY)
  
  Stacker.CreateAndStack _
          Stackables:=Stackables, _
          StartingPoint:=StartingPoint, _
          MaxWidth:=ActivePage.SizeWidth
  
Finally:
  lib_elvin.BoostFinish
  Exit Sub

Catch:
  VBA.MsgBox VBA.Err.Description, vbCritical, "Ошибка"
  Resume Finally

End Sub

'===============================================================================

Private Function GetBitmaps() As ShapeRange
  Set GetBitmaps = ActivePage.FindShapes(Type:=cdrBitmapShape).All
End Function

Private Function ProcessBitmap(ByVal Shape As Shape) As Stackable
  Dim TrimBox As Shape
  Set TrimBox = ActiveLayer.CreateRectangle2(0, 0, TrimBoxWidth, TrimBoxHeight)
  TrimBox.Fill.ApplyUniformFill CreateColor(TrimBoxBackgroundColor)
  If IsAlmostFit(Shape) Then
    FillInside Shape, TrimBox.BoundingBox
  Else
    AddBackground(Shape, TrimBox).AddToPowerClip TrimBox
    FitInside Shape, TrimBox.BoundingBox
  End If
  Shape.AddToPowerClip TrimBox
  Set ProcessBitmap = Stackable.Create(TrimBox)
End Function

Private Function IsAlmostFit(ByVal Shape As Shape) As Boolean
  Dim TargetHeight As Double
  TargetHeight = _
    lib_elvin.GetHeightKeepProportions(Shape.BoundingBox, TrimBoxWidth)
  If TrimBoxHeight > TargetHeight - (TargetHeight * FactorToFit) And _
     TrimBoxHeight < TargetHeight + (TargetHeight * FactorToFit) Then _
  IsAlmostFit = True
End Function

Private Function AddBackground _
                 (ByVal Shape As Shape, TrimBox As Shape) As Shape
  Set AddBackground = Shape.Duplicate
  With AddBackground
    .Bitmap.ApplyBitmapEffect _
      "GaussianBlur", _
      "GaussianBlurEffect GaussianBlurRadius=4230,GaussianBlurResampled=0"
    FillInside AddBackground, TrimBox.BoundingBox
    .SetSize .SizeWidth * BackgroundEnlargeMult
  End With
End Function

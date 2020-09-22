Attribute VB_Name = "modPrintDeclare"
Option Explicit

Public nClassItem As Integer
Public strPage() As String
Public bCustomPage As Boolean

Private Type qtPrinterPageInfo
  Width As Single
  Height As Single
  AvailWidth As Single
  AvailHeight As Single
  LeftM As Single
  RightM As Single
  TopM As Single
  BottomM As Single
  HeaderH As Single
  FooterH As Single
  HFAvailHeight As Single
End Type

Public mTwipMLeft As Single
Public mTwipMRight As Single
Public mTwipMTop As Single
Public mTwipMBottom As Single
Public qPage As qtPrinterPageInfo
Public bPageChange As Boolean


Public Function ConvertToTwip(ByVal eScale As qePrinterScale, ByVal sValue As Single, Optional PercentWidth As Boolean = True) As Single
   Dim sNewValue As Single
   
   Select Case eScale
      Case qePrinterScale.eTwip: sNewValue = sValue
      Case qePrinterScale.eInch: sNewValue = sValue * 1440
      Case qePrinterScale.eCentimetre: sNewValue = sValue * 567
      Case qePrinterScale.eMillimetre: sNewValue = sValue * 56.7
      Case qePrinterScale.ePercentage: If PercentWidth Then sNewValue = Printer.ScaleWidth / 100 * sValue Else sNewValue = Printer.ScaleHeight / 100 * sValue
   End Select
   
   ConvertToTwip = sNewValue
End Function

Public Function ConvertFromTwip(ByVal eScale As qePrinterScale, ByVal sValue As Single, Optional PercentWidth As Boolean = True) As Single
   Dim sNewValue As Single
   
   Select Case eScale
      Case qePrinterScale.eTwip: sNewValue = sValue
      Case qePrinterScale.eInch: sNewValue = sValue / 1440
      Case qePrinterScale.eCentimetre: sNewValue = sValue / 567
      Case qePrinterScale.eMillimetre: sNewValue = sValue / 56.7
      Case qePrinterScale.ePercentage: If PercentWidth Then sNewValue = (100 / Printer.ScaleWidth) * sValue Else sNewValue = (100 / Printer.ScaleHeight) * sValue
   End Select

   ConvertFromTwip = sNewValue
End Function

Public Function ConvertHTMColor(ByVal HTMColor) As Long
   Dim lRed, lGreen, lBlue As Long
   Dim sHexRed, sHexBlue, sHexGreen As String
   Dim lColor As Long
   
   HTMColor = UCase(HTMColor)
   
   If HTMColor Like "[#][0-9A-F][0-9A-F][0-9A-F][0-9A-F][0-9A-F][0-9A-F]" Then
      sHexRed = "&H" & Mid$(HTMColor, 2, 2)
      sHexGreen = "&H" & Mid$(HTMColor, 4, 2)
      sHexBlue = "&H" & Mid$(HTMColor, 6, 2)
      lBlue = CDec(sHexBlue) * &H10000
      lGreen = CDec(sHexGreen) * &H100
      lRed = CDec(sHexRed)
      lColor = lBlue + lGreen + lRed
   Else
      lColor = 0
   End If
   
   ConvertHTMColor = lColor
End Function


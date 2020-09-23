Attribute VB_Name = "Module1"
Option Explicit

 Private Declare Function CHOOSECOLOR Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As CHOOSECOLOR) As Long
   
   Private Type CHOOSECOLOR
   lStructSize As Long
   hwndOwner As Long
   hInstance As Long
   rgbResult As Long
   lpCustColors As String
   flags As Long
   lCustData As Long
   lpfnHook As Long
   lpTemplateName As String
End Type

Public CustomColors() As Byte
Public strArrayColor() As String
Dim cc As CHOOSECOLOR

Public Function ShowColor() As Long
   
   'set the structure size
   cc.lStructSize = Len(cc)
   'Set the owner
   cc.hwndOwner = Form1.hWnd
   'set the application's instance
   cc.hInstance = App.hInstance
   'set the custom colors (converted to Unicode)
   cc.lpCustColors = StrConv(CustomColors, vbUnicode)
   'no extra flags
   cc.flags = 0  'set to 0 = define custom colors unselected. 2= define custom colors selected
   
   'Show the 'Select Color'-dialog
   If CHOOSECOLOR(cc) <> 0 Then
      ShowColor = (cc.rgbResult)
      CustomColors = StrConv(cc.lpCustColors, vbFromUnicode)
   Else
      ShowColor = -1
   End If
   
End Function

Public Sub Fillit()
   Dim i As Integer
   
   ReDim CustomColors(0 To 16 * 4 - 1) As Byte
  
   For i = LBound(CustomColors) To UBound(CustomColors)
      CustomColors(i) = 0
   Next i
End Sub

Public Sub Save_Color(Optional ftype As Integer = 0)
   Dim FileName As String
   Dim Free As Long
   
   Free = FreeFile
   If ftype = 0 Then
      FileName = App.Path & "\ColorPal.txt"    'name of file to be saved to
      If FileName <> "" Then
         Open FileName For Binary As #Free
         Put #Free, , CustomColors   'this is the array to be saved
         Close #Free
      End If
   Else
      FileName = App.Path & "\CalColor.txt"
      Open FileName For Output As #Free
      Print #Free, Form1.ColorString
      Close #Free
   End If
End Sub

Public Sub Load_CustomColor()
   Dim FileName As String
   Dim Y As Integer
   Dim x As Long
   
   FileName = App.Path & "\" & "ColorPal.txt"  'file where array is saved

   If FileName <> "" Then
      Dim Free As Long

      Free = FreeFile

      Open FileName For Binary As #Free
      Get #Free, , CustomColors   'the array to load
      Close #Free
   End If

   'put values in array
   For Y = 0 To 15
      x = CustomColors(Y)
   Next Y
  
 End Sub

Public Sub LoadColorArray()
   
   Dim textfile As String
   Dim Free As Integer
   Dim x As Integer
   
   Free = FreeFile
   Open App.Path & "\CalColor.txt" For Input As #Free
   
   Do While Not EOF(Free)
      Line Input #Free, textfile
      strArrayColor = Split(textfile, ",")
      Loop
      
      Close #Free
      For x = LBound(strArrayColor) To UBound(strArrayColor)
          Form1.lblMonColor(x).BackColor = strArrayColor(x)
      Next x
   End Sub

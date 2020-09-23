Attribute VB_Name = "modColorOnly"
Option Explicit
   
   Const FW_NORMAL = 400
   Const DEFAULT_CHARSET = 1
   Const OUT_DEFAULT_PRECIS = 0
   Const CLIP_DEFAULT_PRECIS = 0
   Const DEFAULT_QUALITY = 0
   Const DEFAULT_PITCH = 0
   Const FF_ROMAN = 16
   Const CF_PRINTERFONTS = &H2
   Const CF_SCREENFONTS = &H1
   Const CF_BOTH = (CF_SCREENFONTS Or CF_PRINTERFONTS)
   Const CF_EFFECTS = &H100&
   Const CF_FORCEFONTEXIST = &H10000
   Const CF_INITTOLOGFONTSTRUCT = &H40&
   Const CF_LIMITSIZE = &H2000&
   Const REGULAR_FONTTYPE = &H400
   Const LF_FACESIZE = 32
   Const CCHDEVICENAME = 32
   Const CCHFORMNAME = 32
   Const GMEM_MOVEABLE = &H2
   Const GMEM_ZEROINIT = &H40
   Const DM_DUPLEX = &H1000&
   Const DM_ORIENTATION = &H1&
   Const PD_PRINTSETUP = &H40
   Const PD_DISABLEPRINTTOFILE = &H80000
   Private Type POINTAPI
   X As Long
   Y As Long
End Type

Private Type RECT
Left As Long
top As Long
Right As Long
Bottom As Long
End Type

Private Type PRINTDLG_TYPE
lStructSize As Long
hwndOwner As Long
hDevMode As Long
hDevNames As Long
hDC As Long
flags As Long
nFromPage As Integer
nToPage As Integer
nMinPage As Integer
nMaxPage As Integer
nCopies As Integer
hInstance As Long
lCustData As Long
lpfnPrintHook As Long
lpfnSetupHook As Long
lpPrintTemplateName As String
lpSetupTemplateName As String
hPrintTemplate As Long
hSetupTemplate As Long
End Type
Private Type DEVNAMES_TYPE
wDriverOffset As Integer
wDeviceOffset As Integer
wOutputOffset As Integer
wDefault As Integer
extra As String * 100
End Type
Private Type DEVMODE_TYPE
dmDeviceName As String * CCHDEVICENAME
dmSpecVersion As Integer
dmDriverVersion As Integer
dmSize As Integer
dmDriverExtra As Integer
dmFields As Long
dmOrientation As Integer
dmPaperSize As Integer
dmPaperLength As Integer
dmPaperWidth As Integer
dmScale As Integer
dmCopies As Integer
dmDefaultSource As Integer
dmPrintQuality As Integer
dmColor As Integer
dmDuplex As Integer
dmYResolution As Integer
dmTTOption As Integer
dmCollate As Integer
dmFormName As String * CCHFORMNAME
dmUnusedPadding As Integer
dmBitsPerPel As Integer
dmPelsWidth As Long
dmPelsHeight As Long
dmDisplayFlags As Long
dmDisplayFrequency As Long
End Type

Private Declare Function PrintDialog Lib "comdlg32.dll" Alias "PrintDlgA" (pPrintdlg As PRINTDLG_TYPE) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

Private Const CB_GETITEMHEIGHT      As Long = &H154

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
Public strArrayHoliday() As String
Dim cc As CHOOSECOLOR
'strArrayHoliday assignments
'strArrayHoliday(0) = Day of the week
'strArrayHoliday(1) = Month (Integer)
'strArrayHoliday(2) = either Option selected or day entered under Custom
'strArrayHoliday(3) = ID Tag FirstofMonth,LastofMonth etc
'strArrayHoliday(4) = Holiday name

Public Function ShowColor() As Long
   
   'set the structure size
   cc.lStructSize = Len(cc)
   'Set the owner
   cc.hwndOwner = Form1.hwnd
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

Public Sub ShowPrinter(frmOwner As Form, Optional PrintFlags As Long)
   '-> Code by Donald Grover
   Dim PrintDlg As PRINTDLG_TYPE
   Dim DevMode As DEVMODE_TYPE
   Dim DevName As DEVNAMES_TYPE
   
   Dim lpDevMode As Long, lpDevName As Long
   Dim bReturn As Integer
   Dim objPrinter As Printer, NewPrinterName As String
   
   ' Use PrintDialog to get the handle to a memory
   ' block with a DevMode and DevName structures
   
   PrintDlg.lStructSize = Len(PrintDlg)
   PrintDlg.hwndOwner = frmOwner.hwnd
   
   PrintDlg.flags = PrintFlags
   On Error Resume Next
   'Set the current orientation and duplex setting
   DevMode.dmDeviceName = Printer.DeviceName
   DevMode.dmSize = Len(DevMode)
   DevMode.dmFields = DM_ORIENTATION Or DM_DUPLEX
   DevMode.dmPaperWidth = Printer.Width
   DevMode.dmOrientation = Printer.Orientation
   DevMode.dmPaperSize = Printer.PaperSize
   DevMode.dmDuplex = Printer.Duplex
   On Error GoTo 0
   
   'Allocate memory for the initialization hDevMode structure
   'and copy the settings gathered above into this memory
   PrintDlg.hDevMode = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, Len(DevMode))
   lpDevMode = GlobalLock(PrintDlg.hDevMode)
   If lpDevMode > 0 Then
      CopyMemory ByVal lpDevMode, DevMode, Len(DevMode)
      bReturn = GlobalUnlock(PrintDlg.hDevMode)
   End If
   
   'Set the current driver, device, and port name strings
   With DevName
      .wDriverOffset = 8
      .wDeviceOffset = .wDriverOffset + 1 + Len(Printer.DriverName)
      .wOutputOffset = .wDeviceOffset + 1 + Len(Printer.Port)
      .wDefault = 0
   End With
   
   With Printer
      DevName.extra = .DriverName & Chr(0) & .DeviceName & Chr(0) & .Port & Chr(0)
   End With
   
   'Allocate memory for the initial hDevName structure
   'and copy the settings gathered above into this memory
   PrintDlg.hDevNames = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, Len(DevName))
   lpDevName = GlobalLock(PrintDlg.hDevNames)
   If lpDevName > 0 Then
      CopyMemory ByVal lpDevName, DevName, Len(DevName)
      bReturn = GlobalUnlock(lpDevName)
   End If
   
   'Call the print dialog up and let the user make changes
   If PrintDialog(PrintDlg) <> 0 Then
      
      'First get the DevName structure.
      lpDevName = GlobalLock(PrintDlg.hDevNames)
      CopyMemory DevName, ByVal lpDevName, 45
      bReturn = GlobalUnlock(lpDevName)
      GlobalFree PrintDlg.hDevNames
      
      'Next get the DevMode structure and set the printer
      'properties appropriately
      lpDevMode = GlobalLock(PrintDlg.hDevMode)
      CopyMemory DevMode, ByVal lpDevMode, Len(DevMode)
      bReturn = GlobalUnlock(PrintDlg.hDevMode)
      GlobalFree PrintDlg.hDevMode
      NewPrinterName = UCase$(Left(DevMode.dmDeviceName, InStr(DevMode.dmDeviceName, Chr$(0)) - 1))
      If Printer.DeviceName <> NewPrinterName Then
         For Each objPrinter In Printers
            If UCase$(objPrinter.DeviceName) = NewPrinterName Then
               Set Printer = objPrinter
               'set printer toolbar name at this point
            End If
            Next
         End If
         
         On Error Resume Next
         'Set printer object properties according to selections made
         'by user
         Printer.Copies = DevMode.dmCopies
         Printer.Duplex = DevMode.dmDuplex
         Printer.Orientation = DevMode.dmOrientation
         Printer.PaperSize = DevMode.dmPaperSize
         Printer.PrintQuality = DevMode.dmPrintQuality
         Printer.ColorMode = DevMode.dmColor
         Printer.PaperBin = DevMode.dmDefaultSource
         On Error GoTo 0
      End If
   End Sub

Public Sub Fillit()
   Dim i As Integer
   
   ReDim CustomColors(0 To 16 * 4 - 1) As Byte
   
   For i = LBound(CustomColors) To UBound(CustomColors)
      CustomColors(i) = 0
   Next i
End Sub

Public Sub Save_HolidayList()
   Dim FileName As String
   Dim Free As Long
   Free = FreeFile
   FileName = App.Path & "\HolidayList.txt"
   Open FileName For Output As #Free
   Print #Free, Form1.HolidayString
   Close #Free
End Sub

Public Sub Load_HolidayList()
   Dim textfile As String
   Dim Free As Integer
   Dim X As Integer
   
   Form1.lstHolList.Clear
   Free = FreeFile
   Open App.Path & "\HolidayList.txt" For Input As #Free
   
   Do While Not EOF(Free)
      Line Input #Free, textfile
      strArrayHoliday = Split(textfile, ",")
   Loop
      
   Close #Free
      For X = 0 To UBound(strArrayHoliday) - 1 Step 5
         Form1.lstHolList.AddItem strArrayHoliday(X) & "," & strArrayHoliday(X + 1) & "," & strArrayHoliday(X + 2) & "," & strArrayHoliday(X + 3) & "," & strArrayHoliday(X + 4)
      Next X
      Form1.lblTotal.Caption = "Total: " & Form1.lstHolList.ListCount
      Form1.Label14.Caption = Form1.lblTotal.Caption
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
   Dim X As Long
   
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
      X = CustomColors(Y)
   Next Y
End Sub

Public Sub LoadColorArray()
   
   Dim textfile As String
   Dim Free As Integer
   Dim X As Integer
   
   Free = FreeFile
   Open App.Path & "\CalColor.txt" For Input As #Free
   
   Do While Not EOF(Free)
      Line Input #Free, textfile
      strArrayColor = Split(textfile, ",")
      Loop
      
      Close #Free
      For X = 0 To 11
         Form1.lblMonColor(X).BackColor = strArrayColor(X)
      Next X
      Form1.lblHolColor.BackColor = strArrayColor(12)
      Form1.lblFontColor.BackColor = strArrayColor(13)
      Form1.lblBkColor.BackColor = strArrayColor(14)
      Form1.txtDateFrom.Text = strArrayColor(15)
      Form1.txtDateTo.Text = strArrayColor(16)
   End Sub

Public Sub SetDropDownHeight(pobjForm As Form, _
   pobjCombo As ComboBox, _
   plngNumItemsToDisplay As Long)
   
   Dim pt              As POINTAPI
   Dim rc              As RECT
   Dim lngSavedWidth   As Long
   Dim lngNewHeight    As Long
   Dim lngOldScaleMode As Long
   Dim lngItemHeight   As Long
   
   lngSavedWidth = pobjCombo.Width
   
   lngOldScaleMode = pobjForm.ScaleMode
   pobjForm.ScaleMode = vbPixels
   
   lngItemHeight = SendMessage(pobjCombo.hwnd, CB_GETITEMHEIGHT, 0, ByVal 0)
   
   lngNewHeight = lngItemHeight * (plngNumItemsToDisplay + 2)
   
   Call GetWindowRect(pobjCombo.hwnd, rc)
   pt.X = rc.Left
   pt.Y = rc.top
   
   Call ScreenToClient(pobjForm.hwnd, pt)
   
   Call MoveWindow(pobjCombo.hwnd, pt.X, pt.Y, pobjCombo.Height, lngNewHeight, True)
   
   pobjForm.ScaleMode = lngOldScaleMode
   pobjCombo.Width = lngSavedWidth
End Sub

Public Function GetFirstLastDate(ByVal fnDay As String, fnMonth As Integer, fnYear As Integer, fnFirstLast As Byte) As Date
   Dim tmpDate As Date, dLoop As Integer, addDate As Date, tmpLastDate As Date
   addDate = DateSerial(fnYear, fnMonth, 1)
   
   Select Case fnFirstLast
      Case 0
         
         If WeekdayName(Weekday(addDate)) = fnDay Then
            GetFirstLastDate = addDate
            Exit Function
         End If
         
         For dLoop = 1 To 7
            tmpDate = DateAdd("w", dLoop, addDate)
            
            If WeekdayName(Weekday(tmpDate)) = fnDay Then
               GetFirstLastDate = tmpDate
               Exit For
            End If
         Next dLoop
      Case 1
         tmpLastDate = DateAdd("d", -1, DateAdd("m", 1, addDate))
         
         If WeekdayName(Weekday(tmpLastDate)) = fnDay Then
            GetFirstLastDate = tmpLastDate
            Exit Function
         End If
         
         For dLoop = 7 To 1 Step -1
            tmpDate = DateAdd("w", -dLoop, tmpLastDate)
            
            If WeekdayName(Weekday(tmpDate)) = fnDay Then
               GetFirstLastDate = tmpDate
               Exit For
            End If
         Next dLoop
   End Select
End Function

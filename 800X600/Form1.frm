VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Single Page 12 Month Calendar"
   ClientHeight    =   7635
   ClientLeft      =   1770
   ClientTop       =   870
   ClientWidth     =   8700
   LinkTopic       =   "Form1"
   ScaleHeight     =   509
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   580
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton OptLabel 
      Caption         =   "Top Middle"
      Enabled         =   0   'False
      Height          =   255
      Index           =   1
      Left            =   7395
      TabIndex        =   41
      Top             =   5940
      Width           =   1095
   End
   Begin VB.OptionButton OptLabel 
      Caption         =   "Top Left"
      Enabled         =   0   'False
      Height          =   255
      Index           =   0
      Left            =   6285
      TabIndex        =   40
      Top             =   5925
      Value           =   -1  'True
      Width           =   1065
   End
   Begin VB.PictureBox picCal 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3105
      Left            =   9135
      ScaleHeight     =   190
      ScaleMode       =   0  'User
      ScaleWidth      =   194
      TabIndex        =   39
      Top             =   1215
      Visible         =   0   'False
      Width           =   2940
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   6270
      TabIndex        =   38
      Top             =   5580
      Visible         =   0   'False
      Width           =   2265
   End
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   13035
      Left            =   9285
      ScaleHeight     =   867
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   755
      TabIndex        =   37
      Top             =   -195
      Visible         =   0   'False
      Width           =   11355
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Add Label ( Do This Last)"
      Height          =   225
      Left            =   6285
      TabIndex        =   36
      Top             =   5280
      Width           =   2190
   End
   Begin VB.CheckBox chkColored 
      Caption         =   "Color Calendar"
      Height          =   375
      Left            =   7665
      TabIndex        =   35
      Top             =   90
      Width           =   960
   End
   Begin VB.Frame Frame1 
      Height          =   3540
      Left            =   7785
      TabIndex        =   10
      Top             =   465
      Width           =   705
      Begin VB.Label lblMonColor 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   0
         Left            =   60
         TabIndex        =   34
         Top             =   300
         Width           =   225
      End
      Begin VB.Label lblMonName 
         BackStyle       =   0  'Transparent
         Caption         =   "Jan"
         Height          =   225
         Index           =   0
         Left            =   330
         TabIndex        =   33
         Top             =   315
         Width           =   465
      End
      Begin VB.Label lblMonColor 
         Appearance      =   0  'Flat
         BackColor       =   &H00C00000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   1
         Left            =   60
         TabIndex        =   32
         Top             =   555
         Width           =   225
      End
      Begin VB.Label lblMonName 
         BackStyle       =   0  'Transparent
         Caption         =   "Feb"
         Height          =   225
         Index           =   1
         Left            =   330
         TabIndex        =   31
         Top             =   570
         Width           =   465
      End
      Begin VB.Label lblMonColor 
         Appearance      =   0  'Flat
         BackColor       =   &H00004000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   2
         Left            =   60
         TabIndex        =   30
         Top             =   810
         Width           =   225
      End
      Begin VB.Label lblMonName 
         BackStyle       =   0  'Transparent
         Caption         =   "Mar"
         Height          =   225
         Index           =   2
         Left            =   330
         TabIndex        =   29
         Top             =   825
         Width           =   465
      End
      Begin VB.Label lblMonColor 
         Appearance      =   0  'Flat
         BackColor       =   &H00008080&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   3
         Left            =   60
         TabIndex        =   28
         Top             =   1065
         Width           =   225
      End
      Begin VB.Label lblMonName 
         BackStyle       =   0  'Transparent
         Caption         =   "Apr"
         Height          =   225
         Index           =   3
         Left            =   330
         TabIndex        =   27
         Top             =   1080
         Width           =   465
      End
      Begin VB.Label lblMonColor 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   4
         Left            =   60
         TabIndex        =   26
         Top             =   1320
         Width           =   225
      End
      Begin VB.Label lblMonName 
         BackStyle       =   0  'Transparent
         Caption         =   "May"
         Height          =   225
         Index           =   4
         Left            =   330
         TabIndex        =   25
         Top             =   1335
         Width           =   465
      End
      Begin VB.Label lblMonColor 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   5
         Left            =   60
         TabIndex        =   24
         Top             =   1575
         Width           =   225
      End
      Begin VB.Label lblMonName 
         BackStyle       =   0  'Transparent
         Caption         =   "Jun"
         Height          =   225
         Index           =   5
         Left            =   330
         TabIndex        =   23
         Top             =   1590
         Width           =   465
      End
      Begin VB.Label lblMonColor 
         Appearance      =   0  'Flat
         BackColor       =   &H000000C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   6
         Left            =   60
         TabIndex        =   22
         Top             =   1845
         Width           =   225
      End
      Begin VB.Label lblMonName 
         BackStyle       =   0  'Transparent
         Caption         =   "Jul"
         Height          =   225
         Index           =   6
         Left            =   330
         TabIndex        =   21
         Top             =   1860
         Width           =   465
      End
      Begin VB.Label lblMonColor 
         Appearance      =   0  'Flat
         BackColor       =   &H00808000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   7
         Left            =   60
         TabIndex        =   20
         Top             =   2115
         Width           =   225
      End
      Begin VB.Label lblMonName 
         BackStyle       =   0  'Transparent
         Caption         =   "Aug"
         Height          =   225
         Index           =   7
         Left            =   330
         TabIndex        =   19
         Top             =   2130
         Width           =   465
      End
      Begin VB.Label lblMonColor 
         Appearance      =   0  'Flat
         BackColor       =   &H000040C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   8
         Left            =   60
         TabIndex        =   18
         Top             =   2385
         Width           =   225
      End
      Begin VB.Label lblMonName 
         BackStyle       =   0  'Transparent
         Caption         =   "Sep"
         Height          =   225
         Index           =   8
         Left            =   330
         TabIndex        =   17
         Top             =   2400
         Width           =   465
      End
      Begin VB.Label lblMonColor 
         Appearance      =   0  'Flat
         BackColor       =   &H00404080&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   9
         Left            =   60
         TabIndex        =   16
         Top             =   2655
         Width           =   225
      End
      Begin VB.Label lblMonName 
         BackStyle       =   0  'Transparent
         Caption         =   "Oct"
         Height          =   225
         Index           =   9
         Left            =   330
         TabIndex        =   15
         Top             =   2670
         Width           =   465
      End
      Begin VB.Label lblMonColor 
         Appearance      =   0  'Flat
         BackColor       =   &H00004040&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   10
         Left            =   60
         TabIndex        =   14
         Top             =   2925
         Width           =   225
      End
      Begin VB.Label lblMonName 
         BackStyle       =   0  'Transparent
         Caption         =   "Nov"
         Height          =   225
         Index           =   10
         Left            =   330
         TabIndex        =   13
         Top             =   2925
         Width           =   465
      End
      Begin VB.Label lblMonColor 
         Appearance      =   0  'Flat
         BackColor       =   &H00404000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   11
         Left            =   60
         TabIndex        =   12
         Top             =   3195
         Width           =   225
      End
      Begin VB.Label lblMonName 
         BackStyle       =   0  'Transparent
         Caption         =   "Dec"
         Height          =   225
         Index           =   11
         Left            =   315
         TabIndex        =   11
         Top             =   3210
         Width           =   465
      End
   End
   Begin VB.ComboBox cboYear 
      Height          =   315
      Left            =   6210
      TabIndex        =   6
      Text            =   "Combo1"
      Top             =   3300
      Width           =   1230
   End
   Begin VB.TextBox txtoffsetX 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   255
      Left            =   6975
      TabIndex        =   5
      Text            =   "0"
      Top             =   2070
      Width           =   360
   End
   Begin VB.TextBox txtoffsetY 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   255
      Left            =   6990
      TabIndex        =   4
      Text            =   "0"
      Top             =   2430
      Width           =   360
   End
   Begin VB.CommandButton cmdOffset 
      Caption         =   "Enter Offsets"
      Height          =   285
      Left            =   6255
      TabIndex        =   3
      Top             =   2820
      Width           =   1095
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      Height          =   495
      Left            =   6150
      TabIndex        =   2
      Top             =   240
      Width           =   1305
   End
   Begin VB.CommandButton cmdSaveBitmap 
      Caption         =   "Save as Bitmap"
      Height          =   525
      Left            =   6210
      TabIndex        =   1
      Top             =   1365
      Width           =   1185
   End
   Begin VB.TextBox txtSaveName 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   6195
      TabIndex        =   0
      Top             =   1065
      Width           =   1230
   End
   Begin VB.Shape Shape2 
      Height          =   7110
      Left            =   30
      Top             =   30
      Width           =   6015
   End
   Begin VB.Image Image1 
      Height          =   7110
      Left            =   30
      Stretch         =   -1  'True
      Top             =   30
      Visible         =   0   'False
      Width           =   6015
   End
   Begin VB.Label Label1 
      Caption         =   "X ="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   6645
      TabIndex        =   9
      Top             =   2085
      Width           =   300
   End
   Begin VB.Label Label2 
      Caption         =   "Y ="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   6645
      TabIndex        =   8
      Top             =   2445
      Width           =   330
   End
   Begin VB.Shape Shape1 
      Height          =   1215
      Left            =   6210
      Top             =   1950
      Width           =   1200
   End
   Begin VB.Label Label5 
      Caption         =   "Save as"
      Height          =   180
      Left            =   6420
      TabIndex        =   7
      Top             =   840
      Width           =   720
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

 Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
   Dim posx As Single
   Dim posy As Single
   Dim RanNo(11) As Long
   Public ColorString As String


Private Sub Form_Load()
   
   cboYear.AddItem "2007"
   cboYear.AddItem "2008"
   cboYear.AddItem "2009"
   cboYear.AddItem "2010"
   cboYear.AddItem "2011"
   cboYear.AddItem "2012"
   cboYear.AddItem "2013"
   cboYear.AddItem "2014"
   cboYear.AddItem "2015"
   cboYear.AddItem "2016"
   cboYear.AddItem "2017"
   cboYear.AddItem "2018"
   cboYear.AddItem "2019"
   cboYear.AddItem "2020"
   cboYear.AddItem "2021"
   cboYear.AddItem "2022"
   cboYear.AddItem "2023"
   cboYear.AddItem "2024"
   cboYear.AddItem "2025"
   
   cboYear.ListIndex = 2          'start with the year 2009
   picCal.AutoRedraw = True
   picCal.ScaleMode = 3
   
   Fillit
   Load_CustomColor
   LoadColorArray
   Me.Show
   Generate
   Image1.Visible = True
   Shape2.Visible = True
End Sub

Private Sub Form_Resize()
   picMain.Height = 950
   picCal.ScaleHeight = 190
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Dim x As Integer
   
   For x = 0 To 10
      ColorString = ColorString & lblMonColor(x).BackColor & ","
   Next x
   ColorString = ColorString & lblMonColor(11).BackColor
   Save_Color 0
   Save_Color 1
   Unload Me
End Sub

Private Sub Generate()
   Dim StDat As String
   Dim mMonth As Integer
   Dim yYear As Integer
   Dim StDG As String
   Dim cRow As Integer
   Dim cColumn As Integer
   Dim ctr As Integer
   Dim Goforit As Integer
   Dim xctr As Integer
   Dim sp As Integer
   Dim moname As String
   
   picCal.FontSize = 8
   picCal.FontBold = True
   
   For xctr = 1 To 12
      picCal.Picture = LoadPicture()
      'set name for each month
      Select Case xctr
         Case 1
            moname = "January"
            If chkColored = Checked Then picCal.ForeColor = lblMonColor(0).BackColor
         Case 2
            moname = "February"
            If chkColored = Checked Then picCal.ForeColor = lblMonColor(1).BackColor
         Case 3
            moname = "March"
            If chkColored = Checked Then picCal.ForeColor = lblMonColor(2).BackColor
         Case 4
            moname = "April"
            If chkColored = Checked Then picCal.ForeColor = lblMonColor(3).BackColor
         Case 5
            moname = "May"
            If chkColored = Checked Then picCal.ForeColor = lblMonColor(4).BackColor
         Case 6
            moname = "June"
            If chkColored = Checked Then picCal.ForeColor = lblMonColor(5).BackColor
         Case 7
            moname = "July"
            If chkColored = Checked Then picCal.ForeColor = lblMonColor(6).BackColor
         Case 8
            moname = "August"
            If chkColored = Checked Then picCal.ForeColor = lblMonColor(7).BackColor
         Case 9
            moname = "September"
            If chkColored = Checked Then picCal.ForeColor = lblMonColor(8).BackColor
         Case 10
            moname = "October"
            If chkColored = Checked Then picCal.ForeColor = lblMonColor(9).BackColor
             picCal.Line (0, picCal.ScaleHeight - 1)-(picCal.ScaleWidth, picCal.ScaleHeight - 1)
         Case 11
            moname = "November"
            If chkColored = Checked Then picCal.ForeColor = lblMonColor(10).BackColor
             picCal.Line (0, picCal.ScaleHeight - 1)-(picCal.ScaleWidth, picCal.ScaleHeight - 1)
         Case 12
            moname = "December"
            If chkColored = Checked Then picCal.ForeColor = lblMonColor(11).BackColor
             picCal.Line (0, picCal.ScaleHeight - 1)-(picCal.ScaleWidth, picCal.ScaleHeight - 1)
      End Select
      
      If chkColored = Unchecked Then picCal.ForeColor = vbBlack
      'draw two horizontal lines
      picCal.Line (0, 41)-(192, 41)
      picCal.Line (0, 27)-(192, 27)
      
      'set position to print month and year
      picCal.CurrentX = 57
      picCal.CurrentY = 10
      StDat = moname & "  " & cboYear.Text
      mMonth = Month(StDat)                               'get month value
      yYear = Year(StDat)                                    'get year value
      StDG = DatePart("W", StDat)
      Goforit = False                                             'first of month is set so print the rest
      picCal.Print StDat                                      'print month and year of each calendar
      
      'set position to print weekday header
      picCal.CurrentY = 28
      
      'print weekday header under month name and year
      picCal.Print Tab(0); "Su"; Tab(5); "Mo"; Tab(10); "Tu"; Tab(15); "We"; _
      Tab(20); "Th"; Tab(25); "Fr"; Tab(30); "Sa"
      
      'set position to start printing days
      picCal.CurrentY = 45
      For cRow = 1 To 6
         
         For cColumn = 1 To 7
            If (StDG = cColumn And Not Goforit) Then     'set first day of month
            ctr = 1
            picCal.Print Str(ctr);
            Goforit = True
         Else                                                        'get the rest of the days
            If (Goforit And IsDate(CStr(ctr) & "-" & mMonth & _
            "-" & yYear)) Then picCal.Print Str(ctr);
         End If
         
         Select Case cColumn                                  'set spacing between the days
               Case 1: picCal.Print Tab(5);
               Case 2: picCal.Print Tab(10);
               Case 3: picCal.Print Tab(15);
               Case 4: picCal.Print Tab(20);
               Case 5: picCal.Print Tab(25);
               Case 6: picCal.Print Tab(30);
         End Select
         
         ctr = ctr + 1
      Next cColumn
      picCal.Print            'line space
      picCal.Print            'line space
   Next cRow
 
   picCal.Picture = picCal.Image                                'render picture so we can copy it
   
   'position and print each month
   If xctr = 1 Then BitBlt picMain.hDC, 40 + posx, 40 + posy, picCal.Width, picCal.Height, picCal.hDC, 0, 0, vbSrcCopy
   If xctr = 2 Then BitBlt picMain.hDC, 280 + posx, 40 + posy, picCal.Width, picCal.Height, picCal.hDC, 0, 0, vbSrcCopy
   If xctr = 3 Then BitBlt picMain.hDC, 520 + posx, 40 + posy, picCal.Width, picCal.Height, picCal.hDC, 0, 0, vbSrcCopy
   If xctr = 4 Then BitBlt picMain.hDC, 40 + posx, 240 + posy, picCal.Width, picCal.Height, picCal.hDC, 0, 0, vbSrcCopy
   If xctr = 5 Then BitBlt picMain.hDC, 280 + posx, 240 + posy, picCal.Width, picCal.Height, picCal.hDC, 0, 0, vbSrcCopy
   If xctr = 6 Then BitBlt picMain.hDC, 520 + posx, 240 + posy, picCal.Width, picCal.Height, picCal.hDC, 0, 0, vbSrcCopy
   If xctr = 7 Then BitBlt picMain.hDC, 40 + posx, 440 + posy, picCal.Width, picCal.Height, picCal.hDC, 0, 0, vbSrcCopy
   If xctr = 8 Then BitBlt picMain.hDC, 280 + posx, 440 + posy, picCal.Width, picCal.Height, picCal.hDC, 0, 0, vbSrcCopy
   If xctr = 9 Then BitBlt picMain.hDC, 520 + posx, 440 + posy, picCal.Width, picCal.Height, picCal.hDC, 0, 0, vbSrcCopy
   If xctr = 10 Then BitBlt picMain.hDC, 40 + posx, 640 + posy, picCal.Width, picCal.Height, picCal.hDC, 0, 0, vbSrcCopy
   If xctr = 11 Then BitBlt picMain.hDC, 280 + posx, 640 + posy, picCal.Width, picCal.Height, picCal.hDC, 0, 0, vbSrcCopy
   If xctr = 12 Then BitBlt picMain.hDC, 520 + posx, 640 + posy, picCal.Width, picCal.Height, picCal.hDC, 0, 0, vbSrcCopy
Next xctr
If Check1.Value = Checked Then                              'include name
   If OptLabel(0).Value = True Then
      picMain.CurrentX = 40 + posx
      picMain.CurrentY = 10
   Else
      picMain.CurrentX = 320 + posx
      picMain.CurrentY = 10
   End If
   picMain.Print Text1.Text
End If
Preview
End Sub

Private Sub cboYear_Click()
   picMain.Cls
   Generate
End Sub

Private Sub Preview()
   Image1.Width = picMain.Width / 1.9
   Image1.Height = picMain.Height / 1.9
   Image1.Picture = picMain.Image
   Shape2.Width = Image1.Width
   Shape2.Height = Image1.Height
   Shape2.Visible = Image1.Visible
End Sub

Private Sub cmdOffset_Click()
   posx = Int(txtoffsetX.Text)
   posy = Int(txtoffsetY.Text)
   picMain.Cls
   Generate
End Sub

Private Sub Check1_Click()
   If Check1.Value = Checked Then
      Text1.Visible = True
      Text1.SetFocus
      OptLabel(0).Enabled = True
      OptLabel(1).Enabled = True
   Else
      Text1.Visible = False
      Text1.Text = ""
      OptLabel(0).Enabled = False
      OptLabel(1).Enabled = False
   End If
End Sub

Private Sub chkColored_Click()
   picMain.Cls
   Generate
   'update preview
   Image1.Width = picMain.Width / 1.9
   Image1.Height = picMain.Height / 1.9
   Image1.Picture = picMain.Image
   Shape2.Width = Image1.Width
   Shape2.Height = Image1.Height
End Sub

Private Sub cmdPrint_Click()
   Generate
   picMain.Picture = picMain.Image                  'render the picture
   Printer.PaintPicture picMain.Picture, 0, 0
   Printer.EndDoc
   
   picMain.Picture = LoadPicture()                     'unrender the picture
   Generate
   Check1.Value = Unchecked
   Check1_Click
End Sub

Private Sub cmdSaveBitmap_Click()
   If txtSaveName.Text = "" Then GoTo here
   Generate
   picMain.Picture = picMain.Image
   SavePicture picMain.Picture, App.Path & "\" & txtSaveName.Text & ".bmp"
   picMain.Picture = LoadPicture()
   Generate
   Check1.Value = Unchecked
   Check1_Click
   MsgBox "Picture saved as a bitmap at " & App.Path & "\" & txtSaveName.Text & ".bmp"
   txtSaveName.Text = ""
   Exit Sub
here:
   MsgBox "Calendar not saved, Please enter a Name for it."
End Sub

Private Sub lblMonColor_Click(Index As Integer)
   Dim sure As Long
   sure = ShowColor
   If sure = -1 Then Exit Sub
   lblMonColor(Index).BackColor = sure
   chkColored.Value = 0
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Generate
      Check1.Value = 0
   End If
End Sub

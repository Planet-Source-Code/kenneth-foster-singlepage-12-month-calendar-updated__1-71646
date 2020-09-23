VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Single Page 12 month Calendar by Ken Foster"
   ClientHeight    =   10020
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14250
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   668
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   950
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDateRange 
      Caption         =   "Date Range"
      Height          =   465
      Left            =   13230
      TabIndex        =   84
      Top             =   8910
      Width           =   990
   End
   Begin VB.Frame Frame3 
      Caption         =   "Date Range"
      Height          =   1215
      Left            =   12915
      TabIndex        =   79
      Top             =   1650
      Visible         =   0   'False
      Width           =   1200
      Begin VB.CommandButton cmdChange 
         Caption         =   "Change"
         Height          =   285
         Left            =   435
         TabIndex        =   85
         Top             =   855
         Width           =   675
      End
      Begin VB.TextBox txtDateTo 
         Height          =   285
         Left            =   45
         TabIndex        =   81
         Text            =   "20"
         Top             =   840
         Width           =   330
      End
      Begin VB.TextBox txtDateFrom 
         Height          =   285
         Left            =   435
         TabIndex        =   80
         Text            =   "2006"
         Top             =   315
         Width           =   705
      End
      Begin VB.Label Label16 
         Caption         =   "How many"
         Height          =   225
         Left            =   30
         TabIndex        =   83
         Top             =   615
         Width           =   780
      End
      Begin VB.Label Label15 
         Caption         =   "Start:"
         Height          =   195
         Left            =   30
         TabIndex        =   82
         Top             =   360
         Width           =   540
      End
   End
   Begin VB.CommandButton cmdCreateList 
      Caption         =   "Create Holiday List"
      Height          =   705
      Left            =   13230
      TabIndex        =   66
      Top             =   8115
      Width           =   1005
   End
   Begin VB.CommandButton cmdPrinterSelect 
      Caption         =   "Printer Setup"
      Height          =   435
      Left            =   12900
      TabIndex        =   53
      Top             =   570
      Width           =   1215
   End
   Begin VB.CheckBox chkLatin 
      Caption         =   "Monday as first   day of week"
      Height          =   390
      Left            =   11565
      TabIndex        =   52
      Top             =   3030
      Width           =   1395
   End
   Begin VB.CheckBox chkHol 
      Caption         =   "Show Holidays"
      Height          =   255
      Left            =   11565
      TabIndex        =   45
      Top             =   3465
      Width           =   1410
   End
   Begin VB.CheckBox chkColored 
      Caption         =   "Color Calendar"
      Height          =   375
      Left            =   12075
      TabIndex        =   13
      Top             =   4200
      Width           =   960
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "Print Preview"
      Height          =   450
      Left            =   12885
      TabIndex        =   43
      Top             =   60
      Width           =   1230
   End
   Begin VB.TextBox txtSaveName 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   11520
      TabIndex        =   41
      Top             =   855
      Width           =   1230
   End
   Begin VB.Frame Frame1 
      Height          =   3540
      Left            =   12075
      TabIndex        =   14
      Top             =   4500
      Width           =   705
      Begin VB.Label lblMonName 
         BackStyle       =   0  'Transparent
         Caption         =   "Dec"
         Height          =   225
         Index           =   11
         Left            =   315
         TabIndex        =   38
         Top             =   3210
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
         TabIndex        =   37
         Top             =   3195
         Width           =   225
      End
      Begin VB.Label lblMonName 
         BackStyle       =   0  'Transparent
         Caption         =   "Nov"
         Height          =   225
         Index           =   10
         Left            =   330
         TabIndex        =   36
         Top             =   2925
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
         TabIndex        =   35
         Top             =   2925
         Width           =   225
      End
      Begin VB.Label lblMonName 
         BackStyle       =   0  'Transparent
         Caption         =   "Oct"
         Height          =   225
         Index           =   9
         Left            =   330
         TabIndex        =   34
         Top             =   2670
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
         TabIndex        =   33
         Top             =   2655
         Width           =   225
      End
      Begin VB.Label lblMonName 
         BackStyle       =   0  'Transparent
         Caption         =   "Sep"
         Height          =   225
         Index           =   8
         Left            =   330
         TabIndex        =   32
         Top             =   2400
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
         TabIndex        =   31
         Top             =   2385
         Width           =   225
      End
      Begin VB.Label lblMonName 
         BackStyle       =   0  'Transparent
         Caption         =   "Aug"
         Height          =   225
         Index           =   7
         Left            =   330
         TabIndex        =   30
         Top             =   2130
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
         TabIndex        =   29
         Top             =   2115
         Width           =   225
      End
      Begin VB.Label lblMonName 
         BackStyle       =   0  'Transparent
         Caption         =   "Jul"
         Height          =   225
         Index           =   6
         Left            =   330
         TabIndex        =   28
         Top             =   1860
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
         TabIndex        =   27
         Top             =   1845
         Width           =   225
      End
      Begin VB.Label lblMonName 
         BackStyle       =   0  'Transparent
         Caption         =   "Jun"
         Height          =   225
         Index           =   5
         Left            =   330
         TabIndex        =   26
         Top             =   1590
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
         TabIndex        =   25
         Top             =   1575
         Width           =   225
      End
      Begin VB.Label lblMonName 
         BackStyle       =   0  'Transparent
         Caption         =   "May"
         Height          =   225
         Index           =   4
         Left            =   330
         TabIndex        =   24
         Top             =   1335
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
         TabIndex        =   23
         Top             =   1320
         Width           =   225
      End
      Begin VB.Label lblMonName 
         BackStyle       =   0  'Transparent
         Caption         =   "Apr"
         Height          =   225
         Index           =   3
         Left            =   330
         TabIndex        =   22
         Top             =   1080
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
         TabIndex        =   21
         Top             =   1065
         Width           =   225
      End
      Begin VB.Label lblMonName 
         BackStyle       =   0  'Transparent
         Caption         =   "Mar"
         Height          =   225
         Index           =   2
         Left            =   330
         TabIndex        =   20
         Top             =   825
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
         TabIndex        =   19
         Top             =   810
         Width           =   225
      End
      Begin VB.Label lblMonName 
         BackStyle       =   0  'Transparent
         Caption         =   "Feb"
         Height          =   225
         Index           =   1
         Left            =   330
         TabIndex        =   18
         Top             =   570
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
         TabIndex        =   17
         Top             =   555
         Width           =   225
      End
      Begin VB.Label lblMonName 
         BackStyle       =   0  'Transparent
         Caption         =   "Jan"
         Height          =   225
         Index           =   0
         Left            =   330
         TabIndex        =   16
         Top             =   315
         Width           =   465
      End
      Begin VB.Label lblMonColor 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   0
         Left            =   60
         TabIndex        =   15
         Top             =   300
         Width           =   225
      End
   End
   Begin VB.CheckBox chkLabel 
      Caption         =   "Add Label"
      Height          =   225
      Left            =   12960
      TabIndex        =   12
      Top             =   6015
      Width           =   1050
   End
   Begin VB.CommandButton cmdSaveBitmap 
      Caption         =   "Save as Bitmap"
      Height          =   525
      Left            =   11535
      TabIndex        =   10
      Top             =   1155
      Width           =   1185
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      Height          =   435
      Left            =   11505
      TabIndex        =   9
      Top             =   75
      Width           =   1260
   End
   Begin VB.CommandButton cmdOffset 
      Caption         =   "Enter Offsets"
      Height          =   285
      Left            =   11580
      TabIndex        =   6
      Top             =   2610
      Width           =   1095
   End
   Begin VB.TextBox txtoffsetY 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   255
      Left            =   12315
      TabIndex        =   5
      Text            =   "0"
      Top             =   2220
      Width           =   360
   End
   Begin VB.TextBox txtoffsetX 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   255
      Left            =   12300
      TabIndex        =   4
      Text            =   "0"
      Top             =   1860
      Width           =   360
   End
   Begin VB.VScrollBar VS1 
      Height          =   2430
      LargeChange     =   100
      Left            =   11580
      Max             =   290
      SmallChange     =   50
      TabIndex        =   3
      Top             =   5100
      Width           =   330
   End
   Begin VB.ComboBox cboYear 
      Height          =   315
      Left            =   13215
      TabIndex        =   2
      Top             =   1125
      Width           =   885
   End
   Begin VB.PictureBox picCal 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2880
      Left            =   8145
      ScaleHeight     =   190
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   202
      TabIndex        =   1
      Top             =   345
      Visible         =   0   'False
      Width           =   3060
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
      Left            =   45
      ScaleHeight     =   867
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   755
      TabIndex        =   0
      Top             =   15
      Width           =   11355
      Begin VB.Frame Frame2 
         BackColor       =   &H0080C0FF&
         Caption         =   "Create / Edit List"
         Height          =   3975
         Left            =   5430
         TabIndex        =   54
         Top             =   3285
         Visible         =   0   'False
         Width           =   5505
         Begin VB.TextBox txtHolidayName 
            BackColor       =   &H00C0E0FF&
            Height          =   285
            Left            =   135
            TabIndex        =   75
            Top             =   3195
            Width           =   2250
         End
         Begin VB.CommandButton cmdExit 
            Caption         =   "Exit"
            Height          =   300
            Left            =   4830
            TabIndex        =   74
            Top             =   315
            Width           =   525
         End
         Begin VB.OptionButton optAlwaysDay 
            BackColor       =   &H0080C0FF&
            Caption         =   "Last (select day) of Month"
            Height          =   195
            Index           =   7
            Left            =   165
            TabIndex        =   71
            Top             =   2715
            Width           =   2130
         End
         Begin VB.TextBox txtCustDay 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            Enabled         =   0   'False
            Height          =   270
            Left            =   405
            TabIndex        =   70
            Text            =   "1"
            Top             =   2325
            Width           =   270
         End
         Begin VB.ComboBox cboHolMonth 
            BackColor       =   &H00C0E0FF&
            Height          =   315
            Left            =   3315
            TabIndex        =   69
            Text            =   "1 = Jan"
            Top             =   315
            Width           =   1290
         End
         Begin VB.ComboBox cboHolDay 
            BackColor       =   &H00C0E0FF&
            Height          =   315
            Left            =   1695
            TabIndex        =   68
            Text            =   "Sunday"
            Top             =   315
            Width           =   1305
         End
         Begin VB.CommandButton cmdDeleteHol 
            Caption         =   "Delete"
            Height          =   345
            Left            =   1305
            TabIndex        =   67
            Top             =   3555
            Width           =   1020
         End
         Begin VB.ListBox lstHolList 
            BackColor       =   &H00C0E0FF&
            Height          =   2985
            Left            =   2415
            TabIndex        =   63
            Top             =   705
            Width           =   2925
         End
         Begin VB.CommandButton cmdEnterHolDay 
            Caption         =   "Enter"
            Height          =   345
            Left            =   180
            TabIndex        =   62
            Top             =   3555
            Width           =   1005
         End
         Begin VB.OptionButton optAlwaysDay 
            BackColor       =   &H0080C0FF&
            Caption         =   "Custom"
            Height          =   195
            Index           =   6
            Left            =   165
            TabIndex        =   61
            Top             =   2070
            Width           =   885
         End
         Begin VB.OptionButton optAlwaysDay 
            BackColor       =   &H0080C0FF&
            Caption         =   "Always Last day of Month"
            Height          =   240
            Index           =   5
            Left            =   165
            TabIndex        =   60
            Top             =   1785
            Width           =   2115
         End
         Begin VB.OptionButton optAlwaysDay 
            BackColor       =   &H0080C0FF&
            Caption         =   "Always First day of Month"
            Height          =   240
            Index           =   4
            Left            =   165
            TabIndex        =   59
            Top             =   1530
            Width           =   2130
         End
         Begin VB.OptionButton optAlwaysDay 
            BackColor       =   &H0080C0FF&
            Caption         =   "Always the 4th"
            Height          =   240
            Index           =   3
            Left            =   165
            TabIndex        =   58
            Top             =   1170
            Width           =   1350
         End
         Begin VB.OptionButton optAlwaysDay 
            BackColor       =   &H0080C0FF&
            Caption         =   "Always the 3rd"
            Height          =   240
            Index           =   2
            Left            =   165
            TabIndex        =   57
            Top             =   900
            Width           =   1335
         End
         Begin VB.OptionButton optAlwaysDay 
            BackColor       =   &H0080C0FF&
            Caption         =   "Always the 2nd"
            Height          =   240
            Index           =   1
            Left            =   165
            TabIndex        =   56
            Top             =   630
            Width           =   1380
         End
         Begin VB.OptionButton optAlwaysDay 
            BackColor       =   &H0080C0FF&
            Caption         =   "Always the 1st"
            Height          =   225
            Index           =   0
            Left            =   165
            TabIndex        =   55
            Top             =   360
            Value           =   -1  'True
            Width           =   1395
         End
         Begin VB.Label Label13 
            BackColor       =   &H0080C0FF&
            Caption         =   "Click and drag to arrange list."
            Height          =   195
            Left            =   2835
            TabIndex        =   77
            Top             =   3690
            Width           =   2145
         End
         Begin VB.Label Label9 
            BackColor       =   &H0080C0FF&
            Caption         =   "Description:"
            Height          =   255
            Left            =   150
            TabIndex        =   76
            Top             =   2985
            Width           =   1065
         End
         Begin VB.Label lblTotal 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Total:"
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   1635
            TabIndex        =   73
            Top             =   1245
            Width           =   720
         End
         Begin VB.Label Label12 
            BackColor       =   &H0080C0FF&
            Caption         =   "(Enter Day and Go Select Month)"
            Enabled         =   0   'False
            Height          =   420
            Left            =   765
            TabIndex        =   72
            Top             =   2280
            Width           =   1410
         End
         Begin VB.Shape Shape7 
            Height          =   1455
            Left            =   135
            Top             =   1500
            Width           =   2175
         End
         Begin VB.Shape Shape6 
            Height          =   1185
            Left            =   135
            Top             =   285
            Width           =   1470
         End
         Begin VB.Label Label11 
            BackColor       =   &H0080C0FF&
            Caption         =   "Select Month Here"
            Height          =   180
            Left            =   3315
            TabIndex        =   65
            Top             =   120
            Width           =   1350
         End
         Begin VB.Label Label10 
            BackColor       =   &H0080C0FF&
            Caption         =   "Select Day Here"
            Height          =   195
            Left            =   1770
            TabIndex        =   64
            Top             =   120
            Width           =   1230
         End
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         TabIndex        =   11
         Top             =   45
         Visible         =   0   'False
         Width           =   3465
      End
      Begin VB.Shape Shape2 
         Height          =   7365
         Left            =   300
         Top             =   540
         Visible         =   0   'False
         Width           =   4770
      End
      Begin VB.Image Image1 
         Height          =   8040
         Left            =   285
         Stretch         =   -1  'True
         Top             =   540
         Visible         =   0   'False
         Width           =   5190
      End
   End
   Begin VB.Label Label14 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total:"
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   12015
      TabIndex        =   78
      Top             =   3720
      Width           =   765
   End
   Begin VB.Shape Shape5 
      Height          =   990
      Left            =   11535
      Top             =   3015
      Width           =   1455
   End
   Begin VB.Shape Shape4 
      Height          =   1410
      Left            =   11475
      Top             =   8115
      Width           =   1725
   End
   Begin VB.Shape Shape3 
      Height          =   2535
      Left            =   12825
      Top             =   5490
      Width           =   1365
   End
   Begin VB.Label Label8 
      Caption         =   "Note:Printing calendar with background other than white uses more ink."
      Height          =   795
      Left            =   11535
      TabIndex        =   51
      Top             =   8700
      Width           =   1680
   End
   Begin VB.Label Label7 
      Caption         =   "BackColor"
      Height          =   195
      Left            =   11805
      TabIndex        =   50
      Top             =   8445
      Width           =   810
   End
   Begin VB.Label lblBkColor 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   11535
      TabIndex        =   49
      Top             =   8445
      Width           =   210
   End
   Begin VB.Label lblFontColor 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   11535
      TabIndex        =   48
      Top             =   8175
      Width           =   210
   End
   Begin VB.Label lblFC 
      Caption         =   "Font Color"
      Height          =   240
      Left            =   11790
      TabIndex        =   47
      Top             =   8160
      Width           =   780
   End
   Begin VB.Label lblHolColor 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   11580
      TabIndex        =   46
      Top             =   3765
      Width           =   180
   End
   Begin VB.Label Label6 
      Caption         =   "Once set, you can add another label."
      Height          =   645
      Left            =   12945
      TabIndex        =   44
      Top             =   7380
      Width           =   1200
   End
   Begin VB.Label Label5 
      Caption         =   "Save as"
      Height          =   180
      Left            =   11745
      TabIndex        =   42
      Top             =   630
      Width           =   720
   End
   Begin VB.Label Label4 
      Caption         =   "   DO LABELS           LAST"
      ForeColor       =   &H00FF0000&
      Height          =   405
      Left            =   12900
      TabIndex        =   40
      Top             =   5565
      Width           =   1140
   End
   Begin VB.Label Label3 
      Caption         =   "Left click and hold on name label to drag and drop. Press Enter Bar to set."
      Height          =   1020
      Left            =   12945
      TabIndex        =   39
      Top             =   6300
      Width           =   1140
   End
   Begin VB.Shape Shape1 
      Height          =   1215
      Left            =   11535
      Top             =   1740
      Width           =   1200
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
      Left            =   11970
      TabIndex        =   8
      Top             =   2235
      Width           =   330
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
      Left            =   11970
      TabIndex        =   7
      Top             =   1875
      Width           =   300
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
   'Ken Foster
   '2009
   'use or abuse anyway you like
   
   Option Explicit
   'used with dragging items in listbox
   Private CurrentIndex As Long
   Private LastIndex As Long
   Private LastString As String
   Private LBHwnd As Long
   Private Dragging As Boolean
   Private DragIndex As Long
   Private DragText As String
   Private Const LB_ITEMFROMPOINT As Long = &H1A9
   Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
   'more varibles
   Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
   Dim PosX As Single   'used for the offset of calendars
   Dim PosY As Single   'used for the offset of calendars
   Dim yYear As Integer
   Public ColorString As String
   Public HolidayString As String  'stores holiday info for saving
   Dim Oldx As Single   'used to move textbox
   Dim Oldy As Single   'used to move textbox

Private Sub Form_Load()
   Dim X As Integer
   LoadColorArray
   cboYear.AddItem txtDateFrom.Text
   For X = 1 To txtDateTo.Text - 1
      cboYear.AddItem txtDateFrom.Text + X
   Next X
   cboYear.Text = txtDateFrom.Text
   
   For X = 1 To 7
      cboHolDay.AddItem WeekdayName(X)
   Next X
   
   For X = 1 To 12
      cboHolMonth.AddItem X & " = " & MonthName(X, True)
   Next X
   SetDropDownHeight Me, cboYear, Int(txtDateTo.Text)
   SetDropDownHeight Me, cboHolMonth, 12
   cboHolMonth.top = 315
   cboHolMonth.Left = 3315
   picCal.AutoRedraw = True
   picCal.ScaleMode = 3
   LBHwnd = lstHolList.hwnd    'used with dragging items in listbox
   Fillit
   Load_CustomColor
  ' LoadColorArray
   Load_HolidayList
   picMain.BackColor = lblBkColor.BackColor
   picCal.BackColor = lblBkColor.BackColor
   Generate
End Sub

Private Sub Form_Resize()
   picMain.top = 4
   picMain.Left = 4
   picMain.Height = 950
   picCal.ScaleHeight = 190
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Dim X As Integer
   
   For X = 0 To 11
      ColorString = ColorString & lblMonColor(X).BackColor & ","
   Next X
   ColorString = ColorString & lblHolColor.BackColor & "," & lblFontColor.BackColor & "," & lblBkColor.BackColor & "," & txtDateFrom.Text & "," & txtDateTo.Text
   Save_Color 0
   Save_Color 1
   Unload Me
End Sub

Private Sub SaveHolDay_List()
   Dim X As Integer
   
   If lstHolList.ListCount = -1 Then Exit Sub
   HolidayString = ""
   For X = 0 To lstHolList.ListCount - 1
      If X = lstHolList.ListCount - 1 Then
         HolidayString = HolidayString & lstHolList.List(X)
      Else
         HolidayString = HolidayString & lstHolList.List(X) & ","
      End If
   Next X
   
   Save_HolidayList
   lstHolList.Clear
   Load_HolidayList
End Sub

Private Sub Generate()
   Dim StDat As String
   Dim mMonth As Integer
   Dim StDG As String
   Dim cRow As Integer
   Dim cColumn As Integer
   Dim ctr As Integer
   Dim Goforit As Boolean
   Dim xctr As Integer     'month counter
   Dim moname As String
   Dim i As Integer
   Dim X As Single
   Dim Y As Single
   Dim Hol As Integer
   Dim dn As Integer
   Dim dname As String
   Dim DaysInMonth As Integer
   
   picCal.FontSize = 8
   picCal.FontBold = True
   picCal.ForeColor = lblFontColor.BackColor
   
   For xctr = 1 To 12
      picCal.Picture = LoadPicture()
      'set name for each month
      moname = MonthName(xctr, False)
      picCal.Line (0, picCal.ScaleHeight - 1)-(picCal.ScaleWidth, picCal.ScaleHeight - 1)   'draw line at bottom of each calendar
      If chkColored = Checked Then picCal.ForeColor = lblMonColor(xctr - 1).BackColor
      
      'draw two horizontal lines
      picCal.Line (0, 41)-(192, 41)
      picCal.Line (0, 27)-(192, 27)
      
      'set position to print month and year
      picCal.CurrentX = 57
      picCal.CurrentY = 10
      StDat = moname & "  " & cboYear.Text
      mMonth = Month(StDat)                               'get month value
      yYear = Year(StDat)                                    'get year value
      DaysInMonth = Day(-1 + DateAdd("m", 1, DateSerial(yYear, mMonth, 1)))   'days in month
      
      If chkLatin = 0 Then
         StDG = DatePart("W", StDat, vbSunday)          'first day of week  english
      Else
         StDG = DatePart("W", StDat, vbMonday)         'latin starts day of week on Monday
      End If
      
      Goforit = False                                             'first of month is set so print the rest
      picCal.Print StDat                                      'print month and year of each calendar
      
      'set position to print weekday header
      For i = 0 To 30 Step 5
         picCal.CurrentY = 28
         picCal.CurrentX = 0
         If chkLatin = 0 Then
            picCal.Print Tab(i); Left$(StrConv(Format$((i / 5) + 1, "ddd"), vbProperCase), 2)      'english
         Else
            picCal.Print Tab(i); Left$(StrConv(Format$((i / 5) + 2, "ddd"), vbProperCase), 2)     'latin
         End If
      Next i
      
      picCal.ForeColor = lblFontColor.BackColor
      'set position to start printing days
      picCal.CurrentY = 45
      For cRow = 1 To 6
         
         For cColumn = 1 To 7
            If (StDG = cColumn And Not Goforit) Then     'set first day of month
            ctr = 1
            Goforit = True
         End If                                                        'get the rest of the days
         If (Goforit And IsDate(CStr(ctr) & "-" & mMonth & "-" & yYear)) Then
            If chkHol.Value = 1 Then                            'load the holidays
            For Hol = 0 To UBound(strArrayHoliday) - 1 Step 5
               If strArrayHoliday(Hol) = "FirstDayOfMonth" Then
                  If xctr = strArrayHoliday(Hol + 1) And ctr = 1 Then picCal.ForeColor = lblHolColor.BackColor
               ElseIf strArrayHoliday(Hol) = "LastDayOfMonth" Then
                  If xctr = strArrayHoliday(Hol + 1) And ctr = DaysInMonth Then picCal.ForeColor = lblHolColor.BackColor
               ElseIf strArrayHoliday(Hol) = "Custom" Then
                  If xctr = strArrayHoliday(Hol + 1) And ctr = strArrayHoliday(Hol + 2) Then picCal.ForeColor = lblHolColor.BackColor
               ElseIf strArrayHoliday(Hol + 3) = "EndOfMonthDay" Then
                  If xctr = strArrayHoliday(Hol + 1) And ctr = Int(Format(GetFirstLastDate(strArrayHoliday(Hol), Int(strArrayHoliday(Hol + 1)), yYear, 1), "dd")) Then picCal.ForeColor = lblHolColor.BackColor
               Else
                  If xctr = strArrayHoliday(Hol + 1) And ctr = Int(Format(GetFirstLastDate(strArrayHoliday(Hol), Int(strArrayHoliday(Hol + 1)), yYear, 0), "dd")) And strArrayHoliday(Hol + 2) = 1 Then picCal.ForeColor = lblHolColor.BackColor
                  If xctr = strArrayHoliday(Hol + 1) And ctr = Int(Format(GetFirstLastDate(strArrayHoliday(Hol), Int(strArrayHoliday(Hol + 1)), yYear, 0), "dd")) + 7 And strArrayHoliday(Hol + 2) = 2 Then picCal.ForeColor = lblHolColor.BackColor
                  If xctr = strArrayHoliday(Hol + 1) And ctr = Int(Format(GetFirstLastDate(strArrayHoliday(Hol), Int(strArrayHoliday(Hol + 1)), yYear, 0), "dd")) + 14 And strArrayHoliday(Hol + 2) = 3 Then picCal.ForeColor = lblHolColor.BackColor
                  If xctr = strArrayHoliday(Hol + 1) And ctr = Int(Format(GetFirstLastDate(strArrayHoliday(Hol), Int(strArrayHoliday(Hol + 1)), yYear, 0), "dd")) + 21 And strArrayHoliday(Hol + 2) = 4 Then picCal.ForeColor = lblHolColor.BackColor
               End If
            Next Hol
         End If
         picCal.Print Str(ctr);
         picCal.ForeColor = lblFontColor.BackColor
      End If
      
      If cColumn < 7 Then: picCal.Print Tab(cColumn * 5);
      
      ctr = ctr + 1
   Next cColumn
   picCal.Print            'line space
   picCal.Print            'line space
Next cRow

picCal.Picture = picCal.Image                                'render picture so we can copy it

'position and print each month
X = 40 + X
Y = 40 + Y
picMain.PaintPicture picCal.Picture, X + PosX, Y + PosY
X = IIf(xctr Mod 3 = 0, 0, X + 200)
Y = IIf(xctr Mod 3 = 0, Y + 160, Y - 40)

Next xctr
If chkLabel.Value = Checked Then                              'include name
picMain.CurrentX = Text1.Left + 2
picMain.CurrentY = Text1.top + 2
picMain.Print Text1.Text
End If
End Sub

Private Sub cmdCreateList_Click()
   Frame2.Visible = Not Frame2.Visible
End Sub

Private Sub cboYear_Click()
   picMain.Cls
   Generate
End Sub

Private Sub cmdChange_Click()
Dim X As Integer
   cboYear.Clear
   cboYear.AddItem txtDateFrom.Text
   For X = 1 To txtDateTo.Text - 1
      cboYear.AddItem txtDateFrom.Text + X
   Next X
   cboYear.Text = txtDateFrom.Text
   SetDropDownHeight Me, cboYear, Int(txtDateTo.Text)
   Frame3.Visible = False
End Sub

Private Sub cmdDateRange_Click()
   Frame3.Visible = Not Frame3.Visible
End Sub

Private Sub cmdPreview_Click()
   VS1.Value = 0
   Image1.Width = picMain.Width / 1.5
   Image1.Height = picMain.Height / 1.5
   Image1.Picture = picMain.Image
   Image1.Visible = Not Image1.Visible
   Shape2.Width = Image1.Width
   Shape2.Height = Image1.Height
   Shape2.Visible = Image1.Visible
End Sub

Private Sub cmdDeleteHol_Click()
   Dim X As Long
   Dim Y As Integer
   
   If lstHolList.ListCount = -1 Then Exit Sub
   X = lstHolList.ListCount
   Do While X > 0
      X = X - 1
      If lstHolList.Selected(X) = True Then
         lstHolList.RemoveItem (X)
      End If
   Loop
      
   SaveHolDay_List
   End Sub

Private Sub cmdEnterHolDay_Click()
   Dim X As Integer
   Dim HM As Integer
   HM = Int(Left$(cboHolMonth.Text, 2))
   If txtHolidayName.Text = "" Then txtHolidayName.Text = "*"
   For X = 1 To 8
      If X = 1 And optAlwaysDay(0).Value = True Then lstHolList.AddItem cboHolDay.Text & "," & HM & "," & X & "," & "Always 1st" & "," & txtHolidayName.Text
      If X = 2 And optAlwaysDay(1).Value = True Then lstHolList.AddItem cboHolDay.Text & "," & HM & "," & X & "," & "Always 2nd" & "," & txtHolidayName.Text
      If X = 3 And optAlwaysDay(2).Value = True Then lstHolList.AddItem cboHolDay.Text & "," & HM & "," & X & "," & "Always 3rd" & "," & txtHolidayName.Text
      If X = 4 And optAlwaysDay(3).Value = True Then lstHolList.AddItem cboHolDay.Text & "," & HM & "," & X & "," & "Always 4th" & "," & txtHolidayName.Text
      If X = 5 And optAlwaysDay(4).Value = True Then lstHolList.AddItem "FirstDayOfMonth" & "," & HM & "," & "*" & "," & "*" & "," & txtHolidayName.Text
      If X = 6 And optAlwaysDay(5).Value = True Then lstHolList.AddItem "LastDayOfMonth" & "," & HM & "," & "*" & "," & "*" & "," & txtHolidayName.Text
      If X = 7 And optAlwaysDay(6).Value = True Then lstHolList.AddItem "Custom" & "," & HM & "," & Int(txtCustDay.Text) & "," & "*" & "," & txtHolidayName.Text
      If X = 8 And optAlwaysDay(7).Value = True Then lstHolList.AddItem cboHolDay.Text & "," & HM & "," & "*" & "," & "EndOfMonthDay" & "," & txtHolidayName.Text
   Next X
   txtHolidayName.Text = ""
   SaveHolDay_List
End Sub

Private Sub cmdExit_Click()
   Frame2.Visible = False
End Sub

Private Sub cmdOffset_Click()
   PosX = Int(txtoffsetX.Text)
   PosY = Int(txtoffsetY.Text)
   picMain.Cls
   Generate
End Sub

Private Sub cmdPrint_Click()
   Generate
   picMain.Picture = picMain.Image                  'render the picture
   Printer.PaintPicture picMain.Picture, 0, 0
   Printer.EndDoc
   
   picMain.Picture = LoadPicture()                     'unrender the picture
   Generate
   chkLabel.Value = Unchecked
   chkLabel_Click
End Sub

Private Sub cmdPrinterSelect_Click()
   ShowPrinter Me
End Sub

Private Sub cmdSaveBitmap_Click()
   If txtSaveName.Text = "" Then GoTo here
   Generate
   picMain.Picture = picMain.Image
   SavePicture picMain.Picture, App.Path & "\" & txtSaveName.Text & ".bmp"
   picMain.Picture = LoadPicture()
   Generate
   chkLabel.Value = Unchecked
   chkLabel_Click
   MsgBox "Picture saved as a bitmap at " & App.Path & "\" & txtSaveName.Text & ".bmp"
   txtSaveName.Text = ""
   Exit Sub
here:
   MsgBox "Calendar not saved, Please enter a Name for it."
End Sub

Private Sub chkLabel_Click()
   If chkLabel.Value = Checked Then
      Text1.Visible = True
      Text1.SetFocus
   Else
      Text1.Visible = False
      Text1.Text = ""
   End If
End Sub

Private Sub chkLatin_Click()
   Generate
End Sub

Private Sub chkHol_Click()
   chkColored_Click
End Sub

Private Sub chkColored_Click()
   picMain.Cls
   Generate
   'update print preview
   Image1.Width = picMain.Width / 1.5
   Image1.Height = picMain.Height / 1.5
   Image1.Picture = picMain.Image
   Shape2.Width = Image1.Width
   Shape2.Height = Image1.Height
End Sub

Private Sub Frame2_DblClick()
   Frame2.Visible = False
End Sub

Private Sub Image1_Click()
   Image1.Visible = Not Image1.Visible
   Shape2.Visible = Image1.Visible
End Sub

Private Sub lblBkColor_Click()
   Dim sure As Long
   sure = ShowColor
   If sure = -1 Then Exit Sub
   lblBkColor.BackColor = sure
   picCal.BackColor = lblBkColor.BackColor
   picMain.BackColor = lblBkColor.BackColor
   Generate
End Sub

Private Sub lblFontColor_Click()
   Dim sure As Long
   sure = ShowColor
   If sure = -1 Then Exit Sub
   lblFontColor.BackColor = sure
   picMain.Cls
   Generate
End Sub

Private Sub lblHolColor_Click()
   Dim sure As Long
   sure = ShowColor
   If sure = -1 Then Exit Sub
   lblHolColor.BackColor = sure
   chkHol_Click
End Sub

Private Sub lblMonColor_Click(Index As Integer)
   Dim sure As Long
   sure = ShowColor
   If sure = -1 Then Exit Sub
   lblMonColor(Index).BackColor = sure
   chkColored.Value = 0
End Sub

Private Sub lstHolList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)    're-arrange items in list
   DragIndex = HitTest(Y)                    'item To be dragged
   If DragIndex < 0 Then Exit Sub          'not on an item
   Dragging = True                              'for MouseMove
   DragText = lstHolList.List(DragIndex)  'item text
   LastIndex = DragIndex                     'for first time
   LastString = DragText
End Sub

Private Sub lstHolList_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)      're-arrange items in list
   CurrentIndex = HitTest(Y)                        'what are we over now
   If CurrentIndex < 0 Then Exit Sub             'not an item
   If Dragging Then
      If CurrentIndex <> LastIndex Then         'dragging over different item
      LastString = lstHolList.List(CurrentIndex) 'save this item
      lstHolList.List(CurrentIndex) = DragText  'set current
      lstHolList.List(LastIndex) = LastString     'set previous
      LastIndex = CurrentIndex                     'save last
   End If
End If
End Sub

Private Sub lstHolList_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Dragging = False
End Sub

Private Sub optAlwaysDay_Click(Index As Integer)
   If Index = 6 Then                        'Custom option,enable textbox and label
   txtCustDay.Enabled = True
   Label12.Enabled = True
Else
   txtCustDay.Enabled = False
   Label12.Enabled = False
End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then                     'enter key was pressed
      Generate
      Text1.Visible = False
      chkLabel.Value = 0
   End If
End Sub

Private Sub Text1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Oldx = X
   Oldy = Y
   picMain.ScaleMode = 1
End Sub

Private Sub Text1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = 1 Then                              'drag textbox to new position
      Text1.Left = Text1.Left + (X - Oldx)
      Text1.top = Text1.top + (Y - Oldy)
   End If
End Sub

Private Sub Text1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   picMain.ScaleMode = 3
End Sub

Private Sub txtCustDay_KeyPress(KeyAscii As Integer)
    If Not IsNumeric(Chr(KeyAscii)) Then
        If KeyAscii = 8 Then
            'Allow Backspace And Space
        Else
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub VS1_Change()
   VS1_Scroll
End Sub

Private Sub VS1_Scroll()
   picMain.top = -(VS1.Value) + 4
   cboYear.SetFocus
End Sub

Private Function HitTest(ByVal Y As Long) As Long
   'Gets the listindex from the mouse pos
   HitTest = SendMessage(LBHwnd, LB_ITEMFROMPOINT, ByVal 0&, ByVal (Y \ Screen.TwipsPerPixelY) * 65536)
   If HitTest > lstHolList.ListCount - 1 Then
      'mouse not over an item
      HitTest = -1
   End If
End Function

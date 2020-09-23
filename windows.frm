VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gullu - Windows options page 1"
   ClientHeight    =   5535
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4635
   Icon            =   "windows.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   4635
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00B5752F&
      Height          =   5535
      Left            =   0
      ScaleHeight     =   5475
      ScaleWidth      =   4635
      TabIndex        =   51
      Top             =   0
      Visible         =   0   'False
      Width           =   4695
      Begin VB.Label Label45 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password list"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   210
         Left            =   3360
         TabIndex        =   67
         Top             =   0
         Width           =   990
      End
      Begin VB.Line Line40 
         BorderColor     =   &H00C09634&
         X1              =   3120
         X2              =   3120
         Y1              =   240
         Y2              =   0
      End
      Begin VB.Line Line39 
         BorderColor     =   &H00C09634&
         X1              =   1560
         X2              =   1560
         Y1              =   240
         Y2              =   0
      End
      Begin VB.Image Image4 
         Height          =   240
         Left            =   120
         Picture         =   "windows.frx":628A
         Stretch         =   -1  'True
         Top             =   0
         Width           =   240
      End
      Begin VB.Line Line38 
         BorderColor     =   &H009D6322&
         X1              =   0
         X2              =   4560
         Y1              =   4920
         Y2              =   4920
      End
      Begin VB.Line Line37 
         BorderColor     =   &H009D6322&
         X1              =   0
         X2              =   4560
         Y1              =   4560
         Y2              =   4560
      End
      Begin VB.Line Line36 
         BorderColor     =   &H009D6322&
         X1              =   0
         X2              =   4560
         Y1              =   4200
         Y2              =   4200
      End
      Begin VB.Line Line35 
         BorderColor     =   &H009D6322&
         X1              =   0
         X2              =   4560
         Y1              =   3840
         Y2              =   3840
      End
      Begin VB.Line Line34 
         BorderColor     =   &H009D6322&
         X1              =   0
         X2              =   4560
         Y1              =   3480
         Y2              =   3480
      End
      Begin VB.Line Line33 
         BorderColor     =   &H009D6322&
         X1              =   0
         X2              =   4560
         Y1              =   3120
         Y2              =   3120
      End
      Begin VB.Line Line32 
         BorderColor     =   &H009D6322&
         X1              =   0
         X2              =   4560
         Y1              =   2760
         Y2              =   2760
      End
      Begin VB.Line Line31 
         BorderColor     =   &H009D6322&
         X1              =   0
         X2              =   4560
         Y1              =   2400
         Y2              =   2400
      End
      Begin VB.Line Line30 
         BorderColor     =   &H009D6322&
         X1              =   0
         X2              =   4560
         Y1              =   2040
         Y2              =   2040
      End
      Begin VB.Line Line29 
         BorderColor     =   &H009D6322&
         X1              =   0
         X2              =   4560
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Line Line28 
         BorderColor     =   &H009D6322&
         X1              =   0
         X2              =   4560
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Line Line27 
         BorderColor     =   &H009D6322&
         X1              =   0
         X2              =   4560
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Line Line26 
         BorderColor     =   &H009D6322&
         X1              =   0
         X2              =   4560
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label Label44 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "System properties"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   1380
         TabIndex        =   66
         Top             =   5040
         Width           =   1575
      End
      Begin VB.Label Label43 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sound properties"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   1380
         TabIndex        =   65
         Top             =   4680
         Width           =   1470
      End
      Begin VB.Label Label42 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Regional properties"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   1380
         TabIndex        =   64
         Top             =   4320
         Width           =   1665
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Keyboard properties"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   1440
         TabIndex        =   63
         Top             =   3960
         Width           =   1740
      End
      Begin VB.Label Label40 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Network properties"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   1380
         TabIndex        =   62
         Top             =   3600
         Width           =   1650
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Multimedia properties"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   1380
         TabIndex        =   61
         Top             =   3240
         Width           =   1860
      End
      Begin VB.Label Label38 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mouse properties"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   1380
         TabIndex        =   60
         Top             =   2880
         Width           =   1500
      End
      Begin VB.Label Label37 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Modem properties"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   1380
         TabIndex        =   59
         Top             =   2520
         Width           =   1560
      End
      Begin VB.Label Label36 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Game controller"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   1515
         TabIndex        =   58
         Top             =   2160
         Width           =   1365
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Internet properties"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   1515
         TabIndex        =   57
         Top             =   1800
         Width           =   1590
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Display properties"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   1545
         TabIndex        =   56
         Top             =   1440
         Width           =   1545
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date/time properties"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   1440
         TabIndex        =   55
         Top             =   1080
         Width           =   1740
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Add/remove program"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   1410
         TabIndex        =   54
         Top             =   720
         Width           =   1800
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Add/new hardware"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   1515
         TabIndex        =   53
         Top             =   360
         Width           =   1605
      End
      Begin VB.Label Label30 
         Alignment       =   2  'Center
         BackColor       =   &H009D6322&
         Caption         =   "Launch options"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00EFD8B8&
         Height          =   255
         Left            =   0
         TabIndex        =   52
         Top             =   0
         Width           =   4695
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00B5752F&
      Height          =   5535
      Left            =   0
      ScaleHeight     =   5475
      ScaleWidth      =   4635
      TabIndex        =   18
      Top             =   0
      Visible         =   0   'False
      Width           =   4695
      Begin VB.CommandButton Command10 
         Caption         =   "Sys. Info."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3360
         TabIndex        =   39
         Top             =   4560
         Width           =   1335
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Log off"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2160
         TabIndex        =   40
         Top             =   4560
         Width           =   1215
      End
      Begin VB.CommandButton Command12 
         Caption         =   "Restart"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         TabIndex        =   50
         Top             =   4560
         Width           =   1095
      End
      Begin VB.TextBox Text5 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1560
         TabIndex        =   47
         Top             =   3170
         Width           =   1695
      End
      Begin VB.DriveListBox Drive3 
         Height          =   315
         Left            =   1920
         TabIndex        =   46
         Top             =   2790
         Width           =   2055
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Shut down"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -120
         TabIndex        =   38
         Top             =   4560
         Width           =   1215
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Show"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3480
         TabIndex        =   37
         Top             =   4080
         Width           =   975
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Hide"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2400
         TabIndex        =   36
         Top             =   4080
         Width           =   975
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Disable"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3480
         TabIndex        =   34
         Top             =   3600
         Width           =   975
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Enable"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2400
         TabIndex        =   33
         Top             =   3600
         Width           =   975
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Change"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3480
         TabIndex        =   26
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1920
         TabIndex        =   25
         Top             =   1320
         Width           =   1455
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Add"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3480
         TabIndex        =   23
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1440
         TabIndex        =   22
         Top             =   840
         Width           =   1935
      End
      Begin VB.Image Image3 
         Height          =   480
         Left            =   4080
         Picture         =   "windows.frx":66CC
         Top             =   5040
         Width           =   480
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Set"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00EFD8B8&
         Height          =   210
         Left            =   3840
         TabIndex        =   49
         Top             =   3240
         Width           =   270
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "C:\"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   210
         Left            =   3360
         TabIndex        =   48
         Top             =   3240
         Width           =   195
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Get"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00EFD8B8&
         Height          =   210
         Left            =   1560
         TabIndex        =   45
         Top             =   2880
         Width           =   285
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   120
         Picture         =   "windows.frx":6B0E
         ToolTipText     =   "Previous page"
         Top             =   5040
         Width           =   480
      End
      Begin VB.Line Line25 
         BorderColor     =   &H009D6322&
         X1              =   0
         X2              =   4560
         Y1              =   5040
         Y2              =   5040
      End
      Begin VB.Line Line24 
         BorderColor     =   &H009D6322&
         X1              =   0
         X2              =   4560
         Y1              =   4440
         Y2              =   4440
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hide/show start button:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   210
         Left            =   120
         TabIndex        =   35
         Top             =   4080
         Width           =   1695
      End
      Begin VB.Line Line23 
         BorderColor     =   &H009D6322&
         X1              =   0
         X2              =   4560
         Y1              =   3960
         Y2              =   3960
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Enable/disable Ctrl+Alt+Delete:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   210
         Left            =   120
         TabIndex        =   32
         Top             =   3600
         Width           =   2190
      End
      Begin VB.Line Line22 
         BorderColor     =   &H009D6322&
         X1              =   0
         X2              =   4560
         Y1              =   3480
         Y2              =   3480
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Change disk name:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   210
         Left            =   120
         TabIndex        =   31
         Top             =   3240
         Width           =   1365
      End
      Begin VB.Line Line21 
         BorderColor     =   &H009D6322&
         X1              =   0
         X2              =   4560
         Y1              =   3120
         Y2              =   3120
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Disk serial number:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   210
         Left            =   120
         TabIndex        =   30
         Top             =   2880
         Width           =   1365
      End
      Begin VB.Line Line20 
         BorderColor     =   &H009D6322&
         X1              =   0
         X2              =   4560
         Y1              =   2760
         Y2              =   2760
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Small/large fonts:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   210
         Left            =   120
         TabIndex        =   29
         Top             =   2520
         Width           =   1245
      End
      Begin VB.Line Line19 
         BorderColor     =   &H009D6322&
         X1              =   0
         X2              =   4560
         Y1              =   2400
         Y2              =   2400
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Colors displaying:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   210
         Left            =   120
         TabIndex        =   28
         Top             =   2160
         Width           =   1275
      End
      Begin VB.Line Line18 
         BorderColor     =   &H009D6322&
         X1              =   0
         X2              =   4560
         Y1              =   2040
         Y2              =   2040
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Currency symbol:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   210
         Left            =   120
         TabIndex        =   27
         Top             =   1800
         Width           =   1275
      End
      Begin VB.Line Line17 
         BorderColor     =   &H009D6322&
         X1              =   0
         X2              =   4560
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Change computer name:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   210
         Left            =   120
         TabIndex        =   24
         Top             =   1320
         Width           =   1755
      End
      Begin VB.Line Line16 
         BorderColor     =   &H009D6322&
         X1              =   0
         X2              =   4560
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Add file to recent:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   210
         Left            =   120
         TabIndex        =   21
         Top             =   840
         Width           =   1290
      End
      Begin VB.Line Line15 
         BorderColor     =   &H009D6322&
         X1              =   0
         X2              =   4560
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Windows is running for:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   210
         Left            =   120
         TabIndex        =   20
         Top             =   480
         Width           =   1740
      End
      Begin VB.Line Line14 
         BorderColor     =   &H009D6322&
         X1              =   0
         X2              =   4560
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mouse buttons:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   210
         Left            =   120
         TabIndex        =   19
         Top             =   120
         Width           =   1110
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00B5752F&
      Height          =   5535
      Left            =   0
      ScaleHeight     =   5475
      ScaleWidth      =   4635
      TabIndex        =   0
      Top             =   0
      Width           =   4695
      Begin VB.DriveListBox Drive2 
         Height          =   315
         Left            =   2040
         TabIndex        =   44
         Top             =   3150
         Width           =   1455
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   1680
         TabIndex        =   42
         Top             =   2780
         Width           =   1695
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Move"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3720
         TabIndex        =   9
         Top             =   1680
         Width           =   855
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1440
         TabIndex        =   8
         Top             =   1680
         Width           =   2175
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Change"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3840
         TabIndex        =   3
         Top             =   120
         Width           =   735
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   960
         TabIndex        =   2
         Top             =   120
         Width           =   2775
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Get"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00EFD8B8&
         Height          =   210
         Left            =   1680
         TabIndex        =   43
         Top             =   3240
         Width           =   285
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Get"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00EFD8B8&
         Height          =   210
         Left            =   1200
         TabIndex        =   41
         Top             =   2880
         Width           =   285
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   4080
         Picture         =   "windows.frx":6F50
         ToolTipText     =   "Next page"
         Top             =   4920
         Width           =   480
      End
      Begin VB.Line Line13 
         BorderColor     =   &H009D6322&
         X1              =   0
         X2              =   4560
         Y1              =   4920
         Y2              =   4920
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Double-click time(ms):"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   210
         Left            =   120
         TabIndex        =   17
         Top             =   4680
         Width           =   1575
      End
      Begin VB.Line Line12 
         BorderColor     =   &H009D6322&
         X1              =   0
         X2              =   4560
         Y1              =   4560
         Y2              =   4560
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Modem connection:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   210
         Left            =   120
         TabIndex        =   16
         Top             =   4320
         Width           =   1395
      End
      Begin VB.Line Line11 
         BorderColor     =   &H009D6322&
         X1              =   0
         X2              =   4560
         Y1              =   4200
         Y2              =   4200
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Current Printer name:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   210
         Left            =   120
         TabIndex        =   15
         Top             =   3960
         Width           =   1530
      End
      Begin VB.Line Line10 
         BorderColor     =   &H009D6322&
         X1              =   0
         X2              =   4560
         Y1              =   3840
         Y2              =   3840
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Screen resolution:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   210
         Left            =   120
         TabIndex        =   14
         Top             =   3600
         Width           =   1320
      End
      Begin VB.Line Line9 
         BorderColor     =   &H009D6322&
         X1              =   0
         X2              =   4560
         Y1              =   3480
         Y2              =   3480
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Volume information:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   210
         Left            =   120
         TabIndex        =   13
         Top             =   3240
         Width           =   1425
      End
      Begin VB.Line Line8 
         BorderColor     =   &H009D6322&
         X1              =   0
         X2              =   4560
         Y1              =   3120
         Y2              =   3120
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Free space:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   210
         Left            =   120
         TabIndex        =   12
         Top             =   2880
         Width           =   870
      End
      Begin VB.Line Line7 
         BorderColor     =   &H009D6322&
         X1              =   0
         X2              =   4560
         Y1              =   2760
         Y2              =   2760
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Computer name:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   210
         Left            =   120
         TabIndex        =   11
         Top             =   2520
         Width           =   1170
      End
      Begin VB.Line Line6 
         BorderColor     =   &H009D6322&
         X1              =   0
         X2              =   4560
         Y1              =   2400
         Y2              =   2400
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Username:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   210
         Left            =   120
         TabIndex        =   10
         Top             =   2160
         Width           =   780
      End
      Begin VB.Line Line5 
         BorderColor     =   &H009D6322&
         X1              =   0
         X2              =   4560
         Y1              =   2040
         Y2              =   2040
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Move to recycle:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   210
         Left            =   120
         TabIndex        =   7
         Top             =   1680
         Width           =   1200
      End
      Begin VB.Line Line4 
         BorderColor     =   &H009D6322&
         X1              =   0
         X2              =   4560
         Y1              =   1560
         Y2              =   1560
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Screen saver on/off:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   210
         Left            =   120
         TabIndex        =   6
         Top             =   1320
         Width           =   1515
      End
      Begin VB.Line Line3 
         BorderColor     =   &H009D6322&
         X1              =   0
         X2              =   4560
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "System directory:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   210
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   1275
      End
      Begin VB.Line Line2 
         BorderColor     =   &H009D6322&
         X1              =   0
         X2              =   4560
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Windows directory:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   210
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   1425
      End
      Begin VB.Line Line1 
         BorderColor     =   &H009D6322&
         X1              =   0
         X2              =   4560
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Wallpaper:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   210
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   765
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dblReturn As Double
Dim lngSuccess As Long
Dim strBitmapImage As String
Private Declare Function GetVolumeInformation& Lib "kernel32" Alias "GetVolumeInformationA" _
   (ByVal lpRootPathName As String, ByVal pVolumeNameBuffer As String, _
    ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, _
    lpMaximumComponentLength As Long, lpFileSystemFlags As Long, _
    ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long)
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
Private Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
Private Declare Function GetUserDefaultLCID% Lib "kernel32" ()
Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetTextMetrics Lib "gdi32" Alias "GetTextMetricsA" (ByVal hdc As Long, lpMetrics As TEXTMETRIC) As Long
Private Declare Function SetMapMode Lib "gdi32" (ByVal hdc As Long, ByVal nMapMode As Long) As Long
Private Declare Function SetVolumeLabelA Lib "kernel32" _
   (ByVal lpRootPathName As String, _
   ByVal lpVolumeName As String) As Long
Private Declare Function GetDiskFreeSpace Lib "kernel32" Alias "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, lpSectorsPerCluster As Long, lpBytesPerSector As Long, lpNumberOfFreeClusters As Long, lpTotalNumberOfClusters As Long) As Long
Private Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
Private Type SHFILEOPSTRUCT
    hwnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAnyOperationsAborted As Long
    hNameMappings As Long
    lpszProgressTitle As Long
End Type
Private Type TEXTMETRIC
   tmHeight As Integer
   tmAscent As Integer
   tmDescent As Integer
   tmInternalLeading As Integer
   tmExternalLeading As Integer
   tmAveCharWidth As Integer
   tmMaxCharWidth As Integer
   tmWeight As Integer
   tmItalic As String * 1
   tmUnderlined As String * 1
   tmStruckOut As String * 1
   tmFirstChar As String * 1
   tmLastChar As String * 1
   tmDefaultChar As String * 1
   tmBreakChar As String * 1
   tmPitchAndFamily As String * 1
   tmCharSet As String * 1
   tmOverhang As Integer
   tmDigitizedAspectX As Integer
   tmDigitizedAspectY As Integer
End Type
Private Const MAX_FILENAME_LEN = 256
Private Const MM_TEXT = 1
Private Const PHYSICALOFFSETX = 112
Private Const PHYSICALOFFSETY = 113
Private Const PLANES = 14
Private Const BITSPIXEL = 12
Private Const FO_DELETE = &H3
Private Const LOCALE_SCURRENCY = &H14
Private Const FOF_ALLOWUNDO = &H40
Private Const FS_CASE_IS_PRESERVED = 2
Private Const FS_CASE_SENSITIVE = 1
Private Const FS_UNICODE_STORED_ON_DISK = 4
Private Const FS_PERSISTENT_ACLS = 8
Private Const FS_FILE_COMPRESSION = 16
Private Const FS_VOL_IS_COMPRESSED = 32768
Const HKEY_CURRENT_CONFIG As Long = &H80000005
Private Sub Command1_Click()
On Error Resume Next
strBitmapImage = Text1.Text
lngSuccess = SystemParametersInfo(SPI_SETDESKWALLPAPER, 0, strBitmapImage, 0)
End Sub
Private Sub Command10_Click()
On Error Resume Next
SystemInformation
End Sub
Private Sub Command11_Click()
On Error Resume Next
ShutDownWindows 0
End Sub
Private Sub Command12_Click()
On Error Resume Next
ShutDownWindows 2
End Sub
Private Sub Command2_Click()
On Error Resume Next
Dim typOperation As SHFILEOPSTRUCT
With typOperation
        .wFunc = FO_DELETE
        .pFrom = Text2.Text
        .fFlags = FOF_ALLOWUNDO
    End With
    SHFileOperation typOperation
End Sub
Private Sub Command3_Click()
On Error Resume Next
Dim strNewFile As String
strNewFile = Text3.Text
Call SHAddToRecentDocs(2, strNewFile)
End Sub
Private Sub Command4_Click()
On Error Resume Next
Dim strNewComputerName As String
Dim lngReturn As Long
strNewComputerName = Text4.Text
lngReturn = SetComputerName(strNewComputerName)
End Sub
Private Sub Command5_Click()
On Error Resume Next
DisableCTRLaltDEL False
End Sub
Private Sub Command6_Click()
On Error Resume Next
DisableCTRLaltDEL True
End Sub
Private Sub Command7_Click()
On Error Resume Next
hideStartButton
End Sub
Private Sub Command8_Click()
On Error Resume Next
showStartButton
End Sub
Private Sub Command9_Click()
On Error Resume Next
ShutDownWindows 1
End Sub
Private Sub Form_Load()
On Error Resume Next
Dim Symbol As String
Dim iRet1 As Long
Dim iRet2 As Long
Dim lpLCDataVar As String
Dim Pos As Integer
Dim Locale As Long
Dim GetCurrPrinter As String
Dim lngReturn As Long
Dim PName As String
Dim intWidth As Integer
Dim intHeight As Integer
Dim strRootPathName As String
Dim lngSectorsPerCluster As Long
Dim lngBytesPerSector As Long
Dim lngNumberOfFreeClusters As Long
Dim lngTotalNumberOfClusters As Long
Dim strDrive As String
Dim strMessage As String
Dim lngTotalBytes As Long
Dim lngFreeBytes As Long
Dim lngBufSize As Long
Dim lngStatus As Long
Dim blnReturn As Boolean
Dim blnActive As Boolean
Dim strBuffer As String
Dim strWindowsDirectory As String
strBuffer = Space$(MAX_PATH)
lngReturn = GetWindowsDirectory(strBuffer, MAX_PATH)
strWindowsDirectory = Left$(strBuffer, Len(strBuffer) - 1)
Label2.Caption = "Windows directory:" & Space(1) & strWindowsDirectory
strBuffer = Space$(MAX_PATH)
lngReturn = GetSystemDirectory(strBuffer, MAX_PATH)
strwindowssystemdirectory = Left$(strBuffer, Len(strBuffer) - 1)
Label3.Caption = "System directory: " & Space(1) & strwindowssystemdirectory
Call SystemParametersInfo(SPI_GETSCREENSAVEACTIVE, vbNull, blnReturn, 0)
Label4.Caption = "Screen saver on/off: " & Space(1) & blnReturn
Get_User_Name
  lngBufSize = 255
  strBuffer = String$(lngBufSize, " ")
  lngStatus = GetComputerName(strBuffer, lngBufSize)
  If lngStatus <> 0 Then
     Label7.Caption = ("Computer name: " & Space(1) & Left(strBuffer, lngBufSize))
  End If
  intWidth = Screen.Width \ Screen.TwipsPerPixelX
intHeight = Screen.Height \ Screen.TwipsPerPixelY
Label10.Caption = "Screen resolution::" + Str$(intWidth) + " x" + Str$(intHeight)
GetCurrPrinter = RegGetString$(HKEY_CURRENT_CONFIG, "System\CurrentControlSet\Control\Print\Printers", "Default")
 Label11.Caption = "Current Printer name: " & Space(1) & GetCurrPrinter
 If IsModemConnected = True Then
 Label12.Caption = "Modem connection: " & Space(1) & "Established"
 Else
 Label12.Caption = "Modem connection: " & Space(1) & "Didn't establish"
End If
Label13.Caption = "Double-click time(ms): " & Space(1) & GetDoubleClickTime
Label14.Caption = "Mouse buttons: " & GetSystemMetrics(SM_CMOUSEBUTTONS)
lngReturn = GetTickCount()
Label15.Caption = "Windows is running for: " & (lngReturn / 1000) & " seconds."
      Locale = GetUserDefaultLCID()
      iRet1 = GetLocaleInfo(Locale, LOCALE_SCURRENCY, lpLCDataVar, 0)
      Symbol = String$(iRet1, 0)
      iRet2 = GetLocaleInfo(Locale, LOCALE_SCURRENCY, Symbol, iRet1)
      Pos = InStr(Symbol, Chr$(0))
      If Pos > 0 Then
         Label18.Caption = "Currency symbol: " & Space(1) & Left$(Symbol, Pos - 1)
      End If
Label19.Caption = "Colors displaying: " & Space(1) & GetNColors & " colors"
If SmallFonts = True Then
Label20.Caption = "Small/large fonts: Small fonts are used"
Else
Label20.Caption = "Small/large fonts: Large fonts are used"
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
MsgBox "Thank you for using Gullu Winsource" & vbNewLine & "Made by Gullu" & vbNewLine & "mubassherkamal@hotmail.com" & vbNewLine & "Search Gullu for more at www.psc.com", vbInformation, "About"
End Sub
Private Sub Image1_Click()
Picture1.Visible = False
Picture2.Visible = True
Picture3.Visible = False
Form2.Caption = "Gullu - Windows options page 2"
End Sub
Private Sub Image2_Click()
Picture1.Visible = True
Picture2.Visible = False
Picture3.Visible = False
Form2.Caption = "Gullu - Windows options page 1"
End Sub
Sub Get_User_Name()
Dim lpBuff As String * 25
Dim ret As Long, UserName As String
ret = GetUserName(lpBuff, 25)
UserName = Left(lpBuff, InStr(lpBuff, Chr(0)) - 1)
Label6.Caption = "Username: " & Space(1) & UserName
End Sub
Private Sub Image3_Click()
Picture3.Visible = True
Picture2.Visible = False
Picture1.Visible = False
Form2.Caption = "Gullu - Windows options page 3"
End Sub
Private Sub Image4_Click()
Picture1.Visible = False
Picture3.Visible = False
Picture2.Visible = True
Form2.Caption = "Gullu - Windows options page 2"
End Sub
Private Sub Label25_Click()
On Error Resume Next
      strDrive = Drive1.Drive
    If GetDiskFreeSpace(strDrive, lngSectorsPerCluster, lngBytesPerSector, lngNumberOfFreeClusters, lngTotalNumberOfClusters) = 0 Then
        strMessage = strMessage & vbCrLf & "Cannot find in the current drive"
    Else
        strMessage = strMessage & vbCrLf & "Sectors Per Cluster: " & Format$(lngSectorsPerCluster)
        strMessage = strMessage & vbCrLf & "Bytes Per Sector: " & Format$(lngBytesPerSector)
        strMessage = strMessage & vbCrLf & "Free Clusters: " & Format$(lngNumberOfFreeClusters)
        strMessage = strMessage & vbCrLf & "Total Clusters: " & Format$(lngTotalNumberOfClusters)
        lngTotalBytes = lngTotalNumberOfClusters * lngSectorsPerCluster * lngBytesPerSector
        strMessage = strMessage & vbCrLf & "Total Bytes: " & Format$(lngTotalBytes)
        lngFreeBytes = lngNumberOfFreeClusters * lngSectorsPerCluster * lngBytesPerSector
        strMessage = strMessage & vbCrLf & "Bytes Free: " & Format$(lngFreeBytes)
        strMessage = strMessage & vbCrLf & "Percent Used: " & Format$(1 - (lngFreeBytes / lngTotalBytes), "0.00%")
    End If
    MsgBox (strMessage)
End Sub
Private Sub Label26_Click()
On Error Resume Next
Dim strRootPathName As String
Dim strVolumeNameBuffer As String * 256
Dim lngVolumeNameSize As Long
Dim lngVolumeSerialNumber As Long
Dim lngMaximumComponentLength As Long
Dim lngFileSystemFlags As Long
Dim strFileSystemNameBuffer As String * 256
Dim lngFileSystemNameSize As Long
Dim strMessage As String
strRootPathName = Drive2.Drive
    If GetVolumeInformation(strRootPathName, strVolumeNameBuffer, Len(strVolumeNameBuffer), lngVolumeSerialNumber, lngMaximumComponentLength, lngFileSystemFlags, strFileSystemNameBuffer, Len(strFileSystemNameBuffer)) = 0 Then
        strMessage = "An error occurred!"
    Else
        strMessage = strRootPathName
        strVolumeNameBuffer = Left$(strVolumeNameBuffer, InStr(strVolumeNameBuffer, Chr$(0)) - 1)
        strMessage = strMessage & vbCrLf & "Volume Name: " & strVolumeNameBuffer
        strMessage = strMessage & vbCrLf & "Serial number: " & Format$(lngVolumeSerialNumber)
        strMessage = strMessage & vbCrLf & "Max component length: " & Format$(lngMaximumComponentLength)
        strMessage = strMessage & vbCrLf & "System Flags: "
        If lngFileSystemFlags And FS_CASE_IS_PRESERVED Then strMessage = strMessage & vbCrLf & "    FS_CASE_IS_PRESERVED"
        If lngFileSystemFlags And FS_CASE_SENSITIVE Then strMessage = strMessage & vbCrLf & "    FS_CASE_SENSITIVE"
        If lngFileSystemFlags And FS_UNICODE_STORED_ON_DISK Then strMessage = strMessage & vbCrLf & "    FS_UNICODE_STORED_ON_DISK"
        If lngFileSystemFlags And FS_PERSISTENT_ACLS Then strMessage = strMessage & vbCrLf & "    FS_PERSISTENT_ACLS"
        If lngFileSystemFlags And FS_FILE_COMPRESSION Then strMessage = strMessage & vbCrLf & "    FS_FILE_COMPRESSION"
        If lngFileSystemFlags And FS_VOL_IS_COMPRESSED Then strMessage = strMessage & vbCrLf & "    FS_VOL_IS_COMPRESSED"
        strFileSystemNameBuffer = Left$(strFileSystemNameBuffer, InStr(strFileSystemNameBuffer, Chr$(0)) - 1)
        strMessage = strMessage & vbCrLf & "File System: " & strFileSystemNameBuffer
    End If
MsgBox (strMessage)
End Sub
Public Function GetNColors() As Long
  Dim hSrcDC As Integer
hSrcDC = GetDC(GetDesktopWindow())
  GetNColors = GetDeviceCaps(hSrcDC, PLANES) * 2 ^ GetDeviceCaps(hSrcDC, BITSPIXEL)
  Call ReleaseDC(GetDesktopWindow(), hSrcDC)
End Function
Public Function SmallFonts() As Boolean
   Dim hdc As Long
   Dim hwnd As Long
   Dim PrevMapMode As Long
   Dim tm As TEXTMETRIC
   SmallFonts = True
   hwnd = GetDesktopWindow()
   hdc = GetWindowDC(hwnd)
   If hdc Then
      PrevMapMode = SetMapMode(hdc, MM_TEXT)
      GetTextMetrics hdc, tm
      PrevMapMode = SetMapMode(hdc, PrevMapMode)
      ReleaseDC hwnd, hdc
      If tm.tmHeight > 16 Then SmallFonts = False
   End If
End Function
Public Function SetVolumeName(sDrive As String, n As String) As Boolean
   Dim i As Long
   i = SetVolumeLabelA(sDrive + ":\" & Chr$(0), n & Chr$(0))
   SetVolumeName = IIf(i = 0, False, True)
End Function
Public Sub ShutDownWindows(ByVal uFlags As Long)
Call ExitWindowsEx(uFlags, 0)
End Sub
Public Function GetSerialNumber(sDrive As String) As Long
   Dim ser As Long
   Dim s As String * MAX_FILENAME_LEN
   Dim s2 As String * MAX_FILENAME_LEN
   Dim i As Long
   Dim j As Long
   Call GetVolumeInformation(sDrive + ":\" & Chr$(0), s, MAX_FILENAME_LEN, ser, i, j, s2, MAX_FILENAME_LEN)
   GetSerialNumber = ser
End Function
Private Sub Label27_Click()
On Error Resume Next
MsgBox "The serial number is : " & GetSerialNumber(Drive3.Drive), vbInformation, "Information on serial number"
End Sub
Private Sub Label29_Click()
On Error Resume Next
SetVolumeName "c:\", Text5.Text
End Sub
Sub DisableCTRLaltDEL(huh As Boolean)
GD = SystemParametersInfo(97, huh, CStr(1), 0)
End Sub
Private Sub Label31_Click()
dblReturn = Shell("rundll32.exe shell32.dll,Control_RunDLL sysdm.cpl @1", 5)
End Sub
Private Sub Label32_Click()
dblReturn = Shell("rundll32.exe shell32.dll,Control_RunDLL appwiz.cpl,,1", 5)
End Sub
Private Sub Label33_Click()
dblReturn = Shell("rundll32.exe shell32.dll,Control_RunDLL timedate.cpl", 5)
End Sub
Private Sub Label34_Click()
dblReturn = Shell("rundll32.exe shell32.dll,Control_RunDLL desk.cpl,,0", 5)
End Sub
Private Sub Label35_Click()
dblReturn = Shell("rundll32.exe shell32.dll,Control_RunDLL inetcpl.cpl,,0", 5)
End Sub
Private Sub Label36_Click()
dblReturn = Shell("rundll32.exe shell32.dll,Control_RunDLL joy.cpl", 5)
End Sub
Private Sub Label37_Click()
dblReturn = Shell("rundll32.exe shell32.dll,Control_RunDLL modem.cpl", 5)
End Sub
Private Sub Label38_Click()
dblReturn = Shell("rundll32.exe shell32.dll,Control_RunDLL main.cpl @0", 5)
End Sub
Private Sub Label39_Click()
dblReturn = Shell("rundll32.exe shell32.dll,Control_RunDLL mmsys.cpl,,0", 5)
End Sub
Private Sub Label40_Click()
dblReturn = Shell("rundll32.exe shell32.dll,Control_RunDLL netcpl.cpl", 5)
End Sub
Private Sub Label41_Click()
dblReturn = Shell("rundll32.exe shell32.dll,Control_RunDLL main.cpl @1", 5)
End Sub
Private Sub Label42_Click()
dblReturn = Shell("rundll32.exe shell32.dll,Control_RunDLL intl.cpl,,0", 5)
End Sub
Private Sub Label43_Click()
dblReturn = Shell("rundll32.exe shell32.dll,Control_RunDLL mmsys.cpl @1", 5)
End Sub
Private Sub Label44_Click()
dblReturn = Shell("rundll32.exe shell32.dll,Control_RunDLL sysdm.cpl,,0", 5)
End Sub
Private Sub Label45_Click()
dblReturn = Shell("rundll32.exe shell32.dll,Control_RunDLL password.cpl", 5)
End Sub

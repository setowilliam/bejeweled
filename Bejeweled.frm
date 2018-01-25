VERSION 5.00
Begin VB.Form frmBejeweled 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bejeweled"
   ClientHeight    =   7095
   ClientLeft      =   225
   ClientTop       =   810
   ClientWidth     =   10500
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Palette         =   "Bejeweled.frx":0000
   Picture         =   "Bejeweled.frx":7572
   ScaleHeight     =   7095
   ScaleWidth      =   10500
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picScore 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   360
      ScaleHeight     =   465
      ScaleWidth      =   4305
      TabIndex        =   68
      Top             =   2280
      Width           =   4335
   End
   Begin VB.CommandButton cmdQuitGame 
      Appearance      =   0  'Flat
      Caption         =   "Quit Game"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   67
      Top             =   5280
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1425
      Left            =   120
      Picture         =   "Bejeweled.frx":840C
      ScaleHeight     =   1425
      ScaleWidth      =   4785
      TabIndex        =   66
      Top             =   120
      Width           =   4785
   End
   Begin VB.PictureBox picGem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   510
      Index           =   63
      Left            =   9480
      ScaleHeight     =   510
      ScaleWidth      =   510
      TabIndex        =   64
      Top             =   4800
      Width           =   510
   End
   Begin VB.PictureBox picGem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   510
      Index           =   62
      Left            =   8880
      ScaleHeight     =   510
      ScaleWidth      =   510
      TabIndex        =   63
      Top             =   4800
      Width           =   510
   End
   Begin VB.PictureBox picGem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   510
      Index           =   61
      Left            =   8280
      ScaleHeight     =   510
      ScaleWidth      =   510
      TabIndex        =   62
      Top             =   4800
      Width           =   510
   End
   Begin VB.PictureBox picGem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   510
      Index           =   60
      Left            =   7680
      ScaleHeight     =   510
      ScaleWidth      =   510
      TabIndex        =   61
      Top             =   4800
      Width           =   510
   End
   Begin VB.PictureBox picGem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   510
      Index           =   59
      Left            =   7080
      ScaleHeight     =   510
      ScaleWidth      =   510
      TabIndex        =   60
      Top             =   4800
      Width           =   510
   End
   Begin VB.PictureBox picGem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   510
      Index           =   58
      Left            =   6480
      ScaleHeight     =   510
      ScaleWidth      =   510
      TabIndex        =   59
      Top             =   4800
      Width           =   510
   End
   Begin VB.PictureBox picGem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   510
      Index           =   57
      Left            =   5880
      ScaleHeight     =   510
      ScaleWidth      =   510
      TabIndex        =   58
      Top             =   4800
      Width           =   510
   End
   Begin VB.PictureBox picGem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   510
      Index           =   56
      Left            =   5280
      ScaleHeight     =   510
      ScaleWidth      =   510
      TabIndex        =   57
      Top             =   4800
      Width           =   510
   End
   Begin VB.PictureBox picGem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   510
      Index           =   55
      Left            =   9480
      ScaleHeight     =   510
      ScaleWidth      =   510
      TabIndex        =   56
      Top             =   4200
      Width           =   510
   End
   Begin VB.PictureBox picGem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   510
      Index           =   54
      Left            =   8880
      ScaleHeight     =   510
      ScaleWidth      =   510
      TabIndex        =   55
      Top             =   4200
      Width           =   510
   End
   Begin VB.PictureBox picGem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   510
      Index           =   53
      Left            =   8280
      ScaleHeight     =   510
      ScaleWidth      =   510
      TabIndex        =   54
      Top             =   4200
      Width           =   510
   End
   Begin VB.PictureBox picGem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   510
      Index           =   52
      Left            =   7680
      ScaleHeight     =   510
      ScaleWidth      =   510
      TabIndex        =   53
      Top             =   4200
      Width           =   510
   End
   Begin VB.PictureBox picGem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   510
      Index           =   51
      Left            =   7080
      ScaleHeight     =   510
      ScaleWidth      =   510
      TabIndex        =   52
      Top             =   4200
      Width           =   510
   End
   Begin VB.PictureBox picGem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   510
      Index           =   50
      Left            =   6480
      ScaleHeight     =   510
      ScaleWidth      =   510
      TabIndex        =   51
      Top             =   4200
      Width           =   510
   End
   Begin VB.PictureBox picGem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   510
      Index           =   49
      Left            =   5880
      ScaleHeight     =   510
      ScaleWidth      =   510
      TabIndex        =   50
      Top             =   4200
      Width           =   510
   End
   Begin VB.PictureBox picGem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   510
      Index           =   48
      Left            =   5280
      ScaleHeight     =   510
      ScaleWidth      =   510
      TabIndex        =   49
      Top             =   4200
      Width           =   510
   End
   Begin VB.PictureBox picGem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   510
      Index           =   47
      Left            =   9480
      ScaleHeight     =   510
      ScaleWidth      =   510
      TabIndex        =   48
      Top             =   3600
      Width           =   510
   End
   Begin VB.PictureBox picGem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   510
      Index           =   46
      Left            =   8880
      ScaleHeight     =   510
      ScaleWidth      =   510
      TabIndex        =   47
      Top             =   3600
      Width           =   510
   End
   Begin VB.PictureBox picGem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   510
      Index           =   45
      Left            =   8280
      ScaleHeight     =   510
      ScaleWidth      =   510
      TabIndex        =   46
      Top             =   3600
      Width           =   510
   End
   Begin VB.PictureBox picGem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   510
      Index           =   44
      Left            =   7680
      ScaleHeight     =   510
      ScaleWidth      =   510
      TabIndex        =   45
      Top             =   3600
      Width           =   510
   End
   Begin VB.PictureBox picGem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   510
      Index           =   43
      Left            =   7080
      ScaleHeight     =   510
      ScaleWidth      =   510
      TabIndex        =   44
      Top             =   3600
      Width           =   510
   End
   Begin VB.PictureBox picGem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   510
      Index           =   42
      Left            =   6480
      ScaleHeight     =   510
      ScaleWidth      =   510
      TabIndex        =   43
      Top             =   3600
      Width           =   510
   End
   Begin VB.PictureBox picGem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   510
      Index           =   41
      Left            =   5880
      ScaleHeight     =   510
      ScaleWidth      =   510
      TabIndex        =   42
      Top             =   3600
      Width           =   510
   End
   Begin VB.PictureBox picGem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   510
      Index           =   40
      Left            =   5280
      ScaleHeight     =   510
      ScaleWidth      =   510
      TabIndex        =   41
      Top             =   3600
      Width           =   510
   End
   Begin VB.PictureBox picGem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   510
      Index           =   23
      Left            =   9480
      ScaleHeight     =   510
      ScaleWidth      =   510
      TabIndex        =   24
      Top             =   1800
      Width           =   510
   End
   Begin VB.PictureBox picGem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   510
      Index           =   22
      Left            =   8880
      ScaleHeight     =   510
      ScaleWidth      =   510
      TabIndex        =   23
      Top             =   1800
      Width           =   510
   End
   Begin VB.PictureBox picGem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   510
      Index           =   21
      Left            =   8280
      ScaleHeight     =   510
      ScaleWidth      =   510
      TabIndex        =   22
      Top             =   1800
      Width           =   510
   End
   Begin VB.PictureBox picGem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   510
      Index           =   20
      Left            =   7680
      ScaleHeight     =   510
      ScaleWidth      =   510
      TabIndex        =   21
      Top             =   1800
      Width           =   510
   End
   Begin VB.PictureBox picGem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   510
      Index           =   19
      Left            =   7080
      ScaleHeight     =   510
      ScaleWidth      =   510
      TabIndex        =   20
      Top             =   1800
      Width           =   510
   End
   Begin VB.PictureBox picGem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   510
      Index           =   18
      Left            =   6480
      ScaleHeight     =   510
      ScaleWidth      =   510
      TabIndex        =   19
      Top             =   1800
      Width           =   510
   End
   Begin VB.PictureBox picGem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   510
      Index           =   17
      Left            =   5880
      ScaleHeight     =   510
      ScaleWidth      =   510
      TabIndex        =   18
      Top             =   1800
      Width           =   510
   End
   Begin VB.PictureBox picGem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   510
      Index           =   16
      Left            =   5280
      ScaleHeight     =   510
      ScaleWidth      =   510
      TabIndex        =   17
      Top             =   1800
      Width           =   510
   End
   Begin VB.PictureBox picGem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   510
      Index           =   15
      Left            =   9480
      ScaleHeight     =   510
      ScaleWidth      =   510
      TabIndex        =   16
      Top             =   1200
      Width           =   510
   End
   Begin VB.PictureBox picGem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   510
      Index           =   14
      Left            =   8880
      ScaleHeight     =   510
      ScaleWidth      =   510
      TabIndex        =   15
      Top             =   1200
      Width           =   510
   End
   Begin VB.PictureBox picGem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   510
      Index           =   13
      Left            =   8280
      ScaleHeight     =   510
      ScaleWidth      =   510
      TabIndex        =   14
      Top             =   1200
      Width           =   510
   End
   Begin VB.PictureBox picGem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   510
      Index           =   12
      Left            =   7680
      ScaleHeight     =   510
      ScaleWidth      =   510
      TabIndex        =   13
      Top             =   1200
      Width           =   510
   End
   Begin VB.PictureBox picGem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   510
      Index           =   11
      Left            =   7080
      ScaleHeight     =   510
      ScaleWidth      =   510
      TabIndex        =   12
      Top             =   1200
      Width           =   510
   End
   Begin VB.PictureBox picGem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   510
      Index           =   10
      Left            =   6480
      ScaleHeight     =   510
      ScaleWidth      =   510
      TabIndex        =   11
      Top             =   1200
      Width           =   510
   End
   Begin VB.PictureBox picGem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   510
      Index           =   9
      Left            =   5880
      ScaleHeight     =   510
      ScaleWidth      =   510
      TabIndex        =   10
      Top             =   1200
      Width           =   510
   End
   Begin VB.PictureBox picGem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   510
      Index           =   8
      Left            =   5280
      ScaleHeight     =   510
      ScaleWidth      =   510
      TabIndex        =   9
      Top             =   1200
      Width           =   510
   End
   Begin VB.PictureBox picGem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   510
      Index           =   7
      Left            =   9480
      ScaleHeight     =   510
      ScaleWidth      =   510
      TabIndex        =   8
      Top             =   600
      Width           =   510
   End
   Begin VB.PictureBox picGem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   510
      Index           =   6
      Left            =   8880
      ScaleHeight     =   510
      ScaleWidth      =   510
      TabIndex        =   7
      Top             =   600
      Width           =   510
   End
   Begin VB.PictureBox picGem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   510
      Index           =   5
      Left            =   8280
      ScaleHeight     =   510
      ScaleWidth      =   510
      TabIndex        =   6
      Top             =   600
      Width           =   510
   End
   Begin VB.PictureBox picGem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   510
      Index           =   4
      Left            =   7680
      ScaleHeight     =   510
      ScaleWidth      =   510
      TabIndex        =   5
      Top             =   600
      Width           =   510
   End
   Begin VB.PictureBox picGem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   510
      Index           =   3
      Left            =   7080
      ScaleHeight     =   510
      ScaleWidth      =   510
      TabIndex        =   4
      Top             =   600
      Width           =   510
   End
   Begin VB.PictureBox picGem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   510
      Index           =   2
      Left            =   6480
      ScaleHeight     =   510
      ScaleWidth      =   510
      TabIndex        =   3
      Top             =   600
      Width           =   510
   End
   Begin VB.PictureBox picGem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   510
      Index           =   1
      Left            =   5880
      ScaleHeight     =   510
      ScaleWidth      =   510
      TabIndex        =   2
      Top             =   600
      Width           =   510
   End
   Begin VB.PictureBox picGem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   510
      Index           =   0
      Left            =   5280
      ScaleHeight     =   510
      ScaleWidth      =   510
      TabIndex        =   1
      Top             =   600
      Width           =   510
   End
   Begin VB.Timer tmrIllegalMove 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   4920
      Top             =   6480
   End
   Begin VB.PictureBox picGem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   510
      Index           =   24
      Left            =   5280
      ScaleHeight     =   510
      ScaleWidth      =   510
      TabIndex        =   25
      Top             =   2400
      Width           =   510
   End
   Begin VB.PictureBox picGem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   510
      Index           =   25
      Left            =   5880
      ScaleHeight     =   510
      ScaleWidth      =   510
      TabIndex        =   26
      Top             =   2400
      Width           =   510
   End
   Begin VB.PictureBox picGem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   510
      Index           =   26
      Left            =   6480
      ScaleHeight     =   510
      ScaleWidth      =   510
      TabIndex        =   27
      Top             =   2400
      Width           =   510
   End
   Begin VB.PictureBox picGem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   510
      Index           =   27
      Left            =   7080
      ScaleHeight     =   510
      ScaleWidth      =   510
      TabIndex        =   28
      Top             =   2400
      Width           =   510
   End
   Begin VB.PictureBox picGem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   510
      Index           =   28
      Left            =   7680
      ScaleHeight     =   510
      ScaleWidth      =   510
      TabIndex        =   29
      Top             =   2400
      Width           =   510
   End
   Begin VB.PictureBox picGem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   510
      Index           =   29
      Left            =   8280
      ScaleHeight     =   510
      ScaleWidth      =   510
      TabIndex        =   30
      Top             =   2400
      Width           =   510
   End
   Begin VB.PictureBox picGem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   510
      Index           =   30
      Left            =   8880
      ScaleHeight     =   510
      ScaleWidth      =   510
      TabIndex        =   31
      Top             =   2400
      Width           =   510
   End
   Begin VB.PictureBox picGem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   510
      Index           =   31
      Left            =   9480
      ScaleHeight     =   510
      ScaleWidth      =   510
      TabIndex        =   32
      Top             =   2400
      Width           =   510
   End
   Begin VB.PictureBox picGem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   510
      Index           =   32
      Left            =   5280
      ScaleHeight     =   510
      ScaleWidth      =   510
      TabIndex        =   33
      Top             =   3000
      Width           =   510
   End
   Begin VB.PictureBox picGem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   510
      Index           =   33
      Left            =   5880
      ScaleHeight     =   510
      ScaleWidth      =   510
      TabIndex        =   34
      Top             =   3000
      Width           =   510
   End
   Begin VB.PictureBox picGem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   510
      Index           =   34
      Left            =   6480
      ScaleHeight     =   510
      ScaleWidth      =   510
      TabIndex        =   35
      Top             =   3000
      Width           =   510
   End
   Begin VB.PictureBox picGem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   510
      Index           =   35
      Left            =   7080
      ScaleHeight     =   510
      ScaleWidth      =   510
      TabIndex        =   36
      Top             =   3000
      Width           =   510
   End
   Begin VB.PictureBox picGem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   510
      Index           =   36
      Left            =   7680
      ScaleHeight     =   510
      ScaleWidth      =   510
      TabIndex        =   37
      Top             =   3000
      Width           =   510
   End
   Begin VB.PictureBox picGem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   510
      Index           =   37
      Left            =   8280
      ScaleHeight     =   510
      ScaleWidth      =   510
      TabIndex        =   38
      Top             =   3000
      Width           =   510
   End
   Begin VB.PictureBox picGem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   510
      Index           =   38
      Left            =   8880
      ScaleHeight     =   510
      ScaleWidth      =   510
      TabIndex        =   39
      Top             =   3000
      Width           =   510
   End
   Begin VB.PictureBox picGem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   510
      Index           =   39
      Left            =   9480
      ScaleHeight     =   510
      ScaleWidth      =   510
      TabIndex        =   40
      Top             =   3000
      Width           =   510
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Score:"
      BeginProperty Font 
         Name            =   "@GungsuhChe"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   360
      TabIndex        =   65
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label lblIllegalMove 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ILLEGAL MOVE"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   855
      Left            =   4920
      TabIndex        =   0
      Top             =   5640
      Visible         =   0   'False
      Width           =   5415
   End
   Begin VB.Image imgGemImage 
      Height          =   510
      Index           =   7
      Left            =   4320
      Picture         =   "Bejeweled.frx":1E88E
      Stretch         =   -1  'True
      Top             =   6480
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Image imgGemImage 
      Height          =   510
      Index           =   6
      Left            =   3720
      Picture         =   "Bejeweled.frx":1F728
      Stretch         =   -1  'True
      Top             =   6480
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Image imgGemImage 
      Height          =   510
      Index           =   5
      Left            =   3120
      Picture         =   "Bejeweled.frx":2053A
      Stretch         =   -1  'True
      Top             =   6480
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Image imgGemImage 
      Height          =   510
      Index           =   4
      Left            =   2520
      Picture         =   "Bejeweled.frx":21440
      Stretch         =   -1  'True
      Top             =   6480
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Image imgGemImage 
      Height          =   510
      Index           =   3
      Left            =   1920
      Picture         =   "Bejeweled.frx":222DA
      Stretch         =   -1  'True
      Top             =   6480
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Image imgGemImage 
      Height          =   510
      Index           =   2
      Left            =   1320
      Picture         =   "Bejeweled.frx":230EC
      Stretch         =   -1  'True
      Top             =   6480
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Image imgGemImage 
      Height          =   510
      Index           =   1
      Left            =   720
      Picture         =   "Bejeweled.frx":23F86
      Stretch         =   -1  'True
      Top             =   6480
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Image imgGemImage 
      Height          =   510
      Index           =   0
      Left            =   120
      Picture         =   "Bejeweled.frx":24E20
      Stretch         =   -1  'True
      Top             =   6480
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Image imgGemBackground 
      Height          =   4935
      Left            =   5160
      Picture         =   "Bejeweled.frx":25CBA
      Stretch         =   -1  'True
      Top             =   480
      Width           =   4935
   End
   Begin VB.Image imgBorder 
      Height          =   5415
      Left            =   4920
      Picture         =   "Bejeweled.frx":26B54
      Stretch         =   -1  'True
      Top             =   240
      Width           =   5415
   End
   Begin VB.Image Image1 
      Height          =   7500
      Left            =   0
      Picture         =   "Bejeweled.frx":7322E
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10500
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNewGame 
         Caption         =   "&New Game"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuHighScores 
         Caption         =   "&High Scores"
         Shortcut        =   ^H
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmBejeweled"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Title: Bejeweled
'Author: William Seto
'Date: June 5, 2007
'Files: Bejeweled.bas, Bejeweled.frm, Bejeweled.frx, Bejeweled.vbp, Bejeweled.vbw, Bejeweled High Scores.txt,
'       frmHighScore.frm
'Purpose: The purpose of this program is to allow the user to play a replica of
'         Bejeweled.

Option Explicit

Dim IsMove As Boolean
Dim InitialGem As Integer
Dim SecondaryGem As Integer
Dim IllegalMoveCount As Integer
Dim VValidSwapOne As Boolean
Dim HValidSwapOne As Boolean
Dim VValidSwapTwo As Boolean
Dim HValidSwapTwo As Boolean
Dim Names(1 To 10) As String
Dim Scores(1 To 10) As Long
Dim CurrentScore As Long
Dim Gem(0 To 63) As Integer

Private Sub cmdQuitGame_Click()

    Dim k As Integer
    
    cmdQuitGame.Enabled = False
    
    For k = 0 To 63
        picGem(k).Picture = imgGemImage(7).Picture
    Next k
    EnterHighScore CurrentScore, Names(), Scores()
    
End Sub

Private Sub Form_Load()

    Dim k As Integer
    
    Randomize
    frmHighScore.Show
    InitialGem = -9999
    FillHighScores Names(), Scores()
    frmHighScore.Visible = False
    For k = 0 To 63
        picGem(k).Picture = imgGemImage(7).Picture
    Next k
    
End Sub

Private Sub mnuExit_Click()

    Dim DType As Integer
    Dim DTitle As String
    Dim DMsg As String
    Dim Response As Integer
    
    DType = vbYesNo + vbQuestion
    DTitle = "Termination"
    DMsg = "Are you sure you want to exit?"
    
    Response = MsgBox(DMsg, DType, DTitle)
    
    If Response = vbYes Then
        End
    End If
    
End Sub

Private Sub mnuHighScores_Click()

    Dim k As Integer
    
    frmHighScore.Show
    frmHighScore.Cls
    frmHighScore.Print Tab(5); "Name"; Tab(30); "Score"
    frmHighScore.Print
    
    Open App.Path & "\Bejeweled High Scores.txt" For Output As #1
    
        For k = 1 To 10
            Write #1, Names(k), Scores(k)
            frmHighScore.Print k; Tab(5); Names(k); Tab(30); Scores(k)
        Next k
    
    Close #1
    
End Sub

Private Sub mnuNewGame_Click()

    StartGame Gem(), InitialGem, IsMove
    cmdQuitGame.Enabled = True
    
End Sub

Private Sub picGem_Click(Index As Integer)
    
    Dim k As Integer
    Dim ValidMove As Boolean
    Dim BorderValid As Boolean
    Dim Num As Integer
    
    k = 0
    BorderValid = True
    
    Do
        If picGem(k).BorderStyle = 1 Then
            BorderValid = False
        End If
        k = k + 1
    Loop While BorderValid = True And k < 64
                
    If Gem(Index) = 8 Then
    ElseIf BorderValid = True Then
        picGem(Index).BorderStyle = 1
        InitialGem = Index
    ElseIf BorderValid = False And Index = InitialGem Then
        picGem(Index).BorderStyle = 0
        InitialGem = -9999
    ElseIf BorderValid = False And InitialGem <> -9999 Then
        picGem(InitialGem).BorderStyle = 0
        SecondaryGem = Index
        ValidMove = ValidMoveCheck(InitialGem, SecondaryGem)
        If ValidMove = False Then
            tmrIllegalMove.Enabled = True
        ElseIf ValidMove = True Then
            DestroyGems SecondaryGem, InitialGem, Gem(), HValidSwapOne, HValidSwapTwo, VValidSwapOne, VValidSwapTwo, CurrentScore
            For k = 0 To 63
                If Gem(k) = 8 Then
                    Num = GenerateSingleGem
                    Gem(k) = Num
                    picGem(k).Picture = imgGemImage(Num).Picture
                End If
            Next k
            Delay 0.5
            DestroyDropGem Gem(), CurrentScore
            IsGameOver Gem(), CurrentScore, Names(), Scores()
        End If
    End If
    
End Sub

Private Sub tmrIllegalMove_Timer()
    
    IllegalMoveCount = IllegalMoveCount + 1
    
    lblIllegalMove.Visible = Not lblIllegalMove.Visible
    
    If IllegalMoveCount = 6 Then
        tmrIllegalMove.Enabled = False
        IllegalMoveCount = 0
    End If
    
End Sub

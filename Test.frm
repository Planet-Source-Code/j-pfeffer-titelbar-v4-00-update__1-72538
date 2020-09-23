VERSION 5.00
Begin VB.Form Test 
   BackColor       =   &H00EBE6E4&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7455
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11700
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Test.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   497
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   780
   StartUpPosition =   2  'CenterScreen
   Begin TBar.TitelBar TitelBar4 
      Height          =   555
      Left            =   -45
      TabIndex        =   0
      Top             =   6615
      Width           =   11700
      _ExtentX        =   20638
      _ExtentY        =   979
      BackColor       =   5189670
      Style           =   14
      BackColorCover  =   3
      BackColorV2Begin=   4603449
      BackColorV2End  =   2762275
      BackColorV1Begin=   7760481
      BackColorV1End  =   6050636
      ForeColor       =   16777215
      ShowMaximizedEnabled=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Vista Style"
      BorderNormal    =   2
      IconStyle       =   2
      BorderColorHighLight=   14737632
      BorderColorDarkLight=   4210752
   End
   Begin TBar.TitelBar TitelBar5 
      Height          =   555
      Left            =   0
      TabIndex        =   1
      Top             =   5760
      Width           =   11700
      _ExtentX        =   20638
      _ExtentY        =   979
      BackColor       =   8421504
      Style           =   13
      BackColorCover  =   3
      BackColorV2Begin=   8421504
      BackColorV2End  =   12632256
      BackColorV1Begin=   14737632
      BackColorV1End  =   4210752
      ForeColor       =   16777215
      ShowMaximized   =   -1  'True
      ShowMaximizedEnabled=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Modern Style"
      BorderNormal    =   2
   End
   Begin TBar.TitelBar TitelBar6 
      Height          =   645
      Left            =   0
      TabIndex        =   2
      Top             =   4905
      Width           =   11700
      _ExtentX        =   15028
      _ExtentY        =   1138
      BackColor       =   0
      Style           =   12
      BackColorCover  =   3
      ForeColor       =   16777215
      ShowMaximized   =   -1  'True
      ShowMaximizedEnabled=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PictureButton   =   "Test.frx":0A22
      BorderNormal    =   2
      IconStyle       =   1
   End
   Begin TBar.TitelBar TitelBar7 
      Height          =   645
      Left            =   0
      TabIndex        =   3
      Top             =   4050
      Width           =   11700
      _ExtentX        =   15028
      _ExtentY        =   1138
      BackColor       =   5189670
      Style           =   9
      BackColorCover  =   4
      BackColorV2Begin=   8421376
      BackColorV2End  =   12632064
      BackColorV1Begin=   16761087
      BackColorV1End  =   8388736
      ForeColor       =   16777215
      ShowMaximizedEnabled=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "2 Zones Soft"
      BorderNormal    =   2
   End
   Begin TBar.TitelBar TitelBar8 
      Height          =   645
      Left            =   0
      TabIndex        =   4
      Top             =   3240
      Width           =   11700
      _ExtentX        =   15028
      _ExtentY        =   1138
      BackColor       =   5189670
      BackColorCover  =   7
      BackColorV2Begin=   14737632
      BackColorV2End  =   4210752
      BackColorV1Begin=   6179120
      ForeColor       =   32768
      CaptionColorBack=   4210752
      CaptionColor    =   32768
      ShowMaximized   =   -1  'True
      ShowMaximizedEnabled=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "2 Zones Stripes"
      CaptionPosX     =   2
      PictureButton   =   "Test.frx":12FC
      BorderNormal    =   2
      BorderColorHighLight=   0
      Caption3DWidth  =   3
      CaptionBorder   =   -1  'True
      CaptionBorderColor=   12648384
   End
   Begin TBar.TitelBar TitelBar9 
      Height          =   645
      Left            =   0
      TabIndex        =   5
      Top             =   2430
      Width           =   11700
      _ExtentX        =   15028
      _ExtentY        =   1138
      BackColor       =   5189670
      Style           =   5
      BackColorCover  =   3
      BackColorV1End  =   16576
      ForeColor       =   16777215
      ShowMaximizedEnabled=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "2 Zones simple left"
      BorderNormal    =   2
      Caption3DTop    =   -1  'True
      Caption3DLeft   =   -1  'True
   End
   Begin TBar.TitelBar TitelBar10 
      Height          =   645
      Left            =   0
      TabIndex        =   6
      Top             =   1620
      Width           =   11700
      _ExtentX        =   15028
      _ExtentY        =   1138
      BackColor       =   5189670
      Style           =   4
      BackColorCover  =   2
      ForeColor       =   65535
      CaptionColor    =   65535
      ShowMaximizedEnabled=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "2 Zones right Stripe"
      BorderNormal    =   2
      IconStyle       =   2
   End
   Begin TBar.TitelBar TitelBar11 
      Height          =   645
      Left            =   0
      TabIndex        =   7
      Top             =   855
      Width           =   11700
      _ExtentX        =   15028
      _ExtentY        =   1138
      BackColor       =   5189670
      Style           =   1
      BackColorCover  =   6
      BackColorV2Begin=   16761024
      BackColorV2End  =   12583104
      BackColorV1End  =   32768
      ForeColor       =   16777215
      ShowMaximizedEnabled=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "2 Zones simple right"
      BorderNormal    =   2
   End
   Begin TBar.TitelBar TitelBar12 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   11700
      _ExtentX        =   15028
      _ExtentY        =   1138
      BackColor       =   5189670
      Style           =   0
      BackColorCover  =   3
      ForeColor       =   16777215
      ShowMaximized   =   -1  'True
      ShowMaximizedEnabled=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Simple Style"
      BorderNormal    =   2
   End
End
Attribute VB_Name = "Test"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

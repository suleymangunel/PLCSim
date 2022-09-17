VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8625
   ClientLeft      =   1665
   ClientTop       =   945
   ClientWidth     =   11910
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   162
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8625
   ScaleWidth      =   11910
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   3240
      Top             =   8520
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8475
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   14949
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   5
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Intro"
      TabPicture(0)   =   "Form1.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "PLCSim"
      TabPicture(1)   =   "Form1.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "StatLabel"
      Tab(1).Control(1)=   "Frame4"
      Tab(1).Control(2)=   "Frame5"
      Tab(1).Control(3)=   "Frame6"
      Tab(1).Control(4)=   "Frame9"
      Tab(1).Control(5)=   "Frame1"
      Tab(1).Control(6)=   "Frame3"
      Tab(1).Control(7)=   "Frame7"
      Tab(1).ControlCount=   8
      TabCaption(2)   =   "Ayarlar"
      TabPicture(2)   =   "Form1.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame17"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Açýklama"
      TabPicture(3)   =   "Form1.frx":0496
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "RichTextBox1"
      Tab(3).ControlCount=   1
      Begin RichTextLib.RichTextBox RichTextBox1 
         Height          =   7935
         Left            =   -74880
         TabIndex        =   255
         Top             =   420
         Width           =   11535
         _ExtentX        =   20346
         _ExtentY        =   13996
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         DisableNoScroll =   -1  'True
         FileName        =   "C:\Pr\VB\PLCSim\PLCSim.rtf"
         TextRTF         =   $"Form1.frx":04B2
      End
      Begin VB.Frame Frame7 
         Caption         =   "Reset"
         Enabled         =   0   'False
         Height          =   1395
         Left            =   -65640
         TabIndex        =   233
         Top             =   4260
         Width           =   2295
         Begin VB.CommandButton Command11 
            Height          =   1035
            Left            =   120
            Picture         =   "Form1.frx":059E
            Style           =   1  'Graphical
            TabIndex        =   183
            Top             =   240
            Width           =   2055
         End
      End
      Begin VB.Frame Frame17 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7875
         Left            =   -74880
         TabIndex        =   223
         Top             =   420
         Width           =   11535
         Begin VB.Frame Frame19 
            Caption         =   "Açýklama"
            Height          =   5775
            Left            =   4980
            TabIndex        =   250
            Top             =   1980
            Width           =   6435
            Begin RichTextLib.RichTextBox RichTextBox2 
               Height          =   3435
               Left            =   120
               TabIndex        =   256
               Top             =   660
               Width           =   6195
               _ExtentX        =   10927
               _ExtentY        =   6059
               _Version        =   393217
               Enabled         =   -1  'True
               ScrollBars      =   2
               DisableNoScroll =   -1  'True
               FileName        =   "C:\Pr\VB\PLCSim\Paralel.rtf"
               TextRTF         =   $"Form1.frx":08A8
            End
            Begin VB.PictureBox RichTextBox3 
               Height          =   1095
               Left            =   120
               ScaleHeight     =   1035
               ScaleWidth      =   6135
               TabIndex        =   251
               Top             =   4560
               Width           =   6195
            End
            Begin VB.Label Label30 
               AutoSize        =   -1  'True
               Caption         =   "Dosya Ýþlemleri"
               Height          =   195
               Left            =   120
               TabIndex        =   253
               Top             =   4260
               Width           =   1290
            End
            Begin VB.Label Label29 
               AutoSize        =   -1  'True
               Caption         =   "Paralel Port Adresi"
               Height          =   195
               Left            =   120
               TabIndex        =   252
               Top             =   360
               Width           =   1590
            End
         End
         Begin VB.Frame Frame8 
            Caption         =   "Port Kontrol"
            Height          =   1635
            Left            =   4980
            TabIndex        =   239
            Top             =   240
            Width           =   6435
            Begin VB.CommandButton Command1 
               Caption         =   "Uygula"
               Height          =   1155
               Left            =   3060
               TabIndex        =   249
               Top             =   360
               Width           =   3255
            End
            Begin VB.OptionButton Option3 
               Height          =   195
               Left            =   2460
               TabIndex        =   248
               Top             =   1260
               Width           =   195
            End
            Begin VB.OptionButton Option2 
               Height          =   195
               Left            =   2460
               TabIndex        =   247
               Top             =   840
               Width           =   195
            End
            Begin VB.OptionButton Option1 
               Height          =   195
               Left            =   2460
               TabIndex        =   246
               Top             =   420
               Width           =   195
            End
            Begin VB.TextBox Text4 
               Height          =   315
               Left            =   1440
               TabIndex        =   243
               Top             =   1200
               Width           =   855
            End
            Begin VB.Label Port1 
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Label30"
               Height          =   315
               Left            =   1440
               TabIndex        =   245
               Top             =   780
               Width           =   855
            End
            Begin VB.Label Port0 
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Label29"
               Height          =   315
               Left            =   1440
               TabIndex        =   244
               Top             =   360
               Width           =   855
            End
            Begin VB.Label Label28 
               AutoSize        =   -1  'True
               Caption         =   "Kullanýcý"
               Height          =   195
               Left            =   120
               TabIndex        =   242
               Top             =   1260
               Width           =   735
            End
            Begin VB.Label Label27 
               AutoSize        =   -1  'True
               Caption         =   "Port-1 (LPT2)"
               Height          =   195
               Left            =   120
               TabIndex        =   241
               Top             =   840
               Width           =   1155
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               Caption         =   "Port-0 (LPT1)"
               Height          =   195
               Left            =   120
               TabIndex        =   240
               Top             =   420
               Width           =   1155
            End
         End
         Begin VB.Frame Frame18 
            Caption         =   "Dosya Ýþlemleri"
            Height          =   7515
            Left            =   120
            TabIndex        =   224
            Top             =   240
            Width           =   4755
            Begin VB.CommandButton Command14 
               Caption         =   "Sil"
               Height          =   1155
               Left            =   3240
               Picture         =   "Form1.frx":0994
               Style           =   1  'Graphical
               TabIndex        =   231
               Top             =   5700
               Width           =   1395
            End
            Begin VB.CommandButton Command10 
               Caption         =   "Kaydet"
               Height          =   1155
               Left            =   120
               Picture         =   "Form1.frx":125E
               Style           =   1  'Graphical
               TabIndex        =   230
               Top             =   5700
               Width           =   1395
            End
            Begin VB.CommandButton Command9 
               Caption         =   "Yükle"
               Height          =   1155
               Left            =   1680
               Picture         =   "Form1.frx":1B28
               Style           =   1  'Graphical
               TabIndex        =   229
               Top             =   5700
               Width           =   1395
            End
            Begin VB.TextBox Text1 
               BackColor       =   &H8000000F&
               ForeColor       =   &H80000012&
               Height          =   300
               Left            =   840
               TabIndex        =   228
               Top             =   5220
               Width           =   2475
            End
            Begin VB.FileListBox File1 
               Height          =   4380
               Left            =   2400
               TabIndex        =   227
               Top             =   780
               Width           =   2235
            End
            Begin VB.DirListBox Dir1 
               Height          =   4365
               Left            =   120
               TabIndex        =   226
               Top             =   780
               Width           =   2235
            End
            Begin VB.DriveListBox Drive1 
               Height          =   315
               Left            =   120
               TabIndex        =   225
               Top             =   360
               Width           =   4515
            End
            Begin VB.Label Label16 
               AutoSize        =   -1  'True
               Caption         =   "Dosya:"
               Height          =   195
               Left            =   180
               TabIndex        =   254
               Top             =   5280
               Width           =   600
            End
            Begin VB.Label Label9 
               BackColor       =   &H00000000&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   14.25
                  Charset         =   162
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H0000FF00&
               Height          =   375
               Left            =   120
               TabIndex        =   235
               Top             =   7020
               Width           =   4515
            End
            Begin VB.Label Label3 
               BorderStyle     =   1  'Fixed Single
               Height          =   300
               Left            =   3360
               TabIndex        =   232
               Top             =   5220
               Width           =   1275
            End
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Genel Kontroller"
         Height          =   7395
         Left            =   -67500
         TabIndex        =   219
         Top             =   420
         Width           =   1755
         Begin VB.Frame Frame16 
            Height          =   4935
            Left            =   120
            TabIndex        =   220
            Top             =   300
            Width           =   1515
            Begin VB.CommandButton Command8 
               Caption         =   "Reset Blok"
               Height          =   975
               Left            =   120
               Picture         =   "Form1.frx":23F2
               Style           =   1  'Graphical
               TabIndex        =   182
               Top             =   3300
               Width           =   1275
            End
            Begin VB.CommandButton Command7 
               Caption         =   "Blok Kaydet"
               Height          =   975
               Left            =   120
               Picture         =   "Form1.frx":2834
               Style           =   1  'Graphical
               TabIndex        =   180
               Top             =   1260
               Width           =   1275
            End
            Begin VB.CommandButton Command5 
               Caption         =   "Önceki Blok"
               Height          =   975
               Left            =   120
               Picture         =   "Form1.frx":2B3E
               Style           =   1  'Graphical
               TabIndex        =   179
               Top             =   240
               Width           =   1275
            End
            Begin VB.CommandButton Command6 
               Caption         =   "Sonraki Blok"
               Height          =   975
               Left            =   120
               Picture         =   "Form1.frx":2F80
               Style           =   1  'Graphical
               TabIndex        =   181
               Top             =   2280
               Width           =   1275
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               Caption         =   "Blok No"
               Height          =   195
               Left            =   120
               TabIndex        =   222
               Top             =   4500
               Width           =   690
            End
            Begin VB.Label BlokNo 
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   162
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   375
               Left            =   900
               TabIndex        =   221
               Top             =   4440
               Width           =   495
            End
         End
         Begin VB.CommandButton Command4 
            Height          =   1635
            Left            =   120
            Picture         =   "Form1.frx":33C2
            Style           =   1  'Graphical
            TabIndex        =   177
            Top             =   5640
            Width           =   1515
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "PLCSim"
         Height          =   7395
         Left            =   -74880
         TabIndex        =   204
         Top             =   420
         Width           =   7275
         Begin VB.Frame Frame14 
            Height          =   7095
            Left            =   6000
            TabIndex        =   215
            Top             =   180
            Width           =   30
         End
         Begin VB.Frame Frame13 
            Height          =   7095
            Left            =   4740
            TabIndex        =   214
            Top             =   180
            Width           =   30
         End
         Begin VB.Frame Frame12 
            Height          =   7095
            Left            =   3480
            TabIndex        =   213
            Top             =   180
            Width           =   30
         End
         Begin VB.Frame Frame11 
            Height          =   7095
            Left            =   2220
            TabIndex        =   212
            Top             =   180
            Width           =   30
         End
         Begin VB.Frame Frame10 
            Height          =   7095
            Left            =   1260
            TabIndex        =   217
            Top             =   180
            Width           =   30
         End
         Begin VB.Frame Frame15 
            Height          =   75
            Left            =   120
            TabIndex        =   216
            Top             =   480
            Width           =   7050
         End
         Begin VB.TextBox Out0St 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   2340
            TabIndex        =   4
            Top             =   660
            Width           =   495
         End
         Begin VB.TextBox Out0Pw 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   2880
            TabIndex        =   5
            Top             =   660
            Width           =   495
         End
         Begin VB.TextBox Out1Pw 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   4140
            TabIndex        =   7
            Top             =   660
            Width           =   495
         End
         Begin VB.TextBox Out1St 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   3600
            TabIndex        =   6
            Top             =   660
            Width           =   495
         End
         Begin VB.TextBox Out2Pw 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   5400
            TabIndex        =   9
            Top             =   660
            Width           =   495
         End
         Begin VB.TextBox Out2St 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   4860
            TabIndex        =   8
            Top             =   660
            Width           =   495
         End
         Begin VB.TextBox Out3Pw 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   6660
            TabIndex        =   11
            Top             =   660
            Width           =   495
         End
         Begin VB.TextBox Out3St 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   6120
            TabIndex        =   10
            Top             =   660
            Width           =   495
         End
         Begin VB.TextBox Out3St 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   6120
            TabIndex        =   21
            Top             =   1080
            Width           =   495
         End
         Begin VB.TextBox Out3Pw 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   6660
            TabIndex        =   22
            Top             =   1080
            Width           =   495
         End
         Begin VB.TextBox Out2St 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   4860
            TabIndex        =   19
            Top             =   1080
            Width           =   495
         End
         Begin VB.TextBox Out2Pw 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   5400
            TabIndex        =   20
            Top             =   1080
            Width           =   495
         End
         Begin VB.TextBox Out1St 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   3600
            TabIndex        =   17
            Top             =   1080
            Width           =   495
         End
         Begin VB.TextBox Out1Pw 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   4140
            TabIndex        =   18
            Top             =   1080
            Width           =   495
         End
         Begin VB.TextBox Out0Pw 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   2880
            TabIndex        =   16
            Top             =   1080
            Width           =   495
         End
         Begin VB.TextBox Out0St 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   2340
            TabIndex        =   15
            Top             =   1080
            Width           =   495
         End
         Begin VB.TextBox Out3St 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   2
            Left            =   6120
            TabIndex        =   32
            Top             =   1500
            Width           =   495
         End
         Begin VB.TextBox Out3Pw 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   2
            Left            =   6660
            TabIndex        =   33
            Top             =   1500
            Width           =   495
         End
         Begin VB.TextBox Out2St 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   2
            Left            =   4860
            TabIndex        =   30
            Top             =   1500
            Width           =   495
         End
         Begin VB.TextBox Out2Pw 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   2
            Left            =   5400
            TabIndex        =   31
            Top             =   1500
            Width           =   495
         End
         Begin VB.TextBox Out1St 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   2
            Left            =   3600
            TabIndex        =   28
            Top             =   1500
            Width           =   495
         End
         Begin VB.TextBox Out1Pw 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   2
            Left            =   4140
            TabIndex        =   29
            Top             =   1500
            Width           =   495
         End
         Begin VB.TextBox Out0Pw 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   2
            Left            =   2880
            TabIndex        =   27
            Top             =   1500
            Width           =   495
         End
         Begin VB.TextBox Out0St 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   2
            Left            =   2340
            TabIndex        =   26
            Top             =   1500
            Width           =   495
         End
         Begin VB.TextBox Out3St 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   3
            Left            =   6120
            TabIndex        =   43
            Top             =   1920
            Width           =   495
         End
         Begin VB.TextBox Out3Pw 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   3
            Left            =   6660
            TabIndex        =   44
            Top             =   1920
            Width           =   495
         End
         Begin VB.TextBox Out2St 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   3
            Left            =   4860
            TabIndex        =   41
            Top             =   1920
            Width           =   495
         End
         Begin VB.TextBox Out2Pw 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   3
            Left            =   5400
            TabIndex        =   42
            Top             =   1920
            Width           =   495
         End
         Begin VB.TextBox Out1St 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   3
            Left            =   3600
            TabIndex        =   39
            Top             =   1920
            Width           =   495
         End
         Begin VB.TextBox Out1Pw 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   3
            Left            =   4140
            TabIndex        =   40
            Top             =   1920
            Width           =   495
         End
         Begin VB.TextBox Out0Pw 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   3
            Left            =   2880
            TabIndex        =   38
            Top             =   1920
            Width           =   495
         End
         Begin VB.TextBox Out0St 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   3
            Left            =   2340
            TabIndex        =   37
            Top             =   1920
            Width           =   495
         End
         Begin VB.TextBox Out3St 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   4
            Left            =   6120
            TabIndex        =   54
            Top             =   2340
            Width           =   495
         End
         Begin VB.TextBox Out3Pw 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   4
            Left            =   6660
            TabIndex        =   55
            Top             =   2340
            Width           =   495
         End
         Begin VB.TextBox Out2St 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   4
            Left            =   4860
            TabIndex        =   52
            Top             =   2340
            Width           =   495
         End
         Begin VB.TextBox Out2Pw 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   4
            Left            =   5400
            TabIndex        =   53
            Top             =   2340
            Width           =   495
         End
         Begin VB.TextBox Out1St 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   4
            Left            =   3600
            TabIndex        =   50
            Top             =   2340
            Width           =   495
         End
         Begin VB.TextBox Out1Pw 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   4
            Left            =   4140
            TabIndex        =   51
            Top             =   2340
            Width           =   495
         End
         Begin VB.TextBox Out0Pw 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   4
            Left            =   2880
            TabIndex        =   49
            Top             =   2340
            Width           =   495
         End
         Begin VB.TextBox Out0St 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   4
            Left            =   2340
            TabIndex        =   48
            Top             =   2340
            Width           =   495
         End
         Begin VB.TextBox Out3St 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   5
            Left            =   6120
            TabIndex        =   65
            Top             =   2760
            Width           =   495
         End
         Begin VB.TextBox Out3Pw 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   5
            Left            =   6660
            TabIndex        =   66
            Top             =   2760
            Width           =   495
         End
         Begin VB.TextBox Out2St 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   5
            Left            =   4860
            TabIndex        =   63
            Top             =   2760
            Width           =   495
         End
         Begin VB.TextBox Out2Pw 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   5
            Left            =   5460
            TabIndex        =   64
            Top             =   2760
            Width           =   495
         End
         Begin VB.TextBox Out1St 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   5
            Left            =   3600
            TabIndex        =   61
            Top             =   2760
            Width           =   495
         End
         Begin VB.TextBox Out1Pw 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   5
            Left            =   4140
            TabIndex        =   62
            Top             =   2760
            Width           =   495
         End
         Begin VB.TextBox Out0Pw 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   5
            Left            =   2880
            TabIndex        =   60
            Top             =   2760
            Width           =   495
         End
         Begin VB.TextBox Out0St 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   5
            Left            =   2340
            TabIndex        =   59
            Top             =   2760
            Width           =   495
         End
         Begin VB.TextBox Out3St 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   6
            Left            =   6120
            TabIndex        =   76
            Top             =   3180
            Width           =   495
         End
         Begin VB.TextBox Out3Pw 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   6
            Left            =   6660
            TabIndex        =   77
            Top             =   3180
            Width           =   495
         End
         Begin VB.TextBox Out2St 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   6
            Left            =   4860
            TabIndex        =   74
            Top             =   3180
            Width           =   495
         End
         Begin VB.TextBox Out2Pw 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   6
            Left            =   5400
            TabIndex        =   75
            Top             =   3180
            Width           =   495
         End
         Begin VB.TextBox Out1St 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   6
            Left            =   3600
            TabIndex        =   72
            Top             =   3180
            Width           =   495
         End
         Begin VB.TextBox Out1Pw 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   6
            Left            =   4140
            TabIndex        =   73
            Top             =   3180
            Width           =   495
         End
         Begin VB.TextBox Out0Pw 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   6
            Left            =   2880
            TabIndex        =   71
            Top             =   3180
            Width           =   495
         End
         Begin VB.TextBox Out0St 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   6
            Left            =   2340
            TabIndex        =   70
            Top             =   3180
            Width           =   495
         End
         Begin VB.TextBox Out3St 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   7
            Left            =   6120
            TabIndex        =   87
            Top             =   3600
            Width           =   495
         End
         Begin VB.TextBox Out3Pw 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   7
            Left            =   6660
            TabIndex        =   88
            Top             =   3600
            Width           =   495
         End
         Begin VB.TextBox Out2St 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   7
            Left            =   4860
            TabIndex        =   85
            Top             =   3600
            Width           =   495
         End
         Begin VB.TextBox Out2Pw 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   7
            Left            =   5400
            TabIndex        =   86
            Top             =   3600
            Width           =   495
         End
         Begin VB.TextBox Out1St 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   7
            Left            =   3600
            TabIndex        =   83
            Top             =   3600
            Width           =   495
         End
         Begin VB.TextBox Out1Pw 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   7
            Left            =   4140
            TabIndex        =   84
            Top             =   3600
            Width           =   495
         End
         Begin VB.TextBox Out0Pw 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   7
            Left            =   2880
            TabIndex        =   82
            Top             =   3600
            Width           =   495
         End
         Begin VB.TextBox Out0St 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   7
            Left            =   2340
            TabIndex        =   81
            Top             =   3600
            Width           =   495
         End
         Begin VB.TextBox Out3St 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   8
            Left            =   6120
            TabIndex        =   98
            Top             =   4020
            Width           =   495
         End
         Begin VB.TextBox Out3Pw 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   8
            Left            =   6660
            TabIndex        =   99
            Top             =   4020
            Width           =   495
         End
         Begin VB.TextBox Out2St 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   8
            Left            =   4860
            TabIndex        =   96
            Top             =   4020
            Width           =   495
         End
         Begin VB.TextBox Out2Pw 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   8
            Left            =   5400
            TabIndex        =   97
            Top             =   4020
            Width           =   495
         End
         Begin VB.TextBox Out1St 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   8
            Left            =   3600
            TabIndex        =   94
            Top             =   4020
            Width           =   495
         End
         Begin VB.TextBox Out1Pw 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   8
            Left            =   4140
            TabIndex        =   95
            Top             =   4020
            Width           =   495
         End
         Begin VB.TextBox Out0Pw 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   8
            Left            =   2880
            TabIndex        =   93
            Top             =   4020
            Width           =   495
         End
         Begin VB.TextBox Out0St 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   8
            Left            =   2340
            TabIndex        =   92
            Top             =   4020
            Width           =   495
         End
         Begin VB.TextBox Out3St 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   9
            Left            =   6120
            TabIndex        =   109
            Top             =   4440
            Width           =   495
         End
         Begin VB.TextBox Out3Pw 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   9
            Left            =   6660
            TabIndex        =   110
            Top             =   4440
            Width           =   495
         End
         Begin VB.TextBox Out2St 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   9
            Left            =   4860
            TabIndex        =   107
            Top             =   4440
            Width           =   495
         End
         Begin VB.TextBox Out2Pw 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   9
            Left            =   5400
            TabIndex        =   108
            Top             =   4440
            Width           =   495
         End
         Begin VB.TextBox Out1St 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   9
            Left            =   3600
            TabIndex        =   105
            Top             =   4440
            Width           =   495
         End
         Begin VB.TextBox Out1Pw 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   9
            Left            =   4140
            TabIndex        =   106
            Top             =   4440
            Width           =   495
         End
         Begin VB.TextBox Out0Pw 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   9
            Left            =   2880
            TabIndex        =   104
            Top             =   4440
            Width           =   495
         End
         Begin VB.TextBox Out0St 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   9
            Left            =   2340
            TabIndex        =   103
            Top             =   4440
            Width           =   495
         End
         Begin VB.TextBox Out3St 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   10
            Left            =   6120
            TabIndex        =   120
            Top             =   4860
            Width           =   495
         End
         Begin VB.TextBox Out3Pw 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   10
            Left            =   6660
            TabIndex        =   121
            Top             =   4860
            Width           =   495
         End
         Begin VB.TextBox Out2St 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   10
            Left            =   4860
            TabIndex        =   118
            Top             =   4860
            Width           =   495
         End
         Begin VB.TextBox Out2Pw 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   10
            Left            =   5400
            TabIndex        =   119
            Top             =   4860
            Width           =   495
         End
         Begin VB.TextBox Out1St 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   10
            Left            =   3600
            TabIndex        =   116
            Top             =   4860
            Width           =   495
         End
         Begin VB.TextBox Out1Pw 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   10
            Left            =   4140
            TabIndex        =   117
            Top             =   4860
            Width           =   495
         End
         Begin VB.TextBox Out0Pw 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   10
            Left            =   2880
            TabIndex        =   115
            Top             =   4860
            Width           =   495
         End
         Begin VB.TextBox Out0St 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   10
            Left            =   2340
            TabIndex        =   114
            Top             =   4860
            Width           =   495
         End
         Begin VB.TextBox Out3St 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   11
            Left            =   6120
            TabIndex        =   131
            Top             =   5280
            Width           =   495
         End
         Begin VB.TextBox Out3Pw 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   11
            Left            =   6660
            TabIndex        =   132
            Top             =   5280
            Width           =   495
         End
         Begin VB.TextBox Out2St 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   11
            Left            =   4860
            TabIndex        =   129
            Top             =   5280
            Width           =   495
         End
         Begin VB.TextBox Out2Pw 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   11
            Left            =   5400
            TabIndex        =   130
            Top             =   5280
            Width           =   495
         End
         Begin VB.TextBox Out1St 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   11
            Left            =   3600
            TabIndex        =   127
            Top             =   5280
            Width           =   495
         End
         Begin VB.TextBox Out1Pw 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   11
            Left            =   4140
            TabIndex        =   128
            Top             =   5280
            Width           =   495
         End
         Begin VB.TextBox Out0Pw 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   11
            Left            =   2880
            TabIndex        =   126
            Top             =   5280
            Width           =   495
         End
         Begin VB.TextBox Out0St 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   11
            Left            =   2340
            TabIndex        =   125
            Top             =   5280
            Width           =   495
         End
         Begin VB.TextBox Out0St 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   13
            Left            =   2340
            TabIndex        =   147
            Top             =   6120
            Width           =   495
         End
         Begin VB.TextBox Out0Pw 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   13
            Left            =   2880
            TabIndex        =   148
            Top             =   6120
            Width           =   495
         End
         Begin VB.TextBox Out1Pw 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   13
            Left            =   4140
            TabIndex        =   150
            Top             =   6120
            Width           =   495
         End
         Begin VB.TextBox Out1St 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   13
            Left            =   3600
            TabIndex        =   149
            Top             =   6120
            Width           =   495
         End
         Begin VB.TextBox Out2Pw 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   13
            Left            =   5400
            TabIndex        =   152
            Top             =   6120
            Width           =   495
         End
         Begin VB.TextBox Out2St 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   13
            Left            =   4860
            TabIndex        =   151
            Top             =   6120
            Width           =   495
         End
         Begin VB.TextBox Out0St 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   15
            Left            =   2340
            TabIndex        =   169
            Top             =   6960
            Width           =   495
         End
         Begin VB.TextBox Out0Pw 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   15
            Left            =   2880
            TabIndex        =   170
            Top             =   6960
            Width           =   495
         End
         Begin VB.TextBox Out1Pw 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   15
            Left            =   4140
            TabIndex        =   172
            Top             =   6960
            Width           =   495
         End
         Begin VB.TextBox Out1St 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   15
            Left            =   3600
            TabIndex        =   171
            Top             =   6960
            Width           =   495
         End
         Begin VB.TextBox Out2Pw 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   15
            Left            =   5400
            TabIndex        =   174
            Top             =   6960
            Width           =   495
         End
         Begin VB.TextBox Out2St 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   15
            Left            =   4860
            TabIndex        =   173
            Top             =   6960
            Width           =   495
         End
         Begin VB.TextBox Out3Pw 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   15
            Left            =   6660
            TabIndex        =   176
            Top             =   6960
            Width           =   495
         End
         Begin VB.TextBox Out3St 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   15
            Left            =   6120
            TabIndex        =   175
            Top             =   6960
            Width           =   495
         End
         Begin VB.TextBox Out3St 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   14
            Left            =   6120
            TabIndex        =   164
            Top             =   6540
            Width           =   495
         End
         Begin VB.TextBox Out3Pw 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   14
            Left            =   6660
            TabIndex        =   165
            Top             =   6540
            Width           =   495
         End
         Begin VB.TextBox Out2St 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   14
            Left            =   4860
            TabIndex        =   162
            Top             =   6540
            Width           =   495
         End
         Begin VB.TextBox Out2Pw 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   14
            Left            =   5400
            TabIndex        =   163
            Top             =   6540
            Width           =   495
         End
         Begin VB.TextBox Out1St 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   14
            Left            =   3600
            TabIndex        =   160
            Top             =   6540
            Width           =   495
         End
         Begin VB.TextBox Out1Pw 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   14
            Left            =   4140
            TabIndex        =   161
            Top             =   6540
            Width           =   495
         End
         Begin VB.TextBox Out0Pw 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   14
            Left            =   2880
            TabIndex        =   159
            Top             =   6540
            Width           =   495
         End
         Begin VB.TextBox Out0St 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   14
            Left            =   2340
            TabIndex        =   158
            Top             =   6540
            Width           =   495
         End
         Begin VB.CheckBox Net 
            Caption         =   "Check5"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   1
            Top             =   720
            Width           =   195
         End
         Begin VB.CheckBox Net 
            Caption         =   "Check5"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   12
            Top             =   1140
            Width           =   195
         End
         Begin VB.CheckBox Net 
            Caption         =   "Check5"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   23
            Top             =   1560
            Width           =   195
         End
         Begin VB.CheckBox Net 
            Caption         =   "Check5"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   34
            Top             =   1980
            Width           =   195
         End
         Begin VB.CheckBox Net 
            Caption         =   "Check5"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   4
            Left            =   120
            TabIndex        =   45
            Top             =   2400
            Width           =   195
         End
         Begin VB.CheckBox Net 
            Caption         =   "Check5"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   5
            Left            =   120
            TabIndex        =   56
            Top             =   2820
            Width           =   195
         End
         Begin VB.CheckBox Net 
            Caption         =   "Check5"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   6
            Left            =   120
            TabIndex        =   67
            Top             =   3240
            Width           =   195
         End
         Begin VB.CheckBox Net 
            Caption         =   "Check5"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   7
            Left            =   120
            TabIndex        =   78
            Top             =   3660
            Width           =   195
         End
         Begin VB.CheckBox Net 
            Caption         =   "Check5"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   8
            Left            =   120
            TabIndex        =   89
            Top             =   4080
            Width           =   195
         End
         Begin VB.CheckBox Net 
            Caption         =   "Check5"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   9
            Left            =   120
            TabIndex        =   100
            Top             =   4500
            Width           =   195
         End
         Begin VB.CheckBox Net 
            Caption         =   "Check5"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   10
            Left            =   120
            TabIndex        =   111
            Top             =   4920
            Width           =   195
         End
         Begin VB.CheckBox Net 
            Caption         =   "Check5"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   11
            Left            =   120
            TabIndex        =   122
            Top             =   5340
            Width           =   195
         End
         Begin VB.CheckBox Net 
            Caption         =   "Check5"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   12
            Left            =   120
            TabIndex        =   133
            Top             =   5760
            Width           =   195
         End
         Begin VB.CheckBox Net 
            Caption         =   "Check5"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   13
            Left            =   120
            TabIndex        =   144
            Top             =   6180
            Width           =   195
         End
         Begin VB.CheckBox Net 
            Caption         =   "Check5"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   14
            Left            =   120
            TabIndex        =   155
            Top             =   6600
            Width           =   195
         End
         Begin VB.CheckBox Net 
            Caption         =   "Check5"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   15
            Left            =   120
            TabIndex        =   166
            Top             =   7020
            Width           =   195
         End
         Begin VB.TextBox Inp 
            BackColor       =   &H00FFFF00&
            Height          =   285
            Index           =   0
            Left            =   420
            TabIndex        =   2
            Top             =   660
            Width           =   735
         End
         Begin VB.TextBox Inp 
            BackColor       =   &H00FFFF00&
            Height          =   285
            Index           =   1
            Left            =   420
            TabIndex        =   13
            Top             =   1080
            Width           =   735
         End
         Begin VB.TextBox Inp 
            BackColor       =   &H00FFFF00&
            Height          =   285
            Index           =   2
            Left            =   420
            TabIndex        =   24
            Top             =   1500
            Width           =   735
         End
         Begin VB.TextBox Inp 
            BackColor       =   &H00FFFF00&
            Height          =   285
            Index           =   3
            Left            =   420
            TabIndex        =   35
            Top             =   1920
            Width           =   735
         End
         Begin VB.TextBox Inp 
            BackColor       =   &H00FFFF00&
            Height          =   285
            Index           =   4
            Left            =   420
            TabIndex        =   46
            Top             =   2340
            Width           =   735
         End
         Begin VB.TextBox Inp 
            BackColor       =   &H00FFFF00&
            Height          =   285
            Index           =   5
            Left            =   420
            TabIndex        =   57
            Top             =   2760
            Width           =   735
         End
         Begin VB.TextBox Inp 
            BackColor       =   &H00FFFF00&
            Height          =   285
            Index           =   6
            Left            =   420
            TabIndex        =   68
            Top             =   3180
            Width           =   735
         End
         Begin VB.TextBox Inp 
            BackColor       =   &H00FFFF00&
            Height          =   285
            Index           =   7
            Left            =   420
            TabIndex        =   79
            Top             =   3600
            Width           =   735
         End
         Begin VB.TextBox Inp 
            BackColor       =   &H00FFFF00&
            Height          =   285
            Index           =   8
            Left            =   420
            TabIndex        =   90
            Top             =   4020
            Width           =   735
         End
         Begin VB.TextBox Inp 
            BackColor       =   &H00FFFF00&
            Height          =   285
            Index           =   9
            Left            =   420
            TabIndex        =   101
            Top             =   4440
            Width           =   735
         End
         Begin VB.TextBox Inp 
            BackColor       =   &H00FFFF00&
            Height          =   285
            Index           =   10
            Left            =   420
            TabIndex        =   112
            Top             =   4860
            Width           =   735
         End
         Begin VB.TextBox Inp 
            BackColor       =   &H00FFFF00&
            Height          =   285
            Index           =   11
            Left            =   420
            TabIndex        =   123
            Top             =   5280
            Width           =   735
         End
         Begin VB.TextBox Inp 
            BackColor       =   &H00FFFF00&
            Height          =   285
            Index           =   12
            Left            =   420
            TabIndex        =   134
            Top             =   5700
            Width           =   735
         End
         Begin VB.TextBox Inp 
            BackColor       =   &H00FFFF00&
            Height          =   285
            Index           =   13
            Left            =   420
            TabIndex        =   145
            Top             =   6120
            Width           =   735
         End
         Begin VB.TextBox Inp 
            BackColor       =   &H00FFFF00&
            Height          =   285
            Index           =   14
            Left            =   420
            TabIndex        =   156
            Top             =   6540
            Width           =   735
         End
         Begin VB.TextBox Inp 
            BackColor       =   &H00FFFF00&
            Height          =   285
            Index           =   15
            Left            =   420
            TabIndex        =   167
            Top             =   6960
            Width           =   735
         End
         Begin VB.TextBox Outp 
            BackColor       =   &H0000FF00&
            Height          =   285
            Index           =   0
            Left            =   1380
            TabIndex        =   3
            Top             =   660
            Width           =   735
         End
         Begin VB.TextBox Outp 
            BackColor       =   &H0000FF00&
            Height          =   285
            Index           =   1
            Left            =   1380
            TabIndex        =   14
            Top             =   1080
            Width           =   735
         End
         Begin VB.TextBox Outp 
            BackColor       =   &H0000FF00&
            Height          =   285
            Index           =   2
            Left            =   1380
            TabIndex        =   25
            Top             =   1500
            Width           =   735
         End
         Begin VB.TextBox Outp 
            BackColor       =   &H0000FF00&
            Height          =   285
            Index           =   3
            Left            =   1380
            TabIndex        =   36
            Top             =   1920
            Width           =   735
         End
         Begin VB.TextBox Outp 
            BackColor       =   &H0000FF00&
            Height          =   285
            Index           =   4
            Left            =   1380
            TabIndex        =   47
            Top             =   2340
            Width           =   735
         End
         Begin VB.TextBox Outp 
            BackColor       =   &H0000FF00&
            Height          =   285
            Index           =   5
            Left            =   1380
            TabIndex        =   58
            Top             =   2760
            Width           =   735
         End
         Begin VB.TextBox Outp 
            BackColor       =   &H0000FF00&
            Height          =   285
            Index           =   6
            Left            =   1380
            TabIndex        =   69
            Top             =   3180
            Width           =   735
         End
         Begin VB.TextBox Outp 
            BackColor       =   &H0000FF00&
            Height          =   285
            Index           =   7
            Left            =   1380
            TabIndex        =   80
            Top             =   3600
            Width           =   735
         End
         Begin VB.TextBox Outp 
            BackColor       =   &H0000FF00&
            Height          =   285
            Index           =   8
            Left            =   1380
            TabIndex        =   91
            Top             =   4020
            Width           =   735
         End
         Begin VB.TextBox Outp 
            BackColor       =   &H0000FF00&
            Height          =   285
            Index           =   9
            Left            =   1380
            TabIndex        =   102
            Top             =   4440
            Width           =   735
         End
         Begin VB.TextBox Outp 
            BackColor       =   &H0000FF00&
            Height          =   285
            Index           =   10
            Left            =   1380
            TabIndex        =   113
            Top             =   4860
            Width           =   735
         End
         Begin VB.TextBox Outp 
            BackColor       =   &H0000FF00&
            Height          =   285
            Index           =   11
            Left            =   1380
            TabIndex        =   124
            Top             =   5280
            Width           =   735
         End
         Begin VB.TextBox Outp 
            BackColor       =   &H0000FF00&
            Height          =   285
            Index           =   12
            Left            =   1380
            TabIndex        =   135
            Top             =   5700
            Width           =   735
         End
         Begin VB.TextBox Outp 
            BackColor       =   &H0000FF00&
            Height          =   285
            Index           =   13
            Left            =   1380
            TabIndex        =   146
            Top             =   6120
            Width           =   735
         End
         Begin VB.TextBox Outp 
            BackColor       =   &H0000FF00&
            Height          =   285
            Index           =   14
            Left            =   1380
            TabIndex        =   157
            Top             =   6540
            Width           =   735
         End
         Begin VB.TextBox Outp 
            BackColor       =   &H0000FF00&
            Height          =   285
            Index           =   15
            Left            =   1380
            TabIndex        =   168
            Top             =   6960
            Width           =   735
         End
         Begin VB.TextBox Out3Pw 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   13
            Left            =   6660
            TabIndex        =   154
            Top             =   6120
            Width           =   495
         End
         Begin VB.TextBox Out3St 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   13
            Left            =   6120
            TabIndex        =   153
            Top             =   6120
            Width           =   495
         End
         Begin VB.TextBox Out0St 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   12
            Left            =   2340
            TabIndex        =   136
            Top             =   5700
            Width           =   495
         End
         Begin VB.TextBox Out0Pw 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   12
            Left            =   2880
            TabIndex        =   137
            Top             =   5700
            Width           =   495
         End
         Begin VB.TextBox Out1Pw 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   12
            Left            =   4140
            TabIndex        =   139
            Top             =   5700
            Width           =   495
         End
         Begin VB.TextBox Out1St 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   12
            Left            =   3600
            TabIndex        =   138
            Top             =   5700
            Width           =   495
         End
         Begin VB.TextBox Out2Pw 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   12
            Left            =   5400
            TabIndex        =   141
            Top             =   5700
            Width           =   495
         End
         Begin VB.TextBox Out2St 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   12
            Left            =   4860
            TabIndex        =   140
            Top             =   5700
            Width           =   495
         End
         Begin VB.TextBox Out3Pw 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   12
            Left            =   6660
            TabIndex        =   143
            Top             =   5700
            Width           =   495
         End
         Begin VB.TextBox Out3St 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   12
            Left            =   6120
            TabIndex        =   142
            Top             =   5700
            Width           =   495
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "0,1,2,3"
            Height          =   195
            Left            =   1440
            TabIndex        =   211
            Top             =   300
            Width           =   615
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "0,1,2,3"
            Height          =   195
            Left            =   480
            TabIndex        =   210
            Top             =   300
            Width           =   615
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Out-0"
            Height          =   195
            Left            =   2640
            TabIndex        =   209
            Top             =   300
            Width           =   480
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Out-1"
            Height          =   195
            Left            =   3900
            TabIndex        =   208
            Top             =   300
            Width           =   480
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Out-2"
            Height          =   195
            Left            =   5160
            TabIndex        =   207
            Top             =   300
            Width           =   480
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Out-3"
            Height          =   195
            Left            =   6420
            TabIndex        =   206
            Top             =   300
            Width           =   480
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            Caption         =   "A"
            Height          =   195
            Left            =   180
            TabIndex        =   205
            Top             =   300
            Width           =   135
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "PLCSim Kartý"
         Height          =   2055
         Left            =   -65640
         TabIndex        =   201
         Top             =   5760
         Width           =   2295
         Begin VB.CommandButton Command2 
            Height          =   1635
            Left            =   120
            Picture         =   "Form1.frx":3804
            Style           =   1  'Graphical
            TabIndex        =   178
            Top             =   300
            Width           =   2055
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Giriþler"
         Height          =   1635
         Left            =   -65640
         TabIndex        =   196
         Top             =   420
         Width           =   1095
         Begin VB.Image pInDisp 
            Height          =   225
            Index           =   0
            Left            =   720
            Picture         =   "Form1.frx":3C46
            Top             =   360
            Width           =   225
         End
         Begin VB.Image pInDisp 
            Height          =   225
            Index           =   1
            Left            =   720
            Picture         =   "Form1.frx":3D40
            Top             =   660
            Width           =   225
         End
         Begin VB.Image pInDisp 
            Height          =   225
            Index           =   2
            Left            =   720
            Picture         =   "Form1.frx":3E3A
            Top             =   960
            Width           =   225
         End
         Begin VB.Image pInDisp 
            Height          =   225
            Index           =   3
            Left            =   720
            Picture         =   "Form1.frx":3F34
            Top             =   1260
            Width           =   225
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            Caption         =   "In-0"
            Height          =   195
            Left            =   120
            TabIndex        =   200
            Top             =   360
            Width           =   345
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            Caption         =   "In-1"
            Height          =   195
            Left            =   120
            TabIndex        =   199
            Top             =   660
            Width           =   345
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            Caption         =   "In-2"
            Height          =   195
            Left            =   120
            TabIndex        =   198
            Top             =   960
            Width           =   345
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "In-3"
            Height          =   195
            Left            =   120
            TabIndex        =   197
            Top             =   1260
            Width           =   345
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Çýkýþlar"
         Height          =   1635
         Left            =   -64440
         TabIndex        =   191
         Top             =   420
         Width           =   1095
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Out-3"
            Height          =   195
            Left            =   120
            TabIndex        =   195
            Top             =   1260
            Width           =   480
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "Out-2"
            Height          =   195
            Left            =   120
            TabIndex        =   194
            Top             =   960
            Width           =   480
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "Out-1"
            Height          =   195
            Left            =   120
            TabIndex        =   193
            Top             =   660
            Width           =   480
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Out-0"
            Height          =   195
            Left            =   120
            TabIndex        =   192
            Top             =   360
            Width           =   480
         End
         Begin VB.Image pOutDisp 
            Height          =   225
            Index           =   3
            Left            =   720
            Picture         =   "Form1.frx":402E
            Top             =   1260
            Width           =   225
         End
         Begin VB.Image pOutDisp 
            Height          =   225
            Index           =   2
            Left            =   720
            Picture         =   "Form1.frx":4128
            Top             =   960
            Width           =   225
         End
         Begin VB.Image pOutDisp 
            Height          =   225
            Index           =   1
            Left            =   720
            Picture         =   "Form1.frx":4222
            Top             =   660
            Width           =   225
         End
         Begin VB.Image pOutDisp 
            Height          =   225
            Index           =   0
            Left            =   720
            Picture         =   "Form1.frx":431C
            Top             =   360
            Width           =   225
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Manual"
         Height          =   1995
         Left            =   -65640
         TabIndex        =   190
         Top             =   2220
         Width           =   2295
         Begin VB.CheckBox OutKontrol 
            Caption         =   "Program (Çýkýþ)"
            Height          =   735
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   237
            Top             =   1140
            Width           =   2055
         End
         Begin VB.CheckBox Dahili 
            Caption         =   "Harici (Giriþ)"
            Height          =   735
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   236
            Top             =   300
            Width           =   2055
         End
      End
      Begin VB.Frame Frame2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7875
         Left            =   120
         TabIndex        =   184
         Top             =   420
         Width           =   11535
         Begin VB.PictureBox Picture3 
            AutoSize        =   -1  'True
            Height          =   3735
            Left            =   3660
            Picture         =   "Form1.frx":4416
            ScaleHeight     =   3675
            ScaleWidth      =   4275
            TabIndex        =   218
            Top             =   2100
            Width           =   4335
         End
         Begin VB.PictureBox Picture2 
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            Height          =   480
            Left            =   180
            Picture         =   "Form1.frx":37790
            ScaleHeight     =   480
            ScaleWidth      =   480
            TabIndex        =   203
            Top             =   300
            Width           =   480
         End
         Begin VB.PictureBox Picture1 
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            Height          =   480
            Left            =   10860
            Picture         =   "Form1.frx":37BD2
            ScaleHeight     =   480
            ScaleWidth      =   480
            TabIndex        =   202
            Top             =   300
            Width           =   480
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Dicle Üniversitesi - M.Y.O. Elektronik Bölümü"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   162
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   3000
            TabIndex        =   238
            Top             =   1500
            Width           =   5700
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Süleyman GÜNEL - 2001 "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   9480
            TabIndex        =   189
            Top             =   7500
            Width           =   1935
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Ramazan DAÞ          Murat ÖZTÜRK           Engin ÝLHAN"
            Height          =   195
            Left            =   3300
            TabIndex        =   188
            Top             =   6780
            Width           =   4890
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "- Proje Grubu -"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   162
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   5040
            TabIndex        =   187
            Top             =   6300
            Width           =   1515
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Süleyman GÜNEL"
            Height          =   195
            Left            =   5040
            TabIndex        =   186
            Top             =   7200
            Width           =   1515
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Entegre PLC Simulasyon Sistemi"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   162
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   4140
            TabIndex        =   185
            Top             =   5940
            Width           =   3390
         End
      End
      Begin VB.Label StatLabel 
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   375
         Left            =   -74880
         TabIndex        =   234
         Top             =   7980
         Width           =   11535
      End
   End
   Begin VB.Image PLCSimOff 
      Height          =   480
      Left            =   2640
      Picture         =   "Form1.frx":38014
      Top             =   8520
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image PLCSimOn 
      Height          =   480
      Left            =   2040
      Picture         =   "Form1.frx":38456
      Top             =   8460
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image OffLine 
      Height          =   480
      Left            =   1440
      Picture         =   "Form1.frx":38898
      Top             =   8460
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image OnLine 
      Height          =   480
      Left            =   960
      Picture         =   "Form1.frx":38CDA
      Top             =   8460
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image LedOn 
      Height          =   225
      Left            =   540
      Picture         =   "Form1.frx":3911C
      Top             =   8400
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image LedOff 
      Height          =   225
      Left            =   180
      Picture         =   "Form1.frx":39216
      Top             =   8400
      Visible         =   0   'False
      Width           =   225
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Pin2(3), pOut2(3) As Byte
Dim StatusOk, NetworkInitOk, StartCount, PortInitOk As Boolean
Sub PortFind()
Dim BiosAdr, GeciciPort, GeciciPort2 As Long
Dim durum As Boolean
Dim c1, c2, GeciciPort3 As Integer
Dim a1, a2 As String
PortInitOk = False
BiosAdr = &H400 + &H8
durum = GetPhysLong(BiosAdr, GeciciPort)
GeciciPort2 = GeciciPort And 65535: GeciciPort = GeciciPort2
If GeciciPort > 32767 Then GeciciPort3 = -32768 + (GeciciPort - 32768) Else GeciciPort3 = GeciciPort
c1 = GeciciPort3
BiosAdr = &H400 + &HA
durum = GetPhysLong(BiosAdr, GeciciPort)
GeciciPort2 = GeciciPort And 65535: GeciciPort = GeciciPort2
If GeciciPort > 32767 Then GeciciPort3 = -32768 + (GeciciPort - 32768) Else GeciciPort3 = GeciciPort
c2 = GeciciPort3
a1 = Left(Hex(c1), 4): a2 = Left(Hex(c2), 4)
a1 = " " & String(4 - Len(a1), "0") & a1 & "h"
a2 = " " & String(4 - Len(a2), "0") & a2 & "h"
Port0.Caption = a1: Port1.Caption = a2
PortNo0 = c1: PortNo1 = c2
If c1 = 0 And c2 = 0 Then MsgBox ("Herhangi bir Paralel Port bulunamadý")
If c1 = &H378 Then Option1.Value = True: PortNo = c1
If c2 = &H378 Then Option2.Value = True: PortNo = c2
If c2 <> 0 And c1 <> &H378 Then Option2.Value = True: PortNo = c2
If c1 <> 0 And c2 <> &H378 Then Option1.Value = True: PortNo = c1
If c1 = 0 And c2 = 0 Then Option3.Value = True
If c1 = 0 Then Option1.Enabled = False
If c2 = 0 Then Option2.Enabled = False
PortInitOk = True
End Sub
Sub PortOut(Komut As Integer)
 Dim Deger, OutByte As Long
 If Komut = -1 Then
  durum = InitializeWinIo(): If durum = 0 Then StatLabel.Caption = "Sistem Kapatma Hatasý !"
  durum = SetPortVal(PortNo, 0, 1): If durum = 0 Then StatLabel.Caption = "Sistem Gönderim Hatasý !"
  Exit Sub
 End If
 If Komut = 0 Then
  durum = InitializeWinIo(): If durum = 0 Then StatLabel.Caption = "Sistem Ayarlama Hatasý !"
  durum = GetPortVal(PortNo, Deger, 1): If durum = 0 Then StatLabel.Caption = "Sistem Okuma Hatasý !"
  OutByte = (Deger And 240) Or (pOut(0) + pOut(1) * 2 + pOut(2) * 4 + pOut(3) * 8)
  durum = InitializeWinIo(): If durum = 0 Then StatLabel.Caption = "Sistem Ayarlama Hatasý !"
  durum = SetPortVal(PortNo, OutByte, 1): If durum = 0 Then StatLabel.Caption = "Sistem Gönderim Hatasý !"
  Exit Sub
 End If
 If Komut = 20 Then
  durum = InitializeWinIo(): If durum = 0 Then StatLabel.Caption = "Sistem Ayarlama Hatasý !"
  durum = GetPortVal(PortNo, Deger, 1): If durum = 0 Then StatLabel.Caption = "Sistem Okuma Hatasý !"
  OutByte = Deger And 15
  durum = InitializeWinIo(): If durum = 0 Then StatLabel.Caption = "Sistem Ayarlama Hatasý !"
  durum = SetPortVal(PortNo, OutByte, 1): If durum = 0 Then StatLabel.Caption = "Sistem Gönderim Hatasý !"
  Exit Sub
 End If
 If Komut = 21 Then
  durum = InitializeWinIo(): If durum = 0 Then StatLabel.Caption = "Sistem Ayarlama Hatasý !"
  durum = GetPortVal(PortNo, Deger, 1): If durum = 0 Then StatLabel.Caption = "Sistem Okuma Hatasý !"
  OutByte = Deger Or 240
  durum = InitializeWinIo(): If durum = 0 Then StatLabel.Caption = "Sistem Ayarlama Hatasý !"
  durum = SetPortVal(PortNo, OutByte, 1): If durum = 0 Then StatLabel.Caption = "Sistem Gönderim Hatasý !"
  Exit Sub
 End If
 If Komut = 9 Then
  durum = InitializeWinIo(): If durum = 0 Then StatLabel.Caption = "Sistem Ayarlama Hatasý !"
  durum = SetPortVal(PortNo, 0, 1): If durum = 0 Then StatLabel.Caption = "Sistem Gönderim Hatasý !"
  Exit Sub
 End If
MsgBox ("Hata")
End Sub
Sub PortStat()
 Dim Deger As Long
 If Dahili.Value = Unchecked Then
  durum = InitializeWinIo(): If durum = 0 Then StatLabel.Caption = "Sistem Kapatma Hatasý !"
  durum = GetPortVal((PortNo + 1), Deger, 1): If durum = 0 Then StatLabel.Caption = "Sistem Gönderim Hatasý !"
  If (Deger And 64) = 0 Then pIn(0) = 1 Else pIn(0) = 0
  If (Deger And 128) = 128 Then pIn(1) = 1 Else pIn(1) = 0
  If (Deger And 32) = 0 Then pIn(2) = 1 Else pIn(2) = 0
  If (Deger And 16) = 0 Then pIn(3) = 1 Else pIn(3) = 0
 End If
End Sub

Sub Yasakla(stt As Boolean)
 Frame1.Enabled = stt
 Frame7.Enabled = stt
 Frame8.Enabled = stt
 Frame16.Enabled = stt
 Frame18.Enabled = stt
 Timer1.Enabled = Not (stt)
End Sub

Private Sub Check1_Click()
If Dahili.Value = Checked Then Dahili.Caption = "Dahili (Giriþ)" Else Dahili.Caption = "Harici  (Giriþ)"
End Sub

Private Sub Command1_Click()
PortNoUser = Val(Text4.Text)
If Option1.Value = True Then PortNo = PortNo0: Exit Sub
If Option2.Value = True Then PortNo = PortNo1: Exit Sub
If Option3.Value = True And PortNoUser <> 0 Then
 PortNo = PortNoUser
Else
 If PortNoUser = 0 Then
  If PortNo1 <> 0 And PortNo0 <> &H378 Then Option2.Value = True: PortNo = PortNo1
  If PortNo0 <> 0 And PortNo1 <> &H378 Then Option1.Value = True: PortNo = PortNo0
  MsgBox ("Port Deðeri Geçersiz !")
 End If
End If
End Sub

Private Sub Command10_Click()
Dim Dosya, Konum As String
Konum = Dir1.Path: If Right(Konum, 1) <> "\" Then Konum = Konum + "\"
Dosya = Konum + Trim(Text1.Text)
If Trim(Text1.Text) <> "" Then
 Open Dosya For Random As #1 Len = Len(FileNet)
 For i = 0 To 255
  FileNet.Active = Network(i).Active
  For j = 0 To 3
   FileNet.Inputs(j) = Network(i).Inputs(j)
   FileNet.Outputs4In(j) = Network(i).Outputs4In(j)
   FileNet.OutActive(j) = Network(i).OutActive(j)
   FileNet.OutOnDelay(j) = Network(i).OutOnDelay(j)
   FileNet.OutOffDelay(j) = Network(i).OutOffDelay(j)
   FileNet.OutRet(j) = Network(i).OutRet(j)
   FileNet.OutNC(j) = Network(i).OutNC(j)
  Next j
 Put #1, i + 1, FileNet
 Next
 Close #1
 File1.Refresh
 Label9.Caption = "Dosya baþarýyla kayýt edildi."
End If
End Sub

Private Sub Command11_Click()
secim = MsgBox("Programýn tamamýný silmek istediðinizden eminmisiniz?", vbCritical + vbYesNo, "Uyarý !")
If secim = vbNo Then Exit Sub
 NetworkInit
 NetworkDisplay
End Sub

Private Sub Command14_Click()
Dim Dosya, Konum As String
If Trim(Text1.Text) <> "" Then
 Konum = Dir1.Path: If Right(Konum, 1) <> "\" Then Konum = Konum + "\"
 Dosya = Konum + Trim(Text1.Text)
 cevap = MsgBox(UCase(Dosya) + " Dosyasýný silmek istediðinizden eminmisiniz?", vbYesNo, "Silem Onayý")
 If cevap = vbYes Then
  Kill (Dosya)
  File1.Refresh
  Text1.Text = "": Label3.Caption = ""
  Dir1.Refresh: File1.Refresh
 End If
End If
End Sub

Private Sub Command2_Click()
If Command2.Picture = PLCSimOff.Picture Then
 Command2.Picture = PLCSimOn.Picture
 Frame3.Enabled = True: Frame4.Enabled = True
 If Timer1.Enabled = False Then Frame1.Enabled = True:  Frame7.Enabled = True
Else
 Command2.Picture = PLCSimOff.Picture
 Frame1.Enabled = False:  Frame3.Enabled = False: Frame7.Enabled = False: Frame4.Enabled = False
End If
End Sub
Private Sub Command4_Click()
If Command4.Picture = OffLine.Picture Then
 Command4.Picture = OnLine.Picture
 For i = 0 To 255
  For j = 0 To 3
   If Network(i).Active = True And Network(i).OutActive(j) = True Then If Network(i).OutNC(j) = True Then pOut(j) = 1 Else pOut(j) = 0
   Netwrk2(i).OutChanged_1(j) = False
   Netwrk2(i).OutChanged_2(j) = False
   Netwrk2(i).OutControl(j) = True
   Netwrk2(i).StartCount = False
   Netwrk2(i).OutOnDelay(j) = 0
   Netwrk2(i).OutOnDelayOk(j) = False
   Netwrk2(i).OutOffDelay(j) = 0
   Netwrk2(i).OutOffDelayOk(j) = False
  Next j
 Next
 PortOut (21)
 PortStatusDisplay
 Yasakla (False)
 StatLabel.Caption = "Program Çalýþýyor..."
Else
 PortOut (20)
 Command4.Picture = OffLine.Picture
 Yasakla (True)
 StatLabel.Caption = "Program durduruldu..."
End If
End Sub

Private Sub Command5_Click()
If BlkNo > 0 Then BlkNo = BlkNo - 1: NetworkDisplay
BlokNo.Caption = " " & BlkNo
End Sub

Private Sub Command6_Click()
If BlkNo < 15 Then BlkNo = BlkNo + 1: NetworkDisplay
BlokNo.Caption = " " & BlkNo
End Sub

Private Sub Command7_Click()
j = -1
For i = BlkNo * 16 To BlkNo * 16 + 15
 j = j + 1
 Network(i).Active = Net(j).Value
 Network(i).Inputs(0) = Val(Left(Inp(j).Text, 1))
 Network(i).Inputs(1) = Val(Mid(Inp(j).Text, 3, 1))
 Network(i).Inputs(2) = Val(Mid(Inp(j).Text, 5, 1))
 Network(i).Inputs(3) = Val(Right(Inp(j).Text, 1))
 Network(i).Outputs4In(0) = Val(Left(Outp(j).Text, 1))
 Network(i).Outputs4In(1) = Val(Mid(Outp(j).Text, 3, 1))
 Network(i).Outputs4In(2) = Val(Mid(Outp(j).Text, 5, 1))
 Network(i).Outputs4In(3) = Val(Right(Outp(j).Text, 1))
 If Left(Right(Out0Pw(j), 2), 1) = "!" Then Network(i).OutActive(0) = False Else Network(i).OutActive(0) = True
 Network(i).OutOnDelay(0) = Abs(Val(Out0St(j)))
 Network(i).OutOffDelay(0) = Abs(Val(Out0Pw(j)))
 If Left(Right(Out0Pw(j), 2), 1) = "r" Then Network(i).OutRet(0) = True Else Network(i).OutRet(0) = False
 If Right(Out0Pw(j), 1) = "c" Then Network(i).OutNC(0) = True Else Network(i).OutNC(0) = False
 If Left(Right(Out1Pw(j), 2), 1) = "!" Then Network(i).OutActive(1) = False Else Network(i).OutActive(1) = True
 Network(i).OutOnDelay(1) = Abs(Val(Out1St(j)))
 Network(i).OutOffDelay(1) = Abs(Val(Out1Pw(j)))
 If Left(Right(Out1Pw(j), 2), 1) = "r" Then Network(i).OutRet(1) = True Else Network(i).OutRet(1) = False
 If Right(Out1Pw(j), 1) = "c" Then Network(i).OutNC(1) = True Else Network(i).OutNC(1) = False
 If Left(Right(Out2Pw(j), 2), 1) = "!" Then Network(i).OutActive(2) = False Else Network(i).OutActive(2) = True
 Network(i).OutOnDelay(2) = Abs(Val(Out2St(j)))
 Network(i).OutOffDelay(2) = Abs(Val(Out2Pw(j)))
 If Left(Right(Out2Pw(j), 2), 1) = "r" Then Network(i).OutRet(2) = True Else Network(i).OutRet(2) = False
 If Right(Out2Pw(j), 1) = "c" Then Network(i).OutNC(2) = True Else Network(i).OutNC(2) = False
 If Left(Right(Out3Pw(j), 2), 1) = "!" Then Network(i).OutActive(3) = False Else Network(i).OutActive(3) = True
 Network(i).OutOnDelay(3) = Abs(Val(Out3St(j)))
 Network(i).OutOffDelay(3) = Abs(Val(Out3Pw(j)))
 If Left(Right(Out3Pw(j), 2), 1) = "r" Then Network(i).OutRet(3) = True Else Network(i).OutRet(3) = False
 If Right(Out3Pw(j), 1) = "c" Then Network(i).OutNC(3) = True Else Network(i).OutNC(3) = False
Next
End Sub

Private Sub Command8_Click()
secim = MsgBox("Ekranda görünen tüm verileri sýfýrlamak istediðinizden eminmisiniz?", vbCritical + vbYesNo, "Uyarý !")
If secim = vbNo Then Exit Sub
For i = BlkNo * 16 To BlkNo * 16 + 15
 Network(i).Active = False
 Netwrk2(i).StartCount = False
 For j = 0 To 3
  Network(i).Inputs(j) = 0
  Network(i).Outputs4In(j) = 0
  Network(i).OutActive(j) = False
  Network(i).OutOnDelay(j) = 0
  Network(i).OutOffDelay(j) = 0
  Network(i).OutRet(j) = False
  Network(i).OutNC(j) = False
  
  Netwrk2(i).OutControl(j) = True
  Netwrk2(i).OutOnDelay(j) = 0
  Netwrk2(i).OutOnDelayOk(j) = False
  Netwrk2(i).OutOffDelay(j) = 0
  Netwrk2(i).OutOffDelayOk(j) = False
  Netwrk2(i).OutChanged_1(j) = False
  Netwrk2(i).OutChanged_2(j) = False
  Netwrk2(i).StatusOkChanged_1(j) = False
  Netwrk2(i).StatusOkChanged_2(j) = False
 Next j
Next
NetworkDisplay
End Sub


Private Sub Command9_Click()
Dim Dosya, Konum As String
Konum = Dir1.Path: If Right(Konum, 1) <> "\" Then Konum = Konum + "\"
Dosya = Konum + Trim(Text1.Text)
If Trim(Text1.Text) <> "" Then
 Open Dosya For Random As #1 Len = Len(FileNet)
 For i = 0 To 255
  Get #1, i + 1, FileNet
  Network(i).Active = FileNet.Active
  For j = 0 To 3
   Network(i).Inputs(j) = FileNet.Inputs(j)
   Network(i).Outputs4In(j) = FileNet.Outputs4In(j)
   Network(i).OutActive(j) = FileNet.OutActive(j)
   Network(i).OutOnDelay(j) = FileNet.OutOnDelay(j)
   Network(i).OutOffDelay(j) = FileNet.OutOffDelay(j)
   Network(i).OutRet(j) = FileNet.OutRet(j)
   Network(i).OutNC(j) = FileNet.OutNC(j)
  Next j
 Next
 Close #1
 NetworkDisplay
 Label9.Caption = "Dosya baþarýyla yüklendi."
End If
End Sub

Private Sub Dahili_Click()
If Dahili.Value = Checked Then Dahili.Caption = "Dahili (Giriþ)" Else Dahili.Caption = "Harici  (Giriþ)"
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
Label9.Caption = ""
End Sub

Private Sub Drive1_Change()
On Error GoTo Hata
 Dir1.Path = Drive1.Drive
 Label37.Caption = ""
 Label9.Caption = ""
 Exit Sub
Hata:
If Err.Number = 68 Then Drive1.Drive = "C:\"
Err.Clear
Resume Next
End Sub

Private Sub File1_Click()
Dim Dosya, Konum As String
Text1.Text = File1.filename
Konum = Dir1.Path: If Right(Konum, 1) <> "\" Then Konum = Konum + "\"
Dosya = Konum + Trim(Text1.Text)
If Trim(Text1.Text) <> "" Then Label3.Caption = FileLen(Dosya)
Label9.Caption = ""
End Sub

Private Sub File1_DblClick()
Command9_Click
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If Dahili.Value = Checked Then
 If KeyCode = vbKeyF1 Then pIn(0) = Abs(pIn(0) - 1)
 If KeyCode = vbKeyF2 Then pIn(1) = Abs(pIn(1) - 1)
 If KeyCode = vbKeyF3 Then pIn(2) = Abs(pIn(2) - 1)
 If KeyCode = vbKeyF4 Then pIn(3) = Abs(pIn(3) - 1)
End If

If OutKontrol.Value = Checked Then
 If KeyCode = vbKeyF5 Then pOut(0) = Abs(pOut(0) - 1)
 If KeyCode = vbKeyF6 Then pOut(1) = Abs(pOut(1) - 1)
 If KeyCode = vbKeyF7 Then pOut(2) = Abs(pOut(2) - 1)
 If KeyCode = vbKeyF8 Then pOut(3) = Abs(pOut(3) - 1)
End If
End Sub

Private Sub Form_Load()
PortFind
While PortInitOk = False: Wend
Me.KeyPreview = True
Me.Top = Int(Screen.Height / 2) - Int(Me.Height / 2): Me.Left = Int(Screen.Width / 2) - Int(Me.Width / 2)
Frame1.Enabled = False: Frame3.Enabled = False: Frame4.Enabled = False
Command2.Picture = PLCSimOff.Picture
Command4.Picture = OffLine.Picture
BlkNo = 0: BlokNo.Caption = " 0"
NetworkInit
While NetworkInitOk = False: Wend
PortOut (9)
NetworkDisplay
PortStatusDisplay
StatLabel.Caption = "Sistem Beklemede..."
RichTextBox1.LoadFile "plcsim.rtf"
End Sub

Sub NetworkInit()
NetworkInitOk = False
For i = 0 To 255
 Network(i).Active = False
 Netwrk2(i).StartCount = False
 For j = 0 To 3
  Network(i).Inputs(j) = 0
  Network(i).Outputs4In(j) = 0
  Network(i).OutActive(j) = False
  Network(i).OutOnDelay(j) = 0
  Network(i).OutOffDelay(j) = 0
  Network(i).OutRet(j) = False
  Network(i).OutNC(j) = False
  
  Netwrk2(i).OutControl(j) = True
  Netwrk2(i).OutOnDelay(j) = 0
  Netwrk2(i).OutOnDelayOk(j) = False
  Netwrk2(i).OutOffDelay(j) = 0
  Netwrk2(i).OutOffDelayOk(j) = False
  Netwrk2(i).OutChanged_1(j) = False
  Netwrk2(i).OutChanged_2(j) = False
  Netwrk2(i).StatusOkChanged_1(j) = False
  Netwrk2(i).StatusOkChanged_2(j) = False
 Next j
Next
NetworkInitOk = True
End Sub

Sub NetworkDisplay()
j = -1
For i = BlkNo * 16 To BlkNo * 16 + 15
 j = j + 1
 If Network(i).Active = True Then Net(j).Value = 1 Else Net(j).Value = 0
 Inp(j).Text = Network(i).Inputs(0) & "," & Network(i).Inputs(1) & "," & Network(i).Inputs(2) & "," & Network(i).Inputs(3)
 Outp(j).Text = Network(i).Outputs4In(0) & "," & Network(i).Outputs4In(1) & "," & Network(i).Outputs4In(2) & "," & Network(i).Outputs4In(3)
 Out0St(j).Text = Network(i).OutOnDelay(0)
 Out0Pw(j).Text = Network(i).OutOffDelay(0)
 If Network(i).OutRet(0) = False Then If Network(i).OutActive(0) = True Then Out0Pw(j) = Out0Pw(j) & "." Else Out0Pw(j) = Out0Pw(j) & "!"
 If Network(i).OutRet(0) = True Then Out0Pw(j) = Out0Pw(j) & "r"
 If Network(i).OutNC(0) = True Then Out0Pw(j) = Out0Pw(j) & "c" Else Out0Pw(j) = Out0Pw(j) & "o"
 Out1St(j).Text = Network(i).OutOnDelay(1)
 Out1Pw(j).Text = Network(i).OutOffDelay(1)
 If Network(i).OutRet(1) = False Then If Network(i).OutActive(1) = True Then Out1Pw(j) = Out1Pw(j) & "." Else Out1Pw(j) = Out1Pw(j) & "!"
 If Network(i).OutRet(1) = True Then Out1Pw(j) = Out1Pw(j) & "r"
 If Network(i).OutNC(1) = True Then Out1Pw(j) = Out1Pw(j) & "c" Else Out1Pw(j) = Out1Pw(j) & "o"
 Out2St(j).Text = Network(i).OutOnDelay(2)
 Out2Pw(j).Text = Network(i).OutOffDelay(2)
 If Network(i).OutRet(2) = False Then If Network(i).OutActive(2) = True Then Out2Pw(j) = Out2Pw(j) & "." Else Out2Pw(j) = Out2Pw(j) & "!"
 If Network(i).OutRet(2) = True Then Out2Pw(j) = Out2Pw(j) & "r"
 If Network(i).OutNC(2) = True Then Out2Pw(j) = Out2Pw(j) & "c" Else Out2Pw(j) = Out2Pw(j) & "o"
 Out3St(j).Text = Network(i).OutOnDelay(3)
 Out3Pw(j).Text = Network(i).OutOffDelay(3)
 If Network(i).OutRet(3) = False Then If Network(i).OutActive(3) = True Then Out3Pw(j) = Out3Pw(j) & "." Else Out3Pw(j) = Out3Pw(j) & "!"
 If Network(i).OutRet(3) = True Then Out3Pw(j) = Out3Pw(j) & "r"
 If Network(i).OutNC(3) = True Then Out3Pw(j) = Out3Pw(j) & "c" Else Out3Pw(j) = Out3Pw(j) & "o"
Next
End Sub
Sub PortStatusDisplay()
PortOut (0)
PortStat
For i = 0 To 3
 If pIn(i) = 1 Then pInDisp(i).Picture = LedOn.Picture Else pInDisp(i).Picture = LedOff.Picture
 If pOut(i) = 1 Then pOutDisp(i).Picture = LedOn.Picture Else pOutDisp(i).Picture = LedOff.Picture
Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Command4.Picture = OnLine.Picture Then
 StatLabel.Caption = "Program yürütülürken sistem programýný kapatamazsýnýz !"
 Cancel = -1: Exit Sub
End If
  cevap = MsgBox("Programdan Çýkmak Ýstediðinizden Eminmisiniz ?", vbQuestion + vbYesNo, "Programdan Çýkýþ")
If cevap = vbYes Then
 PortOut (-1)
 End
Else
 Cancel = -1
End If
End Sub

Private Sub OutKontrol_Click()
If OutKontrol.Value = Checked Then OutKontrol.Caption = "Kullanýcý (Çýkýþ)" Else OutKontrol.Caption = "Program  (Çýkýþ)"
End Sub

Private Sub SSTab1_KeyDown(KeyCode As Integer, Shift As Integer)
If Dahili.Value = Checked Then
 If KeyCode = vbKeyF1 Then pIn(0) = Abs(pIn(0) - 1)
 If KeyCode = vbKeyF2 Then pIn(1) = Abs(pIn(1) - 1)
 If KeyCode = vbKeyF3 Then pIn(2) = Abs(pIn(2) - 1)
 If KeyCode = vbKeyF4 Then pIn(3) = Abs(pIn(3) - 1)
End If

If OutKontrol.Value = Checked Then
 If KeyCode = vbKeyF5 Then pOut(0) = Abs(pOut(0) - 1)
 If KeyCode = vbKeyF6 Then pOut(1) = Abs(pOut(1) - 1)
 If KeyCode = vbKeyF7 Then pOut(2) = Abs(pOut(2) - 1)
 If KeyCode = vbKeyF8 Then pOut(3) = Abs(pOut(3) - 1)
End If
End Sub

Private Sub Timer1_Timer()

For i = 0 To 255
sayac = 0
If Network(i).Active = True Then
 For j = 0 To 3
  Pin2(j) = pIn(j): pOut2(j) = pOut(j)
  If Network(i).Inputs(j) = 2 Then Pin2(j) = Network(i).Inputs(j)
  If Network(i).Outputs4In(j) = 2 Then pOut2(j) = Network(i).Outputs4In(j)
  If Pin2(j) = Network(i).Inputs(j) Then sayac = sayac + 1
  If pOut2(j) = Network(i).Outputs4In(j) Then sayac = sayac + 1
 Next j
 If sayac = 8 Then StatusOk = True Else StatusOk = False
 If StatusOk = True And Netwrk2(i).StartCount = False Then Netwrk2(i).StartCount = True

 If StatusOk = True Then
  For j = 0 To 3
   If Network(i).OutActive(j) = True Then
    If Netwrk2(i).OutControl(j) = True Then
     If Network(i).OutRet(j) = False And Network(i).OutOnDelay(j) = 0 And Network(i).OutOffDelay(j) = 0 Then
      If Netwrk2(i).OutChanged_1(j) = False Then pOut(j) = Abs(pOut(j) - 1): Netwrk2(i).OutChanged_1(j) = True: Netwrk2(i).OutChanged_2(j) = False
     End If
     If Network(i).OutRet(j) = True And Network(i).OutOnDelay(j) = 0 And Network(i).OutOffDelay(j) = 0 Then
      pOut(j) = Abs(pOut(j) - 1): Netwrk2(i).OutControl(j) = False
     End If
    End If
   End If
  Next j
 End If
  
 If StatusOk = False Then
  For j = 0 To 3
   If Network(i).OutActive(j) = True Then
    If Netwrk2(i).OutControl(j) = True Then
     If Network(i).OutRet(j) = False And Network(i).OutOnDelay(j) = 0 And Network(i).OutOffDelay(j) = 0 Then
      If Netwrk2(i).OutChanged_2(j) = False And Netwrk2(i).OutChanged_1(j) = True Then pOut(j) = Abs(pOut(j) - 1): Netwrk2(i).OutChanged_2(j) = True: Netwrk2(i).OutChanged_1(j) = False
     End If
    End If
   End If
  Next j
 End If
 
 If Netwrk2(i).StartCount = True Then
  For j = 0 To 3
   If Network(i).OutActive(j) = True Then
    If Netwrk2(i).OutControl(j) = True Then
     If Network(i).OutOnDelay(j) > 0 And Netwrk2(i).OutOnDelayOk(j) = False Then Netwrk2(i).OutOnDelay(j) = Netwrk2(i).OutOnDelay(j) + 1
     If Network(i).OutOffDelay(j) > 0 And Netwrk2(i).OutOffDelayOk(j) = False And (Network(i).OutOnDelay(j) = 0 Or Netwrk2(i).OutOnDelayOk(j) = True) Then Netwrk2(i).OutOffDelay(j) = Netwrk2(i).OutOffDelay(j) + 1
     If Network(i).OutOnDelay(j) > 0 And Netwrk2(i).OutOnDelay(j) >= Network(i).OutOnDelay(j) And Netwrk2(i).OutOnDelayOk(j) = False Then
      Netwrk2(i).OutOnDelayOk(j) = True
      pOut(j) = Abs(pOut(j) - 1)
     End If
     If Network(i).OutOffDelay(j) > 0 And Netwrk2(i).OutOffDelay(j) >= Network(i).OutOffDelay(j) And Netwrk2(i).OutOnDelayOk(j) = True And Netwrk2(i).OutOffDelayOk(j) = False Then
      Netwrk2(i).OutOffDelayOk(j) = True
      pOut(j) = Abs(pOut(j) - 1)
     End If
     If Network(i).OutRet(j) = True And Network(i).OutOnDelay(j) > 0 And Network(i).OutOffDelay(j) > 0 And Netwrk2(i).OutOnDelayOk(j) = True And Netwrk2(i).OutOffDelayOk(j) = True Then
      Netwrk2(i).OutOnDelayOk(j) = False: Netwrk2(i).OutOffDelayOk(j) = False
      Netwrk2(i).OutOnDelay(j) = 0: Netwrk2(i).OutOffDelay(j) = 0
     End If
     If Network(i).OutRet(j) = True And Network(i).OutOnDelay(j) > 0 And Network(i).OutOffDelay(j) = 0 And Netwrk2(i).OutOnDelayOk(j) = True Then Netwrk2(i).OutControl(j) = False
     If Network(i).OutRet(j) = False And Netwrk2(i).OutOnDelayOk(j) = True And Network(i).OutOnDelay(j) > 0 And Network(i).OutOffDelay(j) = 0 Then
      If StatusOk = True Then Netwrk2(i).StatusOkChanged_1(j) = True Else Netwrk2(i).StatusOkChanged_2(j) = True
      If StatusOk = True And Netwrk2(i).StatusOkChanged_1(j) = True And Netwrk2(i).StatusOkChanged_2(j) = True Then
       Netwrk2(i).StatusOkChanged_1(j) = False
       Netwrk2(i).StatusOkChanged_2(j) = False
       Netwrk2(i).OutOnDelay(j) = 0
       Netwrk2(i).OutOnDelayOk(j) = False
       If Network(i).OutNC(j) = True Then pOut(j) = 1 Else pOut(j) = 0
      End If
     End If
     If Network(i).OutRet(j) = False And Netwrk2(i).OutOnDelayOk(j) = True And Network(i).OutOnDelay(j) > 0 And Netwrk2(i).OutOffDelayOk(j) = True And Network(i).OutOffDelay(j) > 0 Then
      If StatusOk = True Then Netwrk2(i).StatusOkChanged_1(j) = True Else Netwrk2(i).StatusOkChanged_2(j) = True
      If StatusOk = True And Netwrk2(i).StatusOkChanged_1(j) = True And Netwrk2(i).StatusOkChanged_2(j) = True Then
       Netwrk2(i).StatusOkChanged_1(j) = False
       Netwrk2(i).StatusOkChanged_2(j) = False
       Netwrk2(i).OutOnDelay(j) = 0
       Netwrk2(i).OutOnDelayOk(j) = False
       Netwrk2(i).OutOffDelay(j) = 0
       Netwrk2(i).OutOffDelayOk(j) = False
       If Network(i).OutNC(j) = True Then pOut(j) = 1 Else pOut(j) = 0
      End If
     End If
    End If
   End If
  Next j
 End If

End If
devam:
Next i
PortStatusDisplay
End Sub


VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Cube Photos by Ken Foster"
   ClientHeight    =   7380
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10560
   LinkTopic       =   "Form1"
   ScaleHeight     =   492
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   704
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton Option3 
      Caption         =   "Large"
      Height          =   225
      Left            =   8580
      TabIndex        =   36
      Top             =   6975
      Value           =   -1  'True
      Width           =   900
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Medium"
      Height          =   195
      Left            =   9375
      TabIndex        =   35
      Top             =   6765
      Width           =   870
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Small"
      Height          =   225
      Left            =   8580
      TabIndex        =   34
      Top             =   6750
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2085
      TabIndex        =   32
      Top             =   6555
      Width           =   2310
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Save as Bitmap"
      Height          =   270
      Left            =   195
      TabIndex        =   31
      Top             =   6555
      Width           =   1740
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Print"
      Height          =   345
      Left            =   7530
      TabIndex        =   30
      Top             =   6645
      Width           =   870
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Load 6"
      Height          =   315
      Left            =   9660
      TabIndex        =   28
      Top             =   6180
      Width           =   720
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Preview 6"
      Height          =   315
      Left            =   8745
      TabIndex        =   27
      Top             =   6180
      Width           =   840
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Load 5"
      Height          =   300
      Left            =   7920
      TabIndex        =   26
      Top             =   6195
      Width           =   720
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Preview 5"
      Height          =   300
      Left            =   7005
      TabIndex        =   25
      Top             =   6195
      Width           =   840
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Load 4"
      Height          =   300
      Left            =   6180
      TabIndex        =   24
      Top             =   6195
      Width           =   720
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Preview 4"
      Height          =   300
      Left            =   5250
      TabIndex        =   23
      Top             =   6195
      Width           =   840
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Load 3"
      Height          =   300
      Left            =   4440
      TabIndex        =   16
      Top             =   6195
      Width           =   720
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Preview 3"
      Height          =   300
      Left            =   3510
      TabIndex        =   15
      Top             =   6195
      Width           =   840
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Load 2"
      Height          =   300
      Left            =   2715
      TabIndex        =   14
      Top             =   6195
      Width           =   720
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Preview 2"
      Height          =   300
      Left            =   1770
      TabIndex        =   13
      Top             =   6195
      Width           =   840
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Load 1"
      Height          =   300
      Left            =   960
      TabIndex        =   12
      Top             =   6195
      Width           =   720
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Preview 1"
      Height          =   300
      Left            =   45
      TabIndex        =   11
      Top             =   6195
      Width           =   840
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Left            =   7185
      TabIndex        =   10
      Top             =   4590
      Width           =   3150
   End
   Begin VB.FileListBox File1 
      Height          =   2430
      Left            =   7155
      TabIndex        =   9
      Top             =   1980
      Width           =   3225
   End
   Begin VB.DirListBox Dir1 
      Height          =   1440
      Left            =   7155
      TabIndex        =   8
      Top             =   465
      Width           =   3225
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   7155
      TabIndex        =   7
      Top             =   105
      Width           =   3240
   End
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4755
      Left            =   60
      ScaleHeight     =   315
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   458
      TabIndex        =   0
      Top             =   120
      Width           =   6900
      Begin VB.PictureBox Picture6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1500
         Left            =   2175
         ScaleHeight     =   1470
         ScaleWidth      =   1470
         TabIndex        =   6
         Top             =   3105
         Width           =   1500
         Begin VB.Image Image6 
            Height          =   1500
            Left            =   -15
            Stretch         =   -1  'True
            Top             =   -15
            Width           =   2010
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "6"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   510
            Left            =   510
            TabIndex        =   22
            Top             =   495
            Width           =   480
         End
      End
      Begin VB.PictureBox Picture5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1500
         Left            =   2175
         ScaleHeight     =   1470
         ScaleWidth      =   1470
         TabIndex        =   5
         Top             =   135
         Width           =   1500
         Begin VB.Image Image5 
            Height          =   1500
            Left            =   -15
            Stretch         =   -1  'True
            Top             =   -15
            Width           =   2010
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "5"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   495
            TabIndex        =   21
            Top             =   450
            Width           =   450
         End
      End
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1500
         Left            =   5145
         ScaleHeight     =   1470
         ScaleWidth      =   1485
         TabIndex        =   4
         Top             =   1620
         Width           =   1515
         Begin VB.Image Image4 
            Height          =   1545
            Left            =   -15
            Stretch         =   -1  'True
            Top             =   -15
            Width           =   2010
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "4"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   510
            TabIndex        =   20
            Top             =   495
            Width           =   465
         End
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1500
         Left            =   3660
         ScaleHeight     =   1470
         ScaleWidth      =   1470
         TabIndex        =   3
         Top             =   1620
         Width           =   1500
         Begin VB.Image Image3 
            Height          =   1500
            Left            =   -15
            Stretch         =   -1  'True
            Top             =   -15
            Width           =   2010
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "3"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   540
            TabIndex        =   19
            Top             =   495
            Width           =   435
         End
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1500
         Left            =   2175
         ScaleHeight     =   1470
         ScaleWidth      =   1470
         TabIndex        =   2
         Top             =   1620
         Width           =   1500
         Begin VB.Image Image2 
            Height          =   1515
            Left            =   -15
            Stretch         =   -1  'True
            Top             =   -15
            Width           =   2010
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "2"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   480
            TabIndex        =   18
            Top             =   480
            Width           =   465
         End
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1500
         Left            =   690
         ScaleHeight     =   1470
         ScaleWidth      =   1470
         TabIndex        =   1
         Top             =   1620
         Width           =   1500
         Begin VB.Image Image1 
            Height          =   1500
            Left            =   -15
            Stretch         =   -1  'True
            Top             =   -15
            Width           =   2010
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   540
            TabIndex        =   17
            Top             =   480
            Width           =   270
         End
      End
      Begin VB.Line Line19 
         X1              =   429
         X2              =   357
         Y1              =   222
         Y2              =   222
      End
      Begin VB.Line Line18 
         X1              =   259
         X2              =   329
         Y1              =   222
         Y2              =   222
      End
      Begin VB.Line Line17 
         X1              =   357
         X2              =   343
         Y1              =   221
         Y2              =   207
      End
      Begin VB.Line Line16 
         X1              =   343
         X2              =   328
         Y1              =   207
         Y2              =   222
      End
      Begin VB.Line Line15 
         X1              =   443
         X2              =   429
         Y1              =   208
         Y2              =   222
      End
      Begin VB.Line Line14 
         X1              =   245
         X2              =   260
         Y1              =   208
         Y2              =   223
      End
      Begin VB.Line Line13 
         X1              =   359
         X2              =   430
         Y1              =   94
         Y2              =   94
      End
      Begin VB.Line Line12 
         X1              =   258
         X2              =   330
         Y1              =   95
         Y2              =   95
      End
      Begin VB.Line Line11 
         X1              =   358
         X2              =   342
         Y1              =   94
         Y2              =   110
      End
      Begin VB.Line Line10 
         X1              =   346
         X2              =   329
         Y1              =   111
         Y2              =   94
      End
      Begin VB.Line Line9 
         X1              =   443
         X2              =   429
         Y1              =   108
         Y2              =   94
      End
      Begin VB.Line Line8 
         X1              =   245
         X2              =   258
         Y1              =   107
         Y2              =   94
      End
      Begin VB.Line Line7 
         X1              =   58
         X2              =   132
         Y1              =   221
         Y2              =   221
      End
      Begin VB.Line Line6 
         X1              =   144
         X2              =   130
         Y1              =   208
         Y2              =   222
      End
      Begin VB.Line Line5 
         X1              =   34
         X2              =   34
         Y1              =   119
         Y2              =   198
      End
      Begin VB.Line Line4 
         X1              =   57
         X2              =   34
         Y1              =   220
         Y2              =   197
      End
      Begin VB.Line Line3 
         X1              =   57
         X2              =   133
         Y1              =   95
         Y2              =   95
      End
      Begin VB.Line Line2 
         X1              =   144
         X2              =   131
         Y1              =   107
         Y2              =   94
      End
      Begin VB.Line Line1 
         X1              =   57
         X2              =   33
         Y1              =   95
         Y2              =   119
      End
   End
   Begin VB.PictureBox Picture7 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4755
      Left            =   60
      ScaleHeight     =   315
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   458
      TabIndex        =   29
      Top             =   120
      Visible         =   0   'False
      Width           =   6900
   End
   Begin VB.Label Label11 
      Caption         =   "Saved bitmaps cannot be printed out in Cube Photo."
      Height          =   255
      Left            =   465
      TabIndex        =   40
      Top             =   6870
      Width           =   3900
   End
   Begin VB.Label Label10 
      Caption         =   "Option: Drag and drop filename to preview box, then drag and drop to cube."
      Height          =   285
      Left            =   90
      TabIndex        =   39
      Top             =   7110
      Width           =   5490
   End
   Begin VB.Label Label9 
      Height          =   255
      Left            =   6630
      TabIndex        =   38
      Top             =   7050
      Width           =   2205
   End
   Begin VB.Label Label8 
      Caption         =   "Size of cube to print."
      Height          =   225
      Left            =   8685
      TabIndex        =   37
      Top             =   6540
      Width           =   1560
   End
   Begin VB.Label Label7 
      Caption         =   "Filename only- - no extension needed - - will be saved as bitmap"
      Height          =   405
      Left            =   4485
      TabIndex        =   33
      Top             =   6525
      Width           =   2820
   End
   Begin VB.Image Image12 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      DragMode        =   1  'Automatic
      Height          =   1185
      Left            =   8715
      Stretch         =   -1  'True
      Top             =   4950
      Width           =   1665
   End
   Begin VB.Image Image11 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      DragMode        =   1  'Automatic
      Height          =   1185
      Left            =   6975
      Stretch         =   -1  'True
      Top             =   4965
      Width           =   1665
   End
   Begin VB.Image Image10 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      DragMode        =   1  'Automatic
      Height          =   1185
      Left            =   5235
      Stretch         =   -1  'True
      Top             =   4965
      Width           =   1665
   End
   Begin VB.Image Image9 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      DragMode        =   1  'Automatic
      Height          =   1185
      Left            =   3495
      Stretch         =   -1  'True
      Top             =   4965
      Width           =   1665
   End
   Begin VB.Image Image8 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      DragMode        =   1  'Automatic
      Height          =   1185
      Left            =   1770
      Stretch         =   -1  'True
      Top             =   4965
      Width           =   1665
   End
   Begin VB.Image Image7 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      DragMode        =   1  'Automatic
      Height          =   1185
      Left            =   30
      Stretch         =   -1  'True
      Top             =   4965
      Width           =   1665
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'*******************************************************************
'**                               Cube  Photos
'**                               Version 1.0.0
'**                               By Ken Foster
'**                                 June  2004
'**                     Freeware--- no copyrights claimed
'*******************************************************************

Option Explicit

Private OldX As Integer
Private OldY As Integer

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Private Sub Dir1_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
 ' Change pointer to no drop.
    If State = 0 Then Source.MousePointer = 12
    ' Use default mouse pointer.
    If State = 1 Then Source.MousePointer = 0
End Sub

Private Sub Drive1_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
' Change pointer to no drop.
    If State = 0 Then Source.MousePointer = 12
    ' Use default mouse pointer.
    If State = 1 Then Source.MousePointer = 0

End Sub

Private Sub File1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Form1.ScaleMode = 1
Dim DY  ' Declare variable.
    DY = TextHeight("A")    ' Get height of one line.
    Label9.Move File1.Left, File1.Top + Y - DY / 2, File1.Width, DY
    Label9.Drag ' Drag label outline.
    
End Sub

Private Sub File1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Form1.ScaleMode = 3
End Sub
Private Sub drive1_change()
   On Error Resume Next
   Dir1.Path = Drive1.Drive
End Sub

Private Sub dir1_change()
   File1.Path = Dir1.Path
End Sub

Private Sub file1_click()
   Dim readystring As String
   
   If Right$(Dir1.Path, 1) = "\" Then 'determine if an  "\" is at the end of the dir1 path
   readystring = Dir1.Path & File1.FileName
Else
   readystring = Dir1.Path & "\" & File1.FileName
End If
Text1.Text = readystring
End Sub
Private Sub Form_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
 ' Change pointer to no drop.
    If State = 0 Then Source.MousePointer = 12
    ' Use default mouse pointer.
    If State = 1 Then Source.MousePointer = 0
End Sub

Private Sub Form_Load()
   Dir1.Path = App.Path & "\ImageFolder"  ' set Directory to image folder
End Sub

Private Sub Command2_Click()  'preview 1
   Image7.Picture = LoadPicture(Text1.Text)
End Sub

Private Sub Command3_Click()  'load 1
   Image1.Picture = Image7.Picture
   Reset_Image Image1, Picture1
End Sub

Private Sub Command4_Click()  'preview 2
   Image8.Picture = LoadPicture(Text1.Text)
End Sub

Private Sub Command5_Click()  'load 2
   Image2.Picture = Image8.Picture
   Reset_Image Image2, Picture2
End Sub

Private Sub Command6_Click()  'preview 3
   Image9.Picture = LoadPicture(Text1.Text)
End Sub

Private Sub Command7_Click()  'load 3
   Image3.Picture = Image9.Picture
   Reset_Image Image3, Picture3
End Sub

Private Sub Command8_Click()  'preview 4
   Image10.Picture = LoadPicture(Text1.Text)
End Sub

Private Sub Command9_Click()  'load 4
   Image4.Picture = Image10.Picture
   Reset_Image Image4, Picture4
End Sub
Private Sub Command1_Click()  'preview 5
   Image11.Picture = LoadPicture(Text1.Text)
End Sub
Private Sub Command10_Click()  'load 5
   Image5.Picture = Image11.Picture
   Reset_Image Image5, Picture5
End Sub

Private Sub Command11_Click()  'preview 6
   Image12.Picture = LoadPicture(Text1.Text)
End Sub

Private Sub Command12_Click()  ' load 6
   Image6.Picture = Image12.Picture
   Reset_Image Image6, Picture6
   Label7.Visible = False
End Sub
Private Sub Command13_Click() 'Print
Dim Sz As Integer  'determines size of printed picture
Form1.ScaleMode = 3
If Image1.Picture = LoadPicture() Then  ' just assumes no pictures at all are loaded
    MsgBox "No pictures loaded"
    Exit Sub
 End If
 
BitBlt Picture7.hDC, 0, 0, Picture7.ScaleWidth, Picture7.ScaleHeight, picMain.hDC, 0, 0, vbSrcCopy
Picture7.Picture = Picture7.Image
' scales picture to printer
If Option1.Value = True Then Sz = 6
If Option2.Value = True Then Sz = 9
If Option3.Value = True Then Sz = 13

'Printer.Orientation = vbPRORPortrait   ' or 1
Printer.Orientation = vbPRORLandscape ' or 2
Printer.ScaleMode = 3
Printer.CurrentX = 0
Printer.CurrentY = 0
Printer.PaintPicture Picture7.Picture, 0, 0, Picture7.Width * Sz, Picture7.Height * Sz, , , , , vbSrcCopy
Printer.EndDoc
Form1.ScaleMode = 1
End Sub
Private Sub Command14_Click()  'save button
If Image1.Picture = LoadPicture() Then   ' assumes no pictures loaded
      MsgBox "Please load pictures"
      Exit Sub
 End If
 
 If Text2.Text = "" Then   ' no filename
      MsgBox "Enter a filename"
      Exit Sub
 End If
 
BitBlt Picture7.hDC, 0, 0, Picture7.Width, Picture7.Height, picMain.hDC, 0, 0, vbSrcCopy
SavePicture Picture7.Image, App.Path & "\ImageFolder\" & Text2.Text & ".bmp"
MsgBox "File saved to  " & App.Path & "\ImageFolder\" & Text2.Text & ".bmp"
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Me
End Sub

Private Sub Image1_DragDrop(Source As Control, X As Single, Y As Single)
If TypeOf Source Is Image Then Image1.Picture = Source.Picture
Reset_Image Image1, Picture1
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   OldX = X
   OldY = Y
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = 1 Then  ' if left mouse button then move picture left or right
   Image1.Left = Image1.Left + (X - OldX)
   Image1.Top = Image1.Top + (Y - OldY)
End If
End Sub

Private Sub Image2_DragDrop(Source As Control, X As Single, Y As Single)
If TypeOf Source Is Image Then Image2.Picture = Source.Picture
Reset_Image Image2, Picture2
End Sub

Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   OldX = X
   OldY = Y
End Sub

Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = 1 Then  ' if left mouse button then move picture left or right
   Image2.Left = Image2.Left + (X - OldX)
   Image2.Top = Image2.Top + (Y - OldY)
End If
End Sub

Private Sub Image3_DragDrop(Source As Control, X As Single, Y As Single)
If TypeOf Source Is Image Then Image3.Picture = Source.Picture
Reset_Image Image3, Picture3
End Sub

Private Sub Image3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   OldX = X
   OldY = Y
End Sub

Private Sub Image3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = 1 Then    ' if left mouse button then move picture left or right
   Image3.Left = Image3.Left + (X - OldX)
   Image3.Top = Image3.Top + (Y - OldY)
End If
End Sub

Private Sub Image4_DragDrop(Source As Control, X As Single, Y As Single)
If TypeOf Source Is Image Then Image4.Picture = Source.Picture
Reset_Image Image4, Picture4
End Sub

Private Sub Image4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   OldX = X
   OldY = Y
End Sub

Private Sub Image4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = 1 Then     ' if left mouse button then move picture left or right
   Image4.Left = Image4.Left + (X - OldX)
   Image4.Top = Image4.Top + (Y - OldY)
End If
End Sub

Private Sub Image5_DragDrop(Source As Control, X As Single, Y As Single)
If TypeOf Source Is Image Then Image5.Picture = Source.Picture
Reset_Image Image5, Picture5
End Sub

Private Sub Image5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   OldX = X
   OldY = Y
End Sub

Private Sub Image5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = 1 Then    ' if left mouse button then move picture left or right
   Image5.Left = Image5.Left + (X - OldX)
   Image5.Top = Image5.Top + (Y - OldY)
End If
End Sub

Private Sub Image6_DragDrop(Source As Control, X As Single, Y As Single)
If TypeOf Source Is Image Then Image6.Picture = Source.Picture
Reset_Image Image6, Picture6
End Sub

Private Sub Image6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   OldX = X
   OldY = Y
End Sub

Private Sub Image6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = 1 Then    ' if left mouse button then move picture left or right
   Image6.Left = Image6.Left + (X - OldX)
   Image6.Top = Image6.Top + (Y - OldY)

End If
End Sub

Private Sub Image7_DragDrop(Source As Control, X As Single, Y As Single)
On Error Resume Next
    Image7.Picture = LoadPicture(File1.Path + "\" + File1.FileName)

If Err Then MsgBox "The picture file can't be loaded."
End Sub

Private Sub Image8_DragDrop(Source As Control, X As Single, Y As Single)
On Error Resume Next
    Image8.Picture = LoadPicture(File1.Path + "\" + File1.FileName)

If Err Then MsgBox "The picture file can't be loaded."
End Sub


Private Sub Image9_DragDrop(Source As Control, X As Single, Y As Single)
On Error Resume Next
    Image9.Picture = LoadPicture(File1.Path + "\" + File1.FileName)

If Err Then MsgBox "The picture file can't be loaded."
End Sub
Private Sub Image10_DragDrop(Source As Control, X As Single, Y As Single)
On Error Resume Next
    Image10.Picture = LoadPicture(File1.Path + "\" + File1.FileName)

If Err Then MsgBox "The picture file can't be loaded."
End Sub

Private Sub Image11_DragDrop(Source As Control, X As Single, Y As Single)
On Error Resume Next
    Image11.Picture = LoadPicture(File1.Path + "\" + File1.FileName)

If Err Then MsgBox "The picture file can't be loaded."
End Sub

Private Sub Image12_DragDrop(Source As Control, X As Single, Y As Single)
On Error Resume Next
    Image12.Picture = LoadPicture(File1.Path + "\" + File1.FileName)

If Err Then MsgBox "The picture file can't be loaded."
End Sub
Private Sub Reset_Image(Image As Image, Picture As PictureBox)  ' Maintains aspect ratio of images
Picture.ScaleMode = 3
Image.Visible = False
If Image.Picture Then
   '~ this is used in case the image change ' s
   '~ if it's not used, the image control is
  '~ still the same size as the previous pic
        Image.Height = Image.Picture.Height
        Image.Width = Image.Picture.Width
       
                If Image.Picture.Height > Image.Picture.Width Then
                    '~ the Pic is taller than wide
                    Image.Height = Picture.Height
                    Image.Width = Image.Picture.Width     'Image.Width / (Image.Picture.Height / Image.Height)
                               If Image.Width > Picture.Width Then
                                           '~ If the PictureBox isn't square, the p ' ic still may be larger than it
                                            Image.Width = Picture.Width
                                             Image.Height = Image.Picture.Height / (Image.Picture.Width / Image.Width)
                                 End If
                      End If
              If Image.Picture.Width > Image.Picture.Height Then
                     '~ Image is wider than tall
                     Image.Width = Picture.Width
                      Image.Height = Image.Picture.Height    '   Image.Height / (Image.Picture.Width / Image.Width)
               If Image.Height > Picture.Height Then
                    Image.Height = Picture.Height
                     Image.Width = Image.Picture.Width / (Image.Picture.Height / Image.Height)
              End If
       End If
       '~ Center Image1 within Picture1
  Image.Left = (Picture.Width / 2) - (Image.Width / 2)
 Image.Top = (Picture.Height / 2) - (Image.Height / 2)
Image.Visible = True
End If
Picture.ScaleMode = 1
End Sub

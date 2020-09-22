VERSION 5.00
Begin VB.Form M0 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5070
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   3780
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   338
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   252
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.FileListBox File1 
      Height          =   195
      Left            =   3540
      TabIndex        =   1
      Top             =   4980
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Timer tmrHold 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   1320
      Top             =   2280
   End
   Begin VB.PictureBox picWhiteArrow 
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   720
      Picture         =   "frmBPrograms.frx":0000
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   5
      Top             =   4440
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picBlackArrow 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   420
      Picture         =   "frmBPrograms.frx":058A
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   4
      Top             =   4440
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picTemp 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   0
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   22
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.PictureBox picItem 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   0
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   4035
      TabIndex        =   2
      Top             =   0
      Width           =   4035
   End
   Begin VB.DirListBox Dir1 
      Height          =   315
      Left            =   2280
      TabIndex        =   0
      Top             =   4920
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "M0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim OldIndex As Integer
Dim Over() As Boolean
Public MenuIndex As Integer
Dim MaxLen As Integer



Private Sub Form_Load()
MenuIndex = -1
End Sub

Private Sub picItem_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
    If Index > Dir1.ListCount - 1 Then
        ShellExecute Me.hWnd, "open", Dir1.path & "\" & Right(picItem(Index).Tag, Len(picItem(Index).Tag) - 8), "", "", 1
        M0.HideMe
        HideStartMenu
        s_Playsound "select"
    End If
ElseIf Button = vbRightButton Then
    
End If
End Sub

Private Sub picItem_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Dir1.ListCount + File1.ListCount = 0 Then Exit Sub
If Over(Index) = False Then
    picItem(Index).BackColor = vbHighlight
    picItem(Index).Cls
    picItem(Index).ForeColor = vbHighlightText
    picItem(Index).Print picItem(Index).Tag
    If Index < Dir1.ListCount Then BitBlt picItem(Index).hdc, Me.ScaleWidth - 18, 1, 16, 16, picWhiteArrow.hdc, 0, 0, vbSrcCopy
    If Index > Dir1.ListCount - 1 Then
        DrawIcon Dir1.path & "\" & File1.List(Index - Dir1.ListCount), Index, False
        picTemp.ForeColor = vb3DHighlight
        picTemp.Line (0, 0)-(19, 0)
        picTemp.Line (0, 0)-(0, 19)
        picTemp.ForeColor = vbButtonShadow
        picTemp.Line (0, 19)-(19, 19)
        picTemp.Line (19, 0)-(19, 19)
        BitBlt picItem(Index).hdc, 0, 0, 21, 20, picTemp.hdc, 0, 0, vbSrcCopy
    Else
        DrawIcon Dir1.List(Index), Index, False
        picTemp.ForeColor = vb3DHighlight
        picTemp.Line (0, 0)-(19, 0)
        picTemp.Line (0, 0)-(0, 19)
        picTemp.ForeColor = vbButtonShadow
        picTemp.Line (0, 19)-(19, 19)
        picTemp.Line (19, 0)-(19, 19)
        BitBlt picItem(Index).hdc, 0, 0, 21, 20, picTemp.hdc, 0, 0, vbSrcCopy
    End If
    Over(Index) = True
    If Index <> OldIndex Then
        picItem(OldIndex).BackColor = vbButtonFace
        picItem(OldIndex).Cls
        picItem(OldIndex).ForeColor = vbButtonText
        picItem(OldIndex).Print picItem(OldIndex).Tag
        picTemp.Cls
        If OldIndex < Dir1.ListCount Then BitBlt picItem(OldIndex).hdc, Me.ScaleWidth - 18, 0, 18, 18, picBlackArrow.hdc, 0, 0, vbSrcCopy
        Over(OldIndex) = False
        If OldIndex > Dir1.ListCount - 1 Then
            DrawIcon Dir1.path & "\" & File1.List(OldIndex - Dir1.ListCount), OldIndex
        Else
            DrawIcon Dir1.List(OldIndex), OldIndex
        End If
    End If
    
    If Index < Dir1.ListCount Then
        Dim f As New M0
        f.Top = Me.Top + picItem(Index).Top * Screen.TwipsPerPixelX
        f.Left = Me.Left + Me.Width - 50
        f.GetMenu Dir1.path & "\" & Right(picItem(Index).Tag, Len(picItem(Index).Tag) - 8)
    End If
    If Index <> OldIndex Then
        If MenuIndex <> -1 And OldIndex < Dir1.ListCount Then
            Forms(MenuIndex).HideMe
            MenuIndex = -1
        End If
        OldIndex = Index
    End If
    If Index < Dir1.ListCount Then
        MenuIndex = Forms.Count - 1
        s_Playsound "open"
    Else
        s_Playsound "hover"
    End If
End If
End Sub

Public Sub GetMenu(path As String)
Dir1.path = path
File1.path = path
If File1.ListCount + Dir1.ListCount = 0 Then
    picItem(0).Print "[ Empty ]"
    MaxLen = 10
Else
    If Dir1.ListCount > 0 Then
        For i = 1 To Dir1.ListCount + File1.ListCount - 1
            Load picItem(i)
            picItem(i).Visible = True
            picItem(i).Top = 20 * i
        Next
        For i = 0 To Dir1.ListCount - 1
            DrawIcon Dir1.List(i), i
            picItem(i).Print "        " & ExtractFileName(Dir1.List(i))
            picItem(i).Tag = "        " & ExtractFileName(Dir1.List(i))
            If Len(picItem(i).Tag) > MaxLen Then MaxLen = Len(picItem(i).Tag)
        Next
        For i = 0 To File1.ListCount - 1
            picTemp.BackColor = vbButtonFace
            DrawIcon Dir1.path & "\" & File1.List(i), i + Dir1.ListCount
            picItem(i + Dir1.ListCount).Print "        " & Left(File1.List(i), Len(File1.List(i)) - 4)
            picItem(i + Dir1.ListCount).Tag = "        " & Left(File1.List(i), Len(File1.List(i)) - 4)
            If Len(picItem(i + Dir1.ListCount).Tag) > MaxLen Then MaxLen = Len(picItem(i + Dir1.ListCount).Tag)
        Next
    Else
        For i = 1 To File1.ListCount - 1
            Load picItem(i)
            picItem(i).Visible = True
            picItem(i).Top = 20 * i
        Next
        For i = 0 To File1.ListCount - 1
            DrawIcon Dir1.path & "\" & File1.List(i), i
            picItem(i).Print "        " & Left(File1.List(i), Len(File1.List(i)) - 4)
            picItem(i).Tag = "        " & Left(File1.List(i), Len(File1.List(i)) - 4)
            If Len(picItem(i).Tag) > MaxLen Then MaxLen = Len(picItem(i).Tag)
        Next
    End If
    ReDim Over(Dir1.ListCount + File1.ListCount - 1)
End If
Me.Width = MaxLen * 5 * Screen.TwipsPerPixelX + 350
Me.Height = picItem.Count * 20 * Screen.TwipsPerPixelY + 10
SetWindowPos Me.hWnd, -1, Me.Left / Screen.TwipsPerPixelX, Me.Top / Screen.TwipsPerPixelY, Me.ScaleWidth, Me.ScaleHeight + 10, SWP_NOREPOSITION
For i = 0 To Dir1.ListCount - 1
    BitBlt picItem(i).hdc, Me.ScaleWidth - 18, 1, 16, 16, picBlackArrow.hdc, 0, 0, vbSrcCopy
Next

For i = 0 To picItem.Count - 1
    picItem(i).Width = Me.ScaleWidth
Next
Me.Show
Me.Refresh
End Sub

Public Sub HideMe()
    If MenuIndex > Forms.Count Then Exit Sub
    If MenuIndex > -1 Then Forms(MenuIndex).HideMe
    Unload Me
End Sub

Sub DrawIcon(path, Index, Optional blt = True)
    hImgLarge& = SHGetFileInfo(path, 0&, shinfo, Len(shinfo), _
    BASIC_SHGFI_FLAGS Or SHGFI_SMALLICON)
    picTemp.Cls
    If blt Then
        ImageList_Draw hImgLarge&, shinfo.iIcon, picTemp.hdc, 2, 2, ILD_TRANSPARENT
        BitBlt picItem(Index).hdc, 0, 0, 20, 20, picTemp.hdc, 0, 0, vbSrcCopy
    Else
        ImageList_Draw hImgLarge&, shinfo.iIcon, picTemp.hdc, 2, 2, ILD_TRANSPARENT
    End If
End Sub



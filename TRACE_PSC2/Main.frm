VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   Caption         =   "PaperTrace"
   ClientHeight    =   6630
   ClientLeft      =   165
   ClientTop       =   465
   ClientWidth     =   11130
   ForeColor       =   &H00000000&
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   442
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   742
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdUndo 
      Caption         =   "Undo Last"
      Height          =   405
      Left            =   75
      TabIndex        =   17
      Top             =   5445
      Width           =   1350
   End
   Begin VB.OptionButton optTools 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Cirlipise"
      ForeColor       =   &H00FF0000&
      Height          =   300
      Index           =   4
      Left            =   105
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   1665
      Width           =   1260
   End
   Begin VB.OptionButton optTools 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Rectangle"
      ForeColor       =   &H00FF0000&
      Height          =   300
      Index           =   3
      Left            =   105
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   1350
      Width           =   1260
   End
   Begin VB.OptionButton optTools 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Large Rubber"
      ForeColor       =   &H00FF0000&
      Height          =   300
      Index           =   8
      Left            =   105
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2925
      Width           =   1260
   End
   Begin VB.OptionButton optTools 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Curvy-Line"
      ForeColor       =   &H00FF0000&
      Height          =   300
      Index           =   6
      Left            =   105
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2295
      Width           =   1260
   End
   Begin VB.HScrollBar HSAlpha 
      Height          =   195
      Left            =   105
      Max             =   9
      Min             =   1
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   4350
      Value           =   5
      Width           =   1260
   End
   Begin VB.OptionButton optTools 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Poly-Line"
      ForeColor       =   &H00FF0000&
      Height          =   300
      Index           =   5
      Left            =   105
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1980
      Width           =   1260
   End
   Begin VB.OptionButton optTools 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Line"
      ForeColor       =   &H00FF0000&
      Height          =   300
      Index           =   2
      Left            =   105
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1035
      Width           =   1260
   End
   Begin VB.OptionButton optTools 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Small Rubber"
      ForeColor       =   &H00FF0000&
      Height          =   300
      Index           =   7
      Left            =   105
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2610
      Width           =   1260
   End
   Begin VB.OptionButton optTools 
      BackColor       =   &H00C0FFFF&
      Caption         =   "FreeDraw"
      ForeColor       =   &H00FF0000&
      Height          =   300
      Index           =   1
      Left            =   105
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   720
      Width           =   1260
   End
   Begin VB.OptionButton optTools 
      BackColor       =   &H00C0FFFF&
      Caption         =   "None"
      ForeColor       =   &H00FF0000&
      Height          =   300
      Index           =   0
      Left            =   105
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   405
      Width           =   1260
   End
   Begin VB.CommandButton cmdPaper 
      Caption         =   "Toggle On/Off"
      Height          =   300
      Left            =   105
      TabIndex        =   2
      Top             =   3645
      Width           =   1260
   End
   Begin VB.PictureBox picPAPER 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      DrawMode        =   7  'Invert
      ForeColor       =   &H00FFFFFF&
      Height          =   2280
      Left            =   2670
      ScaleHeight     =   152
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   137
      TabIndex        =   1
      Top             =   930
      Width           =   2055
   End
   Begin VB.PictureBox picDRAW 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      DrawMode        =   7  'Invert
      ForeColor       =   &H00FFFFFF&
      Height          =   6000
      Left            =   1740
      ScaleHeight     =   400
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   600
      TabIndex        =   0
      Top             =   255
      Width           =   9000
      Begin VB.Shape shpSmallRubber 
         Height          =   105
         Left            =   75
         Top             =   45
         Width           =   105
      End
      Begin VB.Shape shpLargeRubber 
         Height          =   225
         Left            =   285
         Top             =   270
         Width           =   225
      End
   End
   Begin VB.Shape shpBorder 
      Height          =   315
      Left            =   4470
      Top             =   225
      Width           =   345
   End
   Begin VB.Label LabPaper 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Paper"
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   15
      TabIndex        =   14
      Top             =   3375
      Width           =   1440
   End
   Begin VB.Label LabTools 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tools"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   15
      TabIndex        =   11
      Top             =   150
      Width           =   1440
   End
   Begin VB.Label LabNote 
      BackColor       =   &H00FF8080&
      Caption         =   "Note: changing transparency loses Undo Last."
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   45
      TabIndex        =   10
      Top             =   4575
      Width           =   1365
   End
   Begin VB.Label LabAlpha 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Transparency"
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   15
      TabIndex        =   9
      Top             =   4065
      Width           =   1440
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileOps 
         Caption         =   "&Open picture file"
         Index           =   0
      End
      Begin VB.Menu mnuFileOps 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuFileOps 
         Caption         =   "&Save paper as 2 color BMP"
         Index           =   2
      End
      Begin VB.Menu mnuFileOps 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
   End
   Begin VB.Menu mnuFileInfo 
      Caption         =   "File Info"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' PaperTrace  by  Robert Rayment  Aug 2006

Option Explicit

' Update 10 Aug:  LoadPicture Function simplified a bit

' Update 14 Aug:  Removed XOR wipe out in finished drawing
'                 by saving draw coords.
'                 Corrected a FreeDraw error.
'
' Update 17 Aug   Varying transparency no longer loses drawing.


' ImageData()  Loaded image

' picDRAW      Tracing paper to draw on
' DrawData()

' picPAPER     White paper: copied drawing to save
' PaperData()

' For XP manifest
Private Declare Sub InitCommonControls Lib "comctl32" ()
Private Declare Function LoadLibrary Lib "Kernel32" Alias "LoadLibraryA" ( _
    ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "Kernel32" ( _
   ByVal hLibModule As Long) As Long
Private m_hMod As Long
''
' Drawing
Private aMouseDown As Boolean
Private aDrawDone As Boolean
Private xs As Single, ys As Single
Private xp As Single, yp As Single
Private PointCount As Long
Private xsto() As Integer, ysto() As Integer
Private PaperOffOn As Long

Private NumFree As Long

Private zAlpha As Single   ' Transparency

Private aPicLoaded As Boolean

' Files
Private PathSpec$, CurrPath$, FileSpec$
Private SavePath$, SaveSpec$

' Screen.Twips
Private STX As Long, STY As Long

Dim CommonDialog1 As OSDialog


Private Sub cmdUndo_Click()
   If Not aPicLoaded Then
      Exit Sub
   End If
   If aDrawDone Then
      DrawData() = saveDRAWData()
      PaperData() = savePAPERData()
      DISPLAY picDRAW, DrawData()
      DISPLAY picPAPER, PaperData()
      aMouseDown = False
      aDrawDone = False
      PaperOffOn = 1
      cmdPaper_Click
   End If
End Sub


'#### TOOLS & DRAWING ####
Private Sub optTools_Click(Index As Integer)
   Tool = Index
   shpLargeRubber.Visible = False
   If Tool = SmallRubber Then
      shpSmallRubber.Visible = True
   End If
   If Tool = LargeRubber Then
      shpLargeRubber.Visible = True
   End If
   LabTools = "Tool: " & optTools(Index).Caption
   shpSmallRubber.Left = -7
   shpSmallRubber.Top = -7
   shpLargeRubber.Left = -15
   shpLargeRubber.Top = -15
End Sub

Private Sub picDRAW_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim C As Long
   
   If Not aPicLoaded Then
      Exit Sub
   End If
   
   If Tool = 0 Then Exit Sub
   
   If Button = vbLeftButton Then
      
   ' SavePictures   ' Undo, Restore picture, 1 level of Undo
   ' For Poly & Curvy lines only save if PointCount=1
   ' picboxes or arrays ?
      If Tool = PolyLine Or Tool = Curvyline Then
         If PointCount = 1 Then
            saveDRAWData() = DrawData()
            savePAPERData() = PaperData()
         End If
      Else
            saveDRAWData() = DrawData()
            savePAPERData() = PaperData()
      End If
      aDrawDone = True
      aMouseDown = True
      xs = x
      ys = y
      xp = x
      yp = y
      Select Case Tool
      Case None
      Case FreeDraw
         xsto(PointCount) = x
         ysto(PointCount) = y
         PointCount = PointCount + 1
         If PointCount > UBound(xsto()) Then
            ReDim Preserve xsto(1 To PointCount + 10), ysto(1 To PointCount + 10)
         End If
         picDRAW.PSet (x, y)
      Case ALine
         picDRAW.Line (xs, ys)-(xp, yp)
      Case Rectangle
         picDRAW.Line (xs, ys)-(xp, yp), vbWhite, B
         picPAPER.Line (xs, ys)-(xp, yp), vbWhite, B
      Case Cirlipse
         ModPublics.EvalCirlipse xs, ys, xp, yp
         picDRAW.Circle (xs, ys), zrad, vbWhite, , , zratio
      Case PolyLine
         xsto(PointCount) = x
         ysto(PointCount) = y
         picDRAW.Line (xs, ys)-(xp, yp)
         PointCount = PointCount + 1
         If PointCount > UBound(xsto()) Then
            ReDim Preserve xsto(1 To PointCount + 10), ysto(1 To PointCount + 10)
         End If
         xsto(PointCount) = xp
         ysto(PointCount) = yp
      Case Curvyline
         xsto(PointCount) = x
         ysto(PointCount) = y
         picDRAW.Line (xs, ys)-(xp, yp)
         PointCount = PointCount + 1
         If PointCount > UBound(xsto()) Then
            ReDim Preserve xsto(1 To PointCount + 10), ysto(1 To PointCount + 10)
         End If
         xsto(PointCount) = xp
         ysto(PointCount) = yp
      
      Case SmallRubber
         ' Hide cursor
         Do
           C = ShowCursor(0)
         Loop Until C < 0
         'Sleep 5
         xs = x - shpSmallRubber.Width
         ys = y - shpSmallRubber.Height
         shpSmallRubber.Left = xs
         shpSmallRubber.Top = ys
         ' Reform DrawData() under rubber
         FADER zAlpha, CLng(xs), CLng(H - ys - shpSmallRubber.Height), _
            CLng(shpSmallRubber.Width), CLng(shpSmallRubber.Height)
         xsto(PointCount) = xs
         ysto(PointCount) = ys
         PointCount = PointCount + 1
         If PointCount > UBound(xsto()) Then
            ReDim Preserve xsto(1 To PointCount + 10), ysto(1 To PointCount + 10)
         End If
      Case LargeRubber
         ' Hide cursor
         Do
           C = ShowCursor(0)
         Loop Until C < 0
         'Sleep 5
         xs = x - shpLargeRubber.Width
         ys = y - shpLargeRubber.Height
         shpLargeRubber.Left = xs
         shpLargeRubber.Top = ys
         ' Reform DrawData() under rubber
         FADER zAlpha, CLng(xs), CLng(H - ys - shpLargeRubber.Height), _
            CLng(shpLargeRubber.Width), CLng(shpLargeRubber.Height)
         xsto(PointCount) = xs
         ysto(PointCount) = ys
         PointCount = PointCount + 1
         If PointCount > UBound(xsto()) Then
            ReDim Preserve xsto(1 To PointCount + 10), ysto(1 To PointCount + 10)
         End If
      End Select
   Else
      aMouseDown = False
   End If
End Sub

Private Sub picDRAW_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim k As Long
   If Not aPicLoaded Then
      Exit Sub
   End If
   
   If Tool = 0 Then Exit Sub
   
   If aMouseDown Then
      Select Case Tool
      Case None
      Case FreeDraw
         xsto(PointCount) = x
         ysto(PointCount) = y
         picDRAW.Line -(x, y)
         PointCount = PointCount + 1
         If PointCount > UBound(xsto()) Then
            ReDim Preserve xsto(1 To PointCount + 10), ysto(1 To PointCount + 10)
         End If
      Case ALine
         picDRAW.Line (xs, ys)-(xp, yp)
         xp = x
         yp = y
         picDRAW.Line (xs, ys)-(xp, yp)
      Case Rectangle
         picDRAW.Line (xs, ys)-(xp, yp), vbWhite, B
         xp = x
         yp = y
         picDRAW.Line (xs, ys)-(xp, yp), vbWhite, B
      Case Cirlipse
         picDRAW.Circle (xs, ys), zrad, vbWhite, , , zratio
         xp = x
         yp = y
         ModPublics.EvalCirlipse xs, ys, xp, yp
         picDRAW.Circle (xs, ys), zrad, vbWhite, , , zratio
      Case PolyLine
         For k = 1 To PointCount - 1
            picDRAW.Line (xsto(k), ysto(k))-(xsto(k + 1), ysto(k + 1)), vbWhite
         Next k
         xsto(PointCount) = x
         ysto(PointCount) = y
         For k = 1 To PointCount - 1
            picDRAW.Line (xsto(k), ysto(k))-(xsto(k + 1), ysto(k + 1)), vbWhite
         Next k
      Case Curvyline
         For k = 1 To PointCount - 1
            picDRAW.Line (xsto(k), ysto(k))-(xsto(k + 1), ysto(k + 1)), vbWhite
         Next k
         xsto(PointCount) = x
         ysto(PointCount) = y
         For k = 1 To PointCount - 1
            picDRAW.Line (xsto(k), ysto(k))-(xsto(k + 1), ysto(k + 1)), vbWhite
         Next k
      Case SmallRubber
         xs = x - shpSmallRubber.Width
         ys = y - shpSmallRubber.Height
         shpSmallRubber.Left = xs
         shpSmallRubber.Top = ys
         ' Reform DrawData() under rubber
         FADER zAlpha, CLng(xs), CLng(H - ys - shpSmallRubber.Height), _
            CLng(shpSmallRubber.Width), CLng(shpSmallRubber.Height)
         xsto(PointCount) = xs
         ysto(PointCount) = ys
         PointCount = PointCount + 1
         If PointCount > UBound(xsto()) Then
            ReDim Preserve xsto(1 To PointCount + 10), ysto(1 To PointCount + 10)
         End If
      Case LargeRubber
         xs = x - shpLargeRubber.Width
         ys = y - shpLargeRubber.Height
         shpLargeRubber.Left = xs
         shpLargeRubber.Top = ys
         ' Reform DrawData() under rubber
         FADER zAlpha, CLng(xs), CLng(H - ys - shpLargeRubber.Height), _
            CLng(shpLargeRubber.Width), CLng(shpLargeRubber.Height)
         xsto(PointCount) = xs
         ysto(PointCount) = ys
         PointCount = PointCount + 1
         If PointCount > UBound(xsto()) Then
            ReDim Preserve xsto(1 To PointCount + 10), ysto(1 To PointCount + 10)
         End If
      End Select
   Else
      If Tool = SmallRubber Then
         shpSmallRubber.Left = x - shpSmallRubber.Width
         shpSmallRubber.Top = y - shpSmallRubber.Height
      End If
      If Tool = LargeRubber Then
         shpLargeRubber.Left = x - shpLargeRubber.Width
         shpLargeRubber.Top = y - shpLargeRubber.Height
      End If
   End If
End Sub

Private Sub picDRAW_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim k As Long
   If Not aPicLoaded Then
      Exit Sub
   End If
   
   If Tool = 0 Then Exit Sub
   
   ' Showcursor
   Do
      k = ShowCursor(1)
   Loop Until k >= 0
   'Sleep 5
   
   Select Case Tool
   Case FreeDraw
      aMouseDown = False
      ' Clear old points/lines
      picDRAW.PSet (xsto(1), ysto(1))
      For k = 2 To PointCount - 1
         picDRAW.Line -(xsto(k), ysto(k)), vbWhite
      Next k
      picDRAW.DrawMode = vbCopyPen
      picDRAW.PSet (xsto(1), ysto(1)), 0
      picPAPER.PSet (xsto(1), ysto(1)), 0
         For k = 2 To PointCount - 1
            picDRAW.Line -(xsto(k), ysto(k)), 0
            picPAPER.Line -(xsto(k), ysto(k)), 0
         Next k
      picDRAW.DrawMode = vbXorPen
      PointCount = 1
      ReDim xsto(1 To 20), ysto(1 To 20)
      FillDataArrays x, y
   Case ALine
      aMouseDown = False
      picDRAW.Line (xs, ys)-(xp, yp)   ' Clear old line
      picDRAW.DrawMode = vbCopyPen
      picDRAW.Line (xs, ys)-(xp, yp), 0
      picPAPER.Line (xs, ys)-(xp, yp), 0
      picDRAW.DrawMode = vbXorPen
      FillDataArrays x, y
   Case Rectangle
      aMouseDown = False
      picDRAW.Line (xs, ys)-(xp, yp), vbWhite, B ' Clear old box
      picDRAW.DrawMode = vbCopyPen
      picDRAW.Line (xs, ys)-(xp, yp), 0, B
      picPAPER.Line (xs, ys)-(xp, yp), 0, B
      picDRAW.DrawMode = vbXorPen
      FillDataArrays x, y
   Case Cirlipse
      aMouseDown = False
      picDRAW.Circle (xs, ys), zrad, vbWhite, , , zratio ' Clear old Cirlipse
      ModPublics.EvalCirlipse xs, ys, xp, yp
      picDRAW.DrawMode = vbCopyPen
      picDRAW.Circle (xs, ys), zrad, 0, , , zratio
      picPAPER.Circle (xs, ys), zrad, 0, , , zratio
      picDRAW.DrawMode = vbXorPen
      FillDataArrays x, y
   Case PolyLine, Curvyline
      If Button = vbRightButton Then  ' Finish off
          aMouseDown = False
         ' Clear old lines
         For k = 1 To PointCount - 1
            picDRAW.Line (xsto(k), ysto(k))-(xsto(k + 1), ysto(k + 1)), vbWhite
         Next k
         '  & redraw poly-lines
         picDRAW.DrawMode = vbCopyPen
         If Tool = Curvyline Then
            ModPublics.GenerateCurvyPoints xsto(), ysto(), PointCount
         End If
         ' Draw poly or poly-curvy lines
         For k = 1 To PointCount - 1
            picDRAW.Line (xsto(k), ysto(k))-(xsto(k + 1), ysto(k + 1)), 0
            picPAPER.Line (xsto(k), ysto(k))-(xsto(k + 1), ysto(k + 1)), 0
         Next k
         picDRAW.DrawMode = vbXorPen
         PointCount = 1
         ReDim xsto(1 To 20), ysto(1 To 20)
         FillDataArrays x, y
      End If
   Case SmallRubber
      aMouseDown = False
      ' Rub out on paper
      For k = 1 To PointCount - 1
         ys = ysto(k)
         xs = xsto(k)
         picPAPER.Line (xsto(k), ysto(k))-(xsto(k) + shpSmallRubber.Width - 1, ysto(k) _
            + shpSmallRubber.Height - 1), vbWhite, BF ' or paper color
      Next k
      PointCount = 1
      ReDim xsto(1 To 20), ysto(1 To 20)
      FillDataArrays x, y
   Case LargeRubber
      aMouseDown = False
      ' Rub out on paper
      For k = 1 To PointCount - 1
         picPAPER.Line (xsto(k), ysto(k))-(xsto(k) + shpLargeRubber.Width - 1, ysto(k) _
            + shpLargeRubber.Height - 1), vbWhite, BF ' or paper color
      Next k
      PointCount = 1
      ReDim xsto(1 To 20), ysto(1 To 20)
      FillDataArrays x, y
   End Select
End Sub

Private Sub FillDataArrays(x As Single, y As Single)
   ' Confine x,y else GetDIBIts will fail
   If x < 0 Then x = 0
   If x >= W Then x = W - 1
   If y < 0 Then y = 0
   If y >= H Then y = H - 1
      
   ' Transfer pixel colors to DrawData() & PaperData()
   If GetDIBits(Form1.hdc, picDRAW.Image, 0, _
   H, DrawData(0, 0, 0), BHI, 0) = 0 Then
      MsgBox "DIB ERROR"
      Exit Sub
   End If
   If GetDIBits(Form1.hdc, picPAPER.Image, 0, _
      H, PaperData(0, 0, 0), BHI, 0) = 0 Then
      MsgBox "DIB ERROR"
      Exit Sub
   End If
End Sub
'#### END TOOLS & DRAWING ####


Private Sub cmdPaper_Click()
   If Not aPicLoaded Then
      Exit Sub
   End If
   PaperOffOn = 1 - PaperOffOn
   
   Select Case PaperOffOn
   Case 0   ' Hide paper
      picPAPER.Visible = False
   Case 1   ' Show paper
      picPAPER.Visible = True
      DISPLAY picPAPER, PaperData()
      picPAPER.ZOrder   ' picPAPER on top
   End Select
End Sub


' #### FORM STUFF ####
Private Sub Form_Initialize()
   m_hMod = LoadLibrary("shell32.dll")
   InitCommonControls
End Sub


Private Sub Form_Load()
' Public W As Long, H As Long     ' Image width & height
   
   aPicLoaded = False
   
   STX = Screen.TwipsPerPixelX
   STY = Screen.TwipsPerPixelY
   
   'Initial Form size
   Me.Width = 750 * STX    ' 11250 when STX=15
   Me.Height = 500 * STY   ' 7500
   
   PathSpec$ = App.Path
   If Right$(PathSpec$, 1) <> "\" Then PathSpec$ = PathSpec$ & "\"
   CurrPath$ = PathSpec$
   SavePath$ = PathSpec$
   
   zAlpha = 0.7  ' .5 - .9
   HSAlpha.Value = 7

   aMouseDown = False
   W = picDRAW.Width
   H = picDRAW.Height
   
   SizeClsPICS
   
   ' Ensure picbox properties
   picDRAW.Width = 600
   picDRAW.Height = 400
   
   picPAPER.Width = 600
   picPAPER.Height = 400
   
   picDRAW.Visible = True
   picPAPER.Visible = False '<<<<<<<<<<<<

   picDRAW.AutoRedraw = True
   picPAPER.AutoRedraw = True
   
   picDRAW.ScaleMode = vbPixels
   picPAPER.ScaleMode = vbPixels
   
   picDRAW.DrawMode = vbXorPen
   picPAPER.DrawMode = vbXorPen
   
   picDRAW.ForeColor = vbWhite
   picPAPER.ForeColor = vbWhite
   
   'picDRAW.BackColor = &HE0E0E0
   picPAPER.BackColor = vbWhite
   
   ' Line up
   picPAPER.Top = picDRAW.Top
   picPAPER.Left = picDRAW.Left    ' + 200  '<<<<<<
   
   picDRAW.DrawMode = vbXorPen
   picPAPER.DrawMode = vbCopyPen

   
   ' For Poly/CurvyLines
   PointCount = 1
   ReDim xsto(1 To 20), ysto(1 To 20)
   PaperOffOn = 0 ' ie OFF
   NumFree = 0
   aDrawDone = False
   
   shpSmallRubber.Left = -7
   shpSmallRubber.Top = -7
   shpSmallRubber.Visible = False
   
   shpLargeRubber.Left = -15
   shpLargeRubber.Top = -15
   shpLargeRubber.Visible = False
   
   Tool = 0
   LabTools = "Tool: None"
   LabNote.BackColor = RGB(100, 130, 220)
   Panel
End Sub

Private Sub Form_Resize()
   If WindowState <> vbMinimized Then
      Panel
   End If
End Sub

Private Sub mnuExit_Click()
   Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim C As Long
   Erase ImageData(), DrawData(), PaperData()
   Erase saveDRAWData(), savePAPERData()
   ' Showcursor
   Do
      C = ShowCursor(1)
   Loop Until C >= 0
   'Sleep 5
   FreeLibrary m_hMod
   Set Form1 = Nothing
   End
End Sub

Private Sub Panel()
Dim XX As Long
   XX = LabTools.Left + LabTools.Width
   Line (0, 0)-(XX + 1, Me.Height \ STY), RGB(100, 130, 220), BF
   Line (XX + 3, 0)-(XX + 3, Me.Height \ STY), 0
   Line (XX + 4, 0)-(XX + 4, Me.Height \ STY), vbWhite
   PicBorder
End Sub

Private Sub PicBorder()
   With shpBorder
      .Left = picDRAW.Left - 1
      .Top = picDRAW.Top - 1
      .Width = picDRAW.Width + 2
      .Height = picDRAW.Height + 2
   End With
End Sub

Private Sub SizeClsPICS()
   picDRAW.Cls
   picDRAW.Width = W
   picDRAW.Height = H
   picPAPER.Cls
   picPAPER.Width = W
   picPAPER.Height = H
End Sub
' #### FORM STUFF ####


'#### FILE STUFF ####
Private Sub mnuFileOps_Click(Index As Integer)
' Public W As Long, H As Long     ' Image width & height
' Public ImageData() As Byte
' Public DrawData() As Byte
Dim Title$, Filt$, InDir$
Dim FIndex As Long
Dim zAspect As Single
Dim WORG As Long, HORG As Long
Dim A$

   Select Case Index
   Case 0
      Title$ = "Load a picture file"
      Filt$ = "Pics bmp,jpg,gif|*.bmp;*.jpg;*.gif"
      FileSpec$ = ""
      InDir$ = CurrPath$ 'Pathspec$
      Set CommonDialog1 = New OSDialog
      CommonDialog1.ShowOpen FileSpec$, Title$, Filt$, InDir$, "", Me.hWnd, FIndex
      'FIndex = 1 bmp
      'FIndex = 2 jpg
      Set CommonDialog1 = Nothing
      If Len(FileSpec$) = 0 Then
         Screen.MousePointer = vbDefault
         Exit Sub
      End If
      CurrPath$ = GetPath(FileSpec$)
      If Not LoadThePicture(FileSpec$, ImageData()) Then
         Exit Sub
      End If
      
      If W <= 2 Or H <= 2 Then
         MsgBox "Picture size too small ie <= 2x2 ", vbCritical, "m-IP"
         W = 20
         H = 20
         ReDim ImageData(1, W - 1, H - 1)
         Exit Sub
      Else
         aMouseDown = False
         aDrawDone = False
         PaperOffOn = 1
         cmdPaper_Click
         optTools(0).Value = True
         WORG = W
         HORG = H
         zAspect = W / H
         If W > 600 Or H > 400 Then
            ' Resize to fit keeping aspect ratio
            ' NB Fullzise picDraw is 600x400, Aspect W/H = 1.5
            If W >= H Then ' zAspect >= 1
               If W > 600 And zAspect >= 1.5 Then
                     picDRAW.Width = 600
                     picDRAW.Height = 600 / zAspect
               Else ' W <= 600
                  picDRAW.Height = 400
                  picDRAW.Width = 400 * zAspect
               End If
            Else  ' W < H   ' zAspect < 1
               If H > 400 Then
                  picDRAW.Height = 400
                  picDRAW.Width = 400 * zAspect
               End If
            End If

            SetStretchBltMode picDRAW.hdc, HALFTONE
            Call StretchDIBits(picDRAW.hdc, _
            0, 0, _
            picDRAW.Width, picDRAW.Height, _
            0, 0, _
            W, H, _
            ImageData(0, 0, 0), _
            BHI, 0, vbSrcCopy)
            picDRAW.Refresh
                     
            W = picDRAW.Width
            H = picDRAW.Height
            
            ReDim ImageData(0 To 3, 0 To W - 1, 0 To H - 1)
            BHI.biWidth = W
            BHI.biHeight = H
            BHI.biBitCount = 32
            If GetDIBits(Form1.hdc, picDRAW.Image, 0, _
               H, ImageData(0, 0, 0), BHI, 0) = 0 Then
               MsgBox "DIB SHRINK ERROR"
               Exit Sub
            End If
         Else  ' Image smaller than 600x400
            picDRAW.Width = W
            picDRAW.Height = H
         End If
         ReDim DrawData(0 To 3, 0 To W - 1, 0 To H - 1)
         ReDim PaperData(0 To 3, 0 To W - 1, 0 To H - 1)
         FillMemory PaperData(0, 0, 0), 4 * W * H, 255
         SizeClsPICS
         
         ReDim saveDRAWData(0 To 3, 0 To W - 1, 0 To H - 1)
         ReDim savePAPERData(0 To 3, 0 To W - 1, 0 To H - 1)
         
      End If
      picDRAW.Visible = True
      ' Transfer faded image to picDRAW
      ' using ImageData() & DrawData()
      ' zAlpha 0.1 to 0.9
      FADER zAlpha, 0, 0, W, H
      aPicLoaded = True
      PicBorder
      A$ = " " & GetFileName(FileSpec$) & " "
      A$ = A$ & "WxH =" & Str$(WORG) & " x" & Str$(HORG) & " "
      A$ = A$ & Str$(FileLen(FileSpec$)) & "  B."
      A$ = A$ & " Size after loading =" & Str$(W) & " x" & Str$(H)
      mnuFileInfo.Caption = A$
   Case 1   ' Break
   Case 2   ' Save B/W BMP
      If Not aPicLoaded Then
         Exit Sub
      End If
      Title$ = "Save As 2 Color BMP"
      Filt$ = "Pics bmp|*.bmp"
      SaveSpec$ = ""
      InDir$ = SavePath$
      Set CommonDialog1 = New OSDialog
      CommonDialog1.ShowSave SaveSpec$, Title$, Filt$, InDir$, "", Me.hWnd, FIndex
      Set CommonDialog1 = Nothing
      
      If Len(SaveSpec$) = 0 Then
         Exit Sub
      End If
      FixExtension SaveSpec$, ".bmp"
      SavePath$ = GetPath(SaveSpec$)
      picPAPER.Picture = picPAPER.Image
      BHI.biWidth = W
      BHI.biHeight = H
      BHI.biBitCount = 32
      If GetDIBits(Form1.hdc, picPAPER.Image, 0, _
         H, PaperData(0, 0, 0), BHI, 0) = 0 Then
         MsgBox "DIB PAPER ERROR"
         Exit Sub
      End If
      If Not SaveBMP2(SaveSpec$, PaperData(), W, H) Then
         MsgBox "Saving Paper - failed", vbCritical, "PaperTrace"
      End If
   Case 3   ' Break
   End Select
End Sub

Private Function GetFileName(FSpec$) As String
' VB5 also
Dim L As Long
Dim k As Long
   GetFileName = ""
   L = Len(FSpec$)
   If L < 1 Then Exit Function
   For k = L To 1 Step -1
      If Mid$(FSpec$, k, 1) = "\" Then Exit For
   Next k
   If k = 0 Then
      GetFileName = FSpec$
   Else
      GetFileName = Right$(FSpec$, L - k)
   End If
End Function

Private Function GetPath(FSpec$) As String
' VB5 also
Dim L As Long
Dim k As Long
   GetPath = ""
   L = Len(FSpec$)
   If L < 1 Then Exit Function
   For k = L To 1 Step -1
      If Mid$(FSpec$, k, 1) = "\" Then Exit For
   Next k
   If k <> 0 Then
      GetPath = Left$(FSpec$, k)  ' NB includes last \
   End If
End Function

Private Sub FixExtension(FSpec$, Ext$)
' In: FileSpec$ & Ext$ (".xxx")
Dim p As Long
   If Len(FSpec$) = 0 Then Exit Sub
   Ext$ = LCase$(Ext$)
   p = InStr(1, FSpec$, ".")
   If p = 0 Then
      FSpec$ = FSpec$ & Ext$
   Else
      FSpec$ = Mid$(FSpec$, 1, p - 1) & Ext$
   End If
End Sub
' #### END FILE STUFF ####


' #### TRANSPARENCY ####
Private Sub HSAlpha_Change()
   Call HSAlpha_Scroll
End Sub

Private Sub HSAlpha_Scroll()
   zAlpha = HSAlpha.Value / 10
   LabAlpha = "Trans =" & Str$((1 - zAlpha) * 100) & " %"
   If aPicLoaded Then
      FADER zAlpha, 0, 0, W, H, True   ' Keeps drawing
      
      ' Commented out so that picture not erased
      'picPAPER.Picture = LoadPicture
      'ReDim PaperData(0 To 3, 0 To W - 1, 0 To H - 1)
      'FillMemory PaperData(0, 0, 0), 4 * W * H, 255
      
      PointCount = 1
      ' This loses the Undo Last
      saveDRAWData() = DrawData()
      savePAPERData() = PaperData()
      aDrawDone = False
      aMouseDown = False
      PaperOffOn = 1
      cmdPaper_Click
      optTools(0).Value = True
   End If
End Sub

Private Sub FADER(Alpha As Single, ixs As Long, iys As Long, wid As Long, hit As Long, Optional aKeep As Boolean = False)
' Public W As Long, H As Long     ' Image width & height
' Public ImageData() As Byte
' Public DrawData() As Byte

' aKeep = True keeps drawing, False loses drawing (default)

Dim ix As Long, iy As Long
Dim ixlo As Long, ixhi As Long
Dim iylo As Long, iyhi As Long

Dim B As Long, G As Long, R As Long
   iylo = iys: iyhi = iys + hit - 1
   ixlo = ixs: ixhi = ixs + wid - 1
   If iylo < 0 Then iylo = 0
   If ixlo < 0 Then ixlo = 0
   If iyhi > H - 1 Then iyhi = H - 1
   If ixhi > W - 1 Then ixhi = W - 1
   
   For iy = iylo To iyhi
   For ix = ixlo To ixhi
      B = ImageData(0, ix, iy)
      G = ImageData(1, ix, iy)
      R = ImageData(2, ix, iy)
      B = Alpha * (255 - B) + B
      G = Alpha * (255 - G) + G
      R = Alpha * (255 - R) + R
      DrawData(0, ix, iy) = B
      DrawData(1, ix, iy) = G
      DrawData(2, ix, iy) = R
      
      If aKeep Then
         ' Transfer any drawing on PaperData
         ' to Drawdata
         If PaperData(0, ix, iy) = 0 Then ' ie black
            DrawData(0, ix, iy) = 0
            DrawData(1, ix, iy) = 0
            DrawData(2, ix, iy) = 0
         End If
      End If
      
   Next ix
   Next iy
   DISPLAY picDRAW, DrawData()
End Sub
' #### END TRANSPARENCY ####

Private Sub DISPLAY(PB As PictureBox, DT() As Byte)
' Public W As Long, H As Long     ' Image width & height
' Public BHI As BITMAPINFOHEADER
   SetStretchBltMode PB.hdc, COLORONCOLOR
   Call StretchDIBits(PB.hdc, _
   0, 0, _
   W, H, _
   0, 0, _
   W, H, _
   DT(0, 0, 0), _
   BHI, 0, vbSrcCopy)
   
   PB.Refresh
End Sub

Private Sub mnuFileInfo_Click()
   MsgBox "Information about opened file", vbInformation, "Paper Trace by Robert Rayment"
End Sub

Private Sub mnuHelp_Click()
Dim A$
Dim C$
   C$ = vbCrLf
   
   A$ = "Paper Trace by Robert Rayment" & C$ & C$
   A$ = A$ & "Notes:" & C$
   A$ = A$ & "Pictures larger than 600x400 are scaled down." & C$ & C$
   A$ = A$ & "All drawing is done with MouseDown except" & C$
   A$ = A$ & "Poly-Lines & Curvy-Lines where LeftClick starts" & C$
   A$ = A$ & "a new segment and RightClick ends it." & C$
   A$ = A$ & "The final Curvy-Line is drawn at the RightClick." & C$ & C$
   A$ = A$ & "Tracings are saved on a white paper sheet" & C$
   A$ = A$ & "which can be saved as a 2 color (B/W) BMP." & C$ & C$
   A$ = A$ & "One level of Undo is catered for." & C$
   
   MsgBox A$, vbInformation, "Paper Trace"

End Sub


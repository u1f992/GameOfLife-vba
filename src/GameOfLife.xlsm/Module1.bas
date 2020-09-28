Attribute VB_Name = "Module1"
'@Folder("VBAProject")
Option Explicit
Public Const px As Long = 10
Public Const Height As Long = 100
Public Const Width As Long = 100
Public OnColor As Byte
Public Const OffColor As Byte = &H0

Sub SetCellsSizeSquare(ByVal Target As Range, ByVal px As Long)
    'Target.Cells(1, 1).Select
    Target.Clear
    
    Dim RatioHeight As Double: RatioHeight = 0.75
    Dim RatioWidth As Double: RatioWidth = 0.118
    
    Target.Cells(1, 1).RowHeight = px * RatioHeight
    Target.Cells(1, 1).ColumnWidth = px * RatioWidth
    
    'ある程度以下の小さな正方形を作ろうとするとRowHeightとRow / ColumnWidthとColumnの値に齟齬が生じるようになるため、キャリブレーション
    Do While Target.Cells(1, 1).Width > Target.Cells(1, 1).Height
        RatioWidth = RatioWidth - 0.001
    
        Target.Cells(1, 1).RowHeight = px * RatioHeight
        Target.Cells(1, 1).ColumnWidth = px * RatioWidth
        
        DoEvents
    Loop
    '補正した比率を選択範囲に適用
    Target.RowHeight = px * RatioHeight
    Target.ColumnWidth = px * RatioWidth
    
End Sub

Sub SetCellsSizeDefault(ByVal ws As Worksheet)
    ws.Cells(1, 1).Select
    ws.Cells.Clear
    ws.Cells.RowHeight = 18.75
    ws.Cells.ColumnWidth = 8.38
End Sub

Sub Default()
    SetCellsSizeDefault ActiveSheet
End Sub
Sub Square()
    UnlockActiveSheet
    SetCellsSizeSquare ActiveSheet.Cells, 10
End Sub

'ゲームを継続させるか判定する
'全てのセルを見て、生存セルが存在する場合にTrue
Private Function GameContinue(ByVal Target As Range) As Boolean
    Dim Cell As Range
    For Each Cell In Target
        If Cell.Interior.Color = HEX2RGB(OnColor) Then
            GameContinue = True
            Exit Function
        End If
    Next Cell
    GameContinue = False
End Function

'ライフゲーム
    '誕生 - 死んでいるセルに隣接する生きたセルがちょうど3つあれば､次の世代が誕生する｡
    '生存 - 生きているセルに隣接する生きたセルが2つか3つならば､次の世代でも生存する｡
    '過疎 - 生きているセルに隣接する生きたセルが1つ以下ならば､過疎により死滅する｡
    '過密 - 生きているセルに隣接する生きたセルが4つ以上ならば､過密により死滅する｡
Sub Main()

    Dim i As Long, j As Long
    Dim seed As Long: seed = 0: OnColor = GenerateRainbow(seed)
    Dim PrevOnColor As Byte: PrevOnColor = OnColor

    'fpsを安定させたい
    Dim t As New Timer
    Dim f As Double: f = 0.05
    
    'シートを初期化
    Initialize False
    Range(Cells(1, 1), Cells(Height, Width)).Interior.Color = HEX2RGB(OffColor)
    Stop
    Cells(1, 1).Select
    LockActiveSheet
    
    Dim Buffer() As Byte: ReDim Buffer(1 To Height, 1 To Width)
    Dim Previous() As Byte: ReDim Previous(1 To Height, 1 To Width)
    Dim GameSpace As Range: Set GameSpace = Range(Cells(LBound(Buffer, 1), LBound(Buffer, 1)), Cells(UBound(Buffer, 1), UBound(Buffer, 2)))
    
    Dim Cell As Range
    
    '初回Previousを用意
    For Each Cell In GameSpace
        If Cell.Interior.Color = HEX2RGB(OnColor) Then Previous(Cell.Row, Cell.Column) = OnColor Else Previous(Cell.Row, Cell.Column) = OffColor
    Next Cell
    
    Do While True 'GameContinue(GameSpace)
        'バッファを初期化
        ReDim Buffer(1 To Height, 1 To Width)
        
        t.StartTimer
        For i = LBound(Previous, 1) To UBound(Previous, 1)
            For j = LBound(Previous, 2) To UBound(Previous, 2)
                
                If Previous(i, j) = PrevOnColor Then '生きている場合
                    Select Case Vicinity(Previous, i, j, PrevOnColor)
                        Case 2, 3
                            Buffer(i, j) = OnColor
                        Case Else
                            Buffer(i, j) = OffColor
                    End Select
    
                ElseIf Previous(i, j) = OffColor Then '死んでいる場合
                    Select Case Vicinity(Previous, i, j, PrevOnColor)
                        Case 3
                            Buffer(i, j) = OnColor
                        Case Else
                            Buffer(i, j) = OffColor
                    End Select
                End If
        
            Next j
        Next i
        
        Do While t.TakeLap < f
            DoEvents
        Loop
        t.StopTimer
        
        'Debug.Print t.StopTimer
        '描画の更新
        
        UpdateScreen Buffer
        
        Previous = Buffer
        
        PrevOnColor = OnColor
        seed = seed + 1
        OnColor = GenerateRainbow(seed)
        
        DoEvents
    Loop
End Sub

'ランダムな初期値を与える場合はTrue
Public Sub Initialize(ByVal flag As Boolean)
    UnlockActiveSheet
    SetCellsSizeSquare ActiveSheet.Cells, px
    LockActiveSheet
    
    If Not flag Then GoTo Dispose
    
    Dim arr() As Byte
    ReDim arr(1 To Height, 1 To Width)
    Dim t As New Timer

    Dim i As Long, j As Long
    For i = 1 To Height
        For j = 1 To Width
            Randomize
            If Rnd > 0.9 Then arr(i, j) = OnColor Else arr(i, j) = OffColor
            'Debug.Print i & ", " & j
        Next j
    Next i
    
    UpdateScreen arr
    
    DoEvents
    
Dispose:
    UnlockActiveSheet
End Sub

Public Sub UpdateScreen(ByRef Buffer() As Byte)
    
    Application.ScreenUpdating = False
    UnlockActiveSheet
    
    Dim t As New Timer
    Dim f As Double: f = 0.3
    Dim dblWait As Double: dblWait = t.StartTimer
    
    Dim i As Long
    Dim str(0 To 255) As String
    
    Dim Cell As Range
    
    For Each Cell In Range(Cells(LBound(Buffer, 1), LBound(Buffer, 1)), Cells(UBound(Buffer, 1), UBound(Buffer, 2)))
        
        For i = 0 To 255
        
            'Onにするアドレス(String)を集める
            'Rangeは255文字までしか受け付けないので、255文字集まる毎にRange.Interior.Colorを変更
            'UnionによるRangeの結合は、ひとつずつ色を変えるより遅い
            If Buffer(Cell.Row, Cell.Column) = i Then
                If Len(str(i) & Cell.Address) <= 255 Then
                    str(i) = str(i) & "," & Cell.Address
                Else
                    'str(i)の先頭のコンマを外す
                    Range(Mid(str(i), 2)).Interior.Color = HEX2RGB(i)
                    str(i) = "," & Cell.Address
                End If
                
                Exit For
                
            End If
        Next i
        
    Next Cell
    
    '余りのRange.Interior.Colorを変更
    For i = 0 To 255
        If str(i) <> "" Then Range(Mid(str(i), 2)).Interior.Color = HEX2RGB(i)
        str(i) = ""
    Next i
    
    'fps安定化
    '描画範囲が広くなると期待した動作をしない
'    Do While t.TakeLap - dblWait < f
'        DoEvents
'    Loop
    
    LockActiveSheet
    Application.ScreenUpdating = True
    
End Sub

Public Sub LockActiveSheet()

    ActiveSheet.ScrollArea = ActiveSheet.UsedRange.Address
    ActiveSheet.Cells(1, 1).Select
    
    ActiveSheet.Cells.Locked = True
    ActiveSheet.Protect
    ActiveSheet.EnableSelection = xlUnlockedCells
End Sub

Public Sub UnlockActiveSheet()
    ActiveSheet.Unprotect
    ActiveSheet.Cells.Locked = False
    
    ActiveSheet.ScrollArea = ""
End Sub

Public Sub RCHidden()
'    Application.DisplayFullScreen = True
    ActiveWindow.DisplayGridlines = False
    Application.DisplayStatusBar = False
    ActiveWindow.DisplayWorkbookTabs = False
    ActiveWindow.DisplayHeadings = False
    Application.DisplayFormulaBar = False
    ActiveWindow.DisplayVerticalScrollBar = False
    ActiveWindow.DisplayHorizontalScrollBar = False
    If Application.CommandBars.GetPressedMso("MinimizeRibbon") = False Then Application.CommandBars.ExecuteMso "MinimizeRibbon"
'    Application.WindowState = xlMaximized
End Sub

Public Sub RCVisible()
    ActiveWindow.DisplayGridlines = True
    Application.DisplayStatusBar = True
    ActiveWindow.DisplayWorkbookTabs = True
    ActiveWindow.DisplayHeadings = True
    Application.DisplayFormulaBar = True
    ActiveWindow.DisplayVerticalScrollBar = True
    ActiveWindow.DisplayHorizontalScrollBar = True
    If Application.CommandBars.GetPressedMso("MinimizeRibbon") Then Application.CommandBars.ExecuteMso "MinimizeRibbon"
'    Application.DisplayFullScreen = False
'    Application.WindowState = xlNormal
End Sub

'近傍の生存セル数
Public Function Vicinity(ByRef Buffer() As Byte, ByVal i As Long, ByVal j As Long, ByVal OnColor As Byte) As Long

    Vicinity = 0
    
    If i > LBound(Buffer, 1) Then
        If j > LBound(Buffer, 2) Then
            If Buffer(i - 1, j - 1) = OnColor Then Vicinity = Vicinity + 1 '左上
        End If
        
        If j < UBound(Buffer, 2) Then
            If Buffer(i - 1, j + 1) = OnColor Then Vicinity = Vicinity + 1 '右上
        End If
        
        If Buffer(i - 1, j) = OnColor Then Vicinity = Vicinity + 1 '上
    End If
    
    If i < UBound(Buffer, 1) Then
        If j > LBound(Buffer, 2) Then
            If Buffer(i + 1, j - 1) = OnColor Then Vicinity = Vicinity + 1 '左下
        End If
        
        If j < UBound(Buffer, 2) Then
            If Buffer(i + 1, j + 1) = OnColor Then Vicinity = Vicinity + 1 '右下
        End If
        
        If Buffer(i + 1, j) = OnColor Then Vicinity = Vicinity + 1 '下
    End If
    
    If j > LBound(Buffer, 2) Then
        If Buffer(i, j - 1) = OnColor Then Vicinity = Vicinity + 1 '左
    End If
    
    If j < UBound(Buffer, 2) Then
        If Buffer(i, j + 1) = OnColor Then Vicinity = Vicinity + 1 '右
    End If
End Function

'1バイト値をRGBに変換する
Public Function HEX2RGB(ByVal val As Byte) As Long
    '上位3bitをR, 3bitをG, 2bitをBに変換
    'Andでマスク &HE0 = 111 000 00
    '　　　　　　&H1C = 000 111 00
    '　　　　　　&H3 = 000 000 11
    '"\ (2 ^ [桁数])"でシフト
    '表示できる最大数(RG:7, B:3)で割って、255をかける
    Dim r As Long, g As Long, b As Long
    r = Int(CLng((val And &HE0) \ (2 ^ 5)) / 7 * 255)
    g = Int(CLng((val And &H1C) \ (2 ^ 2)) / 7 * 255)
    b = Int(CLng(val And &H3) / 3 * 255)
    HEX2RGB = RGB(r, g, b)
End Function

'虹色を生成する
'42で1周する
'r : 0~7
'g : 0~7
'b : 0~3
Public Function GenerateRainbow(ByVal seed As Long) As Byte
    
    Dim r As Long, g As Long, b As Long
    
    Do While seed >= 42
        seed = seed - 42
    Loop
    
    If 0 <= seed And seed < 7 Then
        r = 7
        g = 0
        b = Int(seed / 2)
    ElseIf 7 <= seed And seed < 14 Then
        r = 14 - seed
        g = 0
        b = 3
    ElseIf 14 <= seed And seed < 21 Then
        r = 0
        g = seed - 14
        b = 3
    ElseIf 21 <= seed And seed < 28 Then
        r = 0
        g = 7
        b = Int((28 - seed) / 2)
    ElseIf 28 <= seed And seed < 35 Then
        r = seed - 28
        g = 7
        b = 0
    ElseIf 35 <= seed And seed < 42 Then
        r = 7
        g = 42 - seed
        b = 0
    End If
    
    GenerateRainbow = (r * (2 ^ 5)) Or (g * (2 ^ 2)) Or b
    
End Function

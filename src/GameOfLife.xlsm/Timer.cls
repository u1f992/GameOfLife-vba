VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Timer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

''' 高精度タイマーモジュール
''' <example>
'''
''' 宣言例
''' <code>
''' Dim clsTimer As New Timer
''' </code>
'''
''' ある処理にかかる時間を計測する例
''' <code>
''' clsTimer.StartTimer
''' ' 計測対象の処理
''' Debug.Print clsTimer.StopTimer
''' </code>
'''
''' 複数区間の経過時間を計測する例
''' <code>
''' clsTimer.StartTimer
''' ' 計測対象の処理_1
''' Debug.Print clsTimer.TakeLap
''' ' 計測対象の処理_2
''' Debug.Print clsTimer.StopTimer
''' </code>
'''
''' 複数区間の計測結果をForループで表示する例
''' <code>
''' Dim i As Long
''' For i = 0 To clsTimer.Length - 1
'''     Debug.Print clsTimer.Laps(i)
''' Next i
''' </code>
''' </example>

Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (ByRef freq As LongPtr) As Long
Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (ByRef procTime As LongPtr) As Long
Private freq As LongPtr

''' <summary>
''' タイマー本体
''' </summary>
Private t As Double

''' <summary>
''' ラップ一覧を格納する配列
''' </summary>
Private arrLap() As Double

''' <summary>
''' タイマーが実行中か判定する
''' </summary>
Private flag As Boolean

''' <summary>
''' タイマーの状況
''' </summary>
Public Property Get Running() As Boolean
    Running = flag
End Property

''' <summary>
''' Lapsプロパティの要素数
''' </summary>
Public Property Get Length()

    If flag = False Then
    
        If UBound(arrLap) = 0 Then
            Err.Raise 1026, "Timer.Length", "タイマー実行前は無効のプロパティです。"
        Else
            Length = UBound(arrLap) + 1
        End If
        
    Else
        Err.Raise 1024, "Timer.Length", "タイマー実行中は無効のプロパティです。"
    End If
    
End Property

''' <summary>
''' タイマーを開始する
''' </summary>
''' <return>記録した開始時間 = 0</return>
Public Function StartTimer()
    
    If flag = False Then
        
        flag = True
    
        ' t, arrLap()をリセット
        t = 0
        ReDim arrLap(0)
        
        t = GetSecond
        arrLap(0) = 0 ' 開始時点のラップ = 0
        StartTimer = 0
    
    Else
        Err.Raise 1024, "Timer.StartTimer", "タイマー実行中は無効のメソッドです。"
    End If
    
End Function

''' <summary>
''' タイマーを終了する
''' </summary>
''' <return>記録した終了時間</return>
Public Function StopTimer() As Double

    If flag = True Then
        StopTimer = GetSecond - t
        
        ' 最後のラップに記録する
        ReDim Preserve arrLap(UBound(arrLap) + 1)
        arrLap(UBound(arrLap)) = StopTimer
    
        flag = False
    
    Else
        Err.Raise 1025, "Timer.StopTimer", "タイマー停止中は無効のメソッドです。"
    End If
    
End Function

''' <summary>
''' ラップを記録する
''' </summary>
''' <return>記録したラップ</return>
Public Function TakeLap() As Double
    
    If flag = True Then
        TakeLap = GetSecond - t
        
        ReDim Preserve arrLap(UBound(arrLap) + 1)
        arrLap(UBound(arrLap)) = TakeLap
    
    Else
        Err.Raise 1025, "Timer.TakeLap", "タイマー停止中は無効のメソッドです。"
    End If
    
End Function

''' <summary>
''' ラップを取得する
''' </summary>
''' <param name="index">要素 0~(Length - 1)</param>
Public Property Get Laps(ByVal index As Long)

    If flag = False Then
        If UBound(arrLap) = 0 Then
            Err.Raise 1026, "Timer.Laps", "タイマー実行前は無効のプロパティです。"
        ElseIf index < LBound(arrLap) Or UBound(arrLap) < index Then
            Err.Raise 1027, "Timer.Laps", "インデックスが有効範囲にありません。"
        End If
        
        Laps = arrLap(index)
    
    Else
        Err.Raise 1024, "Timer.Laps", "タイマー実行中は無効のプロパティです。"
    End If
    
End Property

''' <summary>
''' 指定秒数待機する(DoEvents)
''' </summary>
''' <param name="ms">ミリ秒</param>
Public Sub Sleep(ByVal ms As Long)
    Dim temp As Double: temp = GetMilliSecond
    Do While GetMilliSecond - temp < ms
        DoEvents
    Loop
End Sub

Private Sub Class_Initialize()
    QueryPerformanceFrequency freq
    
    flag = False
    t = 0
    ReDim arrLap(0)
End Sub

''' <summary>
''' 現在の秒数を取得する
''' </summary>
Private Function GetSecond() As Double
    
    Dim procTime As LongPtr
    
    'カウント数を「1マイクロ秒間に増えるカウント数」で割ることで秒を取得できる
    QueryPerformanceCounter procTime
    GetSecond = procTime / freq
    
End Function

''' <summary>
''' 現在の秒数を取得する(ミリ秒)
''' </summary>
Private Function GetMilliSecond() As Double
    GetMilliSecond = GetSecond * 1000
End Function

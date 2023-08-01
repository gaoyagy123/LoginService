VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "服务器端"
   ClientHeight    =   5805
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9555
   LinkTopic       =   "Form1"
   ScaleHeight     =   5805
   ScaleWidth      =   9555
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "清空数据库 "
      Height          =   375
      Left            =   6240
      TabIndex        =   2
      Top             =   0
      Width           =   1335
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   5175
      Left            =   0
      TabIndex        =   1
      Top             =   600
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   9128
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Index           =   0
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock listener 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Index           =   1
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Index           =   2
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Index           =   3
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Index           =   4
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Index           =   5
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Index           =   6
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Index           =   7
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Index           =   8
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Index           =   9
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Index           =   10
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Index           =   11
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Index           =   12
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Index           =   13
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Index           =   14
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Index           =   15
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   720
      TabIndex        =   0
      Top             =   240
      Width           =   3495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Private Declare Sub GetLocalTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)
'定义常量
Const BUSY As Boolean = False
Const FREE As Boolean = True
'定义连接状态
Dim ConnectState() As Boolean
Private Type SYSTEMTIME ' 16 Bytes
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Dim Conn As New ADODB.Connection

Private Sub Command1_Click()
      MsgBox CheckData("gaoyagy", "asdas")

End Sub
Function AddData(UserName As String, Paswd As String, Nick As String, Tname As String, CardID As String, Gold As Long, Icon As Long, Friends As String) As Long
    Dim Rs As New ADODB.Recordset
    Rs.Open "select   *   from userdata", Conn, adOpenDynamic, 3
    AddData = Rs.RecordCount
    Rs.AddNew
    Rs(0).Value = AddData
    Set DataGrid1.DataSource = Rs
    DataGrid1.Columns(1).Value = UserName
    DataGrid1.Columns(2).Value = Paswd
    DataGrid1.Columns(3).Value = Nick
    DataGrid1.Columns(4).Value = Tname
    DataGrid1.Columns(5).Value = CardID
    DataGrid1.Columns(6).Value = Gold
    DataGrid1.Columns(7).Value = Icon
    Rs.Update
    Rs.Close
    Rs.Open "Select * from friend", Conn, adOpenDynamic, 3
    Rs.AddNew
    Rs(0) = UserName
    Rs(1) = Friends
    Rs.Update
    Rs.Close
    Set Rs = Nothing
End Function
Function DeleteAll() As Long
    Dim Rs As New ADODB.Recordset
    Rs.Open "select   *   from userdata", Conn, adOpenDynamic, 3
    Rs.MoveFirst
    Do While Not Rs.EOF
        Rs.Delete adAffectCurrent
        Rs.MoveNext
    Loop
    Rs.Close
    Set Rs = Nothing
End Function

Function GetDataExists(Value As String) As Boolean
    Dim Rs As New ADODB.Recordset
    Rs.Open "Select * from userdata where ID=""" & Value & """", Conn, adOpenDynamic, 3
    
    If Not Rs.EOF Then GetDataExists = True
    Rs.Close
    Set Rs = Nothing
End Function
Function CheckData(Name As String, Passwd As String) As Boolean
    Dim Rs As New ADODB.Recordset
    Rs.Open "Select * from userdata where ID=""" & Name & """", Conn, adOpenDynamic, 3
    If Rs.EOF Then
        Rs.Close
        Set Rs = Nothing
        Exit Function
    End If
    If Rs("password") = Passwd Then CheckData = True
    Rs.Close
    Set Rs = Nothing
End Function
Private Sub Command2_Click()
    DeleteAll
End Sub

Private Sub Form_Load()
    ReDim Preserve ConnectState(0 To 1)
    On Error Resume Next
    ConnectState(0) = FREE
    ConnectState(1) = FREE
    '指定网络端口号
    listener.LocalPort = "5555"
    '开始侦听
    listener.Listen
    InitData
End Sub
Sub InitData()
    Dim CnnStr  As String
    CnnStr = "DRIVER={MySQL ODBC 5.3 ANSI Driver};server=localhost;port=3307;uid=root;pwd=123456;database=mytalk"
    Conn.Open CnnStr
    Conn.CursorLocation = adUseClient
End Sub

Private Sub Form_Unload(Cancel As Integer)
    listener.Close
    Conn.Close
    Set Conn = Nothing
End Sub

Private Sub Listener_ConnectionRequest(ByVal requestID As Long)
    Dim SockIndex As Integer
    Dim SockNum As Integer
    On Error Resume Next
    Label1.Caption = requestID & "连接请求"
    '查找连接的用户数
    SockNum = UBound(ConnectState)
    If SockNum > 14 Then
       ' Exit Sub
    End If
    
    '查找空闲的sock
    SockIndex = FindFreeSocket
    
    '如果已有的sock都忙，而且sock数不超过15个，动态添加sock
    If SockIndex > SockNum Then
        Load Winsock1(SockIndex)
    End If
    ConnectState(SockIndex) = BUSY
    Winsock1(SockIndex).Tag = SockIndex
    '接受请求
    Winsock1(SockIndex).Accept requestID
End Sub


'客户断开，关闭相应的sock
Private Sub Winsock1_Close(Index As Integer)
    Label1.Caption = Winsock1(Index).LocalIP & "断开了"
    If Winsock1(Index).State <> sckClosed Then
        Winsock1(Index).Close
    End If
    ConnectState(Index) = FREE
End Sub

'接收数据
Private Sub Winsock1_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim dx As String
    Dim Uname As String, Passwd As String, Tmp As String
    Dim Nick As String, Tname As String, Tid As String
    Label1.Caption = "数据来自" & Winsock1(Index).LocalIP
    Winsock1(Index).GetData dx, vbString
    If InStr(dx, "&") Then
        Uname = Split(dx, "&")(0)
        Passwd = Split(dx, "&")(1)
        If CheckData(Uname, Passwd) Then
            Winsock1(Index).SendData "LoginSucess"
        Else
            Winsock1(Index).SendData "LoginFaild-2"
        End If
    ElseIf InStr(dx, "|") Then
        Uname = Split(dx, "|")(0)
        If GetDataExists(Uname) Then
            Winsock1(Index).SendData "RegFaild"
            Exit Sub
        End If
        Passwd = Split(dx, "|")(1)
        Nick = Split(dx, "|")(2)
        Tname = Split(dx, "|")(3)
        Tid = Split(dx, "|")(4)
        AddData Uname, Passwd, Nick, Tname, Tid, 0, 0, "admin"
        Winsock1(Index).SendData "RegSucess"
    End If
End Sub

Sub MakeDir(Uname As String)
    On Error GoTo ProErr
    MakeDirBase
    MkDir App.Path & "\Data\" & Uname
    Exit Sub
ProErr:
End Sub
Sub MakeDirBase()
    On Error GoTo ProErr
    MkDir App.Path & "\Data"
    Exit Sub
ProErr:
End Sub
'寻找空闲的sock
Public Function FindFreeSocket()
    Dim SockCount, I As Integer
    SockCount = UBound(ConnectState)
    For I = 0 To SockCount
        If ConnectState(I) = FREE Then
            FindFreeSocket = I
            Exit Function
        End If
    Next I
    ReDim Preserve ConnectState(0 To SockCount + 1)
    FindFreeSocket = UBound(ConnectState)
End Function

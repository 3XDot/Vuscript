VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "VUS"
   ClientHeight    =   2220
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5325
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2220
   ScaleWidth      =   5325
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.TextBox txtZNKZZAll 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tekton Pro Cond"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "Form1.frx":1084A
      Top             =   480
      Width           =   5175
   End
   Begin VB.Label lblZn 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "znkzz(vzpui7@gmail.com)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   2400
      TabIndex        =   3
      ToolTipText     =   "XD"
      Top             =   1560
      Width           =   2535
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblZnkzz 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "���������Ҫ�����ߣ�"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   0
      MousePointer    =   12  'No Drop
      TabIndex        =   2
      Top             =   1560
      Width           =   2400
   End
   Begin VB.Label lblVusVersion 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "vus version��1.5.4 beta 2(R)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   3585
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim ShowErrMsg As Boolean    '���������Ƿ񵯴�
Dim ErrStop As Boolean    '���������Ƿ�END
Dim MyFile As String
Dim TheLine, SLine As Integer    '������������������¼����
Attribute SLine.VB_VarUserMemId = 1073938434
Dim StrTemp As String
Dim VarData(-1 To 16) As Variable '-1������ʱ�������� ,0 �������ͷ ����ʹ��, 1-16 �����������
Dim StrTmp_$()   'err 4
Dim ErrType(1) As Integer, ErrMess_$, ErrSub$
'Dim IsIf_ As Boolean
'****************************************
Private Type Variable
        V_TYPE As Integer
      V_NAME As String
       V_DATA_TEXT As String
      V_DATA_INT As Long
      V_DATA_BOL As Boolean
   End Type

'****************************************

Private Sub Form_Load()

    'If (Environ$("TEMP") = "B:\TEMP") Then MsgBox 11
    'Exit Sub
    On Error GoTo errhand

    'CreateVar 2, "znkzz", "vus by znkzz date 2020-1-1 20:28:10"
If ErrType(1) <> -1 Then MyFile = ".vus"
    #If VB_DEBUG = 1 Then

        
        If (Command = "-v" Or Command = "version" Or Command = "-version") And Len(Command) <> 0 And ErrType(1) <> -1 Then
            Me.Show

            Exit Sub

        ElseIf Len(Command) <> 0 And ErrType(1) <> -1 Then

            Dim Comman As String

            ' MyFile = Mid$(Command, InStr(1, StrTemp, ""), Len(Command) - InStr(1, StrTemp, ""))
            'Debug.Print Mid$(Command, InStr(1, StrTemp, ""), Len(Command) - InStr(1, StrTemp, ""))
            Comman = Replace(Command, """", "")
            MyFile = Comman
      
        End If

        '1.5.4B2A1.exe "C:\Users\Administrator\Desktop\byq\sss\.vus"
    #End If
    '*******************��ʼ������************************
    SLine = 1
    ErrStop = False
    ShowErrMsg = True
     
    '*******************��ʼ������************************
    If Len(MyFile) = 0 Then End     'err x
    If FileLen(MyFile) = 0 Then End     'https://wenda.so.com/q/1365482299062147?src=140

    GetLine    'err 2

    Get_All_Line

GoBack:

    'err x

    ReadLine

    ExecuteCode

    If SLine = TheLine Then
        If ErrType(1) = -1 Then

            Exit Sub

        Else

            End

        End If

    Else
        SLine = SLine + 1    'deb

    End If

    GoTo GoBack

    Exit Sub

errhand:
    ErrType(0) = 1
    ErrMess_ = "Unknow Error"
    ErrSub = "Initialization" + " @" + CStr(Err.Number)
    ShowErr

    End

End Sub

Sub GetLine()
     On Error GoTo errhand
     Dim f
     f = FreeFile
     Open MyFile For Binary As #f
     TheLine = UBound(Split(Input(LOF(1), #1), vbCrLf)) + 1
     Close #f

     Exit Sub
errhand:

     ErrType(0) = 1
     ErrMess_ = "Unknow Error"
     ErrSub = "Initialization" + " @" + CStr(Err.Number)
     ShowErr

End Sub

Sub ExecuteCode()

    ' On Error GoTo errhand
  
    Dim wtf(-2 To 6) As Variable

    If StrTemp = "" Then Exit Sub
 
    Dim iTEMP%

233:
    iTEMP = InStr(4, StrTemp, "$")

    If iTEMP <> 0 Then '

        StrTemp = Replace(StrTemp, Mid$(StrTemp, iTEMP, InStr(iTEMP + 1, StrTemp, "$") - iTEMP + 1), RetVarRealValue(GetVar(Mid$(StrTemp, iTEMP + 1, InStr(iTEMP + 1, StrTemp, "$") - iTEMP - 1))))

        If InStr(4, StrTemp, "$") <> 0 Then GoTo 233
    End If
    
    '---------------------------------------------------------------------------------------------------------------------------
    ''��ȡ����ĳ��� https://wenda.so.com/q/1364416929060667
    ' wtf() = Split(Replace(Join(wtf), "#\A1", "|"), "|")
    '**********************�滻�����ַ�*********** Dim Along%

    ' ********************************Else����*****************************
    If StrTemp = "else" Then

        For i = SLine To TheLine

            If StrTmp_(i) = "end if" Then
                SLine = i: Exit Sub
            End If

        Next

    End If

    '*************************************
    '*************************end if ����********************************

    If StrTemp = "end if" Then

        StrTemp = StrTmp_(SLine)

        Exit Sub

    End If

    '*****************************************************************
    Call SuperClCode(wtf)


    'Along = UBound(wtf()) - LBound(wtf()) + 1

    ' Dim ib As Integer

    'along ��һ��ָ��ĸ���
    ' For ib = 1 To Along - 1
    '    wtf(ib) = Replace(wtf(ib), "#\A1", "|")

    ' Next
   
    '*********************************
    'UBound(a) - LBound(a) + 1
    Select Case wtf(0).V_DATA_TEXT

        Case "MoveFile"
            FileCopy wtf(1).V_DATA_TEXT, wtf(2).V_DATA_TEXT
            Kill wtf(1).V_DATA_TEXT

        Case "KillFile"
            Kill wtf(1).V_DATA_TEXT

        Case "RunCmdCommand"
            Shell wtf(1).V_DATA_TEXT, wtf(2).V_DATA_INT

        Case "ShowMessege"

          '  If Len(wtf(2).V_DATA_TEXT) = 0 Then wtf(2).V_DATA_INT = 0  2020 /1 /27 BUG  FIXED
            If Len(wtf(3).V_DATA_TEXT) = 0 Then wtf(3).V_DATA_TEXT = App.Title
       
            MsgBox wtf(1).V_DATA_TEXT, wtf(2).V_DATA_INT, wtf(3).V_DATA_TEXT


        Case "DowFile"

            If Len(wtf(1).V_DATA_TEXT) = 0 Or Len(wtf(2).V_DATA_TEXT) = 0 Then

                ErrType(0) = 1
                ErrMess_ = "δָ�����ػ��ļ�·��"
                GoTo errhand
            End If

            Dim x, S

            Set x = CreateObject("Microsoft.XMLHTTP")
            x.Open "GET", wtf(1).V_DATA_TEXT, 0
            x.Send
            Set S = CreateObject("ADODB.Stream")
            S.Mode = 3
            S.Type = 1
            S.Open
            S.Write (x.responseBody)
            S.SaveToFile LCase(wtf(2).V_DATA_TEXT), 2
            Set x = Nothing
            Set S = Nothing

        Case "GoTo"
            SLine = wtf(1).V_DATA_INT - 1

        Case "if"

            If (wtf(2).V_DATA_TEXT <> "<>" And wtf(2).V_DATA_TEXT <> "=") Then: Exit Sub

            If CheckCondition(wtf(1).V_DATA_TEXT + wtf(2).V_DATA_TEXT + wtf(3).V_DATA_TEXT) <> True Then

                Dim ia%

                For ia = SLine To TheLine

                    If StrTmp_(ia) = "else" Then
                        SLine = ia

                        Exit Sub

                    End If

                Next

            Else

                Exit Sub

            End If

        Case "System.CSettings.Err2Message"

            If wtf(1).V_DATA_TEXT = "true" Then
                ShowErrMsg = True

                If wtf(2).V_DATA_TEXT = "true" Then ErrStop = True
            Else

                If wtf(2).V_DATA_TEXT = "false" Then ErrStop = False
                ShowErrMsg = False
            End If

        Case "System. Finish"

            End

        Case "LoadOtherVusFile_D" '//�첽

            Dim k   As Integer, redata() As String, revar() As Variable

            Dim fg2 As Integer

            fg2 = SLine
            ReDim redata$(0 To TheLine)
            redata(0) = MyFile

            For k = 1 To UBound(StrTmp_)
                redata(k) = StrTmp_(k)
            Next

            '// �˴��и�Ī�������bug�������
            If Len(VarData(0)) <> 0 Then
     
                Dim flag As Boolean
       
                flag = True
            
                ReDim revar(0 To 16) As Variable

                For k = 0 To 16
                    revar(k) = VarData(k)
                Next
                
            End If

            k = TheLine
           
            '******************************BEGIN*************************
            MyFile = wtf(1).V_DATA_TEXT
            
            ' VarData$(0) = ""
            ErrType(1) = -1

            Dim Const_ As Integer

            For Const_ = 0 To 16
                VarData(k) = FreeVar(VarData(Const_))
            Next
                
            Form_Load
            '******************************END*************************

            ReDim StrTmp_$(1 To k)
            TheLine = k

            For k = 1 To TheLine
                StrTmp_$(k) = redata$(k)
            Next

            If flag = True Then

                For k = 0 To 16
                    VarData(k) = revar(k)
                Next

            End If

            SLine = fg2
            MyFile = redata(0)
            ErrType(1) = 0

            Exit Sub

        Case "LoadOtherVusFile" '//��λ����Ը����Ҷ�Ϊ֮�� date 19.10.29
            Shell App.Path + "\" + App.EXEName + ".exe " + wtf(1).V_DATA_TEXT

        Case "SetSysVar"
            CreateVar CLng(wtf(3).V_DATA_INT), wtf(1).V_DATA_TEXT, RetVarRealValue(wtf(2))

        Case Else

            If InStr(StrTemp, "#") = 1 Then

                Exit Sub

            Else
                ErrType(0) = 1
                ErrMess_ = "δ���庯��" & wtf(0).V_DATA_TEXT
                GoTo errhand
            End If

    End Select

    Exit Sub

errhand:
    ErrSub = "ExecuteCode"
    ShowErr

End Sub

Sub ShowErr()
   

     If ErrStop = False And ShowErrMsg = False Then Exit Sub

     If ShowErrMsg Then
          If ErrStop Then End

     End If

     If ErrType(0) = 1 Then
          MsgBox "Error Type: " & "����ʱ����" & vbCrLf & "Error Description:  " & ErrMess_ & vbCrLf + Err.Description & vbCrLf & "Error Sub:  " & ErrSub & vbCrLf + "Error Line:" + CStr(SLine), 16
     Else
          MsgBox "Error Type: " & "ִ��ʱ����" & vbCrLf & "Error Number: " & Err.Number & vbCrLf & "Error Description: " & Err.Description & vbCrLf & "Error Sub:  " & ErrSub & vbCrLf + "Error Line:" + CStr(SLine), vbExclamation
     End If
     If ErrStop = False And ShowErrMsg = False Then End

End Sub

'***************************************************************************************************************
Sub Get_All_Line()     'GAL FUN

'emmm....û�������ڳ�������XD
     On Error GoTo errhand

     Dim TheNow As Integer

     TheNow = 1

     ReDim StrTmp_(1 To TheLine)     ' ����һ���Ե�ǰ����Ϊ��С������
     Open MyFile For Input As #1

     While (Not EOF(1))

          Line Input #1, StrTmp_(TheNow)
          ' MsgBox StrTmp_(TheNow)
          TheNow = TheNow + 1

     Wend

     Close #1
     Exit Sub
errhand:
     ErrType(0) = 1
     ErrSub = "Get_All_Line"
     ErrMess_ = "���Զ�ȡ�ű��ļ�ʱ���� " + "@" + CStr(Err.Number)
     ShowErr
End Sub

Sub ReadLine()
     On Error GoTo errhand


     StrTemp = StrTmp_(SLine)


     Exit Sub
errhand:
     MsgBox "Error Number: " & vbTab & Err.Number & vbCrLf & "Error Description: " & vbTab & Err.Description & vbCrLf & "Error Sub: " & vbTab & "ReadLine", vbExclamation
End Sub
'***************************************************************************************************************

Private Function CheckCondition(���ʽ$) As Boolean
     On Error GoTo errhand
Dim tg$

     Dim TStr$()     '//�����洢 0���� ��1��ֵ
     If InStr(���ʽ$, "@") <> 0 Then
          If InStr(���ʽ$, "=") = 0 Then
               TStr$() = Split(���ʽ$, "<>")
               TStr$(0) = Replace(TStr$(0), "@", "")
               tg = Environ(TStr$(0))
               If tg <> Replace(TStr$(1), """", "") Then

                    CheckCondition = True
               Else

                    CheckCondition = False
               End If
               Exit Function
       
          Else

             TStr$() = Split(���ʽ$, "=")
               TStr$(0) = Replace(TStr$(0), "@", "")
               tg = Environ(TStr$(0))
               If tg = Replace(TStr$(1), """", "") Then

                    CheckCondition = True
               Else

                    CheckCondition = False
               End If
               Exit Function
          End If
     Else
          If InStr(���ʽ$, "=") = 0 Then
               TStr$() = Split(���ʽ$, "<>")
               TStr$(0) = Replace(TStr$(0), "#", "")
               tg = GetVar(TStr$(0)).V_DATA_TEXT
               '��ת����
               If tg <> TStr$(1) Then

                    CheckCondition = True
               Else

                    CheckCondition = False
               End If
               Exit Function
         
          Else

               TStr$() = Split(���ʽ$, "=")
               TStr$(0) = Replace(TStr$(0), "#", "")
               tg = GetVar(TStr$(0)).V_DATA_TEXT
               

               If tg = TStr$(1) Then
                    CheckCondition = True
               Else
                    CheckCondition = False
               End If
               Exit Function
          End If

     End If

     Exit Function
errhand:
     MsgBox "Error Number: " & vbTab & Err.Number & vbCrLf & "Error Description: " & vbTab & Err.Description & vbCrLf & "Error Function: " & vbTab & "CheckCondition", vbExclamation
End Function

Private Function GetVar(vname$) As Variable
    Dim i As Integer

    For i = 1 To 16
        If VarData(i).V_NAME = vname Then
            GetVar = VarData(i)


            Exit Function
            
        End If
            
    Next


    ErrType(0) = 1
    ErrSub = "GetVar"
    ErrMess_ = "���Զ�ȡ""" + vname + """" + "����ʱ����:δ�ҵ��˱���"
    ShowErr
   
End Function

Private Sub CreateVar(ctype As Integer, vname As String, dat As String)
    
    Dim i%

    For i = 1 To 16

        If Len(VarData(i).V_NAME) = 0 Then

            If ctype = 1 Then
                VarData(i).V_NAME = vname
                VarData(i).V_DATA_INT = CLng(dat)
                VarData(i).V_TYPE = 1
            ElseIf ctype = 2 Then
                VarData(i).V_NAME = vname
                VarData(i).V_DATA_TEXT = CStr(dat)
                VarData(i).V_TYPE = 2
            ElseIf ctype = 3 Then
                VarData(i).V_NAME = vname
                VarData(i).V_DATA_BOL = Str2Bool(dat)
                VarData(i).V_TYPE = 3
            End If

            Exit Sub

        End If

    Next

    ErrType(0) = 1
    ErrSub = "CreateVar"
    ErrMess_ = "����" + """" + vname + """" + "����ʱ���� "
    ShowErr
End Sub

Private Function Str2Bool(text As String) As Boolean
 If Len(text) = 4 Or Len(text) = 5 Then
        If LCase(text) = "true" Then
            Str2Bool = True
        ElseIf LCase(text) = "false" Then
            Str2Bool = False
        End If

        Exit Function

    End If
  ErrType(0) = 1
     ErrSub = "Str2Bool"
     ErrMess_ = "ת������ʱ���� "
     ShowErr
End Function
Private Function FreeVar(vn As Variable) As Variable
vn.V_NAME = ""
FreeVar = vn
End Function

Private Sub SuperClCode(ByRef wtf() As Variable)
    On Error GoTo errhand
    ' date 2020.1.26 ǰ����д���������û���棬����
    ' fun("fff",444,555,$d$)
    Dim ���ֳ� As String


    Dim ��ͼ As String

    Dim ��ͼ2 As String

    Dim ��ͼ_A() As String

    'Dim wtf() As Variable
    Dim ͼ��1 As Integer

    'Dim ͼ��2 As Integer
    Dim ͼ��3 As Integer
    ��ͼ = Mid$(StrTemp, InStr(StrTemp, "(") + 1, InStrRev(StrTemp, ")") - InStr(StrTemp, "(") - 1)
    ��ͼ2 = Left$(StrTemp, InStr(StrTemp, "(") - 1)
    'now is ("fff",444,555,$d$+5)
    wtf(0).V_NAME = "%NONE%"
    wtf(0).V_TYPE = 2
    wtf(0).V_DATA_TEXT = ��ͼ2
  '  If ��ͼ2 = "if" Then
   'If CheckCondition(StrTemp) Then
 'to do��
 '1.ɾ����Чif�������
 '2.��д���˺������еĴ������
 '3.֧��GetVar()��ñ���ֵ
 '4.���Ӷ��ڱ�����Ч�飬��ֹδ��ʼ������
 '5.bug fix
 
 
 
 
 '  End If
  '  If InStr(��ͼ, "<") Then
  '  ��ͼ_A() = Split(��ͼ, "<>")
'End If
    ��1 = 0: ��2 = -1 ' ��ʱ��д�������� fun("f+f,f",cstr(444),555,$d$+5) �Ĵ���
��ͼ_A() = Split(��ͼ, ",")

    'ReDim wtf(0 To UBound(��ͼ_A()) - LBound(��ͼ_A()))
    'GUN:
  
    For ͼ��3 = 0 To (UBound(��ͼ_A()) - LBound(��ͼ_A()))
 ͼ��1 = InStr(��ͼ_A(ͼ��3), """")
        If ͼ��1 <> 0 Then
            wtf(ͼ��3 + 1).V_NAME = "%NONE%"
            wtf(ͼ��3 + 1).V_TYPE = 2
            wtf(ͼ��3 + 1).V_DATA_TEXT = Replace$(��ͼ_A(ͼ��3), """", "")
         
        ElseIf LCase(��ͼ_A(ͼ��3)) = "false" Or LCase(��ͼ_A(ͼ��3)) = "true" Then

            wtf(ͼ��3 + 1).V_NAME = "%NONE%"
            wtf(ͼ��3 + 1).V_TYPE = 3
            wtf(ͼ��3 + 1).V_DATA_BOL = CBool(��ͼ_A(ͼ��3))
        ElseIf IsNumeric(��ͼ_A(ͼ��3)) Then
            wtf(ͼ��3 + 1).V_NAME = "%NONE%"
            wtf(ͼ��3 + 1).V_TYPE = 1
wtf(ͼ��3 + 1).V_DATA_INT = CLng(��ͼ_A(ͼ��3))
        End If

    Next
errhand:
 ErrType(0) = 1
     ErrSub = "SuperClCode"
     ErrMess_ = "һ���������ʱ����:�﷨����? "
     ShowErr
  
End Sub

Private Function RetVarRealValue(data As Variable) As String

    If data.V_TYPE = 1 Then
        RetVarRealValue = (data.V_DATA_INT)
    ElseIf data.V_TYPE = 2 Then
        RetVarRealValue = """" + data.V_DATA_TEXT + """"
    ElseIf data.V_TYPE = 3 Then

        If data.V_DATA_BOL = True Then
            RetVarRealValue = "true"
        Else
            RetVarRealValue = "false"
        End If
    End If

End Function



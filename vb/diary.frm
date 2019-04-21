VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form Form1 
   Caption         =   "轻小说"
   ClientHeight    =   7380
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   18165
   Icon            =   "diary.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7380
   ScaleWidth      =   18165
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton cmd_pan 
      Caption         =   "百度盘！"
      Height          =   255
      Left            =   16980
      TabIndex        =   31
      Top             =   240
      Width           =   915
   End
   Begin VB.TextBox txtdb 
      Height          =   435
      Left            =   120
      TabIndex        =   30
      Text            =   "dbook.xls"
      Top             =   6000
      Width           =   1575
   End
   Begin VB.TextBox Txt_style 
      Height          =   435
      Left            =   4020
      TabIndex        =   29
      Top             =   1920
      Width           =   2295
   End
   Begin VB.CommandButton Cmd_run 
      Caption         =   "*生成网页(&R)"
      Height          =   435
      Left            =   6720
      TabIndex        =   28
      Top             =   5640
      Width           =   1695
   End
   Begin VB.TextBox txtcheck 
      Height          =   435
      Left            =   1560
      TabIndex        =   25
      ToolTipText     =   "URL列表"
      Top             =   240
      Width           =   4755
   End
   Begin VB.TextBox txtbuy 
      Height          =   375
      Left            =   1560
      TabIndex        =   24
      ToolTipText     =   "ASIN组成的URL列表"
      Top             =   660
      Width           =   4755
   End
   Begin VB.CommandButton cmd_go 
      Caption         =   "找链接!"
      Height          =   375
      Left            =   16980
      TabIndex        =   22
      Top             =   540
      Width           =   975
   End
   Begin VB.TextBox txtURL 
      Height          =   435
      Left            =   8520
      TabIndex        =   21
      Text            =   "http://www.amazon.co.jp"
      Top             =   540
      Width           =   8475
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   7800
      Top             =   0
   End
   Begin VB.TextBox txt_kindle 
      Height          =   435
      Left            =   1560
      TabIndex        =   19
      Top             =   1080
      Width           =   4815
   End
   Begin VB.CommandButton Cmd_checkNewBook 
      Caption         =   "查询新书(&T)"
      Height          =   375
      Left            =   6960
      TabIndex        =   18
      Top             =   840
      Width           =   1455
   End
   Begin VB.CommandButton Cmd_checkloss 
      Caption         =   "检查遗漏(&O)"
      Height          =   375
      Left            =   6960
      TabIndex        =   17
      Top             =   420
      Width           =   1455
   End
   Begin VB.CommandButton cmd_FTP 
      Caption         =   "3.FTP上传(&U)"
      Height          =   495
      Left            =   6720
      TabIndex        =   16
      Top             =   5100
      Width           =   1695
   End
   Begin VB.CommandButton cmd_title 
      Caption         =   "按书名作网页"
      Height          =   495
      Left            =   6960
      TabIndex        =   15
      Top             =   1680
      Width           =   1455
   End
   Begin VB.CommandButton Cmd_SP 
      Caption         =   "2.作分类列表(&B)"
      Height          =   555
      Left            =   6660
      TabIndex        =   14
      Top             =   4560
      Width           =   1755
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   5475
      Left            =   8460
      TabIndex        =   13
      Top             =   960
      Width           =   9255
      ExtentX         =   16325
      ExtentY         =   9657
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
      Location        =   "http:///"
   End
   Begin VB.CommandButton Cmd_wenku 
      Caption         =   "文库列表(&P)"
      Height          =   495
      Left            =   6960
      TabIndex        =   12
      Top             =   3540
      Width           =   1455
   End
   Begin VB.CommandButton Cmd_Exit 
      Caption         =   "退出(&E)"
      Height          =   435
      Left            =   6720
      TabIndex        =   11
      Top             =   6000
      Width           =   1695
   End
   Begin VB.CommandButton Cmd_Start 
      Caption         =   "转换txt(&S)"
      Default         =   -1  'True
      Height          =   435
      Left            =   6960
      TabIndex        =   10
      Top             =   2160
      Width           =   1455
   End
   Begin VB.CommandButton CmdCSV 
      Caption         =   "1.作新网页(&C)"
      Height          =   435
      Left            =   6720
      TabIndex        =   9
      Top             =   4140
      Width           =   1695
   End
   Begin VB.CommandButton Cmd_List 
      Caption         =   "Folder列表(&L)"
      Height          =   435
      Left            =   6960
      TabIndex        =   8
      Top             =   3120
      Width           =   1455
   End
   Begin VB.CommandButton Cmd_Lib 
      Caption         =   "读取书目XML(&F)"
      Height          =   435
      Left            =   6960
      TabIndex        =   7
      Top             =   2640
      Width           =   1455
   End
   Begin VB.CommandButton Cmd_buy 
      Caption         =   "购买新书(&D)"
      Height          =   435
      Left            =   6960
      TabIndex        =   6
      Top             =   1200
      Width           =   1455
   End
   Begin VB.FileListBox File1 
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3405
      Left            =   60
      Pattern         =   "*.txt"
      TabIndex        =   5
      Top             =   2520
      Width           =   6255
   End
   Begin VB.TextBox Txt_date 
      Height          =   375
      Left            =   1560
      TabIndex        =   4
      Top             =   1920
      Width           =   2175
   End
   Begin VB.TextBox Txt_path 
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Text            =   "C:\novel\"
      Top             =   1560
      Width           =   4815
   End
   Begin VB.Label Label_buy 
      Caption         =   "新书购买列表"
      Height          =   255
      Left            =   180
      TabIndex        =   27
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label label_new 
      Caption         =   "新书查询列表"
      Height          =   255
      Left            =   180
      TabIndex        =   26
      Top             =   360
      Width           =   1275
   End
   Begin VB.Label Label_web 
      Height          =   435
      Left            =   8580
      TabIndex        =   23
      Top             =   120
      Width           =   9195
   End
   Begin VB.Label Label4 
      Caption         =   "Kindle设备号"
      Height          =   315
      Left            =   180
      TabIndex        =   20
      Top             =   1140
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "更新截止日期"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "路径"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label1 
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   6480
      Width           =   17595
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub Sleep Lib "Kernel32.dll" (ByVal dwMilliseconds As Long)
Private Declare Function timeGetTime Lib "winmm.dll" () As Long
Private Declare Function internetSetOption Lib "wininet.dll" Alias "InternetSetOptionA" (ByVal hinternet As Long, ByVal dwoption As Long, ByRef lpBuffer As Any, ByVal dwbufferlength As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Type INTERNET_PROXY_INFO
dwAccessType As Long
lpszProxy As String
lpszProxyBypass As String
End Type
 

Dim sDomain As String   '域名
Dim maxItem As Integer  '分页显示
Dim filterflag As String * 1 '过滤title中的和谐字 1过滤 其余值不过滤
Dim kindle_1 As String
Dim kindle_2 As String
Dim ServerXmlHttp 'As MSXML2.ServerXmlHttp
Dim checkEndDate As String
Dim EXEPATH As String
Dim lv_free As String
Dim f As FTP
Dim bconn As Boolean
Dim fimg As FTP
Dim bconnimg As Boolean
Dim txt1 As String
Dim txt2 As String
Dim listTxt1 As String
Dim listTxt2 As String
Dim sok As String
Dim sql As String
Dim urlOk As Long
Dim cstyle As String '少年向 乙女向
Dim xlsConn As New ADODB.Connection
Dim xlsConnString As String
Dim imgDomain As String
Dim imgDomai2 As String
Dim transPre As String
Dim transPr2 As String
Dim txt3 As String
Dim proxyAddress As String 'IP:PORT
Dim titleget1 As String
Dim titleget2 As String
Dim urlget1 As String
Dim urlget2 As String
Dim totalnum1 As String
Dim totalnum2 As String
Dim totalnum3 As String
Dim totalnum4 As String

Private Type checkWenkuURLs
'0文库名1日文文库名2乙女向3文库日文地址4百科地址5新书地址6出版社
no As Integer
s0 As String
s1 As String
s2 As String 'style
s3 As String
s4 As String
s5 As String 'new book url
s6 As String 'publisher
s7 As Double 'book number from web query
s8 As Double 'book number from can pre-order
s9 As Double 'book number local query
sa As Double
sb As String
End Type

Private Type Books
isbn As String
id As String
Series As String
series_index As String
identifiers As String
title As String
authors As String
publisher As String
pubdate As String
size As String
cover  As String
author_sort As String
timestamp As String
tags As String
formats As String
rating As String
languages As String
words As String
pages As String
yesno As String
cnseries As String
cnseries_index As String
cntitle As String
vtitle As String
wenku As String
sorttitle As String
filetitle As String
asin As String
txtPath As String
typePath As String
newFileName As String
kindlePrice As String
paperPrice As String
comments As String
mainstyle As String
bookurl As String
coverURL As String
wenkuURL As String  '带<a></a>
seriesURL As String '带<a></a>
translateURL As String '带<a>译文</a>
authorsURL As String '带<a></a>
author1 As String
author2 As String
author1url As String
author2url As String
author1wiki As String
author2wiki As String
lastModify As String
tagMain As String
langA As String
titleURLA As String '带<a></a>
seriesindexURLA As String '带<a></a>
NumURLA As String '带<a></a>
transURLA As String '带<a></a>
titleFiltered As String
buyurlA As String '带<a></a>
finished As String
End Type
Dim URLTab() As checkWenkuURLs

Private Sub Cmd_buy_Click()
Call buyNewbook(txtbuy.Text)
Call writeutf8(EXEPATH & "template\log.txt", Format$(Now, "yyyy-mm-dd"))
End Sub

Public Sub buyNewbook(source As String)
Dim txtURL As String
Dim tURLItem() As String
Dim i As Integer
Dim j As Integer
Dim shtm As String
Dim wrongURL As String
If source = "" Or Dir(source) = "" Then Exit Sub

txtURL = readutf8(source)
If txtURL = "" Then Exit Sub
tURLItem = Split(txtURL, vbCrLf)
'setProxy (proxyAddress)

For i = 0 To UBound(tURLItem)
 '  Call WebBrowser1.Navigate2(tUrlItem(i))
 '  shtm = getHtmlStr(tUrlItem(i))
    shtm = sendHtmlStr(tURLItem(i))
    If shtm = "OK" Then
        j = j + 1
    ElseIf shtm <> "" Then
        wrongURL = wrongURL & tURLItem(i) & vbCrLf
    End If
 '  wait1000 3000
Next i
Label1.Caption = "已经成功访问" & j & "个网页 "
If wrongURL <> "" Then
    Label1.Caption = Label1.Caption & "失败了" & str(i + 1 - j) & "个"
    Call writeutf8(EXEPATH & "template\err.txt", wrongURL, "UTF-8")
Else
 'fileDelete (source)
End If

End Sub

Private Sub Cmd_checkloss_Click()
    Dim filesTab() As String
    Dim files As String
    Dim fileName As String
    Dim bookHtm As String
    Dim url As String
    Dim urlAll As String
    Dim i As Long
    Dim j As Long
    Dim n As Long
    Dim rs As New ADODB.Recordset
    Dim rsOld As New ADODB.Recordset
    File1.Path = "E:\novel\0soft\1\"
    For i = 1 To File1.ListCount
      fileName = File1.List(i - 1)
      files = readutf8(File1.Path & "\" & fileName)
      filesTab = Split(files, vbCrLf)
      For j = 0 To (UBound(filesTab) - 1)
        url = filesTab(j)
        url = Replace(url, "https://www.amazon.co.jp/dp/", "")
        If Len(url) = 10 Then
          urlAll = urlAll & url & vbCrLf
        End If
        DoEvents
      Next j
      DoEvents
    Next i
    writeutf8 EXEPATH & "new.csv", "asin" & vbCrLf & urlAll
    url = ""
    urlAll = ""
    i = 0
    
xlsConn.open xlsConnString
Call wait1000(1000)
Dim csvConn As New ADODB.Connection
csvConn.ConnectionString = "Driver={Microsoft Text Driver (*.txt; *.csv)};DefaultDir=" & EXEPATH
csvConn.open
rs.open "select * from new.csv", csvConn, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly
If rs.recordCount > 0 Then
     rs.MoveFirst
     While Not rs.EOF
     On Error GoTo end1
     url = rs.fields(0)
     rsOld.open "select title from [dbook$] where asin = ' & url & '", xlsConn, adOpenStatic, adLockReadOnly
     If rsOld.recordCount = 0 Then
     urlAll = urlAll & getAmazonTry(url) & vbCrLf
     i = i + 1
     Else
     n = n + 1
     End If
     rsOld.Close
      rs.MoveNext
      DoEvents
      Wend
 End If
end1:
If urlAll <> "" Then
      writeutf8 EXEPATH & "lost.txt", urlAll
      Label1.Caption = Label1.Caption & "缺少" & Trim(str(i)) & "本, 已经下载了" & Trim(str(n)) & "本"
 Else
      Label1.Caption = "已经检查完毕"
End If
rs.Close
csvConn.Close
xlsConn.Close
Set rs = Nothing
Set rsOld = Nothing
Set csvConn = Nothing
Set xlsConn = Nothing
End Sub

Private Sub Cmd_checkNewBook_Click()

Call checkNewBook0(txtcheck.Text)
'Call checkURL(EXEPATH & "template\urlist.txt", "2015-08-15", Now)
End Sub

Private Sub Cmd_exit_Click()
If bconn = True Then

f.关闭连接
Set f = Nothing
End If

Unload Me
End Sub

Private Sub cmd_FTP_Click()
'runFTP
runFTP2
End Sub

Private Sub cmd_go_Click()
Dim i As Integer

Call drill_url(txtURL.Text, Left(Label_web.Caption, 2), i)
Label1.Caption = Label1.Caption & " 发现链接" & i & "页"
End Sub

Private Sub Cmd_List_Click()
'If txtURL.Text <> "" Then
'txtURL.Text = transPre & URLEncode(txtURL.Text)
'End If
End Sub

Private Sub cmd_pan_Click()
Dim i As Integer
txtURL.Text = "https://pan.baidu.com/pcloud/home"
On Error Resume Next
'ShellExecute Me.hwnd, "open", txtURL.Text, "", "", 5
'WebBrowser1.Navigate2 txtURL.Text
'Call drill_url(txtURL.Text, Left(Label_web.Caption, 2), i)
Dim jsstr As String
Dim str2 As String
jsstr = readutf8(App.Path & "\pcloud_feedpage_all.js")
str2 = fromUnicode(jsstr)
writeutf8 App.Path & "\2.js", str2
i = checkURL_pan(App.Path & "\pan.txt")
Label1.Caption = Label1.Caption & " 发现网盘资源" & i & "个"
End Sub

Private Sub Cmd_run_Click()
'Call remove_series_1 '清理只有一本书的系列
'Call readxlsfile(0) '生成网页
'Call readxlsfile(2) '封面瀑布流,图片 生成index.htm
'Call readxlsfile(3) '全部列表,list1-n.htm
'Call generate_zh    '中文书列表,生成zh.htm
'Call generate_author '作者,放到author文件夹,最后汇总生成authorlist.htm
'Call generate_series '系列,放到series文件夹,然后生成serieslist.htm
'Call checkNewBook0(txtcheck.Text) '计算各个文库在售书,预售新书列表生成newbook.htm
'Call generate_wenku '文库,放到wenku文件夹,最后汇总生成wenkulist.htm
'Call create_SiteMaptxt '生成sitemap.txt
'Call readxlsfile(8) 'Sitemap,生成sitemap.xml
'Call generate_whole '全本列表,生成whole.htm
Call CmdCSV_Click '生成网页
Call Cmd_SP_Click '生成封面和分类
Call cmd_FTP_Click 'FTP上传
End Sub

Private Sub Cmd_Start_Click()
Dim str1 As String
Dim sjs As String
Dim sfile As String
Dim simg As String
Dim sname As String
Dim sContent As String
Dim shtm As String
Dim stitle As String
Dim i As Integer
Dim si As Integer
Dim bUp As Boolean
Dim iSuccess As Integer

Dim fso As Object
Set fso = CreateObject("Scripting.FileSystemObject")
' 如果目录不存在,就创建该目录
If fso.FolderExists(EXEPATH & "asin") = "False" Then fso.createfolder (EXEPATH & "novel")
Set fso = Nothing

Label1.Caption = ""

'获取路径下所有txt文件,其文件名就是title
For i = 0 To File1.ListCount - 1
    On Error Resume Next
    sname = File1.List(i)
    If sname = "sitemap.txt" Or sname = "ok.txt" Then GoTo err:
    'Label1.Caption = "正在处理总共" & File1.ListCount & "中的第" & i & "个文件: " & sname
    simg = Replace(sname, ".txt", ".jpg")
    sfile = readutf8(EXEPATH & sname)
    sfile = txtReplace(sfile, lv_free)
    Call remove_spaceline(sfile)
    stitle = Replace(sname, ".txt", "")
    shtm = EXEPATH & "novel\" & stitle & ".htm"
    '文件不存在
    If Dir(shtm) = "" Then
        txt1 = readutf8(EXEPATH & "template\header_bytitle.txt")
        sfile = txt1 + vbCrLf + sfile + vbCrLf + txt2
        sfile = Replace(sfile, "#title#", stitle)
        sfile = Replace(sfile, "#domain#", sDomain)
        sfile = Replace(sfile, "#modifytime#", Format$(Now, "yyyy-mm-dd hh:mm:ss"))
        Call writeutf8(shtm, sfile, "UTF-8")
        Call fileMove(EXEPATH & sname, EXEPATH & "0txt\" & stitle & ".txt")
        'Call fileMove(EXEPATH & simg, EXEPATH & "img\" & stitle & ".jpg")
        'bUp = fileUpload(EXEPATH & "0novel\", "/novel/", stitle & ".htm")
        'bUp = fileUpload(EXEPATH & "img\", "/novel/img/", stitle & ".jpg")
        iSuccess = iSuccess + 1
    End If
err:
Next i

Label1.Caption = Label1.Caption & "页面转换完成了！"
Call create_list("novel", 1)
Label1.Caption = Label1.Caption & " 成功处理好的网页总数:" & iSuccess
End Sub


Private Sub Form_Load()

Dim iniTxt As String
Dim maxi As String
iniTxt = readutf8(App.Path & "\" & "ini.txt")
If iniTxt <> "" Then
  kindle_1 = Fetch(iniTxt, "kindle1=[", "]")
  kindle_2 = Fetch(iniTxt, "kindle2=[", "]")
  txt_kindle.Text = kindle_1
  Txt_path.Text = Fetch(iniTxt, "path=[", "]")
  sDomain = Fetch(iniTxt, "domain=[", "]")
  maxItem = CInt(Fetch(iniTxt, "maxItem=[", "]"))
  filterflag = Fetch(iniTxt, "filterflag=[", "]")
  txtcheck.Text = Fetch(iniTxt, "checkurl=[", "]")
  txtbuy.Text = Fetch(iniTxt, "buyurl=[", "]")
  cstyle = Fetch(iniTxt, "style=[", "]")
  Txt_style.Text = cstyle
  imgDomain = Fetch(iniTxt, "imgdomain=[", "]")
  imgDomai2 = Fetch(iniTxt, "imgdomai2=[", "]")
  transPre = Fetch(iniTxt, "transPre=[", "]")
  transPr2 = Fetch(iniTxt, "transPr2=[", "]")
  proxyAddress = Fetch(iniTxt, "proxy=[", "]")
  titleget1 = Fetch(iniTxt, "pretitle=[", "]")
  titleget2 = Fetch(iniTxt, "posttitle=[", "]")
  urlget1 = Fetch(iniTxt, "preurl=[", "]")
  urlget2 = Fetch(iniTxt, "posturl=[", "]")
  totalnum1 = Fetch(iniTxt, "totalnum1=[", "]")
  totalnum2 = Fetch(iniTxt, "totalnum2=[", "]")
  totalnum3 = Fetch(iniTxt, "totalnum3=[", "]")
  totalnum4 = Fetch(iniTxt, "totalnum4=[", "]")
Else
  Txt_path.Text = App.Path & "\"
  txtcheck.Text = App.Path & "\template\newbookurl.txt"
  txtbuy.Text = App.Path & "\template\buylist.txt"
  sDomain = "/"
  imgDomain = "/"
  maxItem = 50
  filterflag = "1"
  cstyle = "少女向"
  Txt_style.Text = cstyle
  titleget1 = "<a class=""a-link-normal s-access-detail-page  s-color-twister-title-link a-text-normal"" target=""_blank"" title="""
  titleget2 = """"
  urlget1 = "ebook/dp/"
  urlget2 = "/ref="
  totalnum1 = "検索結果 "
  totalnum2 = "のうち"
  totalnum3 = "a-size-base a-spacing-small a-spacing-top-small a-text-normal"">"
  totalnum4 = "件の結果"
End If

Label1.Caption = Format$(Now, "yyyy-mm-dd")


If Txt_path.Text = "" Or Dir(Txt_path.Text) = "" Then Txt_path.Text = App.Path & "\"
EXEPATH = Txt_path.Text
File1.Path = EXEPATH
If Dir(EXEPATH & "template\header.txt") <> "" Then
txt1 = readutf8(EXEPATH & "template\header.txt")
End If

If Dir(EXEPATH & "template\header.txt") <> "" Then
txt2 = readutf8(EXEPATH & "template\footer.txt")
End If

listTxt2 = readutf8(EXEPATH & "template\grouplist2.txt")
If Dir(EXEPATH & "template\free.txt") <> "" Then
lv_free = readutf8(EXEPATH & "template\free.txt")
End If

Me.Caption = Me.Caption & "    一天一个好心情！"
If Dir(EXEPATH & "template\log.txt") <> "" Then
  Txt_date.Text = Format$(readutf8(EXEPATH & "template\log.txt"), "yyyy-mm-dd")
Else
  Txt_date.Text = Format$(Now, "yyyy-mm-dd")
End If

txt3 = readutf8(EXEPATH & "template\grouplist_middle.txt")

WebBrowser1.Navigate "about:blank"
xlsConnString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & EXEPATH & txtdb.Text & ";Extended Properties='Excel 8.0;HDR=Yes;IMEX=1'"

End Sub

Private Sub Form_Unload(Cancel As Integer)
'xlsConn.Close
'Set xlsConn = Nothing
End Sub

Private Sub Label4_Click()
If txt_kindle.Text = kindle_1 Then
  txt_kindle.Text = kindle_2
ElseIf txt_kindle.Text = kindle_2 Then
  txt_kindle.Text = kindle_1
End If
End Sub



Private Sub Txt_date_Change()
Call writeutf8(EXEPATH & "template\log.txt", Format$(Txt_date.Text, "yyyy-mm-dd"), "UTF-8")
End Sub

Private Sub Txt_path_Change()
If Right(Txt_path.Text, 1) = "\" Then
'EXEPATH = Txt_path.Text
File1.Path = Txt_path.Text
End If
End Sub


Public Function getHtmlStr(strURL As String, Optional timeout As Long) As String
Dim stime, ntime
Dim XmlHttp 'As MSXML2.XMLHTTP60
If strURL = "" Or strURL = vbCrLf Or Len(strURL) < 2 Then Exit Function
Set XmlHttp = CreateObject("Microsoft.XMLHTTP")
XmlHttp.open "GET", strURL, False
XmlHttp.setRequestHeader "If-Modified-Since", "0"
On Error GoTo Err_net
stime = Now '获取当前时间
XmlHttp.send
While XmlHttp.ReadyState <> 4
  DoEvents
  ntime = Now '获取循环时间
  If timeout <> 0 Then
    If DateDiff("s", stime, ntime) > timeout Then
      getHtmlStr = ""
      Debug.Print "timeout:" & strURL & vbCrLf
      Exit Function '判断超出3秒即超时退出过程
    End If
  End If
Wend
getHtmlStr = BytesToBstr(XmlHttp.responseBody, "UTF-8")
Set XmlHttp = Nothing
Err_net:
End Function

Public Function sendHtmlStr(strURL As String) As String
Dim XmlHttp 'As MSXML2.XmlHttp
If strURL = "" Or strURL = vbCrLf Then Exit Function
Set XmlHttp = CreateObject("Microsoft.XMLHTTP")
XmlHttp.open "GET", strURL, False
On Error GoTo Err_net
XmlHttp.send
While XmlHttp.ReadyState <> 4
DoEvents
Wend
sendHtmlStr = XmlHttp.StatusText

Set XmlHttp = Nothing
Err_net:
End Function

Public Function getHtmlStr_Async(strURL As String, Optional Sync As Boolean = False, Optional getHtmBody As Boolean = True) As String
'Dim XmlHttp As MSXML2.ServerXmlHttp
If strURL = "" Or strURL = vbCrLf Then Exit Function
Set ServerXmlHttp = CreateObject("MSXML2.ServerXMLHTTP")
ServerXmlHttp.open "GET", strURL, Sync
On Error GoTo Err_net
ServerXmlHttp.send
If Sync Then '异步 用timer来实现
Timer1.Enabled = Sync
getHtmlStr_Async = ""
Else '同步,直接获得结果网页
  While ServerXmlHttp.ReadyState <> 4
    DoEvents
  Wend
  If getHtmBody Then
    getHtmlStr_Async = BytesToBstr(ServerXmlHttp.responseBody, "UTF-8")
  Else
    getHtmlStr_Async = ServerXmlHttp.StatusText
  End If
End If
Set ServerXmlHttp = Nothing
Err_net:
End Function
Private Sub Timer1_Timer()
    If ServerXmlHttp.ReadyState = 4 Then
        Timer1.Enabled = False
        If ServerXmlHttp.Status = 200 Then
          urlOk = urlOk + 1
          Label1.Caption = Label1.Caption & "已经成功访问" & urlOk & "个网页 "
        End If
    End If
End Sub


Private Function BytesToBstr(strBody As String, Optional codeBase As String = "UTF-8") As String
Dim objStream As Object
Set objStream = CreateObject("Adodb.Stream")
objStream.Charset = codeBase
objStream.Type = 1
objStream.mode = 3
objStream.open
objStream.write strBody
objStream.position = 0
objStream.Type = 2

BytesToBstr = objStream.ReadText
objStream.Close
Set objStream = Nothing

End Function

Private Function readutf8(fileName As String, Optional codeBase As String = "UTF-8") As String
'Object.Open(Source,[Mode],[Options],[UserName],[Password])
'Mode 指定打开模式，可不指定，可选参数如下：
 ' adModeRead = 1
  'adModeReadWrite = 3
  'adModeRecursive = 4194304
  'adModeShareDenyNone = 16
  'adModeShareDenyRead = 4
  'adModeShareDenyWrite = 8
  'adModeShareExclusive = 12
  'adModeUnknown = 0
  'adModeWrite = 2
' Options 指定打开的选项，可不指定，可选参数如下：
 ' adOpenStreamAsync = 1
  'adOpenStreamFromRecord = 4
  'adOpenStreamUnspecified = -1
 'UserName 指定用户名，可不指定。
 'Password 指定用户名的密码

Const adTypeText = 2
Dim objStream As Object
Set objStream = CreateObject("Adodb.Stream")
objStream.Charset = codeBase
If Dir(fileName) = "" Then Exit Function
objStream.open
objStream.position = 0
objStream.Type = adTypeText
On Error GoTo line_error
objStream.LoadFromFile fileName
readutf8 = objStream.ReadText

line_error:
objStream.Close
Set objStream = Nothing
End Function

Function writeutf8(filepath As String, str As String, Optional codeBase As String = "UTF-8", Optional compare As Integer = 0)

Const adSaveCreateOverWrite = 2
Const adSaveCreateNotExist = 1
On Error Resume Next
Dim objStream As Object
Dim filefoler As String
Dim fso As Object
Dim newFolder As Object
Dim sfolder As String
Dim oldFile As String
Set fso = CreateObject("Scripting.FileSystemObject")
' 如果目录不存在,就创建该目录
'If fso.FolderExists(filepath) = "" Then fso.createfolder ("filepath")


          
Set objStream = CreateObject("ADODB.Stream")
objStream.Charset = codeBase
objStream.open
objStream.WriteText str
If compare = 0 Then
  If Dir(filepath) <> "" Then fso.DeleteFile filepath
  objStream.SaveToFile filepath, adSaveCreateOverWrite
ElseIf compare = 1 Then
  oldFile = readutf8(filepath)
  If str <> oldFile And str <> "" Then
    objStream.SaveToFile filepath, adSaveCreateOverWrite
  End If
End If

objStream.Close
Set objStream = Nothing
End Function
    
Sub fileMove(strFilename, newpath)

Dim fso As Object
Dim f As Object
Set fso = CreateObject("Scripting.FileSystemObject")
If Not fso.FileExists(strFilename) Then Exit Sub
If fso.FileExists(newpath) Then
fso.DeleteFile newpath
End If
Set f = fso.GetFile(strFilename)
fso.MoveFile strFilename, newpath
Set fso = Nothing

End Sub
Sub fileDelete(strFilename)

Dim fso As Object

Set fso = CreateObject("Scripting.FileSystemObject")
If fso.FileExists(strFilename) Then
fso.DeleteFile strFilename
End If

Set fso = Nothing

End Sub
Sub fileCopy(strFilename, newpath)

Dim fso As Object
Dim f As Object
Set fso = CreateObject("Scripting.FileSystemObject")
If Not fso.FileExists(strFilename) Then Exit Sub
If fso.FileExists(newpath) Then
fso.DeleteFile newpath
End If
Set f = fso.GetFile(strFilename)
fso.copyFile strFilename, newpath
Set fso = Nothing

End Sub


Private Sub create_list(sType As String, Optional iFanYi As Integer = 0)
File1.Pattern = "*.htm"
File1.Path = EXEPATH & sType & "\"
File1.Refresh
Dim i As Integer
Dim sname As String
Dim sline As String
Dim stitle As String
Dim sfile As String
Dim sindex As String
Dim url As String
Dim domainEncode As String
domainEncode = urlEncode(sDomain)
'获取路径下所有htm文件,其文件名就是title
For i = 0 To File1.ListCount - 1
sindex = i + 1
sname = File1.List(i)
stitle = Replace(sname, ".htm", "")
sline = "<tr>"
If iFanYi = 0 Then
sline = sline & "<td>" & sindex & "</td>"
ElseIf iFanYi = 1 Then
url = transPr2 & urlEncode(sDomain & sType & "/" & sname)
sline = sline & "<td><a href=""" & url & """>译文</a></td>"
ElseIf iFanYi = 2 Then
url = sDomain & "author/" & sname & ".htm"
sline = sline & "<td><a href=""" & url & """>sindex</a></td>"
End If

sline = sline & "<td><a href=""" & sType & "/" & stitle & ".htm"">" & stitle & "</a></td>"
sline = sline & "</tr>"
sfile = sfile & sline
Next i
listTxt1 = readutf8(EXEPATH & "template\grouplist_byfolder.txt")
listTxt1 = Replace(listTxt1, "#count#", i)
listTxt1 = Replace(listTxt1, "#modifytime#", Format$(Now, "yyyy-mm-dd hh:mm:ss"))
sfile = listTxt1 + sfile + listTxt2
Call writeutf8(EXEPATH & "list\" & sType & "list.htm", sfile, "UTF-8")
Label1.Caption = Label1.Caption & " " & sType & "列表生成好了,共处理" & str(File1.ListCount) & "个"
End Sub


Private Sub remove_spaceline(str As String)
'删除空行
Dim arr() As String
Dim i As Integer
Dim ss As String
arr = Split(str, vbCrLf)
For i = 0 To UBound(arr)
    If Trim(Replace(arr(i), "　", "")) <> "" Then ss = ss & vbCrLf & arr(i)
Next
str = Mid(ss, 3)
End Sub

Private Function txtReplace(strfile As String, sfree As String) As String
Dim lt_free() As String
Dim i As Long
Dim posEnd As Long
On Error GoTo err
lt_free = Split(sfree, vbCrLf)
For i = LBound(lt_free) To UBound(lt_free)
    lt_free(i) = Trim(lt_free(i))
    If lt_free(i) <> "" Then
      strfile = Replace(strfile, lt_free(i), "")
    End If
Next i

strfile = Replace(strfile, "Table of Contents", "目次")
strfile = Replace(strfile, "contents", "目次")
strfile = Replace(strfile, "prologue", "序章")
strfile = Replace(strfile, "プロローグ", "序章")
strfile = Replace(strfile, "epilogue", "終章")
strfile = Replace(strfile, "エピローグ", "終章")
posEnd = InStr(strfile, "無料サンプルはお楽しみいただけましたか")
If posEnd > 1 Then
posEnd = posEnd - 1
strfile = Left(strfile, posEnd)
End If
err:
txtReplace = strfile
End Function

Private Function fileUpload(localPath As String, remotePath As String, sfile As String) As Boolean
If bconn = False Then
Set f = New FTP
bconn = f.连接服务器("qxw1147100127.my3w.com", 21, "qxw1147100127", "P68243267")
End If
If bconn = True Then
fileUpload = f.上传文件(localPath & sfile, remotePath & sfile)
End If
End Function

Private Function imgUpload(localPath As String, remotePath As String, sfile As String) As Boolean
Dim lv_remotePath As String
If cstyle = "少年向" Then
lv_remotePath = remotePath
Else
lv_remotePath = "\2" & remotePath
End If
If bconnimg = False Then
Set fimg = New FTP
bconnimg = fimg.连接服务器("qxu1606410412.my3w.com", 21, "qxu1606410412", "QBiPmLdHfsfs")
End If
If bconnimg = True Then
imgUpload = fimg.上传文件(localPath & sfile, lv_remotePath & sfile)
End If
End Function


Private Function gOpenUrl(xxURl As String, xx As String) As Boolean

End Function

Private Function Fetch(LinkText As String, s1 As String, s2 As String) As String
On Error GoTo err2
Dim LinkStart As Double '从一个字符串的两个标记中间截取文字
Dim LinkEnd As Double
Dim TempVar As String
If InStr(1, LinkText, s1) > 0 And InStr(1, LinkText, s2) > 0 Then
LinkStart = InStr(1, LinkText, s1)
LinkText = Mid$(LinkText, LinkStart + Len(s1))
LinkEnd = InStr(1, LinkText, s2)
If LinkEnd <= 1 Then GoTo err2
    TempVar = Mid$(LinkText, 1, LinkEnd - 1)
    Fetch = Trim$(TempVar)
    TempVar = ""
    'LinkText = Mid$(LinkText, LinkEnd + Len(s2))
err2:
Else:
    Fetch = ""
End If
End Function


Private Function strSub(sourceText As String, s1 As String, i As Integer) As String
On Error GoTo err2
Dim sourceStart As Integer
Dim sourceEnd As Integer
Dim TempVar As String
If InStr(1, sourceText, s1) > 0 Then
sourceStart = InStr(1, sourceText, s1)
strSub = Mid$(sourceText, sourceStart + Len(s1), i)
err2:   Else: strSub = ""
End If
End Function

Private Function ListtxtToHtm(tpPath As String, encode As String, strStart As String, strEnd As String)
Dim str1 As String
Dim sjs As String
Dim sfile As String
Dim sTemplate As String
Dim sTemplate_head As String
Dim sTemplate_foot As String
Dim sname As String
Dim shtm As String
Dim stitle As String
Dim i As Integer
Dim iSuccess As Integer
Dim posBeg As Integer
Dim posEnd As Integer

Label1.Caption = ""
sTemplate = readutf8(EXEPATH & "template\header_fromtxt.txt")
posBeg = InStr(sTemplate, "<!--begin_txt-->")
If posBeg > 0 Then
posBeg = posBeg + 16
sTemplate_head = Left(sTemplate_head, posBeg)
'sTemplate_foot = Replace(sTemplate, sTemplate_head, "")
End If

posEnd = InStr(sTemplate, "<!--end_txt-->")
If posEnd > 0 And posEnd < Len(sTemplate) Then
sTemplate_foot = Right(sTemplate, Len(sTemplate) - posEnd)
End If

'获取路径下所有txt文件,其文件名就是title
For i = 0 To File1.ListCount - 1
    On Error Resume Next
    sname = File1.List(i)
    'Label1.Caption = "正在处理总共" & File1.ListCount & "中的第" & i & "个文件: " & sname
    sfile = readutf8(EXEPATH & sname)
    sfile = txtReplace(sfile, lv_free)
    Call remove_spaceline(sfile)
    stitle = Replace(sname, ".txt", "")
    shtm = EXEPATH & "htm\" & stitle & ".htm"
'文件不存在
If Dir(shtm) = "" Then
sfile = sTemplate_head + vbCrLf + sfile + vbCrLf + sTemplate_foot
sfile = Replace(sfile, "#title#", stitle)

Call writeutf8(shtm, sfile, "UTF-8")
Call fileMove(EXEPATH & sname, EXEPATH & "0txt\" & stitle & ".txt")
iSuccess = iSuccess + 1
End If
Next i

Label1.Caption = "页面转换完成了！"
'Call create_list("htm",1)
Label1.Caption = Label1.Caption & " 成功处理好的网页总数:" & iSuccess
End Function

Private Function txtToHtm(sTxt As String, tpPath As String, encode As String, strStart As String, strEnd As String)

Dim sTemplate As String
Dim sTemplate_head As String
Dim sTemplate_foot As String
Dim iSuccess As Integer
Dim posBeg As Integer
Dim posEnd As Integer

On Error Resume Next
sTemplate = readutf8(tpPath, encode)
If sTemplate <> "" Then
   posBeg = InStr(sTemplate, strStart)
    If posBeg > 0 Then
    posBeg = posBeg + 16
    sTemplate_head = Left(sTemplate, posBeg)
    sTemplate_foot = Replace(sTemplate, sTemplate_head, "")
    End If
Else
sTemplate_head = txt1
sTemplate_foot = txt2
End If

'posEnd = InStr(sTemplate, strEnd)
'If posEnd > 0 And posEnd > posbeg Then
'sTemplate_foot = Right(sTemplate, Len(sTemplate) - posEnd)
'Else
'sTemplate_foot = Text2.Text
'End If


txtToHtm = sTemplate_head + vbCrLf + sTxt + vbCrLf + sTemplate_foot
End Function

Private Function htmToTxt(shtm As String)
htmToTxt = Fetch(shtm, "<!--begin_txt-->", "<!--end_txt-->")
End Function


Private Sub Txt_style_Change()
cstyle = Txt_style.Text
End Sub

Private Sub txtcheck_DblClick()
ReDim URLTab(0)
End Sub

Private Sub txtdb_Change()
xlsConnString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & EXEPATH & txtdb.Text & ";Extended Properties='Excel 8.0;HDR=Yes;IMEX=1'"
End Sub

Private Sub txtURL_DblClick()
Label_web.Caption = ""
If txtURL.Text <> "" Then
  If Left(txtURL.Text, 4) <> "http" Then txtURL.Text = "http://" & txtURL.Text
  WebBrowser1.Navigate2 txtURL.Text
End If
End Sub

Private Sub WebBrowser1_DocumentComplete(ByVal pDisp As Object, url As Variant)
'防止多iframe多次执行DocumentComplete事件
'If Not (pDisp Is WebBrowser1.Object) Then Exit Sub
'If (pDisp Is WebBrowser1.Application) Then ' status = READYSTATE_COMPLETE
'    inComplete = 1
'    inInitial = inInitial + 1
'    txtURL.Text = URL
'End If
'Dim wdoc As HTMLDocument
'Set wdoc = WebBrowser1.Document
'Label_web.Caption = "完成了"
End Sub

Private Sub WebBrowser1_NavigateComplete2(ByVal pDisp As Object, url As Variant)
'Dim obj As HTMLDocument
'Dim objwin As WebBrowser
'Set objwin = pDisp
'objwin.Silent = True
'Set obj = objwin.Document
'obj.parentWindow.execScript "function showModalDialog(){return;}" '对showModalDialog引起的对话框进行确定
'Text1.Text = WebBrowser1.Document.documentElement.outerHTML
Dim wdoc As HTMLDocument
Set wdoc = WebBrowser1.Document
'Call writeutf8(EXEPATH & "cookie.txt", wdoc.cookie)
txtURL.Text = wdoc.URLUnencoded
Label_web.Caption = wdoc.title
Label_web.Caption = Replace(Label_web.Caption, "Amazon.co.jp", "")
Label_web.Caption = Replace(Label_web.Caption, "amazon.co.jp", "")
Label_web.Caption = Replace(Label_web.Caption, ":", " ")
Label_web.Caption = Replace(Label_web.Caption, "|", " ")
Label_web.Caption = Replace(Label_web.Caption, "?", " ")
Label_web.Caption = Trim(Label_web.Caption)
End Sub

Public Sub wait1000(HaoMiao As Long)
Dim t1 As Long
t1 = timeGetTime
While (timeGetTime - t1) < HaoMiao
DoEvents
Wend
End Sub
 
Private Sub cmd_title_Click()
'Call readCSVfile(1)
'Call readxlsfile(1)
Call create_list("novel", 1)

End Sub

Private Sub CmdCSV_Click()
'Call readCSVfile(0)
'Call readCSVfile(2) '封面瀑布流,图片 生成index.htm
Call remove_series_1 '清理只有一本书的系列
Call readxlsfile(0) '生成网页
End Sub

Private Sub Cmd_SP_Click()
Call readxlsfile(2) '封面瀑布流,图片 生成index.htm
'csv格式
'Call checkNewBook0   '计算各个文库在售书,预售新书列表生成newbook.htm
'Call readCSVfile(3) '文库,放到wenku文件夹,最后汇总生成wenkulist.htm
''Call readCSVfile(4) '系列,放到series文件夹,然后生成serieslist.htm
'Call readCSVfile(5) '全部列表,list1-n.htm
'Call readCSVfile(6) '中文书列表,生成zh.htm
''Call readCSVfile(7) '作者,放到author文件夹,最后汇总生成authorlist.htm
'generate_author
'generate_series
'Call create_SiteMaptxt '生成sitemap.txt
'Call readCSVfile(8) 'Sitemap,生成sitemap.xml

'xls格式
Call readxlsfile(3) '全部列表,list1-n.htm
Call generate_zh    '中文书列表,生成zh.htm
Call generate_author '作者,放到author文件夹,最后汇总生成authorlist.htm
Call generate_series '系列,放到series文件夹,然后生成serieslist.htm
Call checkNewBook0(txtcheck.Text) '计算各个文库在售书,预售新书列表生成newbook.htm
Call generate_wenku '文库,放到wenku文件夹,最后汇总生成wenkulist.htm
Call create_SiteMaptxt '生成sitemap.txt
Call readxlsfile(8) 'Sitemap,生成sitemap.xml
Call generate_whole '全本列表,生成whole.htm
End Sub

Private Sub Cmd_wenku_Click()
'Call readxlsfile(3) '全部列表,list1-n.htm
'Call checkNewBook0(txtcheck.Text)
'generate_wenku
'generate_zh
'generate_series
'generate_author
'Call readxlsfile(2)
'Call create_SiteMaptxt '生成sitemap.txt
'Call create_SiteMaptxt '生成sitemap.txt
'Call readxlsfile(8) 'Sitemap,生成sitemap.xml
'Call generate_whole '全本列表,生成whole.htm
Call readxlsfile(2)
End Sub

Private Function createHTM(rs As ADODB.Recordset, sType As Integer, Optional isNew As Integer = 0) As Integer

Dim sfile As String
Dim sModify As String
Dim shtm As String
Dim snewTxt As String
Dim sMetaData As String

Dim sNow As String
Dim iModify As Integer

Dim bookItem As Books
createHTM = 0
isNew = 0
bookItem = getBook(rs, sType, sMetaData)
If Dir(bookItem.txtPath) = "" Then 'Calibre folder isn't exit
  bookItem.txtPath = EXEPATH & "txt\" & bookItem.asin & ".txt"
  If Dir(bookItem.txtPath) = "" Then
    bookItem.txtPath = EXEPATH & "0txt\" & bookItem.filetitle & "_" & bookItem.asin & ".txt"
  End If
End If

sfile = readutf8(bookItem.txtPath)
sfile = txtReplace(sfile, lv_free)
Call remove_spaceline(sfile)
If sfile = "" Then
  sok = sok & bookItem.filetitle & "_" & bookItem.asin & " :read empty" & vbCrLf
End If
'If sfile = "" Then Exit Function "空书也要生成
If cstyle = "乙女向" Then
   If bookItem.cnseries = "" Then
     If InStr(bookItem.tagMain, "BL") > 0 Then
       If Len(sfile) > 500 Then sfile = Mid(sfile, 1, 500)
     ElseIf InStr(bookItem.tagMain, "TL") > 0 Then
       If Len(sfile) > 500 Then sfile = Mid(sfile, 1, 500)
     ElseIf Len(sfile) > 1000 Then
       sfile = Mid(sfile, 1, 1000)
     End If
  ElseIf bookItem.languages = "日文" Then
     If Len(sfile) > 2000 Then sfile = Mid(sfile, 1, 2000)
  End If
ElseIf cstyle = "少年向" Then
  sfile = Mid(sfile, 1, 1000)
End If
shtm = EXEPATH & bookItem.typePath & "\" & bookItem.newFileName & ".htm"
sModify = Mid(bookItem.lastModify, 1, 10)
sModify = Replace(sModify, ".", "-")
sNow = Format$(Now, "yyyy-mm-dd")
iModify = DateDiff("d", CDate(sModify), CDate(sNow), vbMonday, vbFirstFourDays)

On Error GoTo err
'If (iModify <= 1 Or Dir(shtm) = "") And sfile <> "" Then
If (iModify <= 1 Or Dir(shtm) = "") Then '限于保存今天或者昨天修改的, 或者不存在的新文件
'If (iModify <= 0 Or Dir(shtm) = "") Then '限于保存今天修改的, 或者不存在的新文件
'If (Dir(shtm) = "") Then                 '限于保存文件不存在的
    snewTxt = sfile
    sMetaData = Replace(sMetaData, "'", "")
    sMetaData = Replace(sMetaData, """", "")
    sMetaData = Replace(sMetaData, ";", ",")
    sfile = txt1 + vbCrLf + sfile + vbCrLf + txt2
    sfile = Replace(sfile, "#title#", bookItem.titleFiltered)
    sfile = Replace(sfile, "#metadata#", sMetaData)
    sfile = Replace(sfile, "#asin#", bookItem.asin)
    sfile = Replace(sfile, "#filepath#", bookItem.txtPath)
    sfile = Replace(sfile, "#filetitle#", bookItem.filetitle)
    sfile = Replace(sfile, "#domain#", sDomain)
    'If bookItem.id > 8206 And cstyle = "乙女向" Then
    'sfile = Replace(sfile, "#imgdomain#", imgDomai2)
    'Else
    sfile = Replace(sfile, "#imgdomain#", imgDomain)
    'End If
    sfile = Replace(sfile, "#domainEncode#", urlEncode(sDomain))
    sfile = Replace(sfile, "#authors#", bookItem.authors)
    sfile = Replace(sfile, "#author1url#", bookItem.author1url)
    sfile = Replace(sfile, "#series#", bookItem.Series)
    sfile = Replace(sfile, "#seriesurl#", bookItem.seriesURL)
    sfile = Replace(sfile, "#series_index#", bookItem.series_index)
    sfile = Replace(sfile, "#wenku#", bookItem.wenku)
    sfile = Replace(sfile, "#cntitle#", bookItem.cntitle)
    sfile = Replace(sfile, "#kindlePrice#", bookItem.kindlePrice)
    sfile = Replace(sfile, "#paperPrice#", bookItem.paperPrice)
    sfile = Replace(sfile, "#publisher#", bookItem.publisher)
    sfile = Replace(sfile, "#pubdate#", bookItem.pubdate)
    sfile = Replace(sfile, "#language#", bookItem.languages)
    sfile = Replace(sfile, "#rating#", bookItem.rating)
    sfile = Replace(sfile, "#tags#", bookItem.tags)
    sfile = Replace(sfile, "#pages#", bookItem.pages)
    sfile = Replace(sfile, "#comments#", bookItem.comments)
    sfile = Replace(sfile, "#modifytime#", Format$(Now, "yyyy-mm-dd hh:mm:ss"))
    sfile = Replace(sfile, "#authorsurl#", bookItem.authorsURL)
    sfile = Replace(sfile, "#tagMain#", bookItem.tagMain)
    sfile = Replace(sfile, "#langA#", bookItem.langA)
    sfile = Replace(sfile, "#buyurlA#", bookItem.buyurlA)
    sfile = Replace(sfile, "#translateURL#", bookItem.translateURL)
    sfile = Replace(sfile, "#seriesindexURLA#", bookItem.seriesindexURLA)
    If bookItem.seriesindexURLA <> "" Then
     sfile = Replace(sfile, "#label2#", "丛书:")
    Else
     sfile = Replace(sfile, "#label2#", "")
    End If
    
'    sok = sok & bookItem.asin & " 【" & bookItem.filetitle & "】:open" & vbCrLf
    Call writeutf8(shtm, sfile, "UTF-8")
    Call fileCopy(bookItem.cover, EXEPATH & "240_un\backup\" & bookItem.asin & ".jpg")
    '@@Call fileCopy(bookItem.cover, EXEPATH & "img\" & bookItem.asin & ".jpg")
    '@@Call writeutf8(EXEPATH & "0txt\" & bookItem.filetitle & "_" & bookItem.asin & ".txt", snewTxt)
    createHTM = 1
    isNew = 1
Else
    isNew = -1
'    sok = sok & bookItem.asin & " 【" & bookItem.filetitle & "】:duplicate" & vbCrLf
End If
Exit Function

err:
isNew = -2
sok = sok & bookItem.asin & " 【" & bookItem.filetitle & "】:error" & vbCrLf
End Function

Private Function createHTM_old(rs As ADODB.Recordset, sType As Integer, Optional isNew As Integer = 0) As Integer

Dim sfile As String
Dim simg As String
Dim sname As String, newName As String, stitle As String, snewFileName As String
Dim sid As String
Dim sContent As String
Dim shtm As String
Dim i As Integer
Dim sMetaData As String
Dim asin As String
Dim Series As String
Dim series_index As String
Dim authors As String
Dim author() As String
Dim publisher As String
Dim stime As String
Dim tags As String
Dim tag() As String
Dim sYesno As String
Dim shortTitle As String
Dim wenku As String
Dim pubdate As String
Dim filetitle As String
Dim sPath As String
Dim sPat As String
Dim snewTxt As String
Dim stypePath As String
Dim id As String
Dim sRating As String
Dim sLanguage As String
Dim sWords As String
Dim sPages As String
Dim sCNSeries As String
Dim sCNSries_index As String
Dim sCNTitle As String
Dim sComments As String
Dim sKindlePrice As String
Dim sPaperPrice As String
Dim sModify As String
Dim sNow As String
Dim iModify As Integer
Dim author1url As String
Dim author1 As String
Dim author2 As String
Dim seriesURL As String
Dim lastModify As String
Dim transURL As String

createHTM_old = 0
isNew = 0

sMetaData = ""
For i = 0 To (rs.fields.Count - 1)
On Error Resume Next
Select Case rs.fields(i).Name
Case "id"
sid = rs.fields(i).Value
'sok = sok & "{" & rs.fields(i).Value & "}:start"
Case "series"
    Series = rs.fields(i).Value
Case "series_index"
    If Series <> "" Then
    series_index = rs.fields(i).Value
    Call removeMark(Series)
    seriesURL = "丛书:<a href=""series/" & Series & ".htm"">" & Series & "(" & series_index & ")</a>"
    sMetaData = sMetaData & "," & Series & rs.fields(i).Value
    End If
Case "identifiers" 'like amazon_jp:B00HWLJLAK,mobi-asin:B00HWLJLAK
    asin = strSub(rs.fields(i).Value, "mobi-asin:", 10)
    If asin = "" Then
    asin = strSub(rs.fields(i).Value, "amazon_jp:", 10)
    End If
    If sType = 0 Then 'use asin as newfilename
        snewFileName = asin
        stypePath = "asin"
    End If
    sMetaData = sMetaData & "," & asin
Case "title"
  stitle = rs.fields(i).Value
  stitle = Replace(stitle, "?", " ")
  Call removeMark(stitle)
  sMetaData = sMetaData & "," & stitle
Case "authors"
    authors = rs.fields(i).Value
    Call getAuthorDetail(authors, author1, author2, author1url)
    Call removeMark(author1)
    author1url = "作者:<a href=""" & author1url & """>" & author1 & "</a>"
    Call removeMark(authors)
    sMetaData = sMetaData & "," & authors
Case "publisher"
publisher = rs.fields(i).Value
sMetaData = sMetaData & "," & publisher
Case "pubdate"
pubdate = Mid(rs.fields(i).Value, 1, 10)
Case "size"
Case "cover"
    simg = rs.fields(i).Value
    sPath = Replace(simg, "cover.jpg", "")
    sPat = sPath & "*.azw"
    sname = Dir(sPat)
    If sname <> "" Then
    sname = sPath & Replace(sname, ".azw", ".txt")
    Else
     sPat = sPath & "*.azw3"
     sname = Dir(sPat)
     sname = sPath & Replace(sname, ".azw3", ".txt")
    End If
    If Right(sname, 1) = "3" Then
    sname = Mid$(sname, 1, (Len(sname) - 1))
    End If
    If Dir(sname) = "" Then
    sname = "empty.txt"
    End If
Case "timestamp"
stime = rs.fields(i).Value
Case "tags"
tags = rs.fields(i).Value
tags = Replace(tags, "R18", "")
Call removeMark(tags)
sMetaData = sMetaData & "," & tags
Case "formats"
Case "rating"
sRating = rs.fields(i).Value
Case "languages"
sLanguage = rs.fields(i).Value
Case "#words"
sWords = rs.fields(i).Value
Case "#pages"
sPages = rs.fields(i).Value
Case "#yesno"
sYesno = LCase(rs.fields(i).Value)
Case "#cnseries"
sCNSeries = rs.fields(i).Value
Case "#cnseries_index"
sCNSries_index = rs.fields(i).Value
Case "#cntitle"
sCNTitle = rs.fields(i).Value
Case "#vtitle"
shortTitle = rs.fields(i).Value

Case "#wenku"
wenku = rs.fields(i).Value
sMetaData = sMetaData & "," & wenku
Case "wenku"
wenku = rs.fields(i).Value
sMetaData = sMetaData & "," & wenku
Case "#sorttitle"
Case "#filetitle"
filetitle = rs.fields(i).Value
If sType = 1 Then '以书名作为文件名
stypePath = "htm"
snewFileName = filetitle
End If
Case "#kindleprice"
sKindlePrice = rs.fields(i).Value
If sKindlePrice <> "" Then
sKindlePrice = "￥" & sKindlePrice
End If
Case "#paperprice"
sPaperPrice = rs.fields(i).Value
If sPaperPrice <> "" Then
sPaperPrice = "￥" & sPaperPrice
End If
Case "comments"
sComments = rs.fields(i).Value
Case "#lastmodify"
lastModify = Format$(rs.fields(i).Value, "yyyy-mm-dd hh:mm:ss")
End Select
Next i
    
publisher = getPubfromWenku(wenku, stitle)
sfile = readutf8(sname, "UTF-8")
sfile = txtReplace(sfile, lv_free)
Call remove_spaceline(sfile)
If sType = 1 Then
shtm = EXEPATH & "htm\" & snewFileName & ".htm"
transURL = sDomain & "htm/" & snewFileName & ".htm"
Else
shtm = EXEPATH & "asin\" & snewFileName & ".htm"
transURL = sDomain & "asin/" & snewFileName & ".htm"
End If


'If sfile = "" Then Exit Function
If sLanguage <> "zho" Then
If InStr(tags, "全本") > 0 Then
transURL = transPr2 & urlEncode(transURL)
Else
transURL = transPre & urlEncode(transURL)
End If
Else
transURL = sDomain
End If

sModify = Mid(lastModify, 1, 10)
sNow = Format$(Now, "yyyy-mm-dd")
iModify = DateDiff("d", CDate(sModify), CDate(sNow), vbMonday, vbFirstFourDays)
'限于保存今天或者昨天修改的, 或者文件不存在的
On Error GoTo err
'If (iModify <= 1 Or Dir(shtm) = "") And sfile <> "" Then
If (iModify <= 1 Or Dir(shtm) = "") Then
'If (iModify <= 0 Or Dir(shtm) = "") Then
    snewTxt = sfile
    sMetaData = Replace(sMetaData, "'", "")
    sMetaData = Replace(sMetaData, """", "")
    sfile = txt1 + vbCrLf + sfile + vbCrLf + txt2
    sfile = Replace(sfile, "#title#", stitle)
    sfile = Replace(sfile, "#metadata#", sMetaData)
    sfile = Replace(sfile, "#asin#", asin)
    sfile = Replace(sfile, "#filepath#", stypePath)
    sfile = Replace(sfile, "#filetitle#", snewFileName)
    sfile = Replace(sfile, "#domain#", sDomain)
    'If sid > 8206 And cstyle = "乙女向" Then
    'sfile = Replace(sfile, "#imgdomain#", imgDomai2)
    'Else
    sfile = Replace(sfile, "#imgdomain#", imgDomain)
    'End If
    sfile = Replace(sfile, "#domainEncode#", urlEncode(sDomain))
    sfile = Replace(sfile, "#authors#", authors)
    sfile = Replace(sfile, "#author1url#", author1url)
    sfile = Replace(sfile, "#series#", Series)
    sfile = Replace(sfile, "#seriesurl#", seriesURL)
    sfile = Replace(sfile, "#series_index#", series_index)
    sfile = Replace(sfile, "#wenku#", wenku)
    sfile = Replace(sfile, "#cntitle#", sCNTitle)
    sfile = Replace(sfile, "#kindlePrice#", sKindlePrice)
    sfile = Replace(sfile, "#paperPrice#", sPaperPrice)
    sfile = Replace(sfile, "#publisher#", publisher)
    sfile = Replace(sfile, "#pubdate#", pubdate)
    sfile = Replace(sfile, "#language#", sLanguage)
    sfile = Replace(sfile, "#rating#", sRating)
    sfile = Replace(sfile, "#tags#", tags)
    sfile = Replace(sfile, "#pages#", sPages)
    sfile = Replace(sfile, "#comments#", sComments)
    sfile = Replace(sfile, "#modifytime#", Format$(Now, "yyyy-mm-dd hh:mm:ss"))
    sfile = Replace(sfile, "#trans#", transURL)
    'sfile = Replace(sfile, "#authorsurl#", authorsURL)
    'sfile = Replace(sfile, "#tagMain#", mainstyle)
    'sfile = Replace(sfile, "#langA", langA)
    'sfile = Replace(sfile, "#buyurlA#", buyurlA)
    'sfile = Replace(sfile, "#translateURL#", "")
    'If seriesindexURLA <> "" Then
    ' sfile = Replace(sfile, "#label2#", "丛书")
    ' sfile = Replace(sfile, "#seriesindexURLA#", seriesindexURLA)
    'Else
    ' sfile = Replace(sfile, "#label2#", "")
    ' sfile = Replace(sfile, "#seriesindexURLA#", "")
    'End If
    
    'sok = sok & asin & " 【" & shortTitle & "】:open" & vbCrLf
    Call writeutf8(shtm, sfile, "UTF-8")
    Call fileCopy(simg, EXEPATH & "240_un\" & asin & ".jpg")
    '@@Call fileCopy(simg, EXEPATH & "img\" & asin & ".jpg")
    '@@Call writeutf8(EXEPATH & "0txt\" & filetitle & "_" & asin & ".txt", snewTxt, "UTF-8")
    createHTM_old = 1
    isNew = 1
Else
    isNew = -1
    'sok = sok & asin & " 【" & shortTitle & "】:duplicate" & vbCrLf
End If
Exit Function

err:
isNew = -2
'sok = sok & asin & " 【" & shortTitle & "】:error" & vbCrLf
End Function


Private Function createPIN(rs As ADODB.Recordset, iType As Integer, txtDIV As String) As Integer

Dim txtPin As String
Dim bookItem As Books
Dim tags As String
Dim yesno As String
Dim mainstyle As String
bookItem = getBook(rs, iType)
tags = bookItem.tags
yesno = bookItem.yesno
mainstyle = bookItem.mainstyle
If InStr(tags, "lowcover") <> 0 Or yesno = "X" Or InStr(tags, "ボーイズラブノベルス") <> 0 Then Exit Function

If Dir(EXEPATH & "template\pin.txt") <> "" Then
txtPin = readutf8(EXEPATH & "template\pin.txt", "UTF-8")
Else
Exit Function
End If

txtPin = Replace(txtPin, "#typepath#", bookItem.typePath)
txtPin = Replace(txtPin, "#asin#", bookItem.asin)
txtPin = Replace(txtPin, "#domain#", sDomain)
txtPin = Replace(txtPin, "#domainEncode#", urlEncode(sDomain))
txtPin = Replace(txtPin, "#filetitle#", bookItem.newFileName)
txtPin = Replace(txtPin, "#vtitle#", bookItem.vtitle)
txtPin = Replace(txtPin, "#authors#", bookItem.authors)
txtPin = Replace(txtPin, "#wenku#", bookItem.wenku)
txtPin = Replace(txtPin, "#publisher#", bookItem.publisher)
txtPin = Replace(txtPin, "#series#", bookItem.Series)
txtPin = Replace(txtPin, "#series_index#", bookItem.series_index)
txtPin = Replace(txtPin, "#pubdate#", bookItem.pubdate)
txtPin = Replace(txtPin, "#kindleprice#", bookItem.kindlePrice)
txtPin = Replace(txtPin, "#paperprice#", bookItem.paperPrice)
txtPin = Replace(txtPin, "#bookurl#", bookItem.bookurl)
txtPin = Replace(txtPin, "#translateURL#", bookItem.translateURL)
createPIN = 1
txtDIV = txtDIV & txtPin
End Function

Private Function getBook(rs As ADODB.Recordset, Optional iType As Integer = 0, Optional sMetaData As String = "") As Books

Dim sname As String, newName As String, stitle As String, snewFileName As String
Dim i As Integer
Dim sPat As String
Dim sPath As String
Dim txtPin As String
Dim bookItem As Books
Dim url As String

On Error Resume Next

For i = 0 To (rs.fields.Count - 1)
Select Case rs.fields(i).Name
Case "id"
    bookItem.id = rs.fields(i).Value
Case "series"
    bookItem.Series = rs.fields(i).Value
    Call removeMark(bookItem.Series)
Case "series_index"
    If bookItem.Series <> "" Then
    bookItem.series_index = rs.fields(i).Value
    sMetaData = sMetaData & "," & bookItem.Series & rs.fields(i).Value
    End If
Case "identifiers" 'like amazon_jp:B00HWLJLAK,mobi-asin:B00HWLJLAK
    bookItem.asin = strSub(rs.fields(i).Value, "mobi-asin:", 10)
    If bookItem.asin = "" Then
    bookItem.asin = strSub(rs.fields(i).Value, "amazon_jp:", 10)
    End If
    sMetaData = sMetaData & "," & bookItem.asin
Case "title"
  bookItem.title = rs.fields(i).Value
  Call removeMark(bookItem.title)
  sMetaData = sMetaData & "," & stitle
Case "authors"
    bookItem.authors = rs.fields(i).Value
    Call getAuthorDetail(bookItem.authors, bookItem.author1, bookItem.author2, bookItem.author1url, bookItem.author2url, bookItem.authorsURL)
    Call removeMark(bookItem.authors)
    sMetaData = sMetaData & "," & bookItem.authors
Case "publisher"
bookItem.publisher = rs.fields(i).Value
Call removeMark(bookItem.publisher)
sMetaData = sMetaData & "," & bookItem.publisher
Case "pubdate"
bookItem.pubdate = Mid$(rs.fields(i).Value, 1, 10)
Case "size"
bookItem.size = rs.fields(i).Value
Case "cover"
    bookItem.cover = rs.fields(i).Value
    sPath = Replace(bookItem.cover, "cover.jpg", "")
    sPat = sPath & "*.azw"
    sname = Dir(sPat)
    If sname <> "" Then
    sname = sPath & Replace(sname, ".azw", ".txt")
    Else
     sPat = sPath & "*.azw3"
     sname = Dir(sPat)
     sname = sPath & Replace(sname, ".azw3", ".txt")
    End If
    If Right(sname, 1) = "3" Then
    sname = Mid$(sname, 1, (Len(sname) - 1))
    End If
    bookItem.txtPath = sname
Case "timestamp"
bookItem.timestamp = rs.fields(i).Value
Case "tags"
bookItem.tags = rs.fields(i).Value
Call removeMark(bookItem.tags)
sMetaData = sMetaData & "," & bookItem.tags
Case "formats"
bookItem.formats = rs.fields(i).Value
Case "rating"
bookItem.rating = rs.fields(i).Value
Case "languages"
bookItem.languages = rs.fields(i).Value
If bookItem.languages = "jpn" Then
bookItem.languages = "日文"
bookItem.langA = "日文"
ElseIf bookItem.languages = "zho" Then
bookItem.languages = "中文"
bookItem.langA = "<b>中文</b>"
End If
Case "#words"
bookItem.words = rs.fields(i).Value
Case "#pages"
bookItem.pages = rs.fields(i).Value
Case "#yesno"
    If IsNull(rs.fields(i).Value) Then
        bookItem.yesno = ""
    ElseIf rs.fields(i).Value = False Then
        bookItem.yesno = ""
    ElseIf rs.fields(i).Value = True Then
        bookItem.yesno = "X"
    ElseIf Trim(rs.fields(i).Value) = "TRUE" Then
        bookItem.yesno = "X"
    End If
Case "#cnseries"
bookItem.cnseries = rs.fields(i).Value
Case "#cnseries_index"
bookItem.cnseries_index = rs.fields(i).Value
Case "#cntitle"
bookItem.cntitle = rs.fields(i).Value
Case "#vtitle"
bookItem.vtitle = rs.fields(i).Value
Case "#wenku"
bookItem.wenku = rs.fields(i).Value
Call removeMark(bookItem.wenku)
sMetaData = sMetaData & "," & bookItem.wenku
Case "wenku"
bookItem.wenku = rs.fields(i).Value
Call removeMark(bookItem.wenku)
sMetaData = sMetaData & "," & bookItem.wenku
Case "#sorttitle"
bookItem.sorttitle = rs.fields(i).Value
Case "#filetitle"
bookItem.filetitle = rs.fields(i).Value
Case "#kindleprice"
bookItem.kindlePrice = rs.fields(i).Value
If bookItem.kindlePrice <> "" Then
bookItem.kindlePrice = "￥" & bookItem.kindlePrice
End If
Case "#paperprice"
bookItem.paperPrice = rs.fields(i).Value
If bookItem.paperPrice <> "" Then
bookItem.paperPrice = "￥" & bookItem.paperPrice
End If
Case "comments"
bookItem.comments = rs.fields(i).Value
Case "#lastmodify"
bookItem.lastModify = Format$(rs.fields(i).Value, "yyyy-mm-dd hh:mm:ss")
Case "#finish"
    If IsNull(rs.fields(i).Value) Then
        bookItem.finished = ""
    ElseIf rs.fields(i).Value = False Then
        bookItem.finished = ""
    ElseIf rs.fields(i).Value = True Then
        bookItem.finished = "X"
    ElseIf Trim(rs.fields(i).Value) = "TRUE" Then
        bookItem.finished = "X"
    End If
End Select
Next i
    
If iType = 0 Then
bookItem.typePath = "asin"
bookItem.newFileName = bookItem.asin
ElseIf iType = 1 Then
bookItem.typePath = "htm"
bookItem.newFileName = bookItem.filetitle
End If
bookItem.bookurl = sDomain & bookItem.typePath & "/" & bookItem.newFileName & ".htm"
'If bookItem.id > 8206 And cstyle = "乙女向" Then
'bookItem.coverURL = imgDomai2 & "asin/img/" & bookItem.asin & ".jpg"
'Else
bookItem.coverURL = imgDomain & "asin/img/" & bookItem.asin & ".jpg"
'End If

If bookItem.wenku <> "" Then
url = "wenku/" & urlEncode(bookItem.wenku) & ".htm"
Call removeMark(url)
bookItem.wenkuURL = "<a href=""" & url & """>" & bookItem.wenku & "</a>"
End If

If bookItem.Series <> "" Then
url = "series/" & urlEncode(bookItem.Series) & ".htm"
Call removeMark(url)
If bookItem.cnseries <> "" Then
    bookItem.seriesURL = "<a href=""" & url & """ title=""" & bookItem.Series & """>" & bookItem.cnseries & "</a>"
Else
    bookItem.seriesURL = "<a href=""" & url & """ title=""" & bookItem.cnseries & """>" & bookItem.Series & "</a>"
End If

End If


If InStr(bookItem.tags, "全本") > 0 Then
bookItem.translateURL = transPr2 & urlEncode(bookItem.bookurl)
Else
bookItem.translateURL = transPre & urlEncode(bookItem.bookurl)
End If

If bookItem.languages = "中文" Then
bookItem.translateURL = bookItem.bookurl
End If

bookItem.publisher = getPubfromWenku(bookItem.wenku, bookItem.title)

If bookItem.yesno = "X" Then
    bookItem.mainstyle = cstyle & " TL"
ElseIf InStr(bookItem.tags, "ボーイズラブノベルス") <> 0 Then
    bookItem.mainstyle = cstyle & " BL"
ElseIf InStr(bookItem.tags, "日常") <> 0 Then
    bookItem.mainstyle = cstyle & " 日常"
Else
    bookItem.mainstyle = cstyle
End If

bookItem.tagMain = bookItem.mainstyle
If InStr(bookItem.tags, "书籍样本") <> 0 Then
bookItem.tagMain = bookItem.tagMain & " 书籍样本"
ElseIf InStr(bookItem.tags, "全本") <> 0 Then
bookItem.tagMain = bookItem.tagMain & " 全本"
End If

If InStr(bookItem.tags, "短篇集") <> 0 Then
bookItem.tagMain = bookItem.tagMain & " 短篇集"
End If

If bookItem.Series <> "" Then
bookItem.seriesindexURLA = bookItem.seriesURL & "(" & bookItem.series_index & ")"
End If
 
bookItem.NumURLA = "<a href=""" & bookItem.translateURL & """ target=""_blank"">译文</a>"

If bookItem.cntitle <> "" Then
  url = bookItem.cntitle
Else
  url = bookItem.filetitle
End If

bookItem.titleFiltered = ReplaceY(url)
bookItem.titleURLA = "<a href=""" & bookItem.bookurl & """ title=""" & bookItem.cntitle & """ target=""_blank"">" & ReplaceY(url) & "</a>"

bookItem.buyurlA = "<a href=""https://www.amazon.co.jp/dp/" & bookItem.asin & """ title=""日本亚马逊"" target=""_blank"">原版</a>"

getBook = bookItem
End Function



Private Function createWenku(rs As ADODB.Recordset, iType As Integer, sOutput As String, Optional iXLS As Integer = 0) As Integer
If iXLS = 1 Then
  If rs.fields("wenku").Value = sql Then
  Call createRow2(rs, iType, sOutput, txt3)
  createWenku = 1
  End If
Else
  If rs.fields("#wenku").Value = sql Then
  Call createRow2(rs, iType, sOutput, txt3)
  createWenku = 1
  End If
End If
End Function

Private Function createSeries(rs As ADODB.Recordset, iType As Integer, sOutput As String) As Integer
If rs.fields("series").Value = sql Then
Call createRow2(rs, iType, sOutput, txt3)
createSeries = 1
End If
End Function

Private Function createSeriesRow(rs As ADODB.Recordset, iType As Integer, lastASIN As Books, sOutput As String) As Integer
Dim bookItem As Books
Dim sn As String

If rs.fields("series").Value = sql Then
bookItem = getBook(rs, iType)
If bookItem.languages = "中文" Then
sn = "<b>" & bookItem.series_index & "</b>"
Else
sn = bookItem.series_index
End If
If lastASIN.asin = "" Then lastASIN = bookItem '因为pubdate是倒序的,所以只取一次
sOutput = sOutput & "@" & sn & "|" & bookItem.asin '格式为 序号|asin@序号|asin
createSeriesRow = 1
End If
End Function

Private Function createAuthorRow(rs As ADODB.Recordset, iType As Integer, lastASIN As Books, sOutput As String) As Integer
Dim bookItem As Books
Dim author1 As String
author1 = getAuthorDetail(rs.fields("authors").Value)

If author1 = sql Then
    bookItem = getBook(rs, iType)
    If lastASIN.asin = "" Then lastASIN = bookItem '因为pubdate是倒序的,所以只取一次
    sOutput = sOutput & "@" & bookItem.asin '格式为 asin@asin
    createAuthorRow = 1
End If
End Function
Private Function createAuthor(rs As ADODB.Recordset, iType As Integer, sOutput As String) As Integer

Dim author1 As String
author1 = getAuthorDetail(rs.fields("authors").Value)

If author1 = sql Then
Call createRow2(rs, iType, sOutput, txt3)
createAuthor = 1
End If
End Function

Private Function createZH(rs As ADODB.Recordset, iType As Integer, sOutput As String) As Integer
If rs.fields("languages").Value = "zho" Then
Call createRow2(rs, 3, sOutput, txt3)
createZH = 1
End If
End Function


Private Function createList(rs As ADODB.Recordset, iType As Integer, sOutput As String) As Integer
Call createRow2(rs, iType, sOutput, txt3)
createList = 1
End Function


Private Function createRow_old(rs As ADODB.Recordset, iType As Integer, sOutput As String) As String
Dim bookItem As Books
Dim sTxt As String
Dim ses As String
Dim tag As String
Dim url As String
Dim lang As String
Dim url2 As String
Dim imode As Integer
imode = iType

If iType = 3 Then imode = 0
bookItem = getBook(rs, imode)

tag = bookItem.mainstyle
If InStr(bookItem.tags, "书籍样本") <> 0 Then
tag = tag & " 书籍样本"
ElseIf InStr(bookItem.tags, "全本") <> 0 Then
tag = tag & " 全本"
End If

If InStr(bookItem.tags, "短篇集") <> 0 Then
tag = tag & " 短篇集"
End If

If bookItem.languages = "jpn" Or bookItem.languages = "日文" Then
lang = "日文"
ElseIf bookItem.languages = "zho" Or bookItem.languages = "中文" Then
lang = "<strong>中文</strong>"
End If

If bookItem.Series <> "" Then
ses = bookItem.seriesURL & "(" & bookItem.series_index & ")"
End If

url = bookItem.typePath & "/" & bookItem.newFileName
'url2 = UTF8UrlEncode(url)
url2 = bookItem.typePath & "%2F" & bookItem.newFileName

sTxt = "<tr>"
If iType = 3 Then
    sTxt = sTxt & "<td class=""tbno"">" & "&nbsp;" & "</td>"
Else
    sTxt = sTxt & "<td class=""tbno""><a href=""" & bookItem.translateURL & """ target=""_blank"">译文</a></td>"
End If
bookItem.filetitle = ReplaceY(bookItem.filetitle)
If bookItem.cntitle <> "" Then
    If iType = 3 Then
        sTxt = sTxt & "<td class=""tbtitle""><a href=""" & url & ".htm"" title=""" & bookItem.filetitle & """>" & bookItem.cntitle & "</a></td>"
    Else
        sTxt = sTxt & "<td class=""tbtitle""><a href=""" & url & ".htm"" title=""" & bookItem.cntitle & """>" & bookItem.filetitle & "</a></td>"
    End If
Else
    sTxt = sTxt & "<td class=""tbtitle""><a href=""" & url & ".htm"">" & bookItem.filetitle & "</a></td>"
End If
sTxt = sTxt & "<td class=""tbauthor"">" & bookItem.authorsURL & "</td>"
sTxt = sTxt & "<td class=""tbwenku"">" & bookItem.wenkuURL & "</td>"
sTxt = sTxt & "<td class=""tbseries"">" & ses & "</td>"
sTxt = sTxt & "<td class=""tbpublisher"">" & bookItem.publisher & "</td>"
sTxt = sTxt & "<td class=""tbpubdate"">" & bookItem.pubdate & "</td>"
sTxt = sTxt & "<td class=""tbtags"">" & tag & "</td>"
sTxt = sTxt & "<td class=""tbpages"">" & bookItem.pages & "</td>"
sTxt = sTxt & "<td class=""tblanguage"">" & lang & "</td>"
sTxt = sTxt & "<td class=""tbkindleprice"">" & bookItem.kindlePrice & "</td>"
sTxt = sTxt & "<td class=""tbwenkuprice"">" & bookItem.paperPrice & "</td>"
sTxt = sTxt & "<td class=""tbbuy"">" & "<a href=""https://www.amazon.co.jp/dp/" & bookItem.asin & """>原版</a></td>"
sTxt = sTxt & "</tr>" & vbCrLf
sOutput = sOutput & sTxt
End Function

Private Function createRow2(rs As ADODB.Recordset, iType As Integer, sOutput As String, Optional txtTemplate As String = "") As String
Dim bookItem As Books
Dim txtPin As String
Dim sTxt As String
Dim imode As Integer
imode = iType
txtPin = txtTemplate

If txtPin = "" Then txtPin = readutf8(EXEPATH & "template\grouplist_middle.txt")
If iType = 3 Then imode = 0
bookItem = getBook(rs, imode)

If iType = 3 Then
  bookItem.NumURLA = ""
End If

txtPin = Replace(txtPin, "#label1#", "作者：")

txtPin = Replace(txtPin, "#bookurl#", bookItem.bookurl)
txtPin = Replace(txtPin, "#lastmod#", bookItem.lastModify)
txtPin = Replace(txtPin, "#typepath#", bookItem.typePath)
txtPin = Replace(txtPin, "#asin#", bookItem.asin)
txtPin = Replace(txtPin, "#filetitle#", bookItem.filetitle)
txtPin = Replace(txtPin, "#title#", bookItem.title)
txtPin = Replace(txtPin, "#vtitle#", bookItem.vtitle)
txtPin = Replace(txtPin, "#domain#", sDomain)
txtPin = Replace(txtPin, "#domainEncode#", urlEncode(sDomain))
txtPin = Replace(txtPin, "#mainstyle#", bookItem.mainstyle)
txtPin = Replace(txtPin, "#authors#", bookItem.authors)
txtPin = Replace(txtPin, "#pubdate#", bookItem.pubdate)
txtPin = Replace(txtPin, "#coverurl#", bookItem.coverURL)
txtPin = Replace(txtPin, "#wenku#", bookItem.wenku)
txtPin = Replace(txtPin, "#languages#", bookItem.languages)
txtPin = Replace(txtPin, "#publisher#", bookItem.publisher)
txtPin = Replace(txtPin, "#series#", bookItem.Series)
txtPin = Replace(txtPin, "#series_index#", bookItem.series_index)
txtPin = Replace(txtPin, "#words#", bookItem.words)
txtPin = Replace(txtPin, "#pages#", bookItem.pages)
txtPin = Replace(txtPin, "#tags#", bookItem.tags)
txtPin = Replace(txtPin, "#comments#", bookItem.comments)
txtPin = Replace(txtPin, "#kindlePrice#", bookItem.kindlePrice)
txtPin = Replace(txtPin, "#paperPrice#", bookItem.paperPrice)
txtPin = Replace(txtPin, "#author1#", bookItem.author1)
txtPin = Replace(txtPin, "#author2#", bookItem.author2)
txtPin = Replace(txtPin, "#authorsurl#", bookItem.authorsURL)
txtPin = Replace(txtPin, "#author1url#", bookItem.author1url)
txtPin = Replace(txtPin, "#id#", bookItem.id)
txtPin = Replace(txtPin, "#cnseries#", bookItem.cnseries)
txtPin = Replace(txtPin, "#cnseries_index#", bookItem.cnseries_index)
txtPin = Replace(txtPin, "#cntitle#", bookItem.cntitle)
txtPin = Replace(txtPin, "#langA#", bookItem.langA)
txtPin = Replace(txtPin, "#NumURLA#", bookItem.NumURLA)
txtPin = Replace(txtPin, "#seriesURL#", bookItem.seriesURL)
txtPin = Replace(txtPin, "#seriesindexURLA#", bookItem.seriesindexURLA)
If bookItem.seriesindexURLA = "" Then
txtPin = Replace(txtPin, "#label2#", "")
Else
txtPin = Replace(txtPin, "#label2#", "丛书：")
End If
txtPin = Replace(txtPin, "#tagMain#", bookItem.tagMain)
txtPin = Replace(txtPin, "#titleURLA#", bookItem.titleURLA)
txtPin = Replace(txtPin, "#translateURL#", bookItem.translateURL)
txtPin = Replace(txtPin, "#transURLA#", bookItem.transURLA)
txtPin = Replace(txtPin, "#wenkuurl#", bookItem.wenkuURL)
txtPin = Replace(txtPin, "#titleF#", bookItem.titleFiltered)
txtPin = Replace(txtPin, "#buyurlA#", bookItem.buyurlA)

sOutput = sOutput & txtPin & vbCrLf


End Function
Public Function urlEncode_old(ByVal strParameter As String) As String
  
Dim s As String
Dim i As Integer
Dim intValue As Integer
  
Dim TmpData() As Byte
  
    s = ""
    TmpData = StrConv(strParameter, vbFromUnicode)
    For i = 0 To UBound(TmpData)
      intValue = TmpData(i)
      If (intValue >= 48 And intValue <= 57) Or _
        (intValue >= 65 And intValue <= 90) Or _
        (intValue >= 97 And intValue <= 122) Then
        s = s & Chr(intValue)
      ElseIf intValue = 32 Then
        s = s & "+"
      Else
        s = s & "%" & Hex(intValue)
      End If
    Next i
    urlEncode_old = s
  
End Function

Function urlEncode(ByVal szInput As String) As String 'UTF8Encode_ForJs
       Dim wch  As String
       Dim uch As String
       Dim szRet As String
       Dim x As Long
       Dim inputLen As Long
       Dim nAsc  As Long
       Dim nAsc2 As Long
       Dim nAsc3 As Long
          
       If szInput = "" Then
           urlEncode = szInput
           Exit Function
       End If
       inputLen = Len(szInput)
       For x = 1 To inputLen
       '得到每个字符
           wch = Mid(szInput, x, 1)
           '得到相应的UNICODE编码
           nAsc = AscW(wch)
       '对于<0的编码　其需要加上65536
           If nAsc < 0 Then nAsc = nAsc + 65536
       '对于<128位的ASCII的编码则无需更改
           If (nAsc And &HFF80) = 0 Then
               szRet = szRet & wch
           Else
               If (nAsc And &HF000) = 0 Then
               '真正的第二层编码范围为000080 - 0007FF
               'Unicode在范围D800-DFFF中不存在任何字符，基本多文种平面中约定了这个范围用于UTF-16扩展标识辅助平面（两个UTF-16表示一个辅助平面字符）.
               '当然，任何编码都是可以被转换到这个范围，但在unicode中他们并不代表任何合法的值。
          
                   uch = "%" & Hex(((nAsc \ 2 ^ 6)) Or &HC0) & Hex(nAsc And &H3F Or &H80)
                   szRet = szRet & uch
                      
     
     
     
               Else
               '第三层编码00000800 – 0000FFFF
               '首先取其前四位与11100000进行或去处得到UTF-8编码的前8位
               '其次取其前10位与111111进行并运算，这样就能得到其前10中最后6位的真正的编码　再与10000000进行或运算来得到UTF-8编码中间的8位
               '最后将其与111111进行并运算，这样就能得到其最后6位的真正的编码　再与10000000进行或运算来得到UTF-8编码最后8位编码
                   uch = "%" & Hex((nAsc \ 2 ^ 12) Or &HE0) & "%" & _
                   Hex((nAsc \ 2 ^ 6) And &H3F Or &H80) & "%" & _
                   Hex(nAsc And &H3F Or &H80)
                   szRet = szRet & uch
               End If
           End If
       Next
          
       urlEncode = szRet
End Function

Function UTF8Encode(ByVal szInput As String) As String
       Dim wch  As String
       Dim uch As String
       Dim szRet As String
       Dim x As Long
       Dim inputLen As Long
       Dim nAsc  As Long
       Dim nAsc2 As Long
       Dim nAsc3 As Long
          
       If szInput = "" Then
           UTF8Encode = szInput
           Exit Function
       End If
       inputLen = Len(szInput)
       For x = 1 To inputLen
       '得到每个字符
           wch = Mid(szInput, x, 1)
           '得到相应的UNICODE编码
           nAsc = AscW(wch)
       '对于<0的编码　其需要加上65536
           If nAsc < 0 Then nAsc = nAsc + 65536
       '对于<128位的ASCII的编码则无需更改
           If (nAsc And &HFF80) = 0 Then
               szRet = szRet & wch
           Else
               If (nAsc And &HF000) = 0 Then
               '真正的第二层编码范围为000080 - 0007FF
               'Unicode在范围D800-DFFF中不存在任何字符，基本多文种平面中约定了这个范围用于UTF-16扩展标识辅助平面（两个UTF-16表示一个辅助平面字符）.
               '当然，任何编码都是可以被转换到这个范围，但在unicode中他们并不代表任何合法的值。
          
                   uch = "%" & Hex(((nAsc \ 2 ^ 6)) Or &HC0) & Hex(nAsc And &H3F Or &H80)
                   szRet = szRet & uch
                      
     
     
     
               Else
               '第三层编码00000800 – 0000FFFF
               '首先取其前四位与11100000进行或去处得到UTF-8编码的前8位
               '其次取其前10位与111111进行并运算，这样就能得到其前10中最后6位的真正的编码　再与10000000进行或运算来得到UTF-8编码中间的8位
               '最后将其与111111进行并运算，这样就能得到其最后6位的真正的编码　再与10000000进行或运算来得到UTF-8编码最后8位编码
                   uch = "%" & Hex((nAsc \ 2 ^ 12) Or &HE0) & "%" & _
                   Hex((nAsc \ 2 ^ 6) And &H3F Or &H80) & "%" & _
                   Hex(nAsc And &H3F Or &H80)
                   szRet = szRet & uch
               End If
           End If
       Next
          
       UTF8Encode = szRet
End Function
Function UTF8UrlEncode(szInput As String) As String
Dim toUTF8
    Dim wch, uch, szRet
    Dim x
    Dim nAsc, nAsc2, nAsc3
    '如果输入参数为空，则退出函数
    If szInput = "" Then
        toUTF8 = szInput
        Exit Function
    End If
    '开始转换
     For x = 1 To Len(szInput)
        '利用mid函数分拆GB编码文字
        wch = Mid(szInput, x, 1)
        '利用ascW函数返回每一个GB编码文字的Unicode字符代码
        '注：asc函数返回的是ANSI 字符代码，注意区别
        nAsc = AscW(wch)
        If nAsc < 0 Then nAsc = nAsc + 65536
  
        If (nAsc And &HFF80) = 0 Then
            szRet = szRet & wch
        Else
            If (nAsc And &HF000) = 0 Then
                uch = "%" & Hex(((nAsc2 ^ 6)) Or &HC0) & Hex(nAsc And &H3F Or &H80)
                szRet = szRet & uch
            Else
               'GB编码文字的Unicode字符代码在0800 - FFFF之间采用三字节模版
                uch = "%" & Hex((nAsc2 ^ 12) Or &HE0) & "%" & Hex((nAsc2 ^ 6) And &H3F Or &H80) & "%" & Hex(nAsc And &H3F Or &H80)
                szRet = szRet & uch
            End If
        End If
    Next
      
    toUTF8 = szRet
UTF8UrlEncode = toUTF8
End Function

Sub removeMark(sInput As String)
sInput = Replace(sInput, """", "_")
sInput = Replace(sInput, "'", "")
sInput = Replace(sInput, ":", "_")
sInput = Replace(sInput, "?", "？")
End Sub


Private Function create_urlTxt(sType As String) As String
File1.Pattern = "*.htm"
File1.Path = EXEPATH & sType & "\"
File1.Refresh
Dim i As Integer
Dim sname As String
Dim sline As String
Dim stitle As String
Dim sfile As String
Dim sindex As String

'获取路径下所有htm文件,其文件名就是title
For i = 0 To File1.ListCount - 1
sname = File1.List(i)
sname = UTF8Encode(sname)
sfile = sfile & sDomain & sType & "/" & sname & vbCrLf
Next i
create_urlTxt = sfile
End Function

Private Sub create_SiteMaptxt()
Dim sTxt As String
sTxt = sTxt & sDomain & "index.htm" & vbCrLf
sTxt = sTxt & create_urlTxt("asin")
sTxt = sTxt & create_urlTxt("list")
sTxt = sTxt & create_urlTxt("series")
sTxt = sTxt & create_urlTxt("wenku")
sTxt = sTxt & create_urlTxt("author")
sTxt = sTxt & create_urlTxt("novel")
Call writeutf8(EXEPATH & "sitemap.txt", sTxt, "UTF-8")
Label1.Caption = Label1.Caption & " URL列表sitemap.txt生成好了"
End Sub

Private Function create_urlXML(rs As ADODB.Recordset, iType As Integer, sOutput As String, s2 As String) As Integer

Dim bookItem As Books
Dim txtPin As String
Dim lastmod As String
Dim txtXML As String
bookItem = getBook(rs, iType)
'##bookItem.comments = getComments(bookItem.asin)
bookItem.comments = "轻小说"
bookItem.filetitle = ReplaceY(bookItem.filetitle)
'lastmod = Format$(bookItem.timestamp, "yyyy-mm-ddThh:mm:ss")
lastmod = Mid(bookItem.timestamp, 1, 19)
'lastmod = bookItem.timestamp 百度不支持+8:00这样的格式
txtPin = readutf8(EXEPATH & "template\sitemap.txt", "UTF-8")
txtPin = ReplaceX(txtPin, "#bookurl#", bookItem.bookurl)
txtPin = ReplaceX(txtPin, "#lastmod#", lastmod)
txtPin = ReplaceX(txtPin, "#typepath#", bookItem.typePath)
txtPin = ReplaceX(txtPin, "#asin#", bookItem.asin)
txtPin = ReplaceX(txtPin, "#filetitle#", bookItem.filetitle)
txtPin = ReplaceX(txtPin, "#title#", bookItem.title)
txtPin = ReplaceX(txtPin, "#vtitle#", bookItem.vtitle)
txtPin = ReplaceX(txtPin, "#domain#", sDomain)
txtPin = ReplaceX(txtPin, "#domainEncode#", urlEncode(sDomain))
txtPin = ReplaceX(txtPin, "#mainstyle#", bookItem.mainstyle)
txtPin = ReplaceX(txtPin, "#authors#", bookItem.authors)
txtPin = ReplaceX(txtPin, "#pubdate#", bookItem.pubdate)
txtPin = ReplaceX(txtPin, "#coverurl#", bookItem.coverURL)
txtPin = ReplaceX(txtPin, "#wenku#", bookItem.wenku)
txtPin = ReplaceX(txtPin, "#languages#", bookItem.languages)
txtPin = ReplaceX(txtPin, "#publisher#", bookItem.publisher)
txtPin = ReplaceX(txtPin, "#series#", bookItem.Series)
txtPin = ReplaceX(txtPin, "#series_index#", bookItem.series_index)
txtPin = Replace(txtPin, "#words#", bookItem.words)
txtPin = Replace(txtPin, "#pages#", bookItem.pages)
txtPin = ReplaceX(txtPin, "#tags#", bookItem.tags)
txtPin = ReplaceX(txtPin, "#comments#", bookItem.comments)
txtPin = ReplaceX(txtPin, "#kindlePrice#", bookItem.kindlePrice)
txtPin = ReplaceX(txtPin, "#paperPrice#", bookItem.paperPrice)
sOutput = sOutput & txtPin & vbCrLf

txtXML = readutf8(EXEPATH & "template\sitemap1.txt", "UTF-8")
txtXML = ReplaceX(txtXML, "#bookurl#", bookItem.bookurl)
txtXML = ReplaceX(txtXML, "#lastmod#", lastmod)
s2 = s2 & txtXML & vbCrLf
End Function


Private Sub create_SiteMapXML(sContent As String, ssContet As String)
Dim sTxt As String
Dim sTx1 As String
sTxt = "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbCrLf
sTxt = sTxt & "<urlset xmlns=""http://www.sitemaps.org/schemas/sitemap/0.9"""
'sTxt = sTxt & " xmlns:mobile=""http://www.baidu.com/schemas/sitemap-mobile/1/"""
sTxt = sTxt & ">" & vbCrLf

sTx1 = sTxt & ssContet & "</urlset>" '简化版sitemap
sTxt = sTxt & sContent & "</urlset>" '结构化数据

Call writeutf8(EXEPATH & "sitemap.xml", sTxt, "UTF-8")
Call writeutf8(EXEPATH & "sitemap1.xml", sTx1, "UTF-8")
Label1.Caption = Label1.Caption & " URL列表xml生成好了"
End Sub

Private Function markXML(str1 As String)
str1 = Replace(str1, "<", "&lt;")
str1 = Replace(str1, ">", "&gt;")
str1 = Replace(str1, "&", "&amp;")
str1 = Replace(str1, "'", "&apos;")
str1 = Replace(str1, """", "&quot;")
markXML = str1
End Function

Private Function ReplaceX(str1 As String, find As String, re As String)
re = markXML(re)
ReplaceX = Replace(str1, find, re)
End Function


Private Function ReplaceY(str1 As String)
If filterflag = "1" Then
    str1 = Replace(str1, "淫", "")
    str1 = Replace(str1, "濡", "")
    str1 = Replace(str1, "囁", "")
    'str1 = Replace(str1, "溺", "")
    'str1 = Replace(str1, "蜜", "")
    'str1 = Replace(str1, "初夜", "")
End If
ReplaceY = str1
End Function


Function removeMarkXML(sInput As String) As String
sInput = Replace(sInput, "None", "")
Call removeMark(sInput)
removeMarkXML = sInput
End Function

Function getComments(sASIN As String)
    Dim xmlDoc 'As DOMDocument
    Dim RootNode 'As IXMLDOMNode
    Dim bNode 'As IXMLDOMNode
    Dim cNode 'As IXMLDOMNode
    Dim dNode 'As IXMLDOMNode
    '#Set xmlDoc = New DOMDocument

    Dim nodeName As String
    Dim i As Integer
    
    xmlDoc.Load EXEPATH & "\dbook.xml"
    If xmlDoc.documentElement Is Nothing Then
        Exit Function
    End If
   
    Set RootNode = xmlDoc.documentElement '第一个节点 calibredb
    'Dim tBook(0 To (RootNode.childNodes.length - 1)) As Books
    
    For i = 0 To RootNode.childNodes.length - 1
        Dim bookItem As Books
        'If i <> 0 Then ReDim tBook(0 To i) As Books
        Set bNode = RootNode.childNodes.Item(i) '第一个节点的下一层节点 (第2层 record)
        For Each cNode In bNode.childNodes  '第3层节点 书的详细信息
        nodeName = LCase(cNode.nodeName)
        Select Case nodeName
        Case "title"
        bookItem.title = cNode.Text
        Case "_words"
        bookItem.words = removeMarkXML(cNode.Text)
        Case "_pages"
        bookItem.pages = removeMarkXML(cNode.Text)
        Case "_yesno"
        If LCase(cNode.Text) = "true" Then
        bookItem.yesno = "X"
        End If
        Case "_cnseries"
        bookItem.cnseries = removeMarkXML(cNode.Text)
        Case "_cnseries_index"
        bookItem.cnseries_index = removeMarkXML(cNode.Text)
        Case "_cntitle"
        bookItem.cntitle = removeMarkXML(cNode.Text)
        Case "_vtitle"
        bookItem.vtitle = removeMarkXML(cNode.Text)
        Case "_wenku"
        bookItem.wenku = removeMarkXML(cNode.Text)
        Case "_sorttitle"
        bookItem.sorttitle = removeMarkXML(cNode.Text)
        Case "_filetitle"
        bookItem.filetitle = removeMarkXML(cNode.Text)
        Case "_asin"
        bookItem.asin = cNode.Text
        Case "_sortid"
        Case "_kindleprice"
        bookItem.kindlePrice = removeMarkXML(cNode.Text)
        If bookItem.kindlePrice <> "" Then
        bookItem.kindlePrice = "￥" & bookItem.kindlePrice
        End If
        Case "_paperprice"
        bookItem.paperPrice = removeMarkXML(cNode.Text)
        If bookItem.paperPrice <> "" Then
        bookItem.paperPrice = "￥" & bookItem.paperPrice
        End If
        Case "id"
        bookItem.id = cNode.Text
        Case "publisher"
        bookItem.publisher = removeMarkXML(cNode.Text)
        Case "rating"
        bookItem.rating = removeMarkXML(cNode.Text)
        Case "size"
        bookItem.size = cNode.Text
        Case "identifiers" 'mobi-asin:B00N2CPS3C,amazon_jp:B00N2CPS3C
        bookItem.identifiers = cNode.Text
        Case "authors"
        bookItem.authors = cNode.Text
             For Each dNode In cNode.childNodes  '第4层 数组
             Next 'dNode
    '<authors sort="a, b &amp; c, d">
    '  <author>a, b</author>
    '  <author>c, d</author>
    '</authors>
            
    Case "timestamp" '<timestamp>2014-09-20T11:47:04+08:00</timestamp>
    bookItem.timestamp = cNode.Text
    Case "pubdate" '<pubdate>2014-08-18T08:00:00+08:00</pubdate>
    bookItem.pubdate = removeMarkXML(Mid$(cNode.Text, 1, 10))
    Case "tags"
         For Each dNode In cNode.childNodes  '第4层 数组
         bookItem.tags = bookItem.tags & removeMarkXML(dNode.Text) & ","
         Next 'dNode
      '<tags>
      '<tag>abc</tag>
      '<tag>xyz</tag>
      '<tag>Kindle</tag>
      '</tags>
    Case "languages"
    bookItem.languages = removeMarkXML(cNode.Text)
    Case "comments"
    bookItem.comments = removeMarkXML(cNode.Text)
       '<comments>&lt;h3&gt;XXX&lt;/p&gt;</comments>
    Case "cover"
    bookItem.cover = cNode.Text
      '<cover>C:/Users/lousi/Documents/calibre/sayuri, Qi Fu/Xin Hun Kuang Xiang Qu  Qi Shi Tuan (6)/cover.jpg</cover>
    
    Case "formats"
        For Each dNode In cNode.childNodes  '第4层 数组
        If Right(dNode.Text, 4) = ".txt" Then
        bookItem.txtPath = dNode.Text
        End If
        Next 'dNode
    '<formats>
    '  <format>C:/Users/lousi/Documents/calibre/sayuri, Qi Fu/Xin Hun Kuang Xiang Qu  Qi Shi Tuan (6)/Xin Hun Kuang Xiang Qu  Qi Shi - sayuri, Qi Fu.azw3</format>
    '  <format>C:/Users/lousi/Documents/calibre/sayuri, Qi Fu/Xin Hun Kuang Xiang Qu  Qi Shi Tuan (6)/Xin Hun Kuang Xiang Qu  Qi Shi - sayuri, Qi Fu.epub</format>
    '  <format>C:/Users/lousi/Documents/calibre/sayuri, Qi Fu/Xin Hun Kuang Xiang Qu  Qi Shi Tuan (6)/Xin Hun Kuang Xiang Qu  Qi Shi - sayuri, Qi Fu.htmlz</format>
    '  <format>C:/Users/lousi/Documents/calibre/sayuri, Qi Fu/Xin Hun Kuang Xiang Qu  Qi Shi Tuan (6)/Xin Hun Kuang Xiang Qu  Qi Shi - sayuri, Qi Fu.txt</format>
    '</formats>
    Case "library_name"
    ' <library_name>calibre</library_name>
     
    End Select
    Next 'cNode
    If bookItem.yesno = "X" Then
     bookItem.mainstyle = cstyle & "TL"
    ElseIf InStr(bookItem.tags, "ボーイズラブノベルス") <> 0 Then
     bookItem.mainstyle = cstyle & "BL"
    ElseIf InStr(bookItem.tags, "日常") <> 0 Then
     bookItem.mainstyle = cstyle & "日常"
    Else
     bookItem.mainstyle = cstyle
    End If
    bookItem.typePath = "asin"
    bookItem.newFileName = bookItem.asin
    bookItem.bookurl = sDomain & bookItem.newFileName & ".htm"
    bookItem.coverURL = sDomain & "asin/img/" & bookItem.asin & ".jpg"
    'tBooks(i) = bookItem
    If bookItem.asin = sASIN Then
    getComments = bookItem.comments
    Set cNode = Nothing
    Set bNode = Nothing
    Set RootNode = Nothing
    Set xmlDoc = Nothing
    Exit Function
    End If
    Next 'i bNode:records

End Function


Private Function createSeriesRowtxt_old(bookItem As Books, txtSeriesRow As String, seriesNo As Integer) As String
Dim i As Integer
Dim sindex As Integer
'Dim bookItem As Books
Dim sNo As String
Dim sASIN As String
Dim sTxt As String
Dim url As String
Dim urlAll As String
Dim tSeries() As String
Dim tNo(2) As String
Dim strSes As String
Dim j As Integer
Dim n As Integer
Dim buyURL As String

If Mid(bookItem.asin, 1, 1) = "B" Then
buyURL = "<a href=""https://www.amazon.co.jp/s/ref=series_rw_dp_labf?_encoding=UTF8&field-collection=" & bookItem.Series & "&url=search-alias%3Ddigital-text"">原版</a>"
Else
buyURL = "<a href=""https://www.amazon.co.jp/s/ref=nb_sb_noss?__mk_ja_JP=%E3%82%AB%E3%82%BF%E3%82%AB%E3%83%8A&url=search-alias%3Dstripbooks&field-keywords=" & bookItem.Series & """>原版</a>"
End If

If Left(txtSeriesRow, 1) = "@" Then txtSeriesRow = Right(txtSeriesRow, Len(txtSeriesRow) - 1)
tSeries = Split(txtSeriesRow, "@")
sindex = seriesNo
sTxt = "<tr>"
sTxt = sTxt & "<td class=""tbno"">" & "<a href=""https://ja.wikipedia.org/wiki/" & bookItem.Series & """ title=""wiki"">" & sindex & "</a></td>"

If bookItem.finished <> "" Then
sTxt = sTxt & "<td class=""tbseries0"">" & bookItem.seriesURL & "<img id=""finish"" class=""finish_img"" src=""app/ok.gif"" alt=""已完结"">" & "</td>"
Else
sTxt = sTxt & "<td class=""tbseries0"">" & bookItem.seriesURL & "</td>"
End If


For i = 0 To UBound(tSeries)
  n = UBound(tSeries) - i
  strSes = tSeries(n)
  j = InStr(strSes, "|")
  sNo = Left(strSes, j - 1)
  sASIN = Right(strSes, Len(strSes) - j)
  url = "asin/" & sASIN
  urlAll = urlAll & "<a href=""" & url & ".htm"">" & sNo & "</a>&nbsp;"
  
Next i
sTxt = sTxt & "<td>" & urlAll & "</td>"
sTxt = sTxt & "<td class=""tbauthor"">" & bookItem.authorsURL & "</td>"
sTxt = sTxt & "<td class=""tbwenku"">" & bookItem.wenkuURL & "</td>"
sTxt = sTxt & "<td class=""tbpublisher"">" & bookItem.publisher & "</td>"
sTxt = sTxt & "<td class=""tbpubdate"">" & bookItem.pubdate & "</td>"
sTxt = sTxt & "<td class=""tbtags"">" & bookItem.mainstyle & "</td>"
sTxt = sTxt & "<td class=""tblanguage"">" & bookItem.languages & "</td>"
sTxt = sTxt & "<td class=""tbbuy"">" & buyURL & "</td>"
sTxt = sTxt & "</tr>" & vbCrLf
createSeriesRowtxt_old = sTxt
End Function

Private Function createSeriesRowtxt(bookItem As Books, txtSeriesRow As String, seriesNo As Integer, Optional txtTemplate As String = "") As String
Dim txtPin As String
Dim sTxt As String
txtPin = txtTemplate

Dim i As Integer
Dim j As Integer
Dim n As Integer
Dim tSeries() As String
Dim strSes As String
Dim buyURL As String
Dim sNo As String
Dim sASIN As String
Dim urlAll As String
Dim strSes2 As String

If Mid(bookItem.asin, 1, 1) = "B" Then
buyURL = "<a href=""https://www.amazon.co.jp/s/ref=series_rw_dp_labf?_encoding=UTF8&field-collection=" & bookItem.Series & "&url=search-alias%3Ddigital-text"">原版</a>"
Else
buyURL = "<a href=""https://www.amazon.co.jp/s/ref=nb_sb_noss?__mk_ja_JP=%E3%82%AB%E3%82%BF%E3%82%AB%E3%83%8A&url=search-alias%3Dstripbooks&field-keywords=" & bookItem.Series & """>原版</a>"
End If
If Left(txtSeriesRow, 1) = "@" Then txtSeriesRow = Right(txtSeriesRow, Len(txtSeriesRow) - 1)
tSeries = Split(txtSeriesRow, "@") '注意这里是按照出版顺序倒序的
For i = 0 To UBound(tSeries)
  n = UBound(tSeries) - i
  strSes = tSeries(n)
  j = InStr(strSes, "|")
  sNo = Left(strSes, j - 1)
  sASIN = Right(strSes, Len(strSes) - j)
  urlAll = urlAll & "<a href=""" & "asin/" & sASIN & ".htm"">" & sNo & "</a>&nbsp;"
Next i

If txtPin = "" Then txtPin = readutf8(EXEPATH & "template\grouplist_middle.txt")

If bookItem.cnseries <> "" Then
strSes = bookItem.cnseries
strSes2 = bookItem.Series
Else
strSes = bookItem.Series
strSes2 = bookItem.cnseries
End If

If bookItem.finished <> "" Then
strSes = strSes & "<img id=""finish"" class=""finish_img"" src=""app/ok.gif"" alt=""已完结"">"
End If
txtPin = Replace(txtPin, "#label1#", "作者：")
txtPin = Replace(txtPin, "#label2#", "系列：")
txtPin = Replace(txtPin, "#cntitle#", strSes2)
txtPin = Replace(txtPin, "#bookurl#", "series/" & bookItem.Series & ".htm")
txtPin = Replace(txtPin, "#titleF#", strSes)
txtPin = Replace(txtPin, "#seriesindexURLA#", urlAll)

txtPin = Replace(txtPin, "#lastmod#", bookItem.lastModify)
txtPin = Replace(txtPin, "#typepath#", bookItem.typePath)
txtPin = Replace(txtPin, "#asin#", bookItem.asin)
txtPin = Replace(txtPin, "#filetitle#", bookItem.filetitle)
txtPin = Replace(txtPin, "#title#", bookItem.title)
txtPin = Replace(txtPin, "#vtitle#", bookItem.vtitle)
txtPin = Replace(txtPin, "#domain#", sDomain)
txtPin = Replace(txtPin, "#domainEncode#", urlEncode(sDomain))
txtPin = Replace(txtPin, "#mainstyle#", bookItem.mainstyle)
txtPin = Replace(txtPin, "#authors#", bookItem.authors)
txtPin = Replace(txtPin, "#pubdate#", bookItem.pubdate)
txtPin = Replace(txtPin, "#coverurl#", bookItem.coverURL)
txtPin = Replace(txtPin, "#wenku#", bookItem.wenku)
txtPin = Replace(txtPin, "#languages#", bookItem.languages)
txtPin = Replace(txtPin, "#publisher#", bookItem.publisher)
txtPin = Replace(txtPin, "#series#", bookItem.Series)
txtPin = Replace(txtPin, "#series_index#", bookItem.series_index)
txtPin = Replace(txtPin, "#words#", bookItem.words)
txtPin = Replace(txtPin, "#pages#", bookItem.pages)
txtPin = Replace(txtPin, "#tags#", bookItem.tags)
txtPin = Replace(txtPin, "#comments#", bookItem.comments)
txtPin = Replace(txtPin, "#kindlePrice#", bookItem.kindlePrice)
txtPin = Replace(txtPin, "#paperPrice#", bookItem.paperPrice)
txtPin = Replace(txtPin, "#author1#", bookItem.author1)
txtPin = Replace(txtPin, "#author2#", bookItem.author2)
txtPin = Replace(txtPin, "#author1url#", bookItem.author1url)
txtPin = Replace(txtPin, "#authorsurl#", bookItem.authorsURL)
txtPin = Replace(txtPin, "#id#", bookItem.id)
txtPin = Replace(txtPin, "#cnseries#", bookItem.cnseries)
txtPin = Replace(txtPin, "#cnseries_index#", bookItem.cnseries_index)
txtPin = Replace(txtPin, "#langA#", bookItem.langA)
txtPin = Replace(txtPin, "#NumURLA#", bookItem.NumURLA)
txtPin = Replace(txtPin, "#seriesURL#", bookItem.seriesURL)
txtPin = Replace(txtPin, "#tagMain#", bookItem.tagMain)
txtPin = Replace(txtPin, "#titleURLA#", bookItem.titleURLA)
txtPin = Replace(txtPin, "#translateURL#", bookItem.translateURL)
txtPin = Replace(txtPin, "#transURLA#", bookItem.transURLA)
txtPin = Replace(txtPin, "#wenkuurl#", bookItem.wenkuURL)
txtPin = Replace(txtPin, "#buyurlA#", buyURL)

createSeriesRowtxt = txtPin & vbCrLf


End Function

Private Function createAuthorRowtxt_old(bookItem As Books, txtAuthorRow As String, Optional no As Integer = 1) As String
Dim i As Integer
Dim sindex As Integer
Dim sNo As String
Dim sASIN As String
Dim sTxt As String
Dim url As String
Dim urlAll As String
Dim tAuthors() As String
Dim j As Integer
Dim strSes As String

If Left(txtAuthorRow, 1) = "@" Then txtAuthorRow = Right(txtAuthorRow, Len(txtAuthorRow) - 1)
tAuthors = Split(txtAuthorRow, "@")
j = UBound(tAuthors) + 1

sindex = no
sTxt = "<tr>"
sTxt = sTxt & "<td class=""tbno"">" & "<a href=""https://ja.wikipedia.org/wiki/" & bookItem.author1 & """ title=""wiki"">" & sindex & "</a></td>"
sTxt = sTxt & "<td class=""tbauthor0"">" & "<a href=""" & bookItem.author1url & """>" & bookItem.author1 & "</a>" & "</td>"
sTxt = sTxt & "<td>" & str(j) & "</td>"
bookItem.filetitle = ReplaceY(bookItem.filetitle)
If bookItem.cntitle <> "" Then
sTxt = sTxt & "<td class=""tbtitle""><a href=""" & bookItem.bookurl & """ title=""" & bookItem.cntitle & """>" & bookItem.filetitle & "</a></td>"
Else
sTxt = sTxt & "<td class=""tbtitle""><a href=""" & bookItem.bookurl & """>" & bookItem.filetitle & "</a></td>"
End If
sTxt = sTxt & "<td class=""tbwenku"">" & bookItem.wenkuURL & "</td>"
sTxt = sTxt & "<td class=""tbseries"">" & bookItem.seriesURL & "</td>"
sTxt = sTxt & "<td class=""tbpublisher"">" & bookItem.publisher & "</td>"
sTxt = sTxt & "<td class=""tbpubdate"">" & bookItem.pubdate & "</td>"
sTxt = sTxt & "<td class=""tbtags"">" & bookItem.mainstyle & "</td>"
sTxt = sTxt & "<td class=""tbpages"">" & bookItem.pages & "</td>"
sTxt = sTxt & "<td class=""tblanguage"">" & bookItem.languages & "</td>"
sTxt = sTxt & "<td class=""tbkindleprice"">" & bookItem.kindlePrice & "</td>"
sTxt = sTxt & "<td class=""tbwenkuprice"">" & bookItem.paperPrice & "</td>"
sTxt = sTxt & "<td class=""tbbuy"">" & "<a href=""https://www.amazon.co.jp/s/ref=dp_byline_sr_book_1?ie=UTF8&field-author=" & bookItem.author1 & "&search-alias=books-jp&text=" & bookItem.author1 & "&sort=relevancerank"">" & "原版" & "</a></td>"
sTxt = sTxt & "</tr>" & vbCrLf
createAuthorRowtxt_old = sTxt
End Function

Private Function createAuthorRowtxt(bookItem As Books, txtAuthorRow, Optional no As Integer = 1, Optional txtTemplate As String = "") As String
Dim txtPin As String
Dim sTxt As String
txtPin = txtTemplate

Dim j As Integer
Dim strSes As String
Dim tAuthors() As String
Dim buyURL As String
buyURL = "<a href=""https://www.amazon.co.jp/s/ref=dp_byline_sr_book_1?ie=UTF8&field-author=" & bookItem.author1 & "&search-alias=books-jp&text=" & bookItem.author1 & "&sort=relevancerank"">" & "原版" & "</a>"

If Left(txtAuthorRow, 1) = "@" Then txtAuthorRow = Right(txtAuthorRow, Len(txtAuthorRow) - 1)
tAuthors = Split(txtAuthorRow, "@")
j = UBound(tAuthors) + 1
 strSes = j & "本书"
If txtPin = "" Then txtPin = readutf8(EXEPATH & "template\grouplist_middle.txt")

txtPin = Replace(txtPin, "#label1#", "最新：")
txtPin = Replace(txtPin, "#label2#", "总计：")
txtPin = Replace(txtPin, "#cntitle#", "")
txtPin = Replace(txtPin, "#bookurl#", bookItem.author1url)
txtPin = Replace(txtPin, "#titleF#", bookItem.author1)
txtPin = Replace(txtPin, "#authorsurl#", bookItem.titleFiltered)
txtPin = Replace(txtPin, "#seriesindexURLA#", strSes)

txtPin = Replace(txtPin, "#lastmod#", bookItem.lastModify)
txtPin = Replace(txtPin, "#typepath#", bookItem.typePath)
txtPin = Replace(txtPin, "#asin#", bookItem.asin)
txtPin = Replace(txtPin, "#filetitle#", bookItem.filetitle)
txtPin = Replace(txtPin, "#title#", bookItem.title)
txtPin = Replace(txtPin, "#vtitle#", bookItem.vtitle)
txtPin = Replace(txtPin, "#domain#", sDomain)
txtPin = Replace(txtPin, "#domainEncode#", urlEncode(sDomain))
txtPin = Replace(txtPin, "#mainstyle#", bookItem.mainstyle)
txtPin = Replace(txtPin, "#authors#", bookItem.authors)
txtPin = Replace(txtPin, "#pubdate#", bookItem.pubdate)
txtPin = Replace(txtPin, "#coverurl#", bookItem.coverURL)
txtPin = Replace(txtPin, "#wenku#", bookItem.wenku)
txtPin = Replace(txtPin, "#languages#", bookItem.languages)
txtPin = Replace(txtPin, "#publisher#", bookItem.publisher)
txtPin = Replace(txtPin, "#series#", bookItem.Series)
txtPin = Replace(txtPin, "#series_index#", bookItem.series_index)
txtPin = Replace(txtPin, "#words#", bookItem.words)
txtPin = Replace(txtPin, "#pages#", bookItem.pages)
txtPin = Replace(txtPin, "#tags#", bookItem.tags)
txtPin = Replace(txtPin, "#comments#", bookItem.comments)
txtPin = Replace(txtPin, "#kindlePrice#", bookItem.kindlePrice)
txtPin = Replace(txtPin, "#paperPrice#", bookItem.paperPrice)
txtPin = Replace(txtPin, "#author1#", bookItem.author1)
txtPin = Replace(txtPin, "#author2#", bookItem.author2)
txtPin = Replace(txtPin, "#author1url#", bookItem.author1url)
txtPin = Replace(txtPin, "#id#", bookItem.id)
txtPin = Replace(txtPin, "#cnseries#", bookItem.cnseries)
txtPin = Replace(txtPin, "#cnseries_index#", bookItem.cnseries_index)
txtPin = Replace(txtPin, "#langA#", bookItem.langA)
txtPin = Replace(txtPin, "#NumURLA#", bookItem.NumURLA)
txtPin = Replace(txtPin, "#seriesURL#", bookItem.seriesURL)
txtPin = Replace(txtPin, "#tagMain#", bookItem.tagMain)
txtPin = Replace(txtPin, "#titleURLA#", bookItem.titleURLA)
txtPin = Replace(txtPin, "#translateURL#", bookItem.translateURL)
txtPin = Replace(txtPin, "#transURLA#", bookItem.transURLA)
txtPin = Replace(txtPin, "#wenkuurl#", bookItem.wenkuURL)
txtPin = Replace(txtPin, "#buyurlA#", buyURL)

createAuthorRowtxt = txtPin & vbCrLf


End Function

Private Function getAuthorDetail(sAuthor As String, Optional author1 As String = "", Optional author2 As String = "", Optional author1url As String = "", Optional author2url As String = "", Optional authorsURL As String = "", Optional author1wiki As String = "", Optional author2wiki As String = "")
Dim i As Integer
Dim j As Integer
Dim otherAuthor As String
sAuthor = Replace(sAuthor, "&", ",")
i = InStr(sAuthor, ",")
If i > 0 Then
  author1 = Left(sAuthor, i - 1)
  author1 = Trim(author1)
  otherAuthor = Right(sAuthor, Len(sAuthor) - i)
  otherAuthor = Trim(otherAuthor)
'  j = InStr(otherauther, ",")
'  If j > 0 Then
'     author2 = Left(otherAuthor, j - 1)
'     otherAuthor = Right(otherAuthor, Len(otherAuthor) - i)
     If otherAuthor <> "" Then otherAuthor = "," & otherAuthor
'   End If
Else
 author1 = sAuthor '作者只有一位
End If
getAuthorDetail = author1
Call removeMark(author1)
author1url = "author/" & urlEncode(author1) & ".htm"
author1wiki = "https://www.amazon.co.jp/s/ref=dp_byline_sr_book_1?ie=UTF8&field-author=" & author1 & "&search-alias=books-jp&text=" & author1 & "&sort=relevancerank"
'https://www.amazon.co.jp/1/e/B004LVIX0S/ref=ntt_athr_dp_pel_1
authorsURL = "<a href=""" & author1url & """>" & author1 & "</a>"
If author2 <> "" Then
  Call removeMark(author2)
  author2url = "author/" & urlEncode(author2) & ".htm"
  authorsURL = authorsURL & "," & "<a href=""" & author2url & """>" & author2 & "</a>"
  author2wiki = "https://www.amazon.co.jp/s/ref=ntt_athr_dp_sr_2?_encoding=UTF8&field-author=" & author2 & "search-alias=digital-text&sort=relevancerank"
End If

authorsURL = authorsURL & otherAuthor
End Function

Public Function GetWenkuwikiURL(sname As String, newbookurl As String) As String
Dim i As Integer
getURLTab
For i = 0 To UBound(URLTab)
If URLTab(i).s0 = sname Then
    newbookurl = URLTab(i).s5
    GetWenkuwikiURL = URLTab(i).s1
    Exit For
End If
Next i
End Function

Public Function getPubfromWenku(wenku As String, Optional stitle As String) As String
Dim str As String
Dim i As Integer
getURLTab
For i = 0 To UBound(URLTab)
    If URLTab(i).s0 = wenku Then
        str = URLTab(i).s6
        If InStr(stitle, "さらさ文庫") > 0 Then
            str = "Redox media"
        End If
        Exit For
    End If
Next i
getPubfromWenku = str
End Function

    
    
Public Sub nextPage(i As Integer, txtlist As String, groupName As String, recordCount As Integer, pagelistURL As String, Optional pageMax As Integer = 100)
  Dim j As Integer
  Dim txtPath As String
  Dim pageNumber As Integer
  
       '分页 每100页
      If (i Mod pageMax) = 0 Or i = recordCount Then
      j = i / pageMax
      
      listTxt1 = readutf8(EXEPATH & "template\pagelist.txt", "UTF-8")
      listTxt1 = Replace(listTxt1, "#count#", Trim(str((j - 1) * maxItem)) & "-" & Trim(str(i)))
      listTxt1 = Replace(listTxt1, "#modifytime#", Format$(Now, "yyyy-mm-dd hh:mm:ss"))
      listTxt1 = Replace(listTxt1, "#pagelistURL#", pagelistURL)
      txtlist = listTxt1 & txtlist & listTxt2
      If j >= 2 Then
        txtPath = EXEPATH & groupName & "\" & groupName & Trim(str(j)) & ".htm"
        pagelistURL = pagelistURL & "<a href=""" & groupName & "/" & groupName & Trim(str(j)) & ".htm"">[" & Trim(str(j)) & "]</a>&nbsp;"
      Else
        txtPath = EXEPATH & groupName & "\" & groupName & ".htm"
        pagelistURL = pagelistURL & "<a href=""" & groupName & "/" & groupName & ".htm"">[1]</a>&nbsp;"
      End If
      Call writeutf8(txtPath, txtlist, "UTF-8")
     
      txtlist = ""
      End If
End Sub

Public Sub nextPage1(i As Integer, txtlist As String, folder As String, groupName As String, pagelistURL As String, recordCount As Integer, pageMax As Integer, Optional compare As Integer = 0)
  Dim p As Double
  Dim j As Integer
  Dim txtPath As String
  Dim group As String
  Dim eachPageURL As String
  Dim noStart As Integer
  Dim selfURL As String
  
  j = i \ pageMax
  p = i Mod pageMax
  If p <> 0 Then '余数不是0, 那么多加一页
    j = j + 1
  End If
  
  If txtlist = "" Then Exit Sub
  If pagelistURL = "" Then
    eachPageURL = ""
  Else
    noStart = (j - 1) * pageMax + 1
   eachPageURL = "第" & noStart & "-" & i & "项&nbsp;" & pagelistURL
  End If
  
  group = Left(groupName, 6)
  If group = "series" Then
'@   listTxt1 = readutf8(EXEPATH & "template\serieslist.txt", "UTF-8")
'@   listTxt2 = readutf8(EXEPATH & "template\grouplist2.txt", "UTF-8")
      listTxt1 = readutf8(EXEPATH & "template\grouplist_img.txt", "UTF-8")
      listTxt2 = readutf8(EXEPATH & "template\grouplist_imgend.txt", "UTF-8")
      listTxt1 = Replace(listTxt1, "#groupTitle#", "连载汇总列表")
  ElseIf group = "author" Then
'    listTxt1 = readutf8(EXEPATH & "template\authorslist.txt", "UTF-8")
'    listTxt2 = readutf8(EXEPATH & "template\grouplist2.txt", "UTF-8")
      listTxt1 = readutf8(EXEPATH & "template\grouplist_img.txt", "UTF-8")
      listTxt2 = readutf8(EXEPATH & "template\grouplist_imgend.txt", "UTF-8")
      listTxt1 = Replace(listTxt1, "#groupTitle#", "作者汇总列表")
  Else
    If j = 1 Then
      listTxt1 = readutf8(EXEPATH & "template\grouplist_img.txt", "UTF-8")
      listTxt2 = readutf8(EXEPATH & "template\grouplist_imgend.txt", "UTF-8")
    Else
      listTxt1 = readutf8(EXEPATH & "template\grouplist.txt", "UTF-8") '###
      listTxt2 = readutf8(EXEPATH & "template\grouplist2.txt", "UTF-8")
    End If
  End If
     
     listTxt1 = Replace(listTxt1, "#newbook#", "")
     listTxt1 = Replace(listTxt1, "#count#", recordCount)
     listTxt1 = Replace(listTxt1, "#PageNo#", eachPageURL)
     listTxt2 = Replace(listTxt2, "#PageNo#", eachPageURL)
     listTxt1 = Replace(listTxt1, "#domain#", sDomain)
     listTxt1 = Replace(listTxt1, "#modifytime#", Format$(Now, "yyyy-mm-dd hh:mm:ss"))
     txtlist = listTxt1 & txtlist & listTxt2
     Call removeMark(groupName)
     If j >= 2 Then
        txtPath = EXEPATH & folder & "\" & groupName & j & ".htm"
        selfURL = sDomain & folder & "/" & groupName & j & ".htm"
      Else
        txtPath = EXEPATH & folder & "\" & groupName & ".htm"
        selfURL = sDomain & folder & "/" & groupName & ".htm"
      End If
     txtlist = Replace(txtlist, "#self#", selfURL)
     Call writeutf8(txtPath, txtlist, "UTF-8", compare)
     txtlist = ""

End Sub


Public Sub nextPage2(i As Integer, txtlist As String, groupType As String, groupName As String, pagelistURL As String, recordCount As Integer, pageMax As Integer, Optional compare As Integer = 0)
  Dim p As Double
  Dim j As Integer
  Dim txtPath As String
  Dim pageNumber As Integer
  Dim grouptit As String
  Dim sqlurl As String
  Dim eachPageURL As String
  Dim noStart As Integer
  Dim selfURL As String
  
  j = i \ pageMax '整除
  p = i Mod pageMax
  If p <> 0 Then '余数不是0, 那么多加一页
    j = j + 1
  End If
  
  If txtlist = "" Then Exit Sub
  If pagelistURL = "" Then
    eachPageURL = ""
  Else
    noStart = (j - 1) * pageMax + 1
   eachPageURL = "第" & noStart & "-" & i & "项&nbsp;" & pagelistURL
  End If

     If j = 1 Then '第一页是list+img
         listTxt1 = readutf8(EXEPATH & "template\grouplist_img.txt", "UTF-8")
         listTxt2 = readutf8(EXEPATH & "template\grouplist_imgend.txt", "UTF-8")
     Else '第二页后是tab
         listTxt1 = readutf8(EXEPATH & "template\grouplist.txt", "UTF-8")
         listTxt2 = readutf8(EXEPATH & "template\grouplist2.txt", "UTF-8")
     End If
     listTxt1 = Replace(listTxt1, "#count#", recordCount)
     listTxt1 = Replace(listTxt1, "#PageNo#", eachPageURL)
     listTxt1 = Replace(listTxt1, "#modifytime#", Format$(Now, "yyyy-mm-dd hh:mm:ss"))
     listTxt1 = Replace(listTxt1, "#pagelistURL#", pagelistURL)
     grouptit = GetWenkuwikiURL(groupName, sqlurl)
     listTxt1 = Replace(listTxt1, "#groupTitle#", grouptit)
     listTxt1 = Replace(listTxt1, "#newbook#", sqlurl)
     listTxt1 = Replace(listTxt1, "#domain#", sDomain)
     listTxt2 = Replace(listTxt2, "#PageNo#", eachPageURL)
     txtlist = listTxt1 & txtlist & listTxt2
     Call removeMark(groupName)
     If j >= 2 Then
        txtPath = EXEPATH & groupType & "\" & groupName & Trim(str(j)) & ".htm"
        selfURL = sDomain & groupType & "/" & groupName & Trim(str(j)) & ".htm"
      Else
        txtPath = EXEPATH & groupType & "\" & groupName & ".htm"
        selfURL = sDomain & groupType & "/" & groupName & ".htm"
      End If
     txtlist = Replace(txtlist, "#self#", selfURL)
     Call writeutf8(txtPath, txtlist, "UTF-8", compare)
     txtlist = ""

End Sub

Public Function getPageURL(recordCount As Integer, groupName As String, groupValue As String, Optional pageMax As Integer = 50, Optional currentNo As Integer = 0) As String
Dim cNo As Integer
Dim n As Integer
Dim i As Integer
Dim currentPage As Integer
Dim url As String
If recordCount <= pageMax Then '只有一页 无需分页信息
    getPageURL = ""
    Exit Function
End If

n = recordCount \ pageMax '整除
i = recordCount Mod pageMax
If i <> 0 Then n = n + 1 '余数不是0, 那么多加一页

If currentNo > 0 Then
    currentPage = currentNo \ pageMax
    If (currentNo Mod pageMax) <> 0 Then currentPage = currentPage + 1
Else
    currentPage = 1
End If

cNo = currentPage

If n < 9 Then '全部显示,无需省略
currentPage = 1
'ElseIf n >= 9 Then
End If


Call removeMark(groupValue)

For i = currentPage To n
    If i = cNo Then '当前页不加链接
        url = url & "[" & i & "]&nbsp;"
    ElseIf i = 1 Then
        url = url & "<a href=""" & groupName & "/" & groupValue & ".htm"">[1]</a>&nbsp;"
        
    ElseIf i >= (currentPage + 1) And i <= (currentPage + 4) Then
        url = url & "<a href=""" & groupName & "/" & groupValue & i & ".htm"">[" & i & "]</a>&nbsp;"
      
    ElseIf i = (currentPage + 5) Then
      If i < n - 2 Then
        url = url & "..."
      Else
        url = url & "<a href=""" & groupName & "/" & groupValue & i & ".htm"">[" & i & "]</a>&nbsp;"
      End If
      
    ElseIf i > (currentPage + 5) And i >= (n - 2) Then
        url = url & "<a href=""" & groupName & "/" & groupValue & i & ".htm"">[" & i & "]</a>&nbsp;"
    End If
Next i
getPageURL = url
End Function


Public Function getGroupCount(groupName As String, groupValue As String, level As Integer) As Integer
  Dim Conn As New ADODB.Connection
  Dim rs As New ADODB.Recordset
  Dim sql As String
  Dim txtGroup As String
  Dim i As Integer
  Dim j As Integer
  getGroupCount = 0
  
  Conn.open xlsConnString
  
  Select Case level
  Case 0
  rs.open "Select * from [dbook$] order by pubdate desc", Conn, 1, 3
  Case 1
  rs.open "Select DISTINCT(" & groupName & ") from [dbook$]", Conn, 1, 3
  Case 2
  sql = "Select * from [dbook$] where " & groupName & "='" & groupValue & "'"
  rs.open sql, Conn, 1, 3
  Case 3 'same as 2 but sort by
  rs.open "Select * from [dbook$] order by series pubdate where " & groupName & "='" & groupValue & "'", Conn, 1, 3
  End Select
  If rs.recordCount > 0 Then
  getGroupCount = rs.recordCount
  Else
  getGroupCount = 0
  End If
  rs.Close
  Set rs = Nothing
  Conn.Close
  Set Conn = Nothing
  
End Function

Public Function getGroupRecord(groupName As String, groupValue As String, level As Integer) As ADODB.Recordset
  Dim adoConn As New ADODB.Connection
  Dim rs As New ADODB.Recordset
  Dim sql As String
  Dim txtGroup As String
  Dim i As Integer
  Dim j As Integer
  
  adoConn.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & EXEPATH & "dbook.xls;Extended Properties='Excel 8.0;HDR=Yes;IMEX=1'"
  
  Select Case level
  Case 0
  rs.open "Select * from [dbook$] order by pubdate desc", adoConn, 1, 3
  Case 1
  rs.open "Select DISTINCT(" & groupName & ") from [dbook$]", adoConn, 1, 3
  Case 2
  sql = "Select * from [dbook$] where " & groupName & "='" & groupValue & "'"
  rs.open sql, adoConn, 1, 3
  Case 3 'same as 2 but sort by
  rs.open "Select * from [dbook$] order by series pubdate where " & groupName & "='" & groupValue & "'", adoConn, 1, 3
  End Select
  
  Set getGroupRecord = rs
  rs.Close
  Set rs = Nothing
  adoConn.Close
  Set adoConn = Nothing
  
End Function

Public Function generate_zh() As Integer
 
 Dim rs As New ADODB.Recordset
 Dim pageURL As String
 Dim txtZH As String
 Dim i As Integer
 Dim iOk As Integer
 Dim n As Integer
 generate_zh = 0
  
  xlsConn.open xlsConnString
  rs.open "Select * from [dbook$] where languages='zho' order by timestamp desc", xlsConn, 1, 3
  If rs.recordCount > 0 Then
     n = rs.recordCount
     pageURL = getPageURL(n, "list", "zh", maxItem)
     rs.MoveFirst
     While Not rs.EOF
      If i = maxItem Then
        txt3 = readutf8(EXEPATH & "template\grouplist_tab.txt")
      End If
      iOk = createList(rs, 0, txtZH) '###createZH
      i = i + iOk
      If (i Mod maxItem) = 0 Then
        Call nextPage2(i, txtZH, "list", "zh", pageURL, n, maxItem, 1)
      End If
      rs.MoveNext
      DoEvents
     Wend
   Call nextPage2(i, txtZH, "list", "zh", pageURL, n, maxItem, 1)
   Label1.Caption = Label1.Caption & "已经列出了" & i & "本中文书"
  End If
  generate_zh = i
   rs.Close
  Set rs = Nothing
  xlsConn.Close
  txt3 = readutf8(EXEPATH & "template\grouplist_middle.txt")
End Function

Public Function generate_whole() As Integer
 
 Dim rs As New ADODB.Recordset
 Dim pageURL As String
 Dim txtWhole As String
 Dim i As Integer
 Dim iOk As Integer
 Dim n As Integer
 generate_whole = 0
  On Error GoTo err:
  xlsConn.open xlsConnString
  'rs.open "Select * from [dbook$] where languages='jpn' and tags like '%全本%' order by pubdate desc,series asc ", xlsConn, 1, 3
  rs.open "Select * from [dbook$] where whole='x' order by timestamp desc,pubdate desc,series asc, series_index desc", xlsConn, 1, 3
  If rs.recordCount > 0 Then
     n = rs.recordCount
     pageURL = getPageURL(n, "list", "whole", maxItem)
     rs.MoveFirst
     While Not rs.EOF
      If i = maxItem Then
        txt3 = readutf8(EXEPATH & "template\grouplist_tab.txt")
      End If
      iOk = createList(rs, 0, txtWhole)
      i = i + iOk
      If (i Mod maxItem) = 0 Then
        Call nextPage2(i, txtWhole, "list", "whole", pageURL, n, maxItem, 1)
      End If
      rs.MoveNext
      DoEvents
     Wend
   Call nextPage2(i, txtWhole, "list", "whole", pageURL, n, maxItem, 1)
   Label1.Caption = Label1.Caption & "已经列出了" & i & "本完整本"
  End If
  generate_whole = i
  rs.Close
  Set rs = Nothing
err:
  xlsConn.Close
  txt3 = readutf8(EXEPATH & "template\grouplist_middle.txt")
End Function


Public Sub checkNewBook2(Optional source As String = "")

Dim txtURL As String
Dim tWenkuItem() As String
Dim numFlag As String
Dim i As Integer
Dim j As Integer
Dim m As Integer
Dim n As Integer
Dim sline As checkWenkuURLs
Dim tURLItem() As checkWenkuURLs
Dim num As String
Dim shtm As String
Dim tempHtm As String
Dim sfile As String
Dim sname As String
Dim sWenku As String
Dim newNum As Integer
Dim totalnum As Integer
Dim sPath As String
Dim urlAll As String
Dim wenkuname As String
Dim newBookCanBuy As Integer

'当没有网络的时候停止检查
shtm = getHtmlStr("https://www.amazon.co.jp")
If shtm = "" Then Exit Sub

checkEndDate = "2100-12-31"
If source = "" Then source = txtcheck.Text
If source = "" Or Dir(source) = "" Then Exit Sub
txtURL = readutf8(source, "UTF-8")
If txtURL = "" Then Exit Sub
sWenku = readutf8(EXEPATH & "template\wenkulist1.txt", "UTF-8")
tWenkuItem = Split(txtURL, vbCrLf)
num = (UBound(tWenkuItem) + 1) / 7 - 1
ReDim st(num)
ReDim tURLItem(num)
num = 0
For i = 0 To UBound(tWenkuItem)
 m = i \ 7   '文库结构体序号
 n = i Mod 7 '余数,0文库名1日文文库名2乙女向3文库日文地址4百科地址5新书地址6永远是空格
Select Case n
 Case 0
    sline.s0 = tWenkuItem(i)
 Case 1
    sline.s1 = tWenkuItem(i)
 Case 2
    sline.s2 = tWenkuItem(i)
 Case 3
    sline.s3 = tWenkuItem(i)
 Case 4
    sline.s4 = tWenkuItem(i)
 Case 5
    sline.s5 = tWenkuItem(i)
 Case 6
    sline.s6 = tWenkuItem(i)
    tURLItem(m) = sline
End Select
 
Next i

'tURLItem = Split(txtURL, vbCrLf)
For i = 1 To (UBound(tURLItem) + 1)
   If Left(tURLItem(i - 1).s2, 2) = Left(cstyle, 2) Then '过滤掉少年向和乙女向
   shtm = getHtmlStr(tURLItem(i - 1).s5)
   newNum = 0
   If InStr(shtm, "に一致する商品がありませんでした。すべてのカテゴリーから再検索しています。") > 0 Then shtm = ""
   If shtm <> "" Then
     tempHtm = shtm
     num = Fetch(tempHtm, totalnum1, totalnum2)
     '##num = Fetch(tempHtm, "検索結果 ", "のうち")
     If num = "" Then
       '##num = Fetch(tempHtm, "a-size-base a-spacing-small a-spacing-top-small a-text-normal"">", "件の結果")
       num = Fetch(tempHtm, totalnum3, totalnum4)
     End If
     num = Replace(num, ",", "")
     If num = "" Then num = "1"
     numFlag = "#" & i & "num#"
     sWenku = Replace(sWenku, numFlag, num)
     'Call writeutf8(EXEPATH & "template\tempHtm.txt", tempHtm, "UTF-8")
     wenkuname = Fetch(tempHtm, "<span class=""a-color-state a-text-bold"">&#034;", "&#034;</span>")
     If wenkuname = "" Then
     wenkuname = i
     Else '文库名要去掉()和引号
     wenkuname = Replace(wenkuname, "(", "")
     wenkuname = Replace(wenkuname, ")", "")
     wenkuname = Replace(wenkuname, "&#034;", "")
     End If
     sname = fetchAll(tempHtm, newNum, urlAll, 1, newBookCanBuy)
     totalnum = totalnum + newNum
     sfile = sfile & "<font color=""green""><b>" & wenkuname & "</b></font>共在售" & num & "本,其中新书" & newNum & "本<br>" & vbCrLf & sname & "<br>" & vbCrLf
     numFlag = "#" & i & "pre#"
     sWenku = Replace(sWenku, numFlag, newNum)
     sname = ""
   End If
   End If
Next i

Call writeutf8(EXEPATH & "template\wenkulist.txt", sWenku, "UTF-8")
'Call fileCopy(EXEPATH & "template\wenkulist1_bk.txt", EXEPATH & "template\wenkulist1.txt")
Call writeutf8(txtbuy.Text, urlAll)
Call writeutf8(txtbuy.Text & cstyle & ".txt", urlAll)
sPath = EXEPATH & "list\newbook.htm"
sfile = "检查到共" & totalnum & "本新书在售,其中【" & newBookCanBuy & "】本今日可购买<br>" & sfile
Call writeutf8(sPath, sfile)
sPath = Replace(sPath, "\", "/")
sPath = "file:///" & sPath
Call WebBrowser1.Navigate2(sPath)
Label1.Caption = Label1.Caption & "检查到共" & totalnum & "本新书"
End Sub

Public Sub checkNewBook0(Optional source As String = "")

Dim txtURL As String
Dim tWenkuItem() As String
Dim numFlag As String
Dim i As Integer
Dim j As Integer
Dim m As Integer
Dim n As Integer
Dim sline As checkWenkuURLs
Dim num As String
Dim shtm As String
Dim tempHtm As String
Dim sfile As String
Dim sname As String
Dim sWenku As String
Dim newNum As Integer
Dim totalnum As Integer
Dim sPath As String
Dim urlAll As String
Dim wenkuname As String
Dim newBookCanBuy As Integer

'当没有网络的时候停止检查
shtm = getHtmlStr("https://www.amazon.co.jp")
If shtm = "" Then
Label1.Caption = "无网络连接."
Exit Sub
End If
checkEndDate = "2100-12-31"
sWenku = readutf8(EXEPATH & "template\wenkulist1.txt", "UTF-8")

Call getURLTab(source)

For i = 1 To (UBound(URLTab) + 1)
   If Left(URLTab(i - 1).s2, 2) = Left(cstyle, 2) Then '过滤掉少年向和乙女向
   shtm = getHtmlStr(URLTab(i - 1).s5)
   newNum = 0
   If InStr(shtm, "に一致する商品がありませんでした。すべてのカテゴリーから再検索しています。") > 0 Then shtm = ""
   If shtm <> "" Then
     tempHtm = shtm
     num = Fetch(tempHtm, "検索結果", "のうち")
     If num = "" Then
     num = Fetch(tempHtm, "a-size-base a-spacing-small a-spacing-top-small a-text-normal"">", "件の結果")
     End If
     num = Replace(num, ",", "")
     If num = "" Then num = "0"
     'Call setURLTab(URLTab(i - 1).s0, num)
     numFlag = "#" & URLTab(i - 1).s0 & "num#"
     sWenku = Replace(sWenku, numFlag, num)
     wenkuname = Fetch(tempHtm, "<span class=""a-color-state a-text-bold"">&#034;", "&#034;</span>")
     If wenkuname = "" Then
     wenkuname = i
     Else '文库名要去掉()和引号
     wenkuname = Replace(wenkuname, "(", "")
     wenkuname = Replace(wenkuname, ")", "")
     wenkuname = Replace(wenkuname, "&#034;", "")
     End If
     sname = fetchAll(tempHtm, newNum, urlAll, 1, newBookCanBuy)
     totalnum = totalnum + newNum
     sfile = sfile & "<font color=""green""><b>" & wenkuname & "</b></font>共在售" & num & "本,其中新书" & newNum & "本<br>" & vbCrLf & sname & "<br>" & vbCrLf
     Call setURLTab(URLTab(i - 1).s0, CInt(num), CInt(newNum))
     numFlag = "#" & URLTab(i - 1).s0 & "pre#"
     sWenku = Replace(sWenku, numFlag, newNum)
     sname = ""
   End If
   End If
Next i

'##Call writeutf8(EXEPATH & "template\wenkulist.txt", sWenku, "UTF-8")
Call writeutf8(txtbuy.Text, urlAll)
sPath = EXEPATH & "template\newbook.htm"
sfile = "检查到共" & totalnum & "本新书在售,其中【" & newBookCanBuy & "】本今日可购买<br>" & sfile
Call writeutf8(sPath, sfile)
sPath = Replace(sPath, "\", "/")
sPath = "file:///" & sPath
Call WebBrowser1.Navigate2(sPath)
Label1.Caption = Label1.Caption & "检查到共" & totalnum & "本新书"
End Sub


Public Function getURLTab(Optional source As String = "") As Integer
'利用全局变量URLtab
Dim txtURL As String
Dim shtm As String
Dim tWenkuItem() As String
Dim i As Integer
Dim m As Integer
Dim n As Integer
Dim num As String
Dim sline As checkWenkuURLs

On Error GoTo emptyTab
If UBound(URLTab) > 0 Then '判断数组是否为空
    getURLTab = UBound(URLTab) + 1
    Exit Function
End If

emptyTab:

If source = "" Then source = txtcheck.Text
If source = "" Or Dir(source) = "" Then Exit Function
txtURL = readutf8(source, "UTF-8")
If txtURL = "" Then Exit Function

tWenkuItem = Split(txtURL, vbCrLf)
num = (UBound(tWenkuItem) + 1) \ 8 - 1
ReDim URLTab(num)

For i = 0 To UBound(tWenkuItem)
 m = i \ 8   '文库结构体序号
 n = i Mod 8 '余数,0文库名1日文文库名2乙女向3文库日文地址4百科地址5新书地址6publisher
Select Case n
 Case 0
    sline.s0 = tWenkuItem(i)
 Case 1
    sline.s1 = tWenkuItem(i)
 Case 2
    sline.s2 = tWenkuItem(i)
 Case 3
    sline.s3 = tWenkuItem(i)
 Case 4
    sline.s4 = tWenkuItem(i)
 Case 5
    sline.s5 = tWenkuItem(i)
 Case 6
    sline.s6 = tWenkuItem(i)
    URLTab(m) = sline
End Select
 
Next i
getURLTab = i + 1
End Function


Public Function checkURL(source As String, Optional baseDate As String, Optional untilDate As String, Optional url As String = "") As Integer
Dim txtURL As String
Dim tURLItem() As String
Dim numFlag As String
Dim i As Integer
Dim num As String
Dim shtm As String
Dim tempHtm As String
Dim sfile As String
Dim sname As String
Dim sWenku As String
Dim newNum As Integer
Dim totalnum As Integer
Dim sPath As String
Dim urlAll As String

checkEndDate = untilDate
If baseDate = "" Then '只检查已经开始销售的书,不包含预售书
 baseDate = "1900-01-01"
 If untilDate = "" Then checkEndDate = Format$(Now, "yyyy-mm-dd")
ElseIf untilDate = "" Then '只检查预售的书
 checkEndDate = "2100-12-31"
End If
Txt_date.Text = Format$(baseDate, "yyyy-mm-dd")

If source <> "" Then
    txtURL = readutf8(source, "UTF-8")
    If txtURL = "" Then Exit Function
    tURLItem = Split(txtURL, vbCrLf)
ElseIf url <> "" Then
    ReDim tURLItem(0)
    tURLItem(0) = url
End If

For i = 1 To (UBound(tURLItem) + 1)
   shtm = getHtmlStr(tURLItem(i - 1), 60)
   newNum = 0
   If shtm <> "" Then
     tempHtm = shtm
     '##num = Fetch(tempHtm, "検索結果 ", "のうち")
     num = Fetch(tempHtm, totalnum1, totalnum2)
     If num = "" Then
     '##num = Fetch(tempHtm, "a-size-base a-spacing-small a-spacing-top-small a-text-normal"">", "件の結果")
       num = Fetch(tempHtm, totalnum3, totalnum4)
     End If
     num = Replace(num, ",", "")
     sname = fetchAll(tempHtm, newNum, urlAll, 2)
     sfile = sfile & i & "_" & num & " 新书" & newNum & "本<br>" & vbCrLf & sname & "<br>" & vbCrLf
     sname = ""
     totalnum = totalnum + newNum
     'Call writeutf8(EXEPATH & "list\buylist_" & Right(source, 5) & i & ".txt", urlAll, "UTF-8") '##
     'urlAll = "" '##
   End If
Next i
Call writeutf8(EXEPATH & "list\buylist_" & Right(source, 6), urlAll, "UTF-8")  '##
sPath = EXEPATH & "list\newlist_" & Right(source, 6) & ".htm"
If totalnum > 0 Then
  sfile = "共" & totalnum & "本新书" & vbCrLf & sfile
End If
Call writeutf8(sPath, sfile, "UTF-8")
sPath = Replace(sPath, "\", "/")
sPath = "file:///" & sPath
Call WebBrowser1.Navigate2(sPath)
Txt_date.Text = Format$(Now, "yyyy-mm-dd")
Label1.Caption = Label1.Caption & "检查完了,共" & totalnum & "本新书"
checkURL = totalnum
End Function

Private Function fetchAll(sourceHTM As String, Optional urlNum As Integer = 0, Optional urlAll As String, Optional imode As Integer = 1, Optional newBookCanBuy As Integer) As String
Dim tempHtm As String
Dim title As String, url As String, update As String
Dim num As Integer
Dim numOK As Integer
Do
num = FetchURL(sourceHTM, title, url, update, tempHtm, numOK, urlAll, imode, newBookCanBuy)
urlNum = urlNum + numOK
Loop Until num = 0
fetchAll = tempHtm
End Function

Private Function FetchURL(sourceHTM As String, title As String, url As String, update As String, tempHtm As String, Optional newNum As Integer = 0, Optional urlAll As String, Optional imode As Integer = 1, Optional newBookCanBuy As Integer) As Integer
Dim pDate As String
Dim beginDate As String
Dim endDate As String
Dim today As String
Dim price1 As String
Dim price2 As String

today = Format$(Now, "yyyymmdd")
title = ""
url = ""
update = ""
newNum = 0

'##title = Fetch(sourceHTM, "<a class=""a-link-normal s-access-detail-page  s-color-twister-title-link a-text-normal"" target=""_blank"" title=""", """")
title = Fetch(sourceHTM, titleget1, titleget2)
title = fromUnicode(title)
'url = Fetch(sourceHTM, "ebook/dp/", "/ref=")
url = Fetch(sourceHTM, urlget1, urlget2)
update = Fetch(sourceHTM, "<span class=""a-size-small a-color-secondary"">", "</span>")
price1 = Fetch(sourceHTM, "<span class=""a-size-base a-color-price s-price a-text-bold"">", "<")
price2 = Fetch(sourceHTM, "span class=""a-letter-space""></span>-<span class=""a-letter-space""></span>", "</span>")
price1 = Replace(price1, "￥ ", "￥")
price2 = Replace(price2, "￥ ", "")
If price2 <> "" Then price2 = "-" & price2

If title <> "" And Len(url) = 10 Then
  'url = "https://www.amazon.co.jp/dp/" & url
  If InStr(title, "&#28961;&#26009;") > 0 Then '无料
    url = getAmazonBuyFree(url)
  Else
    url = getAmazonTry(url)
  End If
  FetchURL = 1
  beginDate = Format$(Txt_date.Text, "yyyymmdd")
  endDate = Format$(checkEndDate, "yyyymmdd")
  pDate = Format$(update, "yyyymmdd")
  If pDate > beginDate And pDate <= endDate Then
    'If InStr(title, "&#65288;&#65314;&#65324;&#65289;") > 0 Then '(BL)
    'If imode = 1 Or (imode = 2 And InStr(title, "&#19968;&#36805;&#31038;&#25991;&#24235;&#12450;&#12452;&#12522;&#12473;") = 0) Then '一迅社文库ries
     If imode = 1 Or (imode = 2 And InStr(title, "一迅社文庫アイリス") = 0) Then
        If urlAll = "" Or InStr(urlAll, url) = 0 Then '过滤掉重复的链接
          newNum = 1
          tempHtm = tempHtm & "<a href=""" & url & """ target=_blank>" & title & "</a>&nbsp;" & update
          tempHtm = tempHtm & " " & price1 & price2 & "<br>" & vbCrLf
          If pDate <= today Then
            urlAll = urlAll & url & vbCrLf '放到urllist里面的永远是可以购买的
            newBookCanBuy = newBookCanBuy + 1
          End If
        End If
    End If
  End If
End If
End Function


Public Function gg(source As String) As String
Dim txtURL As String
Dim tURLItem() As String
Dim i As Integer
Dim sname As String
Dim shtm As String

txtURL = readutf8(source, "UTF-8")
If txtURL = "" Then Exit Function
tURLItem = Split(txtURL, vbCrLf)

For i = 1 To (UBound(tURLItem) + 1)
   sname = tURLItem(i - 1)
   sname = Replace(sname, kindle_1, kindle_2)
   shtm = shtm & "url = """ & sname & """" & vbCrLf
   sname = ""
Next i
Call writeutf8(source, shtm, "UTF-8")
End Function

Public Sub drill_url(txtURL As String, title As String, Optional urlNum As Integer = 0, Optional urlItems As String)  'txtURL.Text
Dim shtm As String
Dim str1 As String
Dim str2 As String
Dim str3 As String
Dim page2url As String
Dim lastNum As String
Dim i As Integer
Dim pageURL As String
Dim linkAll As String
Dim j As Long
Dim htmTemp As String
Dim fileName As String

If txtURL = "" Then Exit Sub
shtm = getHtmlStr(txtURL)
If shtm = "" Then Exit Sub

linkAll = txtURL
urlNum = 1
str1 = "/s/ref=sr_pg_2?"
str2 = """ >2</a></span>"
page2url = Fetch(shtm, str1, str2)
If page2url = "" Then GoTo drill
str1 = "<span class=""pagnDisabled"">"
str2 = "</span>"
lastNum = Fetch(shtm, str1, str2)
If lastNum = "" Then
 shtm = Replace(shtm, " ", "")
 shtm = Replace(shtm, vbCrLf, "")
 j = InStr(shtm, "<spanclass=""pagnRA"">")
 lastNum = Mid(shtm, (j - 14), 2)
 lastNum = Replace(lastNum, "<", "")
 lastNum = Replace(lastNum, ">", "")
End If
For i = 2 To Int(lastNum)
'数字替换,2替换成i
'pageURL = replace(page2url,"2",i)
str3 = "/s/ref=sr_pg_" & Trim(str(i)) & "?"
pageURL = "https://www.amazon.co.jp" & str3 & page2url
str2 = pageURL
str1 = Fetch(str2, "", "page=2") & "page="
str2 = Right(pageURL, Len(pageURL) - Len(str1) - 1)
pageURL = str1 & Trim(str(i)) & str2
pageURL = Replace(pageURL, "&amp;", "&")
linkAll = linkAll & vbCrLf & pageURL
Next i
urlNum = Int(lastNum)
drill:
fileName = title
fileName = Replace(fileName, """", "")
fileName = Replace(fileName, "'", "")
fileName = Replace(fileName, "?", "_")
fileName = Replace(fileName, "/", "_")
fileName = Replace(fileName, ":", "_")
Call writeutf8(EXEPATH & "\list\drill_" & fileName & ".txt", linkAll, "UTF-8")
Call checkURL(EXEPATH & "\list\drill_" & fileName & ".txt")
End Sub

Public Function checkLess() As Integer
 Dim adoConn1 As New ADODB.Connection
 Dim rs As New ADODB.Recordset
 Dim rs2 As New ADODB.Recordset
 Dim pageURL As String
 Dim i As Integer
 Dim iOk As Integer
 checkLess = 0
  
  adoConn1.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & EXEPATH & "dbook2.xls;Extended Properties='Excel 8.0;HDR=Yes;IMEX=1'"
  rs.open "Select web from [check$]", adoConn1, 1, 3
  rs2.open "Select local from [check$]", adoConn1, 1, 3
  If rs.recordCount > 0 Then
     rs.MoveFirst
     While Not rs.EOF
      If Trim(rs("web")) <> "" Then
       iOk = compare(Trim(rs("web")), rs2)
       If iOk = 0 Then 'do not exist such item
        pageURL = pageURL & getAmazonTry(rs("web")) & vbCrLf
        i = i + 1
       End If
      End If
      rs.MoveNext
      DoEvents
     Wend
  End If
  checkLess = i
  If i > 0 Then
    Call writeutf8(EXEPATH & "list\losturl.txt", pageURL)
    Label1.Caption = Label1.Caption & "共缺" & i & "本书未下载"
  Else
    Label1.Caption = Label1.Caption & "共检查了" & rs.recordCount & "本书都已下载"
  End If
  rs.Close
  Set rs = Nothing
  rs2.Close
  Set rs2 = Nothing
  adoConn1.Close
  Set adoConn1 = Nothing
End Function

Public Function compare(str1 As String, rs2 As ADODB.Recordset) As Integer
compare = 0
If rs2.recordCount > 0 Then
     rs2.MoveFirst
     While Not rs2.EOF
      If str1 = rs2("me") Then
         compare = 1
         Exit Function
      End If
      rs2.MoveNext
     Wend
End If
End Function

Public Function getAmazonTry(asin As String) As String
getAmazonTry = ""
If asin = "" Or Len(asin) <> 10 Then Exit Function
getAmazonTry = "https://www.amazon.co.jp/gp/digital/fiona/buy.html/ref=kics_dp_buybox_oneclick_sample?t=fiona&subtype.0=FREE_CHAPTER&itemCount=1&cor.0=JP&ASIN.0=" & asin & "&target-fiona.0=" & txt_kindle.Text
End Function

Public Function getAmazonBuyFree(asin As String) As String
getAmazonBuyFree = ""
If asin = "" Or Len(asin) <> 10 Then Exit Function
getAmazonBuyFree = "https://www.amazon.co.jp/gp/digital/fiona/buy.html/ref=kics_dp_buybox_oneclick_buy?t=fiona&itemCount=1&cor.0=JP&displayedPrice=0&displayedPriceCurrencyCode=JPY&transactionMode=one-click&isPreorder=&subtype.0=STANDARD&ASIN.0=" & asin & "&target-fiona.0=" & txt_kindle.Text
End Function

Public Sub generate_series()
  Dim rs As New ADODB.Recordset
  Dim rs2 As New ADODB.Recordset
  Dim i As Integer
  Dim k As Integer
  Dim n As Integer
  Dim iOk As Integer
  Dim sSeries As String
  Dim Series As String
  Dim txtSeries As String
  Dim txtPath As String
  Dim txtlist As String
  Dim lastASIN As Books
  Dim txtSeriesRow As String
  Dim SesList As String
  Dim sqlurl As String
  Dim grouptit As String
  Dim pageURL As String
  Dim selfURL As String
  
If Dir(EXEPATH & txtdb.Text) <> "" Then
xlsConn.open xlsConnString
Else
Exit Sub
End If

rs.open "select series,max(pubdate) as pdate,count(*) as num from [dbook$] group by series having count(*) >= 1 and series <> """" order by max(pubdate) desc", xlsConn, 1, 3
rs.MoveFirst
n = rs.recordCount
While Not rs.EOF
    i = i + 1
    sql = rs("series")
    txtSeries = ""
    txtSeriesRow = ""
    rs2.open "Select * from [dbook$] where series ='" & sql & "' order by series_index desc, pubdate desc", xlsConn, 1, 3
    If rs2.recordCount > 0 Then
     rs2.MoveFirst
     While Not rs2.EOF
       k = k + createList(rs2, 0, txtSeries) '###createSeries
       Call createSeriesRow(rs2, 0, lastASIN, txtSeriesRow)
       rs2.MoveNext
       DoEvents
     Wend
    listTxt1 = readutf8(EXEPATH & "template\grouplist_img.txt", "UTF-8") '###
    listTxt2 = readutf8(EXEPATH & "template\grouplist_imgend.txt", "UTF-8") '###
    listTxt1 = Replace(listTxt1, "#count#", Trim(str(k)))
    listTxt1 = Replace(listTxt1, "#modifytime#", Format$(Now, "yyyy-mm-dd hh:mm:ss"))
    listTxt1 = Replace(listTxt1, "#domain#", sDomain)
    listTxt1 = Replace(listTxt1, "#groupTitle#", sql)
    listTxt1 = Replace(listTxt1, "#PageNo#", "")
    Select Case sql
    Case "斎姫異聞シリーズ"
    grouptit = "斎姫異聞"
    Case "銀砂糖師"
    grouptit = "シュガーアップル·フェアリーテイル"
    Case "舞姫恋風伝 新装版"
    grouptit = "舞姫恋風伝"
    Case Else
    grouptit = sql
    End Select
    sqlurl = "https://www.amazon.co.jp/s/ref=series_rw_dp_labf?_encoding=UTF8&field-collection=" & grouptit & "&url=search-alias%3Ddigital-text"
    listTxt1 = Replace(listTxt1, "#newbook#", sqlurl)
    listTxt2 = Replace(listTxt2, "#PageNo#", "")
    txtSeries = listTxt1 & txtSeries & listTxt2
    Call removeMark(sql)
    txtPath = EXEPATH & "series\" & sql & ".htm"
    selfURL = sDomain & "series/" & sql & ".htm"
    txtSeries = Replace(txtSeries, "#self#", selfURL)
    Call writeutf8(txtPath, txtSeries, "UTF-8", 1)
    txtSeries = ""
'@    SesList = SesList & createSeriesRowtxt_old(lastASIN, txtSeriesRow, i)
    SesList = SesList & createSeriesRowtxt(lastASIN, txtSeriesRow, i)
    If (i Mod maxItem) = 0 Then
      pageURL = getPageURL(n, "list", "serieslist", maxItem, i)
      Call nextPage1(i, SesList, "list", "serieslist", pageURL, n, maxItem, 1)
    End If
    txtSeriesRow = ""
    lastASIN.asin = ""
    End If
   rs2.Close
   k = 0
  rs.MoveNext
  DoEvents
Wend
  rs.Close
  Set rs = Nothing
  Set rs2 = Nothing
  xlsConn.Close
  'Set xlsConn = Nothing
 Label1.Caption = Label1.Caption & "已经处理" & n & "个系列"
  '生成最后一页serieslist.htm
 pageURL = getPageURL(n, "list", "serieslist", maxItem, i)
 Call nextPage1(i, SesList, "list", "serieslist", pageURL, n, maxItem, 1)
 Label1.Caption = Label1.Caption & " 已生成系列汇总列表 "
End Sub

Public Sub generate_author()
  Dim rs As New ADODB.Recordset
  Dim rs2 As New ADODB.Recordset
  Dim i As Integer
  Dim k As Integer
  Dim n As Integer
  Dim iOk As Integer
  Dim txtAuthor As String
  Dim sAuthor As String
  Dim lastASIN As Books
  Dim SesList As String
  Dim author1 As String
  Dim author2 As String
  Dim author1url As String
  Dim author2url As String
  Dim author1wiki As String
  Dim author2wiki As String
  Dim auCount As String
  Dim txtPath As String
  Dim txtAuthorRow As String
  Dim sqlurl As String
  Dim grouptit As String
  Dim pageURL As String
  Dim selfURL As String
  
If Dir(EXEPATH & "dbook.xls") <> "" Then
xlsConn.open xlsConnString
Else
Exit Sub
End If
wait1000 1000

rs.open "select author,max(pubdate) as pdate,count(*) as num from [dbook$] group by author having count(*) >= 1 order by max(pubdate) desc", xlsConn, 1, 3
rs.MoveFirst
n = rs.recordCount 'n个author
While Not rs.EOF
    i = i + 1
    txtAuthor = ""
    sql = rs("author")
    rs2.open "Select * from [dbook$] where author ='" & sql & "' order by pubdate desc, series asc, series_index desc", xlsConn, 1, 3
    If rs2.recordCount > 0 Then
     rs2.MoveFirst
     While Not rs2.EOF
       k = k + createAuthor(rs2, 0, txtAuthor)
       Call createAuthorRow(rs2, 0, lastASIN, txtAuthorRow)
      rs2.MoveNext
      DoEvents
     Wend
    listTxt1 = readutf8(EXEPATH & "template\grouplist_img.txt", "UTF-8") '###
    listTxt2 = readutf8(EXEPATH & "template\grouplist_imgend.txt", "UTF-8")
    listTxt1 = Replace(listTxt1, "#groupTitle#", sql)
    listTxt1 = Replace(listTxt1, "#modifytime#", Format$(Now, "yyyy-mm-dd hh:mm:ss"))
    listTxt1 = Replace(listTxt1, "#domain#", sDomain)
    listTxt1 = Replace(listTxt1, "#count#", k)
    listTxt1 = Replace(listTxt1, "#PageNo#", "")
    sqlurl = "https://www.amazon.co.jp/s/ref=dp_byline_sr_book_1?ie=UTF8&field-author=" & sql & "&search-alias=books-jp&text=" & sql & "&sort=relevancerank"
    listTxt1 = Replace(listTxt1, "#newbook#", sqlurl)
    listTxt2 = Replace(listTxt2, "#PageNo#", "")
    txtAuthor = listTxt1 & txtAuthor & listTxt2
    Call removeMark(sql)
    txtPath = EXEPATH & "author\" & sql & ".htm"
    selfURL = sDomain & "author/" & sql & ".htm"
    txtAuthor = Replace(txtAuthor, "#self#", selfURL)
    Call writeutf8(txtPath, txtAuthor, "UTF-8", 1)
    '计数
    auCount = auCount & sql & "," & k & vbCrLf
    txtAuthor = ""
    'SesList = SesList & createAuthorRowtxt_old(lastASIN, txtAuthorRow, i)
    SesList = SesList & createAuthorRowtxt(lastASIN, txtAuthorRow, i)
    If (i Mod maxItem) = 0 Then
      pageURL = getPageURL(n, "list", "authorlist", maxItem, i)
      Call nextPage1(i, SesList, "list", "authorlist", pageURL, n, maxItem)
    End If
    txtAuthorRow = ""
    lastASIN.asin = ""
    End If

   rs2.Close
   k = 0
  rs.MoveNext
  DoEvents
Wend
  rs.Close
  Set rs = Nothing
  Set rs2 = Nothing
  xlsConn.Close
  'Set xlsConn = Nothing
  txtPath = EXEPATH & "list\aucount.csv"
  auCount = Left(auCount, Len(auCount) - 2)
'##  Call writeutf8(txtPath, auCount, "UTF-8", 1)
  Label1.Caption = Label1.Caption & "已经处理" & n & "个作者"
  '生成最后一页authorslist.htm
  pageURL = getPageURL(n, "list", "authorlist", maxItem, i)
  Call nextPage1(i, SesList, "list", "authorlist", pageURL, n, maxItem)
  Label1.Caption = Label1.Caption & " 已生成作者汇总列表 "
End Sub

Public Sub readxlsfile(imode As Integer)

  Dim rs As New ADODB.Recordset
  Dim i As Integer
  Dim j As Integer
  Dim k As Integer
  Dim m As Integer
  Dim n As Integer
  Dim iOk As Integer
  Dim txtDIV As String
  Dim sWenku As String
  Dim wenku As String
  Dim tWenku() As String
  Dim txtWenku As String
  Dim tAuthor() As String
  Dim txtAuthor As String
  Dim sAuthor As String
  Dim sSeries As String
  Dim Series As String
  Dim tSeries() As String
  Dim txtSeries As String
  Dim txtPath As String
  Dim pinTxt1 As String
  Dim pinTxt2 As String
  Dim pinTxt3 As String
  Dim txtlist As String
  Dim txtZH As String
  Dim sXML As String
  Dim ssXML As String
  Dim lastASIN As Books
  Dim txtSeriesRow As String
  Dim SesList As String
  Dim author1 As String
  Dim author2 As String
  Dim author1url As String
  Dim author2url As String
  Dim author1wiki As String
  Dim author2wiki As String
  Dim auCount As String
  Dim txtAuthorRow As String
  Dim sqlurl As String
  Dim grouptit As String
  Dim pageURL As String
  Dim createHTMOk As Integer
  
  Dim fso As Object
Set fso = CreateObject("Scripting.FileSystemObject")
' 如果目录不存在,就创建该目录
If fso.FolderExists(EXEPATH & "asin") = "False" Then fso.createfolder (EXEPATH & "asin")
If fso.FolderExists(EXEPATH & "img") = "False" Then fso.createfolder (EXEPATH & "img")
If fso.FolderExists(EXEPATH & "175") = "False" Then fso.createfolder (EXEPATH & "175")
If fso.FolderExists(EXEPATH & "240_un") = "False" Then fso.createfolder (EXEPATH & "240_un")
If fso.FolderExists(EXEPATH & "list") = "False" Then fso.createfolder (EXEPATH & "list")
If fso.FolderExists(EXEPATH & "author") = "False" Then fso.createfolder (EXEPATH & "author")
If fso.FolderExists(EXEPATH & "wenku") = "False" Then fso.createfolder (EXEPATH & "wenku")
If fso.FolderExists(EXEPATH & "series") = "False" Then fso.createfolder (EXEPATH & "series")
If fso.FolderExists(EXEPATH & "template") = "False" Then fso.createfolder (EXEPATH & "template")
If fso.FolderExists(EXEPATH & "novel") = "False" Then fso.createfolder (EXEPATH & "novel")
If fso.FolderExists(EXEPATH & "240_un\backup") = "False" Then fso.createfolder (EXEPATH & "240_un\backup")
'If fso.FolderExists(EXEPATH & "240") = "False" Then fso.createfolder (EXEPATH & "240")
If fso.FolderExists(EXEPATH & "0txt") = "False" Then fso.createfolder (EXEPATH & "0txt")
Set fso = Nothing
  
  xlsConn.open xlsConnString
  Call wait1000(1000)
  rs.open "select * from [dbook$] order by pubdate desc, series asc, series_index desc", xlsConn, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly
  If rs.recordCount > 0 Then
     rs.MoveFirst
     While Not rs.EOF
     On Error GoTo nextRS
 '   If i < 200 Then 'control count
     Select Case imode
     Case 0 To 1 '按照asin制作htm, 1按照title制作htm
      iOk = createHTM(rs, imode, createHTMOk) '0 means asin as fileame, 1 means title as filename
      If createHTMOk = 0 And iOk = 0 Then sok = sok & rs.fields("#asin") & " " & rs.fields("title") & "is wrong" & vbCrLf
      i = i + iOk
    
    Case 2 '封面中最多取100项,不然会很慢
       If i >= maxItem Then GoTo rsend
       iOk = createPIN(rs, 0, txtDIV)
       i = i + iOk
    Case 3 '分页列表
      n = rs.recordCount
      If i = maxItem Then '如果不是第一页,就用table格式 否则是list+img格式###
          txt3 = readutf8(EXEPATH & "template\grouplist_tab.txt")
      End If
      iOk = createList(rs, 0, txtlist)
      i = i + iOk
      If (i Mod maxItem) = 0 Then
        pageURL = getPageURL(n, "list", "list", maxItem, i)
        Call nextPage2(i, txtlist, "list", "list", pageURL, n, maxItem)
      End If
    Case 8 'sitemap列表
      iOk = create_urlXML(rs, 0, sXML, ssXML)
      i = i + iOk
    End Select
 '     End If 'end control count
nextRS:
      rs.MoveNext
      DoEvents
Wend
End If
  
rsend:
Select Case imode
  Case 0 To 1 '作网页
    If sok <> "" Then
    Call writeutf8(EXEPATH & "ok.txt", sok, "UTF-8")
    End If
    Label1.Caption = Label1.Caption & "已经处理" & i & "个新网页 "
'    Call writeutf8(EXEPATH & "template\log.txt", Format$(Now, "yyyy-mm-dd"), "UTF-8")
    '调用bat处理jpg变成240和175
    Shell "pic.bat", vbMinimizedNoFocus
    Txt_date.Text = Format$(Now, "yyyy-mm-dd")
    
  Case 2 '作首页
    pinTxt1 = readutf8(EXEPATH & "template\pinheader.txt", "UTF-8")
    pinTxt2 = readutf8(EXEPATH & "template\pinmiddle.txt", "UTF-8")
    pinTxt3 = readutf8(EXEPATH & "template\pinfooter.txt", "UTF-8")
    pinTxt1 = Replace(pinTxt1, "#modifytime#", Format$(Now, "yyyy-mm-dd"))
    pinTxt1 = Replace(pinTxt1, "#count#", str(rs.recordCount))
    txtDIV = pinTxt1 & txtDIV & pinTxt2 & txtDIV & pinTxt3
    Call writeutf8(EXEPATH & "index.htm", txtDIV, "UTF-8")
    Label1.Caption = Label1.Caption & "首页已经处理" & i & "个图像 "
    
  Case 3 '作分页
   pageURL = getPageURL(n, "list", "list", maxItem, i)
   Call nextPage2(i, txtlist, "list", "list", pageURL, n, maxItem)
   Label1.Caption = Label1.Caption & "已经分页列出了" & i & "本书"
   
  Case 8
    Call create_SiteMapXML(sXML, ssXML)
  End Select
 
    rs.Close
    Set rs = Nothing
    xlsConn.Close
 
txt3 = readutf8(EXEPATH & "template\grouplist_middle.txt")
End Sub

Public Sub generate_wenku()
  
  Dim rs As New ADODB.Recordset
  Dim rs2 As New ADODB.Recordset
  Dim j As Integer
  Dim k As Integer
  Dim n As Integer
  Dim pageURL As String
  Dim txtWenku As String
  Dim sfile As String
  Dim wkrowtxt As String
  Dim selfURL As String
  
If Dir(EXEPATH & "dbook.xls") <> "" Then
xlsConn.open xlsConnString
Else
Exit Sub
End If
rs.open "Select DISTINCT(wenku) from [dbook$] where wenku <> """"", xlsConn, 1, 3
j = rs.recordCount
getURLTab
rs.MoveFirst
'##sfile = readutf8(EXEPATH & "template\wenkulist.txt", "UTF-8")
While Not rs.EOF
    sql = rs("wenku")
    txtWenku = ""
    pageURL = ""
    n = 0
    rs2.open "Select * from [dbook$] where wenku='" & sql & "' order by pubdate desc, series asc, series_index desc", xlsConn, 1, 3
    txt3 = readutf8(EXEPATH & "template\grouplist_middle.txt")
    If rs2.recordCount > 0 Then
    n = rs2.recordCount
    rs2.MoveFirst
    While Not rs2.EOF
       If k = maxItem Then '如果不是第一页,就用table格式 否则是list+img格式###
          txt3 = readutf8(EXEPATH & "template\grouplist_tab.txt")
       End If
       k = k + createWenku(rs2, 0, txtWenku, 1)
       If (k Mod maxItem) = 0 Then
           pageURL = getPageURL(n, "wenku", sql, maxItem, k)
           Call nextPage2(k, txtWenku, "wenku", sql, pageURL, n, maxItem, 1)
       End If
       rs2.MoveNext
       DoEvents
    Wend
    pageURL = getPageURL(n, "wenku", sql, maxItem, k)
    Call nextPage2(k, txtWenku, "wenku", sql, pageURL, n, maxItem, 1)

    sfile = Replace(sfile, "#" & sql & "count#", k)
    'Call writeutf8(EXEPATH & "template\wenkulist.txt", listTxt1, "UTF-8")
    Call setURLTab(sql, , , CInt(n), "X")
    End If
    rs2.Close
    k = 0
    rs.MoveNext
    DoEvents
  Wend
  rs.Close
  xlsConn.Close
  Set rs2 = Nothing
  Set rs = Nothing
  'listTxt1 = readutf8(EXEPATH & "template\wenkulist.txt", "UTF-8")
  sfile = Replace(sfile, "#modifytime#", Format$(Now, "yyyy-mm-dd hh:mm:ss"))
  sfile = Replace(sfile, "#domain#", sDomain)
  '#Call writeutf8(EXEPATH & "list\wenkulist0.htm", sfile, "UTF-8")
  'Call fileCopy(EXEPATH & "template\wenkulist_bk.txt", EXEPATH & "template\wenkulist.txt")
  '#########
  sfile = readutf8(EXEPATH & "template\wenkulist0.txt", "UTF-8")
  sfile = Replace(sfile, "#modifytime#", Format$(Now, "yyyy-mm-dd hh:mm:ss"))
  sfile = Replace(sfile, "#domain#", sDomain)
  sfile = Replace(sfile, "#count#", j)
  '对URLTab按照sa列进行降序排列
  Dim i As Integer
  Dim tmax As Integer
  Dim temp As checkWenkuURLs
  Dim sfile1 As String
  n = UBound(URLTab)
  k = j
  For i = 0 To n - 1
    tmax = i
    For j = i + 1 To n
      If URLTab(j).sa > URLTab(tmax).sa Then tmax = j
    Next j
    If tmax <> i Then temp = URLTab(i): URLTab(i) = URLTab(tmax): URLTab(tmax) = temp
  Next i
  
  For i = 0 To n
    If URLTab(i).s9 > 0 Then
    URLTab(i).sb = Replace(URLTab(i).sb, "#no#", CStr(i + 1))
    sfile = sfile & URLTab(i).sb
    End If
  Next i
sfile1 = readutf8(EXEPATH & "template\grouplist2.txt", "UTF-8")
sfile = sfile & sfile1
Call writeutf8(EXEPATH & "list\wenkulist.htm", sfile, "UTF-8")
  
  Label1.Caption = Label1.Caption & "已经处理" & k & "个文库和文库列表"
  txt3 = readutf8(EXEPATH & "template\grouplist_middle.txt")
End Sub

Public Function remove_series_1() As Integer
 Dim newConn As New ADODB.Connection
 Dim rs As New ADODB.Recordset
 Dim rs2 As New ADODB.Recordset
 Dim updateSQL As String
 Dim k As Integer
 Dim m As Integer

newConn.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & EXEPATH & "dbook.xls;Extended Properties='Excel 8.0;HDR=Yes;IMEX=2'"
Call wait1000(2000)
'rs2.open updateSQL, newConn, 3, 3
'rs2.update
rs.open "select series,max(pubdate) as pdate,count(*) as num from [dbook$] group by series having count(*) = 1 and series <> """" order by max(pubdate) desc", newConn, 1, 3
remove_series_1 = rs.recordCount
If rs.recordCount <= 0 Then Exit Function
rs.MoveFirst
While Not rs.EOF
  updateSQL = "UPDATE [dbook$] SET "
  updateSQL = updateSQL & "series=''"
  updateSQL = updateSQL & " WHERE series='" & rs("series") & "'"
  On Error Resume Next
  newConn.Execute updateSQL
  If err <> 0 Then
    m = m + 1
    Debug.Print err.Description
  Else
     k = k + 1
  End If
  rs.MoveNext
  DoEvents
Wend
  rs.Close
  Set rs = Nothing
  newConn.Close
  Set newConn = Nothing
 Label1.Caption = Label1.Caption & "已清理了只有一本书的系列" & k & "个,失败" & m & "个"

End Function

Public Sub readCSVfile(imode As Integer)
  Dim adoConn As New ADODB.Connection
  Dim rs As New ADODB.Recordset
  Dim rs2 As New ADODB.Recordset
  Dim i As Integer
  Dim j As Integer
  Dim k As Integer
  Dim m As Integer
  Dim n As Integer
  Dim iOk As Integer
  Dim txtDIV As String
  Dim sWenku As String
  Dim wenku As String
  Dim tWenku() As String
  Dim txtWenku As String
  Dim tAuthor() As String
  Dim txtAuthor As String
  Dim sAuthor As String
  Dim sSeries As String
  Dim Series As String
  Dim tSeries() As String
  Dim txtSeries As String
  Dim txtPath As String
  Dim pinTxt1 As String
  Dim pinTxt2 As String
  Dim pinTxt3 As String
  Dim txtlist As String
  Dim txtZH As String
  Dim sXML As String
  Dim ssXML As String
  Dim lastASIN As Books
  Dim txtSeriesRow As String
  Dim SesList As String
  Dim author1 As String
  Dim author2 As String
  Dim author1url As String
  Dim author2url As String
  Dim author1wiki As String
  Dim author2wiki As String
  Dim auCount As String
  Dim txtAuthorRow As String
  Dim sqlurl As String
  Dim grouptit As String
  Dim pageURL As String
  Dim createHTMOk As Integer
  
  adoConn.ConnectionString = "Driver={Microsoft Text Driver (*.txt; *.csv)};DefaultDir=" & EXEPATH
  adoConn.open
  rs.open "select * from dbook.csv", adoConn, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly
  If rs.recordCount > 0 Then
     rs.MoveFirst
     While Not rs.EOF
 '   If i < 200 Then 'control count
     Select Case imode
     Case 0 '按照asin制作htm
      iOk = createHTM(rs, 0, createHTMOk) '0 means asin as fileame, 1 means title as filename
      If createHTMOk = 0 And iOk = 0 Then sok = sok & rs.fields("#asin") & " " & rs.fields("title") & "is wrong" & vbCrLf
      i = i + iOk
    Case 1 '按照title制作htm
      iOk = createHTM(rs, 1, createHTMOk)
      If createHTMOk = 0 And iOk = 0 Then sok = sok & rs.fields("#asin") & " " & rs.fields("title") & "is wrong" & vbCrLf
      i = i + iOk
    Case 2 '封面中最多取100项,不然会很慢
       If i >= maxItem Then GoTo rsend
       iOk = createPIN(rs, 0, txtDIV)
       i = i + iOk
    Case 3 'wenku列表
      If rs.fields("#wenku").Value <> "" Then
        sWenku = rs.fields("#wenku").Value
        If wenku = "" Then
           wenku = "@" & sWenku & "@"
        Else
           If InStr(wenku, "@" & sWenku & "@") = 0 Then wenku = wenku & sWenku & "@"
        End If
      End If
    Case 4 'series列表
      If rs.fields("series").Value <> "" Then
        sSeries = rs.fields("series").Value
        If Series = "" Then
           Series = "@" & sSeries & "@"
        Else
           If InStr(Series, "@" & sSeries & "@") = 0 Then Series = Series & sSeries & "@"
        End If
      End If
    Case 5 '分页列表
      n = rs.recordCount
      pageURL = getPageURL(n, "list", "list", maxItem)
      iOk = createList(rs, 0, txtlist)
      i = i + iOk
      If (i Mod maxItem) = 0 Then
        Call nextPage2(i, txtlist, "list", "list", pageURL, n, maxItem)
      End If
    Case 6 '中文书列表
      'n = getGroupCount("languages", "zho", 2)
      'pageURL = getPageURL(n, "list", "zh", maxItem)
      'iOk = createZH(rs, 0, txtZH)
      'i = i + iOk
      'If (i Mod maxItem) = 0 Then
      '  Call nextPage2(i, txtZH, "list", "zh", pageURL, n, maxItem)
      'End If
    Case 7 'author列表
      If rs.fields("authors").Value <> "" Then
        author1 = getAuthorDetail(rs.fields("authors").Value)
        If sAuthor = "" Then
           sAuthor = "@" & author1 & "@"
        Else
           If InStr(sAuthor, "@" & author1 & "@") = 0 Then sAuthor = sAuthor & author1 & "@"
        End If
      End If
     Case 8 'sitemap列表
      iOk = create_urlXML(rs, 0, sXML, ssXML)
      i = i + iOk
    End Select
      
 '     End If 'end control count
      rs.MoveNext
      DoEvents
     Wend
  End If 'end if rs empty
  
rsend:
  Select Case imode
  Case 0 '按ASIN作网页
    If sok <> "" Then
    Call writeutf8(EXEPATH & "ok.txt", sok, "UTF-8")
    End If
    Label1.Caption = Label1.Caption & "已经处理" & i & "个新网页 "
    Call writeutf8(EXEPATH & "template\log.txt", Format$(Now, "yyyy-mm-dd"), "UTF-8")
    Txt_date.Text = Format$(Now, "yyyy-mm-dd")
    
  Case 1 '按title作网页
    If sok <> "" Then
    Call writeutf8(EXEPATH & "ok.txt", sok, "UTF-8")
    End If
    Label1.Caption = Label1.Caption & "已经处理" & i & "个新网页 "
    Call writeutf8(EXEPATH & "template\log.txt", Format$(Now, "yyyy-mm-dd"), "UTF-8")
    Txt_date.Text = Format$(Now, "yyyy-mm-dd")
    
  Case 2 '作首页
    pinTxt1 = readutf8(EXEPATH & "template\pinheader.txt", "UTF-8")
    pinTxt2 = readutf8(EXEPATH & "template\pinmiddle.txt", "UTF-8")
    pinTxt3 = readutf8(EXEPATH & "template\pinfooter.txt", "UTF-8")
    pinTxt1 = Replace(pinTxt1, "#modifytime#", Format$(Now, "yyyy-mm-dd"))
    pinTxt1 = Replace(pinTxt1, "#count#", str(rs.recordCount))
    txtDIV = pinTxt1 & txtDIV & pinTxt2 & txtDIV & pinTxt3
    Call writeutf8(EXEPATH & "index.htm", txtDIV, "UTF-8")
    Label1.Caption = Label1.Caption & "首页已经处理" & i & "个图像 "
    
  Case 3 '文库
  If Left(wenku, 1) = "@" Then wenku = Right(wenku, Len(wenku) - 1)
  If Right(wenku, 1) = "@" Then wenku = Left(wenku, Len(wenku) - 1)
  tWenku = Split(wenku, "@")
  For j = 0 To UBound(tWenku)
    If j = 0 Then
      Call fileCopy(EXEPATH & "template\wenkulist.txt", EXEPATH & "template\wenkulist_bk.txt")
    End If
    sql = tWenku(j)
    txtWenku = ""
    pageURL = ""
    n = 0
    rs2.open "select * from dbook.csv", adoConn, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly
    'If rs2.RecordCount > 0 Then
    n = getGroupCount("wenku", sql, 2)
    pageURL = getPageURL(n, "wenku", sql, maxItem)
    rs2.MoveFirst
    While Not rs2.EOF
       k = k + createWenku(rs2, 0, txtWenku)
       '分页 每XXX
       If (k Mod maxItem) = 0 Then
           Call nextPage2(k, txtWenku, "wenku", sql, pageURL, n, maxItem)
       End If
       rs2.MoveNext
       DoEvents
    Wend
    Call nextPage2(k, txtWenku, "wenku", sql, pageURL, n, maxItem) '这是最后一页
    '每个文库计数
    listTxt1 = readutf8(EXEPATH & "template\wenkulist.txt", "UTF-8")
    listTxt1 = Replace(listTxt1, "#" & sql & "count#", k)
    Call writeutf8(EXEPATH & "template\wenkulist.txt", listTxt1, "UTF-8")
    'End If
    rs2.Close
    k = 0
'    If j = UBound(tWenku) Then
'      listTxt1 = readutf8(EXEPATH & "template\wenkulist.txt", "UTF-8")
'      listTxt1 = Replace(listTxt1, "#modifytime#", Format$(Now, "yyyy-mm-dd hh:mm:ss"))
'      Call writeutf8(EXEPATH & "list\wenkulist.htm", listTxt1, "UTF-8")
'      Call fileCopy(EXEPATH & "template\wenkulist_bk.txt", EXEPATH & "template\wenkulist.txt")
'    End If
  Next j
  listTxt1 = readutf8(EXEPATH & "template\wenkulist.txt", "UTF-8")
  listTxt1 = Replace(listTxt1, "#modifytime#", Format$(Now, "yyyy-mm-dd hh:mm:ss"))
  Call writeutf8(EXEPATH & "list\wenkulist.htm", listTxt1, "UTF-8")
  Call fileCopy(EXEPATH & "template\wenkulist_bk.txt", EXEPATH & "template\wenkulist.txt")
  j = UBound(tWenku) + 1
  Label1.Caption = Label1.Caption & "已经处理" & j & "个文库和文库列表"
  
  Case 4 '作系列
  If Left(Series, 1) = "@" Then Series = Right(Series, Len(Series) - 1)
  If Right(Series, 1) = "@" Then Series = Left(Series, Len(Series) - 1)
  tSeries = Split(Series, "@")
  n = getGroupCount("series", "", 1)
  pageURL = getPageURL(n, "list", "serieslist", maxItem)
  For j = 0 To UBound(tSeries)
    sql = tSeries(j)
    txtSeries = ""
    txtSeriesRow = ""
    rs2.open "select * from dbook.csv", adoConn, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly
    If rs2.recordCount > 0 Then
     rs2.MoveFirst
     While Not rs2.EOF
       k = k + createSeries(rs2, 0, txtSeries)
       Call createSeriesRow(rs2, 0, lastASIN, txtSeriesRow)
      rs2.MoveNext
      DoEvents
     Wend
    listTxt1 = readutf8(EXEPATH & "template\grouplist_img.txt", "UTF-8") '###
    listTxt2 = readutf8(EXEPATH & "template\grouplist_imgend.txt", "UTF-8")
    listTxt1 = Replace(listTxt1, "#count#", Trim(str(k)))
    listTxt1 = Replace(listTxt1, "#modifytime#", Format$(Now, "yyyy-mm-dd hh:mm:ss"))
    listTxt1 = Replace(listTxt1, "#groupTitle#", sql)
    listTxt1 = Replace(listTxt1, "#PageNo#", "")
    Select Case sql
    Case "斎姫異聞シリーズ"
    grouptit = "斎姫異聞"
    Case "銀砂糖師"
    grouptit = "シュガーアップル·フェアリーテイル"
    Case "舞姫恋風伝 新装版"
    grouptit = "舞姫恋風伝"
    Case Else
    grouptit = sql
    End Select
    sqlurl = "https://www.amazon.co.jp/s/ref=series_rw_dp_labf?_encoding=UTF8&field-collection=" & grouptit & "&url=search-alias%3Ddigital-text"
    listTxt1 = Replace(listTxt1, "#newbook#", sqlurl)
    listTxt2 = Replace(listTxt2, "#PageNo#", "")
    txtSeries = listTxt1 & txtSeries & listTxt2
    Call removeMark(sql)
    txtPath = EXEPATH & "series\" & sql & ".htm"
    Call writeutf8(txtPath, txtSeries, "UTF-8")
    txtSeries = ""
    'SesList = SesList & createSeriesRowtxt_old(lastASIN, txtSeriesRow, j)
    SesList = SesList & createSeriesRowtxt(lastASIN, txtSeriesRow, j)
    m = j + 1
    If (m Mod maxItem) = 0 Then
      Call nextPage1(m, SesList, "list", "serieslist", pageURL, n, maxItem, 1)
    End If
    txtSeriesRow = ""
    lastASIN.asin = ""
    End If
   rs2.Close
   k = 0
  Next j
  
  Label1.Caption = Label1.Caption & "已经处理" & m & "个系列"
  '生成最后一页serieslist.htm
  Call nextPage1(m, SesList, "list", "serieslist", pageURL, n, maxItem, 1)
  Label1.Caption = Label1.Caption & " 已生成总系列列表 "
  
  Case 5 '作分页
   Call nextPage2(i, txtlist, "list", "list", pageURL, n, maxItem)
   Label1.Caption = Label1.Caption & "已经分页列出了" & i & "本书"
   
  Case 6 '作中文版
   'Call nextPage2(i, txtZH, "list", "zh", pageURL, n, maxItem)
   'Label1.Caption = Label1.Caption & "已经列出了" & i & "本中文书"
   generate_zh
  Case 7 '作者列表
  If Left(sAuthor, 1) = "@" Then sAuthor = Right(sAuthor, Len(sAuthor) - 1)
  If Right(sAuthor, 1) = "@" Then sAuthor = Left(sAuthor, Len(sAuthor) - 1)
  tAuthor = Split(sAuthor, "@")
  n = getGroupCount("author", "", 1)
  pageURL = getPageURL(n, "list", "authorlist", maxItem)
  For j = 0 To UBound(tAuthor)
    sql = tAuthor(j)
    txtAuthor = ""
    rs2.open "select * from dbook.csv", adoConn, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly
    If rs2.recordCount > 0 Then
     rs2.MoveFirst
     While Not rs2.EOF
       k = k + createAuthor(rs2, 0, txtAuthor)
       Call createAuthorRow(rs2, 0, lastASIN, txtAuthorRow)
      rs2.MoveNext
      DoEvents
     Wend
    listTxt1 = readutf8(EXEPATH & "template\grouplist_img.txt", "UTF-8") '###
    listTxt2 = readutf8(EXEPATH & "template\grouplist_imgend.txt", "UTF-8")
    listTxt1 = Replace(listTxt1, "#groupTitle#", sql)
    listTxt1 = Replace(listTxt1, "#modifytime#", Format$(Now, "yyyy-mm-dd hh:mm:ss"))
    listTxt1 = Replace(listTxt1, "#count#", k)
    listTxt1 = Replace(listTxt1, "#PageNo#", "")
    sqlurl = "https://www.amazon.co.jp/s/ref=dp_byline_sr_book_1?ie=UTF8&field-author=" & sql & "&search-alias=books-jp&text=" & sql & "&sort=relevancerank"
    listTxt1 = Replace(listTxt1, "#newbook#", sqlurl)
    listTxt2 = Replace(listTxt2, "#PageNo#", "")
    txtAuthor = listTxt1 & txtAuthor & listTxt2
    Call removeMark(sql)
    txtPath = EXEPATH & "author\" & sql & ".htm"
    Call writeutf8(txtPath, txtAuthor, "UTF-8")
    '计数
    auCount = auCount & sql & "," & k & vbCrLf
    txtAuthor = ""
    txtSeries = ""
    'SesList = SesList & createAuthorRowtxt_old(lastASIN, txtAuthorRow, j)
    SesList = SesList & createAuthorRowtxt(lastASIN, txtAuthorRow, j)
    m = j + 1
    If (m Mod maxItem) = 0 Then
      Call nextPage1(m, SesList, "list", "authorlist", pageURL, n, maxItem)
    End If
    txtAuthorRow = ""
    lastASIN.asin = ""
    End If
    k = 0
    rs2.Close
  Next j
  txtPath = EXEPATH & "list\aucount.csv"
  auCount = Left(auCount, Len(auCount) - 2)
  Call writeutf8(txtPath, auCount, "UTF-8")
  Label1.Caption = Label1.Caption & "已经处理" & m & "个作者"
  '生成最后一页authorslist.htm
  Call nextPage1(m, SesList, "list", "authorlist", pageURL, n, maxItem)
  Label1.Caption = Label1.Caption & " 已生成作者汇总列表 "
  
  Case 8
    Call create_SiteMapXML(sXML, ssXML)
  End Select
 
    rs.Close
    adoConn.Close
    Set rs = Nothing
    Set rs2 = Nothing
    Set adoConn = Nothing

    End Sub

Public Sub checkNewBook(Optional source As String = "")
Dim txtURL As String
Dim tURLItem() As String
Dim numFlag As String
Dim i As Integer
Dim num As String
Dim shtm As String
Dim tempHtm As String
Dim sfile As String
Dim sname As String
Dim sWenku As String
Dim newNum As Integer
Dim totalnum As Integer
Dim sPath As String
Dim urlAll As String
Dim wenkuname As String
Dim newBookCanBuy As Integer

'当没有网络的时候停止检查
shtm = getHtmlStr("https://www.amazon.co.jp")
If shtm = "" Then Exit Sub

checkEndDate = "2100-12-31"
If source = "" Then source = txtcheck.Text
If source = "" Or Dir(source) = "" Then Exit Sub
txtURL = readutf8(source, "UTF-8")
If txtURL = "" Then Exit Sub
sWenku = readutf8(EXEPATH & "template\wenkulist1.txt", "UTF-8")
tURLItem = Split(txtURL, vbCrLf)
For i = 1 To (UBound(tURLItem) + 1)
   shtm = getHtmlStr(tURLItem(i - 1))
   newNum = 0
   If InStr(shtm, "に一致する商品がありませんでした。すべてのカテゴリーから再検索しています。") > 0 Then shtm = ""
   If shtm <> "" Then
     tempHtm = shtm
     '##num = Fetch(tempHtm, "検索結果 ", "のうち")
     num = Fetch(tempHtm, totalnum1, totalnum2)
     If num = "" Then
    '##num = Fetch(tempHtm, "a-size-base a-spacing-small a-spacing-top-small a-text-normal"">", "件の結果")
       num = Fetch(tempHtm, totalnum3, totalnum4)
     End If
     num = Replace(num, ",", "")
     If num = "" Then num = "1"
     numFlag = "#" & i & "num#"
     sWenku = Replace(sWenku, numFlag, num)
     'Call writeutf8(EXEPATH & "template\tempHtm.txt", tempHtm, "UTF-8")
     wenkuname = Fetch(tempHtm, "<span class=""a-color-state a-text-bold"">&#034;", "&#034;</span>")
     If wenkuname = "" Then
     wenkuname = i
     Else '文库名要去掉()和引号
     wenkuname = Replace(wenkuname, "(", "")
     wenkuname = Replace(wenkuname, ")", "")
     wenkuname = Replace(wenkuname, "&#034;", "")
     End If
     sname = fetchAll(tempHtm, newNum, urlAll, 1, newBookCanBuy)
     totalnum = totalnum + newNum
     sfile = sfile & "<font color=""green""><b>" & wenkuname & "</b></font>共在售" & num & "本,其中新书" & newNum & "本<br>" & vbCrLf & sname & "<br>" & vbCrLf
     numFlag = "#" & i & "pre#"
     sWenku = Replace(sWenku, numFlag, newNum)
     sname = ""
   End If
Next i

Call writeutf8(EXEPATH & "template\wenkulist.txt", sWenku, "UTF-8")
'Call fileCopy(EXEPATH & "template\wenkulist1_bk.txt", EXEPATH & "template\wenkulist1.txt")
Call writeutf8(txtbuy.Text, urlAll)
sPath = EXEPATH & "list\newbook.htm"
sfile = "检查到共" & totalnum & "本新书在售,其中【" & newBookCanBuy & "】本今日可购买<br>" & sfile
Call writeutf8(sPath, sfile)
sPath = Replace(sPath, "\", "/")
sPath = "file:///" & sPath
Call WebBrowser1.Navigate2(sPath)
Label1.Caption = Label1.Caption & "检查到共" & totalnum & "本新书"
End Sub

Public Function setURLTab(sWenku As String, Optional web_num As Double = 0, Optional order_num As Double = 0, Optional local_num As Double = 0, Optional row As String) As Integer
Dim i As Integer
Dim sa As String
Dim rowtxt As String

getURLTab
For i = 0 To UBound(URLTab)
        If URLTab(i).s0 = sWenku Then
            If web_num > 0 Then URLTab(i).s7 = web_num
            If order_num > 0 Then URLTab(i).s8 = order_num
            If local_num > 0 Then
                URLTab(i).s9 = local_num
                sa = CStr(local_num)
                '处理加权序号
                If URLTab(i).s2 = "乙女向" Then
                  URLTab(i).sa = "3" & CStr(local_num \ 10000) & Right(CStr(local_num \ 1000), 1) & Right(CStr(local_num \ 100), 1) & Right(CStr(local_num \ 10), 1) & Right(CStr(local_num), 1)
                ElseIf URLTab(i).s2 = "乙女向TL" Then
                  URLTab(i).sa = "2" & CStr(local_num \ 10000) & Right(CStr(local_num \ 1000), 1) & Right(CStr(local_num \ 100), 1) & Right(CStr(local_num \ 10), 1) & Right(CStr(local_num), 1)
                ElseIf URLTab(i).s2 = "乙女向BL" Then
                  URLTab(i).sa = "1" & CStr(local_num \ 10000) & Right(CStr(local_num \ 1000), 1) & Right(CStr(local_num \ 100), 1) & Right(CStr(local_num \ 10), 1) & Right(CStr(local_num), 1)
                Else
                  URLTab(i).sa = sa
                End If
                
            End If
            If Not IsMissing(row) Then
              rowtxt = rowtxt & "<tr class=""tbno""><td><a href=""https://ja.wikipedia.org/wiki/" & URLTab(i).s1 & """ targer=""_blank"">" & "#no#" & "</a></td>"
              rowtxt = rowtxt & "<td class=""tbpub""><a href=""" & URLTab(i).s4 & """ target=""_blank"">" & URLTab(i).s6 & "</a></td>"
              rowtxt = rowtxt & "<td class=""tbjp""><a href=""" & URLTab(i).s3 & """ target=""_blank"">" & URLTab(i).s1 & "</a></td>"
              rowtxt = rowtxt & "<td class=""tbwebku""><a href=""wenku/" & URLTab(i).s0 & ".htm"" target=""_blank"" title=""" & URLTab(i).s2 & """>" & URLTab(i).s0 & "</a></td>"
              rowtxt = rowtxt & "<td class=""tbnew""><a href=""" & URLTab(i).s5 & """ target=""_blank"">新书</a></td>"
              rowtxt = rowtxt & "<td class=""tbnum"">" & URLTab(i).s9 & "</td>"
              rowtxt = rowtxt & "<td class=""tbnum2"">" & URLTab(i).s7 & "</td>"
              rowtxt = rowtxt & "<td class=""tbnum3"">" & URLTab(i).s8 & "</td></tr>" & vbCrLf
              URLTab(i).sb = rowtxt
            End If
            Exit For
        End If 'if this url
    Next i
End Function


Public Sub runFTP()
Dim i As Integer
Dim sname As String
Dim bUp As Boolean
Dim iSuccess As Integer
Dim iOk As Integer
Dim simg As String
Dim updates
Dim today As String
Dim prex As String
If cstyle = "少年向" Then prex = "/b"
today = Format$(Now, "yyyy-mm-dd")
Txt_date.Text = today
Dim fso, folder, f, fc

'1.上传asin文件夹里的htm,img里的jpg,限于当天修改的

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(EXEPATH & "asin")
    Set fc = folder.files
    For Each f In fc
       updates = f.DateLastModified
       updates = Format$(updates, "yyyy-mm-dd")
       If updates = today Then
       sname = f.Name
       simg = Replace(sname, ".htm", ".jpg")
        bUp = fileUpload(EXEPATH & "asin\", prex & "/asin/", sname)
        'bUp = imgUpload(EXEPATH & "240_un\240\", "/asin/img/", simg)
        'Call fileMove(EXEPATH & "240_un\240\" & simg, EXEPATH & "240_done\" & simg)
        'bUp = fileUpload(EXEPATH & "175\", prex & "/asin/175/", simg)
       iOk = iOk + 1
       End If
    Next


'2.上传wenku
File1.Pattern = "*.htm"
File1.Path = EXEPATH & "wenku"
File1.Refresh
For i = 0 To File1.ListCount - 1
    On Error Resume Next
    sname = File1.List(i)
    bUp = fileUpload(EXEPATH & "wenku\", prex & "/wenku/", sname)
    iSuccess = iSuccess + 1
Next i

'3.上传series
File1.Pattern = "*.htm"
File1.Path = EXEPATH & "series"
File1.Refresh
For i = 0 To File1.ListCount - 1
    On Error Resume Next
    sname = File1.List(i)
    bUp = fileUpload(EXEPATH & "series\", prex & "/series/", sname)
    iSuccess = iSuccess + 1
Next i

'4.上传author
File1.Pattern = "*.htm"
File1.Path = EXEPATH & "author"
File1.Refresh
For i = 0 To File1.ListCount - 1
    On Error Resume Next
    sname = File1.List(i)
    bUp = fileUpload(EXEPATH & "author\", prex & "/author/", sname)
    iSuccess = iSuccess + 1
Next i

'5.上传list
File1.Pattern = "*.htm"
File1.Path = EXEPATH & "list"
File1.Refresh
For i = 0 To File1.ListCount - 1
    On Error Resume Next
    sname = File1.List(i)
    bUp = fileUpload(EXEPATH & "list\", prex & "/list/", sname)
    iSuccess = iSuccess + 1
Next i

'5.上传index.htm dbook.csv
    bUp = fileUpload(EXEPATH, prex & "/", "index.htm")
    iSuccess = iSuccess + 1
'    bUp = fileUpload(EXEPATH, prex & "/", "dbook.csv")
'    bUp = fileUpload(EXEPATH, prex & "/", "dbook.xml")
'    bUp = fileUpload(EXEPATH, prex & "/", "sitemap.txt")
'    bUp = fileUpload(EXEPATH, prex & "/", "sitemap.xml")
'    bUp = fileUpload(EXEPATH, prex & "/", "sitemap1.xml")
    
Label1.Caption = Label1.Caption & " 已经成功FTP上传 " & iOk & "个新书和 " & iSuccess & "个网页"
End Sub

Sub runFTP2()

Dim bUp As Boolean
Dim iSuccess As Integer
Dim iOk As Integer
Dim prex As String
If cstyle = "少年向" Then prex = "/b"


'1.上传asin文件夹里的htm,img里的jpg,限于当天修改的
iOk = iOk + uploadFolder("asin", 1)
Call uploadFolder("175", 1, "asin/175")

'2.上传index.htm
bUp = fileUpload(EXEPATH, prex & "/", "index.htm")
iSuccess = iSuccess + 1

'3.上传wenku
iSuccess = iSuccess + uploadFolder("wenku", 1)

'4.上传series文件夹
iSuccess = iSuccess + uploadFolder("series", 1)

'5.上传author文件夹
iSuccess = iSuccess + uploadFolder("author", 1)

'6.上传list文件夹
'iSuccess = iSuccess + uploadFolder("list", 1)
    
Label1.Caption = Label1.Caption & " 已经成功FTP上传 " & iOk & "个新书和 " & iSuccess & "个网页"
End Sub

Function uploadFolder(folderName As String, Optional compare As Integer = 0, Optional targetFolder As String) As Long
Dim i As Integer
Dim sname As String
Dim bUp As Boolean
Dim iSuccess As Integer
Dim iOk As Integer
Dim simg As String
Dim updates
Dim today As String
Dim prex As String
Dim target As String
target = targetFolder
If target = "" Then target = folderName

If cstyle = "少年向" Then prex = "/b"
today = Format$(Now, "yyyy-mm-dd")

Dim fso, folder, f, fc
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(EXEPATH & folderName)
    Set fc = folder.files
    For Each f In fc
       updates = f.DateLastModified
       updates = Format$(updates, "yyyy-mm-dd")
       If compare = 0 Or (updates = today And compare = 1) Then
        sname = f.Name
        bUp = fileUpload(EXEPATH & folderName & "\", prex & "/" & target & "/", sname)
        If bUp = True Then
          iOk = iOk + 1
        End If
       End If
    Next
uploadFolder = iOk
Set fc = Nothing
Set f = Nothing
Set folder = Nothing
Set fso = Nothing
End Function

Function fromUnicode(source As String) As String
    Dim st$, temp$ '&#23707;每8位变为一个汉字
    Dim n As Long
    Dim i As Long
    st = source
    On Error GoTo err
    Do
        n = InStr(n + 1, st, "&#"): If n = 0 Then Exit Do
        i = InStr(n + 2, st, ";")
        temp = ChrW(Mid(st, n + 2, i - n - 2))
        st = Replace(st, Mid(st, n, i - n + 1), temp)
    Loop
fromUnicode = st
Exit Function
err:
fromUnicode = source
End Function

Sub setProxy(proxyAddr As String, Optional mode As String = "HTTP")

If proxyAddr = "" Then Exit Sub
Const INTERNET_OPEN_TYPE_PRECONFIG = 0 'use registry configuration
Const INTERNET_OPEN_TYPE_DIRECT = 1    'direct to net
Const INTERNET_OPEN_TYPE_PROXY = 3     'via named proxy
Const INTERNET_OPEN_TYPE_PRECONFIG_WITH_NO_AUTOPROXY = 4   'prevent using java/script/INS
Const INTERNET_OPTION_PROXY = 38
Const INTERNET_OPTION_SETTINGS_CHANGED = 39

Dim options As INTERNET_PROXY_INFO
options.dwAccessType = INTERNET_OPEN_TYPE_PROXY
options.lpszProxy = mode & "=" & proxyAddr '"HTTP=IP:PORT"
options.lpszProxyBypass = ""
internetSetOption 0, INTERNET_OPTION_PROXY, options, LenB(options)
internetSetOption 0, INTERNET_OPTION_SETTINGS_CHANGED, 0, 0

End Sub

Public Sub check_loss_old()
Dim i As Integer
Dim j As Integer
Dim txtURLs As String
Dim txtTitle As String
Dim tURLItem() As String
Dim tTitle() As String
Dim sTemp As String


'i = checkLess()
'当没有网络的时候停止检查
sTemp = getHtmlStr("https://www.amazon.co.jp")
If sTemp = "" Then Exit Sub

txtURLs = readutf8(txtcheck.Text)
txtTitle = readutf8(EXEPATH & "template\title.txt")
If txtURLs = "" Then Exit Sub

tURLItem = Split(txtURLs, vbCrLf)
tTitle = Split(txtTitle, vbCrLf)
i = CInt(txtURL.Text)
'For i = 0 To UBound(tUrlItem)
   Call drill_url(tURLItem(i), tTitle(i), j)
'Next i
End Sub

Private Sub check_unlimit()
    Dim bookItem As Books
    Dim tags As String
    Dim bookHtm As String
    Dim url As String
    Dim urlAll As String
    Dim i As Long
    Dim j As Long
    Dim rs As New ADODB.Recordset
    
  xlsConn.open xlsConnString
  Call wait1000(1000)
  rs.open "select * from [dbook$] order by pubdate desc, series asc, series_index desc", xlsConn, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly
  If rs.recordCount = 0 Then Exit Sub
     rs.MoveFirst
     While Not rs.EOF
     On Error GoTo nextRS
      bookItem = getBook(rs)
      url = "https://www.amazon.co.jp/dp/" & bookItem.asin
      bookHtm = getHtmlStr(url)
      If bookHtm <> "" Then
        writeutf8 "1.htm", bookHtm
        If InStr(bookHtm, "Kindle Unlimited会員の方は読み放題でお楽しみいただけます") > 0 Or InStr(bookHtm, "订阅用户免费阅读") > 0 Then
            If InStr(bookItem.tags, "全本") > 0 Or InStr(bookItem.tags, "unlimit") > 0 Then
                i = i + 1
            Else
                urlAll = urlAll & "<a href=""" & url & """ target=""_blank"">" & bookItem.title & "</a>" & vbCrLf
                j = j + 1
            End If
        End If
      End If
nextRS:
      rs.MoveNext
      DoEvents
      Wend
If urlAll <> "" Then
      writeutf8 EXEPATH & "unlimit.htm", urlAll
      Label1.Caption = Label1.Caption & "已借阅过" & str(i) & "本kindle unlimited, 还缺" & str(j) & "本未借阅过"
 Else
      Label1.Caption = "已经检查完毕"
End If
End Sub


Public Function checkURL_pan(source_file As String, Optional url As String = "") As Integer
Dim txtURL As String
Dim tURLItem() As String
Dim numFlag As String
Dim i As Integer
Dim num As String
Dim shtm As String
Dim tempHtm As String
Dim sfile As String
Dim sname As String
Dim sWenku As String
Dim newNum As Integer
Dim totalnum As Integer
Dim sPath As String
Dim urlAll As String

If source_file <> "" Then
    txtURL = readutf8(source_file, "UTF-8")
    If txtURL = "" Then Exit Function
    tURLItem = Split(txtURL, vbCrLf)
ElseIf url <> "" Then
    ReDim tURLItem(0)
    tURLItem(0) = url
End If

For i = 0 To UBound(tURLItem)
   shtm = getHtmlStr_pan(tURLItem(i), 60)
   newNum = 0
   If shtm <> "" Then
     Call writeutf8(EXEPATH & "pan.htm", shtm, "UTF-8")
     'sname = fetchAll_pan(shtm, newNum, urlAll)
     sfile = sfile & i & "_" & num & " 新资源" & newNum & "个<br>" & vbCrLf & sname & "<br>" & vbCrLf
     sname = ""
     totalnum = totalnum + newNum
   End If
Next i
Call writeutf8(EXEPATH & "list\panbuylist_" & Right(source_file, 6), urlAll, "UTF-8")  '##
sPath = EXEPATH & "list\pannewlist_" & Right(source_file, 6) & ".htm"
If totalnum > 0 Then
  sfile = "共" & totalnum & "个新项目" & vbCrLf & sfile
End If
Call writeutf8(sPath, sfile, "UTF-8")
sPath = Replace(sPath, "\", "/")
sPath = "file:///" & sPath
'Call WebBrowser1.Navigate2(sPath)
Label1.Caption = Label1.Caption & "检查完了,共" & totalnum & "项新资源"
checkURL_pan = totalnum
End Function

Private Function fetchAll_pan(sourceHTM As String, Optional urlNum As Integer = 0, Optional urlAll As String) As String
Dim tempHtm As String
Dim title As String, url As String, update As String, author As String, author_name As String
Dim num As Integer
Dim numOK As Integer
Do
num = FetchURL_pan(sourceHTM, title, url, update, tempHtm, numOK, urlAll, author, author_name)
'insert into excel?
urlNum = urlNum + numOK
Loop Until num = 0
fetchAll_pan = tempHtm
End Function

Private Function FetchURL_pan(sourceHTM As String, title As String, url As String, update As String, tempHtm As String, Optional newNum As Integer = 0, Optional urlAll As String, Optional author As String, Optional author_name As String, Optional author_face As String) As Integer

Dim today As String
Dim l_index As String
Dim shareid As String

title = ""
url = ""
update = ""
newNum = 0
author = ""
author_name = ""

today = Format$(Now, "yyyymmdd")
l_index = Fetch(sourceHTM, "<li class=""feed-dynamic-item"" _index=""", """")
shareid = Fetch(sourceHTM, "_shareid=""", """>")
url = Fetch(sourceHTM, "<div class=""feed-dynamic-header""><a class=""title"" href=""", """ target=""_blank"" title=")
title = Fetch(sourceHTM, """", """>")
title = fromUnicode(title)
update = Fetch(sourceHTM, "<span class=""time fr"">", "</span>")
update = today & " " & update
author = Fetch(sourceHTM, "<a hidefocus=""true"" id=""feedShareOwner"" data-uk=""", """")
author_face = Fetch(sourceHTM, "<img class=""face"" src=""", """>")
author_name = Fetch(sourceHTM, "<span class=""name fl breviary"" title=""", """>")
author_name = fromUnicode(author_name)

If title <> "" And Len(url) = 11 Then
  url = "https://pan.baidu.com" & url & "?uk=" & author & "&shareid=" & shareid
  FetchURL_pan = 1
        If urlAll = "" Or InStr(urlAll, url) = 0 Then '过滤掉重复的链接
          newNum = 1
          urlAll = urlAll & url & vbCrLf
          tempHtm = tempHtm & "<a href=""" & url & """ target=_blank>" & title & "</a>&nbsp;" & update
          tempHtm = tempHtm & " @<a href=""https://pan.baidu.com/share/home?uk=" & author & """ target=_blank>" & author_name & "</a>" & "<br>" & vbCrLf
        End If
End If
End Function

Public Function getHtmlStr_pan(strURL As String, Optional timeout As Long) As String
Dim json, records
Dim response_str As String
Dim response_header As String
Dim cookie_str As String
Dim stime, ntime
Dim XmlHttp 'As MSXML2.XMLHTTP60
If strURL = "" Or strURL = vbCrLf Or Len(strURL) < 2 Then Exit Function
cookie_str = readutf8(App.Path & "\cookie.txt")
Set XmlHttp = CreateObject("Microsoft.XMLHTTP")
XmlHttp.open "GET", strURL, False
XmlHttp.setRequestHeader "Host", "pan.baidu.com"
XmlHttp.setRequestHeader "Connection", "keep-alive"
XmlHttp.setRequestHeader "Accept", "application/json, text/javascript, */*; q=0.01"
XmlHttp.setRequestHeader "X-Requested-With", "XMLHttpRequest"
XmlHttp.setRequestHeader "User-Agent", " Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/63.0.3239.84 Safari/537.36"
XmlHttp.setRequestHeader "Referer", "https://pan.baidu.com/pcloud/home"
XmlHttp.setRequestHeader "Accept-Encoding", "gzip, deflate, br"
XmlHttp.setRequestHeader "Accept-Language", "en-US,en;q=0.8,zh-CN;q=0.6,zh;q=0.4,ja;q=0.2,ko;q=0.2,zh-TW;q=0.2,fr-FR;q=0.2,fr;q=0.2"
XmlHttp.setRequestHeader "Cookie", cookie_str
On Error GoTo Err_net
stime = Now '获取当前时间
XmlHttp.send
While XmlHttp.ReadyState <> 4
  DoEvents
  ntime = Now '获取循环时间
  If timeout <> 0 Then
    If DateDiff("s", stime, ntime) > timeout Then
      getHtmlStr_pan = ""
      Debug.Print "timeout:" & strURL & vbCrLf
      Exit Function '判断超出3秒即超时退出过程
    End If
  End If
Wend
If XmlHttp.StatusText = "OK" Then
response_str = BytesToBstr(XmlHttp.responseBody, "UTF-8")
json = JSONParse("errno", response_str)
json = JSONParse("records.length", response_str)
records = JSONParse("records[0]", response_str)
getHtmlStr_pan = ""
End If
Set XmlHttp = Nothing
Err_net:
End Function

Public Function JSONParse(ByVal JSONPath As String, ByVal JSONString As String) As Variant
    Dim json As Object
    Set json = CreateObject("MSScriptControl.ScriptControl")
    json.language = "JScript"
    JSONParse = json.eval("JSON=" & JSONString & ";JSON." & JSONPath & ";")
    Set json = Nothing
End Function

Public Function get_Unix_timestamp(Optional time) As String
Dim vtime
If time = "" Then
vtime = Now()
End If
get_Unix_timestamp = DateDiff("s", "01/01/1970 00:00:00", vtime)
End Function

Public Function retrive_timestamp(Unix_timestamp As String) As String
Dim vtime
vtime = DateAdd("s", Unix_timestamp, "01/01/1970 00:00:00")
retrive_timestamp = vtime
End Function


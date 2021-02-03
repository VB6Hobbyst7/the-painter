Attribute VB_Name = "Main1"
'FIXIT: Use Option Explicit to avoid implicitly creating variables of type Variant         FixIT90210ae-R383-H1984
'FIXIT: Declare 'LangA' with an early-bound data type                                      FixIT90210ae-R1672-R1B8ZE
Public LangA
Public lgT(500) As String
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long '执行文件的声明
Public ver1, newVer
Public ShowTool As Boolean
Public ActiveTool


Sub Main()
On Error GoTo CLOLogo
'Read Lang


        LangA = "lgc"

ver1 = App.Major & "." & App.Minor
lgT(9) = "版本: " & ver1 '版本号

lgT(8) = "" 'Website

'Start Program

    LangMain '语言开始

Unload FLogo
FLogo.Show
Exit Sub
CLOLogo:
Unload FLogo
MsgBox "Error!", vbInformation
End Sub

Public Sub APrun()
On Error Resume Next
Unload FWel
FLoading2.Show
FWhole.mAP.Enabled = False
FWhole.mAP.Checked = True
End Sub

Public Sub ICONrun()
On Error Resume Next
Unload FWel
FWhole.mIC.Enabled = False
FWhole.mIC.Checked = True
Load Form1
End Sub
Public Sub QSGRun()
If Dir(App.Path + "\plugin.exe") = "" Then
MsgBox "请重新执行安装文件，找不到plugin.exe", vbExclamation, "Error"
Else
ShellExecute FTemp1.hWnd, "Open", App.Path + "\Plugin.exe", "sg", App.Path, 1
FWhole.WindowState = 1
End If
End Sub


Public Sub FPrun()
On Error Resume Next
    If Dir(App.Path + "\Fpainter.exe") = "" Then
    MsgBox "请重新执行安装文件，找不到Fpainter.exe", vbCritical, "Error"
    Else
    Unload FWel
    ShellExecute FTemp1.hWnd, "Open", App.Path + "\Fpainter.exe", "StartFP", App.Path, 1
    IconTray = getstring(HKEY_CURRENT_USER, "Software\Sicasoft\Sicapic", "TrayIcon")
If IconTray = "" Then
FWhole.Hide
Else
FWhole.WindowState = 1
End If
    End If
End Sub

Sub LangMain()
'''''''''''''''''''''''''
'语言开始
lgT(0) = "图像已更改，您要在退出前保存图象吗？"
lgT(1) = "您要保存编辑中的图片吗？"
lgT(2) = "像素: "
lgT(3) = "所有图片文件"
lgT(4) = "请选择您需编辑的图片:"
lgT(5) = "完成(&F)"
lgT(6) = "取消(&C)"
lgT(7) = "下次启动程序时不启动本向导"

'=========================
lgT(10) = "文件(&F)"
lgT(11) = "编辑(&E)"
lgT(12) = "图片(&P)"
lgT(13) = "颜色(&C)"
lgT(14) = "一般滤镜(&I)"
lgT(15) = "特殊滤镜(&A)"
lgT(16) = "修饰(&S)"
lgT(17) = "填充(&M)"
lgT(18) = "变形(&D)"
lgT(19) = "渐变边(&G)"
lgT(20) = "文字(&T)"
lgT(21) = "帮助(&H)"
'======================
lgT(22) = "打开文件..."
lgT(23) = "文件另存为..."
lgT(24) = "E"
lgT(25) = "-"
lgT(26) = "打印图片..."
lgT(27) = "启动向导..."
lgT(28) = "-"
lgT(29) = "Lang"
lgT(30) = "退出..."
'1-1 Select lang
lgT(31) = "En"
lgT(32) = "中"


'2

lgT(41) = "撤消" 'cancel1
lgT(42) = ""
lgT(43) = "调整选择区域"
lgT(44) = "取消选择 (鼠标右键)"
lgT(45) = "全选"
'3
lgT(56) = "水平翻转图片"
lgT(57) = "垂直翻转图片"
lgT(58) = "-"
lgT(59) = "左侧水平平分图象"
lgT(60) = "右侧水平平分图象"
lgT(61) = "上部垂直平分图象"
lgT(62) = "下部垂直平分图象"
'4
lgT(73) = "调整色阶..."
lgT(74) = "-"
lgT(75) = "去除颜色"
lgT(76) = "-"
lgT(77) = "转换基调"
lgT(78) = "-"
lgT(79) = "亮度..."
lgT(80) = "对比度..."
lgT(81) = "-"
lgT(82) = "底片反色"
lgT(83) = "-"
lgT(84) = "颠倒颜色"
lgT(85) = "-"
lgT(86) = "灰度"
lgT(87) = ""
lgT(88) = ""
lgT(89) = ""
lgT(90) = ""
lgT(91) = ""
lgT(92) = ""
lgT(93) = ""
lgT(94) = ""
lgT(95) = ""
lgT(96) = ""
lgT(97) = ""
lgT(98) = ""
lgT(99) = ""
'4-1
lgT(100) = "去除红色"
lgT(101) = "去除绿色"
lgT(102) = "去除蓝色"
'4-2 RGB
'4-3
lgT(103) = "颠倒红色"
lgT(104) = "颠倒绿色"
lgT(105) = "颠倒蓝色"
'5
lgT(106) = "浮雕"
lgT(107) = "特殊浮雕"
lgT(108) = "-"
lgT(109) = "雕刻"
lgT(110) = "-"
lgT(111) = "氖化"
lgT(112) = "-"
lgT(113) = "模糊"
lgT(114) = "模糊 (多)"
lgT(115) = "-"
lgT(116) = "锐化"
lgT(117) = "-"
lgT(118) = "添加杂色..."
lgT(119) = "-"
lgT(120) = "腐蚀..."
lgT(121) = "煞风..."
lgT(122) = "添加雾道..."
lgT(123) = "添加噪音"
lgT(124) = "-"
lgT(125) = "冻结"
lgT(126) = "冻结 (多)"
lgT(127) = "-"
lgT(128) = "黑白化"
lgT(129) = "-"
lgT(130) = "软型染色"
lgT(131) = "-"
lgT(132) = "硬型染色"
lgT(133) = ""
lgT(134) = ""
lgT(135) = ""
lgT(136) = ""
lgT(137) = ""
lgT(138) = ""
'5-1
lgT(139) = "突出红色"
lgT(140) = "突出绿色"
lgT(141) = "突出蓝色"
'5-2
lgT(142) = "模式 1"
lgT(143) = "模式 2"
lgT(144) = "模式 3"
lgT(169) = "灰度模式"
'5-3
lgT(145) = "红色"
lgT(146) = "绿色"
lgT(147) = "橙色"
lgT(148) = "黄色"
lgT(149) = "紫色"
'5-4
lgT(150) = "红色"
lgT(151) = "绿色"
lgT(152) = "蓝色"
lgT(153) = "黄色"
'6
lgT(154) = "灰色"
lgT(155) = "茶色"
lgT(156) = "水底"
lgT(157) = "黄色"
lgT(158) = "木炭"
lgT(159) = "夜色"
lgT(160) = "日食"
lgT(161) = "紫色"
lgT(162) = "幽灵"
lgT(163) = "虚幻灰暗"
lgT(164) = "强化冷暖色"
lgT(165) = "惨烈"
lgT(166) = "斑点化"
lgT(167) = "暴光过度"
lgT(168) = "纸质杂色"
'lgT(169) 有位
lgT(170) = ""
lgT(171) = ""
lgT(172) = ""
lgT(173) = ""
lgT(174) = ""
lgT(175) = ""
lgT(176) = ""
lgT(177) = ""
lgT(178) = ""
lgT(179) = ""
'7
lgT(180) = "侧光水平百叶窗..."
lgT(181) = "侧光垂直百叶窗..."
lgT(182) = "平光水平百叶窗..."
lgT(183) = "平光垂直百叶窗..."
lgT(184) = "-"
lgT(185) = "水平线"
lgT(186) = "垂直线"
lgT(187) = "正方形网状线"
lgT(188) = "方形层层深入线"
lgT(189) = "圆形层层深入线"
lgT(190) = "右斜线"
lgT(191) = "左斜线"
lgT(192) = "交叉斜线"
lgT(193) = "水平波浪线"
lgT(194) = "垂直波浪线"
lgT(195) = "向上水平波浪折线"
lgT(196) = "向左垂直波浪折线"
lgT(197) = "向下水平波浪折线"
lgT(198) = "向右垂直波浪折线"
lgT(199) = "-"
lgT(200) = "添加边框"
lgT(201) = ""
lgT(202) = ""
lgT(203) = ""
lgT(204) = ""
lgT(205) = ""
lgT(206) = ""
lgT(207) = ""
lgT(208) = ""
lgT(209) = ""
'7-1
lgT(210) = "单色边框"
lgT(211) = "扩边单色边框"
lgT(212) = "渐变边框 1"
lgT(213) = "扩边渐变边框 1"
lgT(214) = "渐变边框 2"
lgT(215) = "扩边渐变边框 2"
lgT(216) = "-"
lgT(217) = "圆形单色边框"
lgT(218) = "圆形渐变边框 1"
lgT(219) = "圆形渐变边框 2"
'8
lgT(220) = "单色填充"
lgT(221) = "渐变填充 1"
lgT(222) = "渐变填充 2"
lgT(223) = "方形填充 1"
lgT(224) = "方形填充 2"
lgT(225) = "圆形填充 1"
lgT(226) = "圆形填充 2"
lgT(227) = "-"
lgT(228) = "居中填充图片"
lgT(229) = "平铺填充图片"
lgT(230) = ""
lgT(231) = ""
lgT(232) = ""
lgT(233) = ""
lgT(234) = ""
lgT(235) = ""
lgT(236) = ""
lgT(237) = ""
lgT(238) = ""
lgT(239) = ""
'9
lgT(240) = "层层递进"
lgT(241) = "-"
lgT(242) = "马赛克"
lgT(243) = "圆形马赛克"
lgT(244) = "-"
lgT(245) = "水平波浪"
lgT(246) = "完全水平波浪"
lgT(247) = "垂直波浪"
lgT(248) = "完全垂直波浪"
lgT(249) = "-"
lgT(250) = "分割图片"

'10
lgT(251) = "增加渐变边 (少)"
lgT(252) = "增加渐变边 (多)"
lgT(253) = "自定义增加"
lgT(254) = ""
lgT(255) = ""
lgT(256) = ""
lgT(257) = ""
lgT(258) = ""
lgT(259) = ""

'10-1
lgT(260) = "左边 1"
lgT(261) = "左边 2"
lgT(262) = "左边 3"
lgT(263) = "-"
lgT(264) = "右边 1"
lgT(265) = "右边 2"
lgT(266) = "右边 3"
lgT(267) = "-"
lgT(268) = "上边 1"
lgT(269) = "上边 2"
lgT(270) = "上边 3"
lgT(271) = "-"
lgT(272) = "底边 1"
lgT(273) = "底边 2"
lgT(274) = "底边 3"
lgT(275) = "-"
lgT(276) = "左右 1"
lgT(277) = "左右 2"
lgT(278) = "-"
lgT(279) = "上下1"
lgT(280) = "上下2"

'11
lgT(281) = "添加文字"
lgT(282) = ""
lgT(283) = ""
lgT(284) = ""
lgT(285) = ""

'12
lgT(286) = "帮助文档..."
lgT(287) = ""
lgT(288) = "关于..."
lgT(289) = ""
lgT(290) = ""
lgT(291) = ""
lgT(292) = ""
lgT(293) = ""
'==========================
lgT(294) = "放大辅助   "


lgT(295) = "历史操作   "
lgT(296) = "图象信息   "
lgT(297) = "选择区域   "
lgT(298) = "常用操作   "

lgT(299) = "调整、更改图片颜色"
lgT(300) = "添加百叶窗、直线、斜线或网格"
lgT(301) = "添加马赛克或波浪特效"
lgT(302) = "添加文字..."
lgT(303) = "图片的翻转、平分操作"
lgT(304) = "普通滤镜操作"

lgT(305) = "  文件名: "
lgT(306) = "  图象宽度: "
lgT(307) = "  图象高度: "
lgT(308) = "您所编辑的图片尺寸太大，编辑时可能程序会无响应"
lgT(309) = "选择区域："
lgT(310) = "正在读取颜色..."
lgT(311) = "完成"
lgT(312) = "正在应用于图象..."

lgT(313) = ""
'400 is here!
lgT(400) = "欢迎使用《小画家》之 图片编辑" & vbCrLf & vbCrLf & "图片编辑器能美化您现有的图片，" & vbCrLf & "如变形操作、颜色编辑、滤镜、添加阴影文字" & vbCrLf & "以及其他修饰功能" & vbCrLf & vbCrLf & "如需帮助，请按 F1" & vbCrLf & vbCrLf & vbCrLf & "开始编辑图片，请点左侧的按钮"
lgT(314) = "图片编辑"

lgT(315) = "颜色"
lgT(316) = "变形"
lgT(317) = "修饰"

lgT(318) = "对不起，您刚刚没有选定要打开的文件"
lgT(319) = "打印命令：将使用默认打印机的默认纸张大小打印"
lgT(320) = "是否横向打印？"
lgT(321) = "您取消了打印"

lgT(322) = "滤镜"
lgT(323) = "填充颜色"
lgT(324) = "填充图片"

lgT(325) = "没有发现选择区域"
lgT(326) = "打开图片时无相应"
lgT(327) = "载入中，请等待..."
lgT(328) = "打开一个图片"
lgT(329) = "打开一个图片"

''''Copy
lgT(330) = "程序已经在运行"
'FEcho
lgT(331) = "层层递进"
lgT(332) = "设置"
lgT(333) = "默认"
lgT(334) = "数量"
lgT(335) = "递减"
lgT(336) = "偏移中心-水平"
lgT(337) = "偏移中心-垂直"
lgT(338) = "生成预览"
lgT(339) = "确定"
lgT(340) = "撤消"
lgT(341) = "偏移中心"

'FText
lgT(342) = "请在这里输入文字"
lgT(343) = "填加文字"
lgT(344) = "字体设置"
lgT(345) = "大小"
lgT(346) = "文字阴影"
lgT(347) = "文字颜色"
lgT(348) = "阴影颜色"
lgT(349) = "阴影设置"
lgT(350) = "坐标X"
lgT(351) = "坐标Y"
lgT(352) = "位置"
lgT(353) = "居左"
lgT(354) = "居中"
lgT(355) = "居右"

'Fcolor
lgT(356) = "透明度"
lgT(357) = "拉伸至全屏"
lgT(358) = "按原大小居中"
lgT(359) = "居中填充图片"
lgT(360) = "确定"
lgT(361) = "取消"
lgT(362) = "颜色"
lgT(363) = "振幅"
lgT(364) = "距离"
lgT(365) = "蓝"
lgT(366) = "绿"
lgT(367) = "红"
lgT(368) = "高度"
lgT(369) = "宽度"
lgT(370) = "调整完后，按[撤消]返回"
lgT(371) = "例图"
lgT(372) = "波浪"
lgT(373) = "生成预览"
lgT(374) = "请稍候..."

lgT(375) = "浏览(&R)..."
lgT(376) = ""
lgT(377) = ""
lgT(378) = ""
lgT(379) = ""
lgT(380) = ""
lgT(381) = ""
lgT(382) = ""
lgT(383) = "参赛版" '保留位
lgT(384) = ""
lgT(385) = ""
lgT(386) = ""
lgT(387) = ""
lgT(388) = ""
lgT(389) = ""
lgT(390) = ""
lgT(391) = ""
lgT(392) = ""
lgT(393) = ""
lgT(394) = ""
lgT(395) = ""
lgT(396) = ""
lgT(397) = ""
lgT(398) = ""
lgT(399) = "您现在要关闭图片编辑吗？"
'400 401有
lgT(402) = ""
'lgT(403) = "
lgT(404) = "小画家"




'语言结束
End Sub

Sub RegMain()

         lgT(383) = lgT(382)
          lgT(381) = "YO"
End Sub

Public Sub ToolsMini()
If ShowTool = True Then
    If FWhole.WindowState = 1 Or FMain.WindowState = 1 Then
    ToolInvisible
    Else
    ToolVisible
    End If
End If
End Sub

Public Sub ToolVisible()
ToolXY.Visible = True
ToolRedo.Visible = True
ToolZoom.Visible = True
End Sub

Public Sub ToolInvisible()
ToolXY.Visible = False
ToolRedo.Visible = False
ToolZoom.Visible = False
End Sub

Public Sub ToolsClose()
Unload ToolXY
Unload ToolRedo
Unload ToolZoom
End Sub

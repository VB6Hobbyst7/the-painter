Attribute VB_Name = "Main1"
'FIXIT: Use Option Explicit to avoid implicitly creating variables of type Variant         FixIT90210ae-R383-H1984
'FIXIT: Declare 'LangA' with an early-bound data type                                      FixIT90210ae-R1672-R1B8ZE
Public LangA
Public lgT(500) As String
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long 'ִ���ļ�������
Public ver1, newVer
Public ShowTool As Boolean
Public ActiveTool


Sub Main()
On Error GoTo CLOLogo
'Read Lang


        LangA = "lgc"

ver1 = App.Major & "." & App.Minor
lgT(9) = "�汾: " & ver1 '�汾��

lgT(8) = "" 'Website

'Start Program

    LangMain '���Կ�ʼ

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
MsgBox "������ִ�а�װ�ļ����Ҳ���plugin.exe", vbExclamation, "Error"
Else
ShellExecute FTemp1.hWnd, "Open", App.Path + "\Plugin.exe", "sg", App.Path, 1
FWhole.WindowState = 1
End If
End Sub


Public Sub FPrun()
On Error Resume Next
    If Dir(App.Path + "\Fpainter.exe") = "" Then
    MsgBox "������ִ�а�װ�ļ����Ҳ���Fpainter.exe", vbCritical, "Error"
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
'���Կ�ʼ
lgT(0) = "ͼ���Ѹ��ģ���Ҫ���˳�ǰ����ͼ����"
lgT(1) = "��Ҫ����༭�е�ͼƬ��"
lgT(2) = "����: "
lgT(3) = "����ͼƬ�ļ�"
lgT(4) = "��ѡ������༭��ͼƬ:"
lgT(5) = "���(&F)"
lgT(6) = "ȡ��(&C)"
lgT(7) = "�´���������ʱ����������"

'=========================
lgT(10) = "�ļ�(&F)"
lgT(11) = "�༭(&E)"
lgT(12) = "ͼƬ(&P)"
lgT(13) = "��ɫ(&C)"
lgT(14) = "һ���˾�(&I)"
lgT(15) = "�����˾�(&A)"
lgT(16) = "����(&S)"
lgT(17) = "���(&M)"
lgT(18) = "����(&D)"
lgT(19) = "�����(&G)"
lgT(20) = "����(&T)"
lgT(21) = "����(&H)"
'======================
lgT(22) = "���ļ�..."
lgT(23) = "�ļ����Ϊ..."
lgT(24) = "E"
lgT(25) = "-"
lgT(26) = "��ӡͼƬ..."
lgT(27) = "������..."
lgT(28) = "-"
lgT(29) = "Lang"
lgT(30) = "�˳�..."
'1-1 Select lang
lgT(31) = "En"
lgT(32) = "��"


'2

lgT(41) = "����" 'cancel1
lgT(42) = ""
lgT(43) = "����ѡ������"
lgT(44) = "ȡ��ѡ�� (����Ҽ�)"
lgT(45) = "ȫѡ"
'3
lgT(56) = "ˮƽ��תͼƬ"
lgT(57) = "��ֱ��תͼƬ"
lgT(58) = "-"
lgT(59) = "���ˮƽƽ��ͼ��"
lgT(60) = "�Ҳ�ˮƽƽ��ͼ��"
lgT(61) = "�ϲ���ֱƽ��ͼ��"
lgT(62) = "�²���ֱƽ��ͼ��"
'4
lgT(73) = "����ɫ��..."
lgT(74) = "-"
lgT(75) = "ȥ����ɫ"
lgT(76) = "-"
lgT(77) = "ת������"
lgT(78) = "-"
lgT(79) = "����..."
lgT(80) = "�Աȶ�..."
lgT(81) = "-"
lgT(82) = "��Ƭ��ɫ"
lgT(83) = "-"
lgT(84) = "�ߵ���ɫ"
lgT(85) = "-"
lgT(86) = "�Ҷ�"
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
lgT(100) = "ȥ����ɫ"
lgT(101) = "ȥ����ɫ"
lgT(102) = "ȥ����ɫ"
'4-2 RGB
'4-3
lgT(103) = "�ߵ���ɫ"
lgT(104) = "�ߵ���ɫ"
lgT(105) = "�ߵ���ɫ"
'5
lgT(106) = "����"
lgT(107) = "���⸡��"
lgT(108) = "-"
lgT(109) = "���"
lgT(110) = "-"
lgT(111) = "�ʻ�"
lgT(112) = "-"
lgT(113) = "ģ��"
lgT(114) = "ģ�� (��)"
lgT(115) = "-"
lgT(116) = "��"
lgT(117) = "-"
lgT(118) = "�����ɫ..."
lgT(119) = "-"
lgT(120) = "��ʴ..."
lgT(121) = "ɷ��..."
lgT(122) = "������..."
lgT(123) = "�������"
lgT(124) = "-"
lgT(125) = "����"
lgT(126) = "���� (��)"
lgT(127) = "-"
lgT(128) = "�ڰ׻�"
lgT(129) = "-"
lgT(130) = "����Ⱦɫ"
lgT(131) = "-"
lgT(132) = "Ӳ��Ⱦɫ"
lgT(133) = ""
lgT(134) = ""
lgT(135) = ""
lgT(136) = ""
lgT(137) = ""
lgT(138) = ""
'5-1
lgT(139) = "ͻ����ɫ"
lgT(140) = "ͻ����ɫ"
lgT(141) = "ͻ����ɫ"
'5-2
lgT(142) = "ģʽ 1"
lgT(143) = "ģʽ 2"
lgT(144) = "ģʽ 3"
lgT(169) = "�Ҷ�ģʽ"
'5-3
lgT(145) = "��ɫ"
lgT(146) = "��ɫ"
lgT(147) = "��ɫ"
lgT(148) = "��ɫ"
lgT(149) = "��ɫ"
'5-4
lgT(150) = "��ɫ"
lgT(151) = "��ɫ"
lgT(152) = "��ɫ"
lgT(153) = "��ɫ"
'6
lgT(154) = "��ɫ"
lgT(155) = "��ɫ"
lgT(156) = "ˮ��"
lgT(157) = "��ɫ"
lgT(158) = "ľ̿"
lgT(159) = "ҹɫ"
lgT(160) = "��ʳ"
lgT(161) = "��ɫ"
lgT(162) = "����"
lgT(163) = "��ûҰ�"
lgT(164) = "ǿ����ůɫ"
lgT(165) = "����"
lgT(166) = "�ߵ㻯"
lgT(167) = "�������"
lgT(168) = "ֽ����ɫ"
'lgT(169) ��λ
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
lgT(180) = "���ˮƽ��Ҷ��..."
lgT(181) = "��ⴹֱ��Ҷ��..."
lgT(182) = "ƽ��ˮƽ��Ҷ��..."
lgT(183) = "ƽ�ⴹֱ��Ҷ��..."
lgT(184) = "-"
lgT(185) = "ˮƽ��"
lgT(186) = "��ֱ��"
lgT(187) = "��������״��"
lgT(188) = "���β��������"
lgT(189) = "Բ�β��������"
lgT(190) = "��б��"
lgT(191) = "��б��"
lgT(192) = "����б��"
lgT(193) = "ˮƽ������"
lgT(194) = "��ֱ������"
lgT(195) = "����ˮƽ��������"
lgT(196) = "����ֱ��������"
lgT(197) = "����ˮƽ��������"
lgT(198) = "���Ҵ�ֱ��������"
lgT(199) = "-"
lgT(200) = "��ӱ߿�"
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
lgT(210) = "��ɫ�߿�"
lgT(211) = "���ߵ�ɫ�߿�"
lgT(212) = "����߿� 1"
lgT(213) = "���߽���߿� 1"
lgT(214) = "����߿� 2"
lgT(215) = "���߽���߿� 2"
lgT(216) = "-"
lgT(217) = "Բ�ε�ɫ�߿�"
lgT(218) = "Բ�ν���߿� 1"
lgT(219) = "Բ�ν���߿� 2"
'8
lgT(220) = "��ɫ���"
lgT(221) = "������� 1"
lgT(222) = "������� 2"
lgT(223) = "������� 1"
lgT(224) = "������� 2"
lgT(225) = "Բ����� 1"
lgT(226) = "Բ����� 2"
lgT(227) = "-"
lgT(228) = "�������ͼƬ"
lgT(229) = "ƽ�����ͼƬ"
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
lgT(240) = "���ݽ�"
lgT(241) = "-"
lgT(242) = "������"
lgT(243) = "Բ��������"
lgT(244) = "-"
lgT(245) = "ˮƽ����"
lgT(246) = "��ȫˮƽ����"
lgT(247) = "��ֱ����"
lgT(248) = "��ȫ��ֱ����"
lgT(249) = "-"
lgT(250) = "�ָ�ͼƬ"

'10
lgT(251) = "���ӽ���� (��)"
lgT(252) = "���ӽ���� (��)"
lgT(253) = "�Զ�������"
lgT(254) = ""
lgT(255) = ""
lgT(256) = ""
lgT(257) = ""
lgT(258) = ""
lgT(259) = ""

'10-1
lgT(260) = "��� 1"
lgT(261) = "��� 2"
lgT(262) = "��� 3"
lgT(263) = "-"
lgT(264) = "�ұ� 1"
lgT(265) = "�ұ� 2"
lgT(266) = "�ұ� 3"
lgT(267) = "-"
lgT(268) = "�ϱ� 1"
lgT(269) = "�ϱ� 2"
lgT(270) = "�ϱ� 3"
lgT(271) = "-"
lgT(272) = "�ױ� 1"
lgT(273) = "�ױ� 2"
lgT(274) = "�ױ� 3"
lgT(275) = "-"
lgT(276) = "���� 1"
lgT(277) = "���� 2"
lgT(278) = "-"
lgT(279) = "����1"
lgT(280) = "����2"

'11
lgT(281) = "�������"
lgT(282) = ""
lgT(283) = ""
lgT(284) = ""
lgT(285) = ""

'12
lgT(286) = "�����ĵ�..."
lgT(287) = ""
lgT(288) = "����..."
lgT(289) = ""
lgT(290) = ""
lgT(291) = ""
lgT(292) = ""
lgT(293) = ""
'==========================
lgT(294) = "�Ŵ���   "


lgT(295) = "��ʷ����   "
lgT(296) = "ͼ����Ϣ   "
lgT(297) = "ѡ������   "
lgT(298) = "���ò���   "

lgT(299) = "����������ͼƬ��ɫ"
lgT(300) = "��Ӱ�Ҷ����ֱ�ߡ�б�߻�����"
lgT(301) = "��������˻�����Ч"
lgT(302) = "�������..."
lgT(303) = "ͼƬ�ķ�ת��ƽ�ֲ���"
lgT(304) = "��ͨ�˾�����"

lgT(305) = "  �ļ���: "
lgT(306) = "  ͼ����: "
lgT(307) = "  ͼ��߶�: "
lgT(308) = "�����༭��ͼƬ�ߴ�̫�󣬱༭ʱ���ܳ��������Ӧ"
lgT(309) = "ѡ������"
lgT(310) = "���ڶ�ȡ��ɫ..."
lgT(311) = "���"
lgT(312) = "����Ӧ����ͼ��..."

lgT(313) = ""
'400 is here!
lgT(400) = "��ӭʹ�á�С���ҡ�֮ ͼƬ�༭" & vbCrLf & vbCrLf & "ͼƬ�༭�������������е�ͼƬ��" & vbCrLf & "����β�������ɫ�༭���˾��������Ӱ����" & vbCrLf & "�Լ��������ι���" & vbCrLf & vbCrLf & "����������밴 F1" & vbCrLf & vbCrLf & vbCrLf & "��ʼ�༭ͼƬ��������İ�ť"
lgT(314) = "ͼƬ�༭"

lgT(315) = "��ɫ"
lgT(316) = "����"
lgT(317) = "����"

lgT(318) = "�Բ������ո�û��ѡ��Ҫ�򿪵��ļ�"
lgT(319) = "��ӡ�����ʹ��Ĭ�ϴ�ӡ����Ĭ��ֽ�Ŵ�С��ӡ"
lgT(320) = "�Ƿ�����ӡ��"
lgT(321) = "��ȡ���˴�ӡ"

lgT(322) = "�˾�"
lgT(323) = "�����ɫ"
lgT(324) = "���ͼƬ"

lgT(325) = "û�з���ѡ������"
lgT(326) = "��ͼƬʱ����Ӧ"
lgT(327) = "�����У���ȴ�..."
lgT(328) = "��һ��ͼƬ"
lgT(329) = "��һ��ͼƬ"

''''Copy
lgT(330) = "�����Ѿ�������"
'FEcho
lgT(331) = "���ݽ�"
lgT(332) = "����"
lgT(333) = "Ĭ��"
lgT(334) = "����"
lgT(335) = "�ݼ�"
lgT(336) = "ƫ������-ˮƽ"
lgT(337) = "ƫ������-��ֱ"
lgT(338) = "����Ԥ��"
lgT(339) = "ȷ��"
lgT(340) = "����"
lgT(341) = "ƫ������"

'FText
lgT(342) = "����������������"
lgT(343) = "�������"
lgT(344) = "��������"
lgT(345) = "��С"
lgT(346) = "������Ӱ"
lgT(347) = "������ɫ"
lgT(348) = "��Ӱ��ɫ"
lgT(349) = "��Ӱ����"
lgT(350) = "����X"
lgT(351) = "����Y"
lgT(352) = "λ��"
lgT(353) = "����"
lgT(354) = "����"
lgT(355) = "����"

'Fcolor
lgT(356) = "͸����"
lgT(357) = "������ȫ��"
lgT(358) = "��ԭ��С����"
lgT(359) = "�������ͼƬ"
lgT(360) = "ȷ��"
lgT(361) = "ȡ��"
lgT(362) = "��ɫ"
lgT(363) = "���"
lgT(364) = "����"
lgT(365) = "��"
lgT(366) = "��"
lgT(367) = "��"
lgT(368) = "�߶�"
lgT(369) = "���"
lgT(370) = "������󣬰�[����]����"
lgT(371) = "��ͼ"
lgT(372) = "����"
lgT(373) = "����Ԥ��"
lgT(374) = "���Ժ�..."

lgT(375) = "���(&R)..."
lgT(376) = ""
lgT(377) = ""
lgT(378) = ""
lgT(379) = ""
lgT(380) = ""
lgT(381) = ""
lgT(382) = ""
lgT(383) = "������" '����λ
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
lgT(399) = "������Ҫ�ر�ͼƬ�༭��"
'400 401��
lgT(402) = ""
'lgT(403) = "
lgT(404) = "С����"




'���Խ���
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

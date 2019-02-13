VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Randoms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'================================================================================
'名称:
'RandomGenerator
'
'描述:
'随机生成器
'
'作者:
'fangzi
'
'RandDateTime   '生成随机日期时间   09:52 2009-9-26
'RandCnName     '常用姓占90%        10:22 2009-9-28
'
'最后更新:10:22 2009-9-28
'================================================================================

Option Explicit

Private Const BAIJIAXING             As String = "李王张赵杨朱钱刘孙卢周吴方丁郑冯陈苏罗徐马胡杜阮彭蒋沈林许崔何夏吕高蔡郭万田韩魏樊施褚卫梁秦尤孔曹傅" & "陶姜戚谢邹喻柏水窦章云潘葛奚范郎鲁韦昌苗凤花俞任袁柳酆鲍史唐费廉岑薛雷贺倪汤滕殷毕郝金邬安常乐于严时" & "皮卞齐康伍余元卜顾孟平黄和穆萧尹姚邵湛汪祁毛禹狄米贝明臧计伏成戴谈宋茅庞熊纪舒屈项祝董蓝闵席季华麻强" & "贾路娄危江童颜盛刁钟邱骆凌霍虞支柯昝管莫经房裘缪干解应宗宣贲邓郁单杭洪梅包诸左石吉钮龚程嵇邢滑裴陆荣" & "翁荀羊於惠甄曲家封芮羿储靳汲邴糜松井段富巫乌焦巴弓牧隗山谷车侯宓蓬全郗班仰秋仲伊宫宁仇栾暴甘钭厉戎祖" & "武符景詹束龙叶幸司韶郜黎蓟薄印宿白怀蒲邰从鄂索咸籍赖卓蔺屠蒙池乔阴欎胥能苍双闻莘党翟谭贡劳逄姬申扶堵" & "冉宰郦雍郤璩桑桂濮牛寿通边扈燕冀郏浦尚农温别庄晏柴瞿阎充慕连"
Private Const COMMON_NAME_1          As String = "敏伟勇军斌静丽涛芳杰萍强俊明燕磊玲华平鹏健波红丹辉超艳莉刚娟峰婷亮洁颖琳英慧飞霞浩凯宇毅林佳云莹娜晶" & "洋文鑫欣琴宁琼兵青琦翔彬锋阳璐旭蕾剑虹蓉建倩梅宏威博君力龙晨薇雪琪欢荣江炜成庆冰东帆雷楠锐进海凡巍维" & "迪媛玮杨群瑛悦春瑶婧兰茜松爽立瑜睿晖聪帅瑾骏雯晓昊勤新瑞岩星忠志怡坤康航利畅坚雄智萌哲岚洪捷珊恒靖清" & "扬昕乐武玉诚菲锦凤珍晔妍璇胜菁科芬露越彤曦义良鸣芸方月铭光震冬源政虎莎彪蓓钢凌奇卫彦烨可黎川淼惠祥然"
Private Const COMMON_NAME_2_1        As String = "建小晓文志国玉丽永海春金明新德秀红亚伟雪俊桂爱美世正庆学家立淑振云华光惠兴天长艳慧利宏佳瑞凤荣秋继嘉" & "卫燕思维少福忠宝子成月洪东一泽林大素旭宇智锦冬玲雅伯翠传启剑安树良中梦广昌元万清静友宗兆丹克彩绍喜远" & "朝敏培胜祖先菊士向有连军健巧耀莉英方和仁孝梅汉兰松水江益开景运贵祥青芳碧婷龙鹏自顺双书生义跃银佩雨保" & "贤仲鸿浩加定炳飞锡柏发超道怀进其富平全阳吉茂彦诗洁润承治焕如君增善希根应勇宜守会凯育湘凌本敬博延乐三"
Private Const COMMON_NAME_2_2        As String = "华平明英军林萍芳玲红生霞梅文荣珍兰娟峰琴云辉东龙敏伟强丽春杰燕民君波国芬清祥斌婷飞良忠新凤锋成勇刚玉" & "元宇海兵安庆涛鹏亮青阳艳松江莲娜兴光德武香俊秀慧雄才宏群琼胜超彬莉中山富花宁利贵福发义蓉喜娥昌仁志全" & "宝权美琳建金贤星丹根和珠康菊琪坤泉秋静佳顺源珊达欣如莹章浩勤芹容友芝豪洁鑫惠洪旺虎远妮森妹南雯奇健卿" & "虹娇媛怡铭川进博智来琦学聪洋乐年翔然栋凯颖鸣丰瑞奎立堂威雪鸿晶桂凡娣先洲毅雅月旭田晖方恒亚泽风银高贞"
Private Const COMMON_NAME_MALE_1     As String = "伟勇军斌杰强俊明磊华平鹏健波辉超刚峰亮飞浩凯宇毅林兵翔彬锋阳旭剑建宏威博力龙荣江炜成庆东雷锐进海巍迪" & "玮杨春松立瑜睿晖聪帅瑾骏昊勤新岩星忠志坤康航利坚雄智哲洪捷恒清扬昕武诚锦胜菁科芬露越彤曦义良鸣方月铭" & "光震冬政虎彪钢凌奇卫彦烨川祥涛洁云文洋鑫欣宁青琦君晨欢冰凡群爽晓畅岚晔璇淼然黎源瑞乐维"
Private Const COMMON_NAME_FEMALE_1   As String = "敏静丽芳萍燕玲红丹艳莉娟婷颖琳英慧霞佳莹娜晶琴琼璐蕾虹蓉倩梅薇雪琪帆楠媛瑛瑶婧兰茜雯怡萌珊靖玉凤珍妍" & "涛洁云文洋鑫欣宁青琦君晨欢冰凡群爽晓畅岚晔璇淼然黎源瑞乐维菲芸莎惠可蓓悦"
Private Const COMMON_NAME_MALE_2_1   As String = "建志国金明新德亚伟俊桂世正庆学家立振光兴天长利宏荣继嘉卫思维少福忠宝成洪东一泽林大旭宇智锦冬伯传启剑" & "安树良中广昌元万清友宗兆丹克绍喜远朝培胜祖先士向有连军健耀方和仁孝汉松水江益开景运贵祥龙鹏自顺双书生" & "义跃银佩保贤仲鸿浩加定炳飞锡柏发超道怀进其富平全阳吉茂彦润承治增善希根应勇宜守会凯育湘凌本敬博延三小" & "晓文永海春红雪爱云华青雨君瑞子"
Private Const COMMON_NAME_FEMALE_2_1 As String = "乐玉丽秀美淑惠艳慧佳凤秋燕梅芳碧婷焕菊敏彩梦翠玲雅素莉英洁诗静月巧兰如小晓文永海春红雪爱云华青雨君瑞子"
Private Const COMMON_NAME_MALE_2_2   As String = "平明军生荣峰辉东龙伟强杰民君波国祥斌飞良忠新锋成勇刚玉元宇海兵安庆涛鹏亮青阳松江兴光德武俊雄才宏群琼" & "胜超彬中山富利贵福发义昌仁志全宝权建星根和康坤泉顺源达章浩勤容友豪鑫洪旺虎远森南奇健铭川进博智来琦学" & "聪洋年翔栋凯鸣丰瑞奎立堂威鸿凡先洲毅雅月旭田晖方恒亚泽风银高林红文云然颖金贤喜宁桂雪清春欣乐"
Private Const COMMON_NAME_FEMALE_2_2 As String = "华英萍芳玲霞梅珍兰娟琴艳贞敏洁惠珊妮如娣莹晶娇媛凤怡卿虹妹芝芹秋静菊琪珠美琳丹娥蓉莉花莲娜婷芬燕丽佳" & "雯香秀慧林红文云然颖金贤喜宁桂雪清春欣乐"
Private Const BAIJIAXING_LENGTH      As Integer = 330
Private Const COMMON_NAME_LENGTH     As Integer = 200
Private Const ALL_LETTERS            As String = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"
Private Const ALL_ABC                As String = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ"
Private Const LOWER_LETTERS          As String = "abcdefghijklmnopqrstuvwxyz"
Private Const UPPER_LETTERS          As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
Private Const ALL_NUMBERS            As String = "0123456789"
Private Const UNI_START              As Long = 19968
Private Const UNI_END                As Long = 40896
    
Enum LetterStyleConstants
    [LettersLower] = 4&
    [LettersUpper] = 2&
    [Numbers] = 1&
End Enum

Private Sub Class_Initialize()
    Randomize
End Sub

'--------------------------------------------------------------------------------
'描述：
'返回String1中的一个随机字符
'
'作者：fangzi
'日期：2010-04-05 17:55:44
'--------------------------------------------------------------------------------
Public Function RandCharOfString(String1 As String) As String
Attribute RandCharOfString.VB_Description = "返回String1中的一个随机字符"
    If Len(String1) = 0 Then Exit Function
    
    RandCharOfString = Mid$(String1, RandLong(Len(String1), 1), 1)
End Function

'--------------------------------------------------------------------------------
'描述：
'返回随机身份证号
'
'作者：fangzi
'日期：2010-04-05 17:54:52
'--------------------------------------------------------------------------------
Public Function RandIDNumber() As String
Attribute RandIDNumber.VB_Description = "返回随机身份证号，目前仅显云南省"
    Dim sAreaCode As String
    Dim sBirthday As String
    Dim sSequence As String
    Dim sValidate As String
    
    sAreaCode = "530" & RandNumberString(3)
    sBirthday = Format$(RandDate(#12/31/1990#, #1/1/1970#), "yyyyMMdd")
    sSequence = Format$(RandInt(999, 1), "000")
    sValidate = IIf(RandBoolean(0.01), "X", RandCharOfString("1234567890"))
    
    RandIDNumber = sAreaCode & sBirthday & sSequence & sValidate
End Function

'返回lst数组中的元素之一
Public Function RandOfArray(lst As Variant) As Variant
Attribute RandOfArray.VB_Description = "返回lst数组中的元素之一"
    Dim i As Long

    If Not IsArray(lst) Then Exit Function
    i = RandInt(UBound(lst), LBound(lst))
    RandOfArray = lst(i)
End Function

'--------------------------------------------------------------------------------
'描述：
'返回ComboBox或ListBox的列表项中之一的字符串
'
'作者：fangzi
'日期：2010-04-03 15:11:31
'--------------------------------------------------------------------------------
Public Function RandOfList(ListControl As Object) As String
Attribute RandOfList.VB_Description = "返回ComboBox或ListBox的列表项中之一的字符串"
    If TypeOf ListControl Is ListBox Or TypeOf ListControl Is ComboBox Then
        RandOfList = ListControl.List(RandInt(ListControl.ListCount - 1, 0))
    End If
End Function

'--------------------------------------------------------------------------------
'描述：
'随机分数,可能是整数也可能是整数.5
'
'参数：
'FullMark： 满分是多少
'PointRate：得到0.5的几率是多少
'PassMark： 及格分数
'PassRate： 及格的几率
'
'作者：fangzi
'日期：2010-04-03 15:12:00
'--------------------------------------------------------------------------------
Public Function RandMark(Optional FullMark As Integer = 100, Optional PointRate As Single = 0.8, Optional PassMark As Integer = 60, Optional PassRate As Single = 0.9) As Single
Attribute RandMark.VB_Description = "返回表示得分的随机数字,可能是整数也可能是整数.5。FullMark:满分；PointRate:得到0.5的几率；PassMark:及格分数；PassRate:及格的几率"
    Dim mark As Single
    
    If RandBoolean(PassRate) Then
        mark = RandInt(FullMark, PassMark)
    Else
        mark = RandInt(PassMark - 1, 0)
    End If
    
    '带分数
    If RandBoolean(PointRate) Then
        If mark = FullMark Then
            RandMark = FullMark
        Else
            RandMark = mark + 0.5
        End If
    Else
        '不带分数
        RandMark = mark
    End If
End Function

'返回小于等于Upper且大于等于Lower的日期
Public Function RandDate(ByVal Upper As Date, Optional ByVal Lower As Date = #1/1/1900#) As Date
Attribute RandDate.VB_Description = "返回小于等于Upper且大于等于Lower的日期"
    Dim diff As Long
    Dim nd   As Long

    diff = DateDiff("d", Lower, Upper)
    
    nd = RandLong(diff, 0)
    RandDate = DateAdd("d", nd, Lower)
End Function

'返回小于等于Upper且大于等于Lower的日期和时间
Public Function RandDateTime(ByVal Upper As Date, Optional ByVal Lower As Date = #1/1/1900#) As String
Attribute RandDateTime.VB_Description = "返回小于等于Upper且大于等于Lower的日期和时间"
    Dim diff As Long
    Dim nd   As Long

    diff = DateDiff("s", Lower, Upper)
    
    nd = RandLong(diff, 0)
    RandDateTime = Format$(DateAdd("s", nd, Lower), "yyyy-MM-dd")
    RandDateTime = RandDateTime & " " & Format$(RandInt(23, 0), "00") & ":" & Format$(RandInt(59, 0), "00") & ":" & Format$(RandInt(59, 0), "00")
End Function

'返回两个Variant之一
Public Function RandOf2(ByVal v1 As Variant, ByVal v2 As Variant, Optional ByVal v1Rate As Single = 0.5) As Variant
Attribute RandOf2.VB_Description = "返回两个Variant之一。v1Rate表示返回值是第一个值的几率。"
    RandOf2 = IIf(Rnd < v1Rate, v1, v2)
End Function

'返回两个字符串之一
Public Function RandOf2String(ByVal str1 As String, ByVal str2 As String, Optional ByVal str1Rate As Single = 0.5)
Attribute RandOf2String.VB_Description = "返回两个字符串之一"
    RandOf2String = IIf(Rnd < str1Rate, str1, str2)
End Function

'返回boolean
Public Function RandBoolean(Optional trueRate As Single = 0.5) As Boolean
Attribute RandBoolean.VB_Description = "返回boolean"
    RandBoolean = (Rnd < trueRate)
End Function

'返回小于等于Upper且大于等于Lower的整数
Public Function RandInt(ByVal Upper As Integer, Optional ByVal Lower As Integer) As Integer
Attribute RandInt.VB_Description = "返回小于等于Upper且大于等于Lower的整数"
    RandInt = Int((Upper - Lower + 1) * Rnd + Lower)
End Function

'返回int1和int2之一
Public Function RandOf2Integer(ByVal int1 As Integer, ByVal int2 As Integer, Optional ByVal int1Rate As Single = 0.5) As Integer
Attribute RandOf2Integer.VB_Description = "返回int1和int2之一"
    RandOf2Integer = IIf(Rnd >= int1Rate, int2, int1)
End Function

'返回小于等于Upper且大于等于Lower的长整数
Public Function RandLong(ByVal Upper As Long, Optional ByVal Lower As Long) As Long
Attribute RandLong.VB_Description = "返回小于等于Upper且大于等于Lower的长整数"
    RandLong = Int((Upper - Lower + 1) * Rnd + Lower)
End Function

'返回长度为Length的字符串
Public Function RandString(ByVal Length As Long) As String
Attribute RandString.VB_Description = "返回长度为Length的字符串。"
    Dim nIndex As Integer
    Dim i      As Long
    Dim ret    As String
    
    ret = String$(Length, Chr$(0))

    For i = 1 To Length
        nIndex = RandInt(Len(ALL_LETTERS), 1)
        Mid$(ret, i, 1) = Mid$(ALL_LETTERS, nIndex, 1)
    Next

    RandString = ret
End Function

'Public Function RandLetter(ByVal Lower As Integer, ByVal Upper As Integer)
'
'End Function

'返回由数字组成的长度为Length的字符串
Public Function RandNumberString(ByVal Length As Long) As String
Attribute RandNumberString.VB_Description = "返回由数字组成的长度为Length的字符串"
    Dim nIndex As Integer
    Dim i      As Long
    Dim ret    As String
    
    ret = String$(Length, Chr$(0))

    For i = 1 To Length
        nIndex = RandInt(Len(ALL_NUMBERS), 1)
        Mid$(ret, i, 1) = Mid$(ALL_NUMBERS, nIndex, 1)
    Next

    RandNumberString = ret
End Function

'返回长度为Length的unicode汉字
Public Function RandUnicodeCnString(ByVal Length As Long) As String
Attribute RandUnicodeCnString.VB_Description = "返回长度为Length的unicode汉字"
    Dim unicode As Long
    Dim ret     As String
    Dim i       As Long

    ret = String$(Length, Chr$(0))
    
    For i = 1 To Length
        unicode = RandLong(UNI_END, UNI_START)
        Mid$(ret, i, 1) = ChrW$(unicode)
    Next

    RandUnicodeCnString = ret
End Function

'Public Function RandCommonCnString(ByVal Length As Long) As String
'    Dim index As String
'    Dim ret As String
'    Dim s1  As String
'    Dim s2  As String
'    Dim s3  As String
'    Dim s4  As String
'    Dim i   As Long
'
'    ret = String(Length, Chr$(0))
'
'    For i = 1 To Length
'        s1 = RandLetter(2, 4)
'        s2 = Hex$(RandInt(15))
'        s3 = Hex$(RandInt(15, 10))
'        If s3 = "A" Then
'            s4 = Hex$(RandInt(15, 1))
'        ElseIf s3 = "F" Then
'            s4 = Hex$(RandInt(14, 0))
'        Else
'            s4 = Hex$(RandInt(15, 0))
'        End If
'        Mid$(ret, i, 1) = Chr$("&H" & s1 & s2 & s3 & s4)
'    Next
'
'    RandCommonCnString = ret
'End Function
  
'--------------------------------------------------------------------------------
'描述：
'返回随机中国人名
'
'参数：
'Gender: 1 男 2 女 其他 忽略性别
'
'返回：
'
'作者：fangzi
'日期：2010-04-03 16:49:53
'--------------------------------------------------------------------------------
Public Function RandName(Optional ByVal Gender As Long = 0) As String
    Dim nameLen   As Integer
    Dim Index     As Integer
    Dim firstName As String
    Dim midName   As String
    Dim lastName  As String
    Dim isCommon  As Boolean

    isCommon = RandBoolean(0.9)
    
    nameLen = RandOf2Integer(2, 3, 0.4)
    
    If isCommon Then
        firstName = Mid$(Left$(BAIJIAXING, 50), RandInt(50, 1), 1)
    Else
        firstName = Mid$(BAIJIAXING, RandInt(BAIJIAXING_LENGTH, 1), 1)
    End If

    Select Case Gender
    Case 1  '男
        If nameLen = 3 Then
            Index = RandInt(Len(COMMON_NAME_MALE_2_1), 1)
            midName = Mid$(COMMON_NAME_MALE_2_1, Index, 1)
            Index = RandInt(Len(COMMON_NAME_MALE_2_2), 1)
            lastName = Mid$(COMMON_NAME_MALE_2_2, Index, 1)
            RandName = firstName & midName & lastName
        Else
            Index = RandInt(Len(COMMON_NAME_MALE_1), 1)
            lastName = Mid$(COMMON_NAME_MALE_1, Index, 1)
            RandName = firstName & lastName
        End If
    Case 2  '女
        If nameLen = 3 Then
            Index = RandInt(Len(COMMON_NAME_FEMALE_2_1), 1)
            midName = Mid$(COMMON_NAME_FEMALE_2_1, Index, 1)
            Index = RandInt(Len(COMMON_NAME_FEMALE_2_2), 1)
            lastName = Mid$(COMMON_NAME_FEMALE_2_2, Index, 1)
            RandName = firstName & midName & lastName
        Else
            Index = RandInt(Len(COMMON_NAME_FEMALE_1), 1)
            lastName = Mid$(COMMON_NAME_FEMALE_1, Index, 1)
            RandName = firstName & lastName
        End If
    Case Else  '不分
        If nameLen = 3 Then
            Index = RandInt(COMMON_NAME_LENGTH, 1)
            midName = Mid$(COMMON_NAME_2_1, Index, 1)
            Index = RandInt(COMMON_NAME_LENGTH, 1)
            lastName = Mid$(COMMON_NAME_2_2, Index, 1)
            RandName = firstName & midName & lastName
        Else
            Index = RandInt(COMMON_NAME_LENGTH, 1)
            lastName = Mid$(COMMON_NAME_1, Index, 1)
            RandName = firstName & lastName
        End If
    End Select
End Function

'返回RGB颜色
Public Function RandColor() As Long
Attribute RandColor.VB_Description = "返回RGB颜色"
    RandColor = RGB(RandInt(255, 0), RandInt(255, 0), RandInt(255, 0))
End Function

'--------------------------------------------------------------------------------
'描述：
'返回随机性别，MaleRate指定获取男性的几率
'作者：fangzi
'日期：2009-12-24 20:58:28
'--------------------------------------------------------------------------------
Public Function RandGender(Optional ByVal MaleRate As Single = 0.5) As String
Attribute RandGender.VB_Description = "返回随机性别，MaleRate指定获取男性的几率"
    RandGender = IIf(Rnd > MaleRate, "男", "女")
End Function

'--------------------------------------------------------------------------------
'描述：
'返回随机年龄，MinAge和MaxAge指定返回的年龄范围
'作者：fangzi
'日期：2009-12-24 20:59:44
'--------------------------------------------------------------------------------
Public Function RandAge(Optional MinAge As Integer = 1, Optional MaxAge As Integer = 100) As Integer
Attribute RandAge.VB_Description = "返回随机年龄，MinAge和MaxAge指定返回的年龄范围"
    RandAge = RandInt(MaxAge, MinAge)
End Function

'--------------------------------------------------------------------------------
'描述：
'返回随机生日
'作者：fangzi
'日期：2009-12-24 21:00:27
'--------------------------------------------------------------------------------
Public Function RandBirthday(Optional DateFormat As String = "yyyy-MM-dd") As String
Attribute RandBirthday.VB_Description = "返回随机生日，DateFormat指定日期格式"
    Dim d As Date

    d = RandDate(Now, DateAdd("yyyy", -99, Now))
    RandBirthday = Format$(d, DateFormat)
End Function

'--------------------------------------------------------------------------------
'描述：
'返回随机小写字母
'作者：fangzi
'日期：2009-12-24 21:16:38
'--------------------------------------------------------------------------------
Public Function RandLetterLower()
Attribute RandLetterLower.VB_Description = "返回随机小写字母"
    RandLetterLower = Mid$(LOWER_LETTERS, RandInt(26, 1), 1)
End Function

'--------------------------------------------------------------------------------
'描述：
'返回随机大写字母
'作者：fangzi
'日期：2009-12-24 21:16:38
'--------------------------------------------------------------------------------
Public Function RandLetterUpper()
Attribute RandLetterUpper.VB_Description = "返回随机大写字母"
    RandLetterUpper = Mid$(UPPER_LETTERS, RandInt(26, 1), 1)
End Function

'--------------------------------------------------------------------------------
'描述：
'返回随机字母，不分大小写
'作者：fangzi
'日期：2009-12-24 21:16:38
'--------------------------------------------------------------------------------
Public Function RandLetter()
Attribute RandLetter.VB_Description = "返回随机字母，不分大小写"
    RandLetter = Mid$(ALL_ABC, RandInt(52, 1), 1)
End Function

'--------------------------------------------------------------------------------
'描述：
'返回随机Ascii可见字符
'作者：fangzi
'日期：2009-12-24 21:20:02
'--------------------------------------------------------------------------------
Public Function RandChar()
Attribute RandChar.VB_Description = "返回随机Ascii可见字符"
    RandChar = Chr$(RandInt(126, 32))
End Function

'--------------------------------------------------------------------------------
'描述：
'返回随机的格式化字符串
'FormatString:
'   A：表示A~Z
'   a：表示a~z
'   f: 表示a~Z
'   0：表示0~9
'   ?：表示任意单个字符
'   啊：表示任意汉字
'   其他字符：返回原文
'
'   示例：要得到asd-124-ASD我是-a
'         debug.print RandFormatString("aaa-000-AAA啊啊-?")
'作者：fangzi
'日期：2009-12-24 21:04:57
'--------------------------------------------------------------------------------
Public Function RandFormatString(Optional FormatString As String = "A-a~f/0?啊")
Attribute RandFormatString.VB_Description = "返回随机的格式化字符串。\r\n[A:表示A~Z | a:表示a~z | f:表示a~Z | 0:表示0~9 | ?:表示任意单个字符 | 啊:表示任意汉字 | 其他字符:返回原文"
    Dim i    As Long
    Dim s    As String
    Dim c    As String
    Dim sRet As String
    Dim u    As Long
   
    u = Len(FormatString)

    If u = 0 Then Exit Function
    
    sRet = String$(u, Chr$(0))
    
    For i = 1 To u
        c = Mid$(FormatString, i, 1)

        Select Case c
        Case "a":  Mid$(sRet, i, 1) = RandLetterLower
        Case "A":  Mid$(sRet, i, 1) = RandLetterUpper
        Case "f":  Mid$(sRet, i, 1) = RandLetter
        Case "?":  Mid$(sRet, i, 1) = RandChar
        Case "0":  Mid$(sRet, i, 1) = RandInt(9, 0)
        Case "啊": Mid$(sRet, i, 1) = RandUnicodeCnString(1)
        Case Else: Mid$(sRet, i, 1) = c
        End Select
    Next
    
    RandFormatString = sRet
End Function

'--------------------------------------------------------------------------------
'描述：
'返回大于1的随机浮点数
'参数：
'Max:最大值
'PointNumber: 保留的小数点位数
'作者：fangzi
'日期：2009-12-24 22:02:02
'--------------------------------------------------------------------------------
Public Function RandFloat(Max As Long, Optional Min As Long = 0, Optional PointReserved As Integer = 2) As Double
Attribute RandFloat.VB_Description = "返回大于1的随机浮点数，Max表示最大值，Min表示最小值，PointReserved表示保留的小数位数"
    If PointReserved < 0 Then Exit Function
    
    RandFloat = Format$(RandLong(Max, Min) + Rnd, "#." & String$(PointReserved, "0"))
End Function

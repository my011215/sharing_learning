import random
import string
from datetime import datetime, timedelta

import xlsxwriter as xw

str_phone = list()
str_id = list()
str_id_connt = list()
str_phone_connt = list()

name_list = []
set_phone = []
set_ID = []
list_date1 = []
list_date2 = []
list_date3 = []
list_date4 = []
list_date5 = []
list_port1 = []
list_port2 = []
ladingnumber = []
containerid = []
list_containerid = []
min = 0


def randomnum():  # 手机号
    str_phone = []
    str_phone.append(str(15))
    for i in range(0, 9):
        num = random.randint(1, 9)
        str_phone.append(str(num))
    if str_phone not in str_phone_connt:
        return str_phone
    else:
        pass


def deal_day(d):  # 日期处理
    id_day = random.randint(1, d)
    if id_day < 10:
        str_id.append("0")
        str_id.append(str(id_day))
    else:
        str_id.append(str(id_day))


def randid():  # 身份证
    str_id.clear()
    list1 = ['11', '12', '13', '14', '15', '21', '22', '23', '31', '32', '33', '34', '35', '36', '37', '41', '42', '43',
             '44', '45', '46', '50', '51', '52', '53', '54', '61', '62', '63', '64', '65', '81', '82', '83']  # 全国区域代码
    province_num = random.randint(1, len(list1) - 1)
    str_id.append(list1[province_num])
    city_num = random.randint(1000, 9999)
    str_id.append(str(city_num))
    id_year = random.randint(1950, 2010)
    str_id.append(str(id_year))
    id_month = random.randint(1, 12)
    if id_month < 10:
        str_id.append("0")
        str_id.append(str(id_month))
    else:
        str_id.append(str(id_month))
    if id_year % 4 == 0:  # 闰年
        if id_month == 2:
            deal_day(29)
        elif id_month in (1, 3, 5, 7, 8, 10, 12):
            deal_day(31)
        else:
            deal_day(30)
    else:
        if id_month == 2:
            deal_day(28)
        elif id_month in (1, 3, 5, 7, 8, 10, 12):
            deal_day(31)
        else:
            deal_day(30)
    rand_num = random.randint(100, 999)
    str_id.append(str(rand_num))
    check_num = random.randint(1, 10)
    if check_num == 10:
        str_id.append("x")
    else:
        str_id.append(str(check_num))
    return str_id

def randname():  # 姓名
    xing = [
        '赵', '钱', '孙', '李', '周', '吴', '郑', '王', '冯', '陈', '褚', '卫', '蒋', '沈', '韩', '杨', '朱', '秦',
        '尤', '许',
        '何', '吕', '施', '张', '孔', '曹', '严', '华', '金', '魏', '陶', '姜', '戚', '谢', '邹', '喻', '柏', '水',
        '窦', '章',
        '云', '苏', '潘', '葛', '奚', '范', '彭', '郎', '鲁', '韦', '昌', '马', '苗', '凤', '花', '方', '俞', '任',
        '袁', '柳',
        '酆', '鲍', '史', '唐', '费', '廉', '岑', '薛', '雷', '贺', '倪', '汤', '滕', '殷', '罗', '毕', '郝', '邬',
        '安', '常',
        '乐', '于', '时', '傅', '皮', '卞', '齐', '康', '伍', '余', '元', '卜', '顾', '孟', '平', '黄', '和', '穆',
        '萧', '尹',
        '姚', '邵', '堪', '汪', '祁', '毛', '禹', '狄', '贝', '明', '臧', '计', '伏', '成', '戴', '谈', '宋', '茅',
        '庞', '梁']
    ming1 = ['一', '二', '三', '四', '五', '六', '七', '八', '九']
    ming2 = [
        '的', '一', '是', '了', '我', '不', '人', '在', '他', '有', '这', '个', '上', '们', '来', '到', '时', '大',
        '地', '为',
        '子', '中', '你', '说', '生', '国', '年', '着', '就', '那', '和', '要', '她', '出', '也', '得', '里', '后',
        '自', '以',
        '会', '家', '可', '下', '而', '过', '天', '去', '能', '对', '小', '多', '然', '于', '心', '学', '么', '之',
        '都', '好',
        '看', '起', '发', '当', '没', '成', '只', '如', '事', '把', '还', '用', '第', '样', '道', '想', '作', '种',
        '开', '美',
        '总', '从', '无', '情', '己', '面', '最', '女', '但', '现', '前', '些', '所', '同', '日', '手', '又', '行',
        '意', '动',
        '方', '期', '它', '头', '经', '长', '儿', '回', '位', '分', '爱', '老', '因', '很', '给', '名', '法', '间',
        '斯', '知',
        '世', '什', '两', '次', '使', '身', '者', '被', '高', '已', '亲', '其', '进', '此', '话', '常', '与', '活',
        '正', '感',
        '见', '明', '问', '力', '理', '尔', '点', '文', '几', '定', '本', '公', '特', '做', '外', '孩', '相', '西',
        '果', '走',
        '将', '月', '十', '实', '向', '声', '车', '全', '信', '重', '三', '机', '工', '物', '气', '每', '并', '别',
        '真', '打',
        '太', '新', '比', '才', '便', '夫', '再', '书', '部', '水', '像', '眼', '等', '体', '却', '加', '电', '主',
        '界', '门',
        '利', '海', '受', '听', '表', '德', '少', '克', '代', '员', '许', '稜', '先', '口', '由', '死', '安', '写',
        '性', '马',
        '光', '白', '或', '住', '难', '望', '教', '命', '花', '结', '乐', '色', '更', '拉', '东', '神', '记', '处',
        '让', '母',
        '父', '应', '直', '字', '场', '平', '报', '友', '关', '放', '至', '张', '认', '接', '告', '入', '笑', '内',
        '英', '军',
        '候', '民', '岁', '往', '何', '度', '山', '觉', '路', '带', '万', '男', '边', '风', '解', '叫', '任', '金',
        '快', '原',
        '吃', '妈', '变', '通', '师', '立', '象', '数', '四', '失', '满', '战', '远', '格', '士', '音', '轻', '目',
        '条', '呢',
        '病', '始', '达', '深', '完', '今', '提', '求', '清', '王', '化', '空', '业', '思', '切', '怎', '非', '找',
        '片', '罗',
        '钱', '紶', '吗', '语', '元', '喜', '曾', '离', '飞', '科', '言', '干', '流', '欢', '约', '各', '即', '指',
        '合', '反',
        '题', '必', '该', '论', '交', '终', '林', '请', '医', '晚', '制', '球', '决', '窢', '传', '画', '保', '读',
        '运', '及',
        '则', '房', '早', '院', '量', '苦', '火', '布', '品', '近', '坐', '产', '答', '星', '精', '视', '五', '连',
        '司', '巴',
        '奇', '管', '类', '未', '朋', '且', '婚', '台', '夜', '青', '北', '队', '久', '乎', '越', '观', '落', '尽',
        '形', '影',
        '红', '爸', '百', '令', '周', '吧', '识', '步', '希', '亚', '术', '留', '市', '半', '热', '送', '兴', '造',
        '谈', '容',
        '极', '随', '演', '收', '首', '根', '讲', '整', '式', '取', '照', '办', '强', '石', '古', '华', '諣', '拿',
        '计', '您',
        '装', '似', '足', '双', '妻', '尼', '转', '诉', '米', '称', '丽', '客', '南', '领', '节', '衣', '站', '黑',
        '刻', '统',
        '断', '福', '城', '故', '历', '惊', '脸', '选', '包', '紧', '争', '另', '建', '维', '绝', '树', '系', '伤',
        '示', '愿',
        '持', '千', '史', '谁', '准', '联', '妇', '纪', '基', '买', '志', '静', '阿', '诗', '独', '复', '痛', '消',
        '社', '算',
        '义', '竟', '确', '酒', '需', '单', '治', '卡', '幸', '兰', '念', '举', '仅', '钟', '怕', '共', '毛', '句',
        '息', '功',
        '官', '待', '究', '跟', '穿', '室', '易', '游', '程', '号', '居', '考', '突', '皮', '哪', '费', '倒', '价',
        '图', '具',
        '刚', '脑', '永', '歌', '响', '商', '礼', '细', '专', '黄', '块', '脚', '味', '灵', '改', '据', '般', '破',
        '引', '食',
        '仍', '存', '众', '注', '笔', '甚', '某', '沉', '血', '备', '习', '校', '默', '务', '土', '微', '娘', '须',
        '试', '怀',
        '料', '调', '广', '蜖', '苏', '显', '赛', '查', '密', '议', '底', '列', '富', '梦', '错', '座', '参', '八',
        '除', '跑',
        '亮', '假', '印', '设', '线', '温', '虽', '掉', '京', '初', '养', '香', '停', '际', '致', '阳', '纸', '李',
        '纳', '验',
        '助', '激', '够', '严', '证', '帝', '饭', '忘', '趣', '支', '春', '集', '丈', '木', '研', '班', '普', '导',
        '顿', '睡',
        '展', '跳', '获', '艺', '六', '波', '察', '群', '皇', '段', '急', '庭', '创', '区', '奥', '器', '谢', '弟',
        '店', '否',
        '害', '草', '排', '背', '止', '组', '州', '朝', '封', '睛', '板', '角', '况', '曲', '馆', '育', '忙', '质',
        '河', '续',
        '哥', '呼', '若', '推', '境', '遇', '雨', '标', '姐', '充', '围', '案', '伦', '护', '冷', '警', '贝', '著',
        '雪', '索',
        '剧', '啊', '船', '险', '烟', '依', '斗', '值', '帮', '汉', '慢', '佛', '肯', '闻', '唱', '沙', '局', '伯',
        '族', '低',
        '玩', '资', '屋', '击', '速', '顾', '泪', '洲', '团', '圣', '旁', '堂', '兵', '七', '露', '园', '牛', '哭',
        '旅', '街',
        '劳', '型', '烈', '姑', '陈', '莫', '鱼', '异', '抱', '宝', '权', '鲁', '简', '态', '级', '票', '怪', '寻',
        '杀', '律',
        '胜', '份', '汽', '右', '洋', '范', '床', '舞', '秘', '午', '登', '楼', '贵', '吸', '责', '例', '追', '较',
        '职', '属',
        '渐', '左', '录', '丝', '牙', '党', '继', '托', '赶', '章', '智', '冲', '叶', '胡', '吉', '卖', '坚', '喝',
        '肉', '遗',
        '救', '修', '松', '临', '藏', '担', '戏', '善', '卫', '药', '悲', '敢', '靠', '伊', '村', '戴', '词', '森',
        '耳', '差',
        '短', '祖', '云', '规', '窗', '散', '迷', '油', '旧', '适', '乡', '架', '恩', '投', '弹', '铁', '博', '雷',
        '府', '压',
        '超', '负', '勒', '杂', '醒', '洗', '采', '毫', '嘴', '毕', '九', '冰', '既', '状', '乱', '景', '席', '珍',
        '童', '顶',
        '派', '素', '脱', '农', '疑', '练', '野', '按', '犯', '拍', '征', '坏', '骨', '余', '承', '置', '臓', '彩',
        '灯', '巨',
        '琴', '免', '环', '姆', '暗', '换', '技', '翻', '束', '增', '忍', '餐', '洛', '塞', '缺', '忆', '判', '欧',
        '层', '付',
        '阵', '玛', '批', '岛', '项', '狗', '休', '懂', '武', '革', '良', '恶', '恋', '委', '拥', '娜', '妙', '探',
        '呀', '营',
        '退', '摇', '弄', '桌', '熟', '诺', '宣', '银', '势', '奖', '宫', '忽', '套', '康', '供', '优', '课', '鸟',
        '喊', '降',
        '夏', '困', '刘', '罪', '亡', '鞋', '健', '模', '败', '伴', '守', '挥', '鲜', '财', '孤', '枪', '禁', '恐',
        '伙', '杰',
        '迹', '妹', '藸', '遍', '盖', '副', '坦', '牌', '江', '顺', '秋', '萨', '菜', '划', '授', '归', '浪', '听',
        '凡', '预',
        '奶', '雄', '升', '碃', '编', '典', '袋', '莱', '含', '盛', '济', '蒙', '棋', '端', '腿', '招', '释', '介',
        '烧', '误',
        '乾', '坤']
    x = random.randint(0, len(xing) - 1)
    y = random.randint(0, len(ming1) - 1)
    z = random.randint(0, len(ming2) - 1)
    return xing[x] + ming1[y] + ming2[z]

def get_city():
    x = {'河北省': ['石家庄', '唐山', '秦皇岛', '承德'],
         '山东省': ['济南','青岛','淄博','枣庄','东营','烟台','潍坊','济宁','泰安','威海','日照','莱芜','临沂','德州','聊城','滨州','菏泽'],
         '湖南省': ['长沙','株州','湘潭','衡阳','邵阳','岳阳','常德','张家界','益阳','郴州','永州','怀化','娄底','湘西'],
         '江西省': ['南昌','景德镇','萍乡','九江','新余','鹰潭','赣州','吉安','宜春','抚州','上饶'],
         '安徽省': ['合肥', '芜湖', '蚌埠', '淮南', '马鞍山', '淮北', '铜陵'],
         '福建省': ['福州', '厦门', '漳州', '泉州', '三明', '莆田', '南平', '龙岩', '宁德'],
         '山西省': ['太原', '大同', '晋城', '运城', '临汾', '阳泉', '朔州', '晋中', '吕梁'],
         '内蒙古': ['呼和浩特', '包头', '乌海', '鄂尔多斯', '赤峰', '通辽', '呼伦贝尔', '巴彦淖尔', '乌兰察布', '兴安盟', '锡林郭勒盟'],
         '河南省': ['郑州', '开封', '南阳', '洛阳', '平顶山', '安阳', '鹤壁', '新乡', '济源', '濮阳', '许昌', '漯河「三门峡', '商丘', '信阳', '周口', '驻马店'],
         '湖北省': ['武汉', '黄石', '十堰', '宜昌', '襄樊', '鄂州', '荆州', '黄风', '咸宁', '随州', '神农架', '恩施土家族苗族', '仙桃', '潜江', '天门'],
         '广东省': ['广州','深圳','珠海','佛山','江门','肇庆','惠州','东莞','中山','韶关','汕头','湛江','茂名','梅州','汕尾','河源','阳江','清远','潮州','揭阳','云浮'],
         '广西省': ['南宁','柳州','桂林','梧州','北海','防城港','钦州','贵港','玉林','百色','贺州','河池','来宾','崇左']
         }
    for i in range(1):
        s = list(x.keys())  # 省列表
        sheng = random.choice(s)  # 随机选一个省
        city = random.choice(x[sheng])  # 随机选一人市
    return sheng+city

def get_ladingnumber():
    global min
    list = []
    x = ['TKNG', 'HYJZ', 'QPYT', 'XNCX', 'OXVX']
    for j in range(5):
        for i in range(1, 1000):
            str_num = x[j] + "{:07d}".format(i)
            list.append(str_num)
    list1 = list
    while min > len(list):
        list = list + list1
    list = list[: min]
    return list

def get_firmname():
    global min
    list = []
    x = [
'宜昌裕丰国际物流有限公司',
'起航船务代理有限公司',
'振智物流有限公司',
'佳予舜呈物流有限公司',
'明佳物流有限公司',
'明琦商贸有限公司',
'绿时代福佑商贸有限公司',
'天蛟国际物流有限公司',
'桦业国际物流有限公司',
'越洋国际货运代理有限公司',
'乾宇国际货运代理有限公司',
'中亿运国际货运代理有限公司',
'通速货运代理有限公司',
'永昌顺运输有限公司',
'凯天亿方国际货运代理有限公司',
'浩通国际货运代理有限公司',
'近洋国际货运代理有限公司',
'顺泽国通货运代理有限公司',
'知行货运代理有限公司',
'晨光顺达货运代理有限公司'
    ]
    for i in range(min):
        list.append(x[random.randint(0, len(x)-1)])
    return list

def get_containerid():
    global min
    global list_containerid
    x = ['YWCM', 'MWYW', 'FLTI', 'OSIF', 'XBCB', 'FFKG', 'KPOX']
    container_list = []
    list_id = []
    for j in range(7):
        for i in range(1, 10000):
            str_num = x[j] + "{:04d}".format(i)
            list_id.append(str_num)
    for i in range(min):
        list_id1 = set()
        str1 = ""
        for k in range(random.randint(1, 6)):
            list_id1.add(list_id[random.randint(0, 10000)])
        list_id2 = list(list_id1)
        str1 = str(list_id2[0])
        for j in range(1, len(list_id1)):
            str1 = str1 + ',' + str(list_id2[j])
        container_list.append(str1)
        list_containerid.append(list_id2)
    list_containerid = list_containerid[: min]
    print(list_containerid)
    container_list = container_list[: min]
    return container_list

def get_goodsname():
    global min
    list = []
    x = ['稻谷', '麦', '杂粮', '煤炭', '生铁', '钢锭', '纺织品', '食品', '日用百货', '金属轻工业品', '手工业产品', '砖', '黄沙', '石子', '瓦',
         '水泥', '矿石', '茶叶', '塑料制品', '家电', '茄子', '白菜', '大豆', '蛋白粉', '花生', '苹果', '香蕉', '牛奶', '奶粉', '辣椒', '大蒜'
         '仪器', '仪表', '红酒', '白酒', '中药', '零件', '灯具', '电子配件', '分析仪器', '钟表', '医疗器械', '玻璃', '陶器', '瓷器', '书籍',
         '大蒜', '印刷品', '玩具', '西药', '生物制品', '烟草加工品', '罐头', '羊奶', '盐', '味精', '醋', '酱油', '化妆品', '护肤品', '燃料', '树脂'
         ]
    for i in range(min):
        list.append(x[random.randint(0, len(x)-1)])
    return list

def get_goodsweight():
    global min
    list = []
    for i in range(min):
        list.append(random.randint(10, 500))
    return list

def get_Shipping_Companies():
    global min
    list = []
    x = [
'澳国航运',
'美国总统',
'邦拿美',
'波罗的海',
'中波',
'南美邮船',
'智利航运',
'中日轮渡',
'天敬海运',
'达飞轮船',
'京汉航运',
'中远集运',
'朝阳公司',
'达贸国际',
'德国胜利',
'埃及船务',
'香港长荣',
'远东海洋',
'金发船务',
'海华轮船',
'浩洲船务',
'韩进海运',
'香港海运',
'香港明华',
'赫伯罗特',
'现代商船',
'海隆轮船',
'金华航运',
'高丽海运',
'七星轮船',
'育海航运',
'中福轮船',
'山东海丰',
'墨西哥航运',
'天海货运',
'东航船务',
'宁波泛洋',
'阿拉伯轮船',
'立荣海运',
'环球船务',
'万海股份',
'伟航船务',
'威兰德船务',
'阳明公司',
'以星轮船',
'浙江远洋',
'联华航业',
'联丰船务',
'意邮船',
'马国航运',
'商船三井',
'地中海',
'马士基',
'民生神原',
'太古船代',
'渣华邮船',
'新海皇',
'北欧亚航',
'宁波远洋',
'南星海运',
'沙特航运',
'日本邮船',
'东方海外',
'英国铁行',
'泛洲海运',
'太平船务',
'泛洋商船',
'瑞克麦斯',
'美商海陆',
'南非轮船',
'东映海运',
'国际轮渡',
'中海发展',
'长锦公司',
'锦江船代',
'志晓船务',
'中外运',
]
    for i in range(min):
        list.append(x[random.randint(0, len(x)-1)])
    return list

def get_Shippingname():
    global min
    list = []
    x = ['的', '一', '是', '了', '我', '不', '人', '在', '他', '有', '这', '个', '上', '们', '来', '到', '时', '大',
        '地', '为',
        '子', '中', '你', '说', '生', '国', '年', '着', '就', '那', '和', '要', '她', '出', '也', '得', '里', '后',
        '自', '以',
        '会', '家', '可', '下', '而', '过', '天', '去', '能', '对', '小', '多', '然', '于', '心', '学', '么', '之',
        '都', '好',
        '看', '起', '发', '当', '没', '成', '只', '如', '事', '把', '还', '用', '第', '样', '道', '想', '作', '种',
        '开', '美',
        '总', '从', '无', '情', '己', '面', '最', '女', '但', '现', '前', '些', '所', '同', '日', '手', '又', '行',
        '意', '动']
    y = ['夏', '困', '刘', '罪', '亡', '鞋', '健', '模', '败', '伴']
    for i in range(100):
        for j in range(10):
            str_name = x[i] + y[j]
            for k in range(0, 1000):
                list.append(str_name + str(k))
    list = list[: min]
    return list

def get_worktime():
    global list_date1
    global list_date2
    global list_date3
    global list_date4
    global list_date5
    global min
    list_date6 = []
    dt1 = datetime(2018, 1, 1, 0, 0, 0)
    for i in range(min):
        dt2 = dt1 + timedelta(hours=random.randint(1, 36000))
        list_date1.append(dt2)
        dt3 = dt2 + timedelta(hours=random.randint(1, 20))
        list_date2.append(dt3)
        dt4 = dt3 + timedelta(hours=1)
        list_date3.append(dt4)
        dt5 = dt4 + timedelta(days=random.randint(1, 7))
        list_date4.append(dt5)
        dt6 = dt5 + timedelta(hours=random.randint(1, 4))
        list_date5.append(dt6)
        dt7 = dt6 + timedelta(hours=random.randint(1, 20))
        list_date6.append(dt7)
        list_date1[i] = str(list_date1[i])
        list_date2[i] = str(list_date2[i])
        list_date3[i] = str(list_date3[i])
        list_date4[i] = str(list_date4[i])
        list_date5[i] = str(list_date5[i])
        list_date6[i] = str(list_date6[i])

    return list_date5, list_date6



def get_portname():
    global min
    global list_port1
    global list_port2
    x = ['石家庄港', '唐山港', '秦皇岛港',
         '济南港', '青岛港', '淄博港', '枣庄港', '东营港', '烟台港', '潍坊港', '济宁港', '泰安港', '威海港', '日照港', '莱芜港', '临沂港', '德州港', '聊城港', '滨州港',
         '长沙港', '株州港', '湘潭港', '衡阳港', '邵阳港', '岳阳港', '常德港', '张家界港', '益阳港', '郴州港', '永州港', '怀化港', '娄底港',
         '南昌港', '景德镇港', '萍乡港', '九江港', '新余港', '鹰潭港', '赣州港', '吉安港', '宜春港', '抚州港',
         '合肥港', '芜湖港', '蚌埠港', '淮南港', '马鞍山港', '淮北港',
         '福州港', '厦门港', '漳州港', '泉州港', '三明港', '莆田港', '南平港', '龙岩港',
         '太原港', '大同港', '晋城港', '运城港', '临汾港', '阳泉港', '朔州港', '晋中港',
         '呼和浩特港', '包头港', '乌海港', '鄂尔多斯港', '赤峰港', '通辽港', '呼伦贝尔港', '巴彦淖尔港', '乌兰察布港', '兴安盟港',
         '郑州港', '开封港', '南阳港', '洛阳港', '平顶山港', '安阳港', '鹤壁港', '新乡港', '济源港', '濮阳港', '许昌港', '漯河「三门峡港', '商丘港', '信阳港', '周口港',
         '武汉港', '黄石港', '十堰港', '宜昌港', '襄樊港', '鄂州港', '荆州港', '黄风港', '咸宁港', '随州港', '神农架港', '仙桃港', '潜江港',
         '广州港', '深圳港', '珠海港', '佛山港', '江门港', '肇庆港', '惠州港', '东莞港', '中山港', '韶关港', '汕头港', '湛江港', '茂名港', '梅州港', '汕尾港', '河源港',
         '阳江港', '清远港', '潮州港', '揭阳港',
         '南宁港', '柳州港', '桂林港', '梧州港', '北海港', '防城港港', '钦州港', '贵港港', '玉林港', '百色港', '贺州港', '河池港', '来宾港'
         ]
    for i in range(min):
        s = random.randint(0, len(x)-1)
        list_port1.append(x[s])
        list_port2.append(x[s-1])

    return list_port1, list_port2

def get_BoxSize():
    global min
    list = []
    for i in range(min):
        list.append('20')
    return list

def get_position(num):
    postion_list = []
    for i in range(num):
        s = string.ascii_uppercase #所有大写字母(A-Z)
        r = random.choice(s)
        s_num = "{:06d}".format(random.randint(0, 100000))
        x = r + s_num
        postion_list.append(x)
    return postion_list


def get_Port_and_containerDynamics_containerID_Datetime():
    global list_containerid
    global ladingnumber
    global min
    global list_port1
    global list_port2
    global list_date1
    global list_date5
    containerid_now = []
    ladingnumber_now =[]
    port_now = []
    date_now = []
    boxsize_now = []
    storehouse_now = []
    for i in range(min):
        for j in range(len(list_containerid[i])):
            port_now.append(list_port1[i])
            port_now.append(list_port2[i])
            containerid_now.append(list_containerid[i][j])
            containerid_now.append(list_containerid[i][j])
            boxsize_now.append("20")
            boxsize_now.append("20")
            ladingnumber_now.append(ladingnumber[i])
            ladingnumber_now.append(ladingnumber[i])
            storehouse_now.append("出库")
            storehouse_now.append("入库")
            date_now.append(list_date1[i][: 11])
            date_now.append(list_date5[i][: 11])
    postion_list = get_position(len(containerid_now))
    return port_now, containerid_now, boxsize_now, ladingnumber_now, postion_list, storehouse_now, date_now




def xw_toExcel(name, phone, id, city, fileName):  # xlsxwriter库储存数据到excel
    workbook = xw.Workbook(fileName)  # 创建工作簿
    worksheet1 = workbook.add_worksheet("sheet1")  # 创建子表
    worksheet1.activate()  # 激活表
    title = ['客户名称', '客户编号', '手机号', '省市区']  # 设置表
    worksheet1.write_row('A1', title)  # 从A1单元格开始写入表头
    i = 2  # 从第二行开始写入数据
    worksheet1.write_column("A2", name)
    worksheet1.write_column("B2", phone)
    worksheet1.write_column("C2", id)
    worksheet1.write_column("D2", city)
    workbook.close()  # 关闭表

def xw_toExcel1(ladingnumber, name, id, firmname, containerid, goodsname, goodsweight, fileName):  # xlsxwriter库储存数据到excel
    workbook = xw.Workbook(fileName)  # 创建工作簿
    worksheet1 = workbook.add_worksheet("sheet1")  # 创建子表
    worksheet1.activate()  # 激活表
    title = ['提单号', '货主名称', '货主代码', '物流公司', '集装箱箱号', '货物名称', '货重（吨）']  # 设置表头
    worksheet1.write_row('A1', title)  # 从A1单元格开始写入表头
    i = 2  # 从第二行开始写入数据
    worksheet1.write_column("A2", ladingnumber)
    worksheet1.write_column("B2", name)
    worksheet1.write_column("C2", id)
    worksheet1.write_column("D2", firmname)
    worksheet1.write_column("E2", containerid)
    worksheet1.write_column("F2", goodsname)
    worksheet1.write_column("G2", goodsweight)
    workbook.close()  # 关闭表

def xw_toExcel2(firmname, shipname, list_date1, list_date2, list_date3, list_date4, workport, ladingnumber, containerid, boxsize, listport1, listport2, fileName):  # xlsxwriter库储存数据到excel
    workbook = xw.Workbook(fileName)  # 创建工作簿
    worksheet1 = workbook.add_worksheet("sheet1")  # 创建子表
    worksheet1.activate()  # 激活表
    title = ['船公司', '船名称', '作业开始时间', '作业结束时间', '始发时间', '到达时间', '作业港口', '提单号', '集装箱箱号', '箱尺寸（TEU）', '启运地', '目的地']  # 设置表头
    worksheet1.write_row('A1', title)  # 从A1单元格开始写入表头
    i = 2  # 从第二行开始写入数据
    worksheet1.write_column("A2", firmname)
    worksheet1.write_column("B2", shipname)
    worksheet1.write_column("C2", list_date1)
    worksheet1.write_column("D2", list_date2)
    worksheet1.write_column("E2", list_date3)
    worksheet1.write_column("F2", list_date4)
    worksheet1.write_column("G2", workport)
    worksheet1.write_column("H2", ladingnumber)
    worksheet1.write_column("I2", containerid)
    worksheet1.write_column("J2",  boxsize)
    worksheet1.write_column("K2", listport1)
    worksheet1.write_column("L2", listport2)
    workbook.close()  # 关闭表

def xw_toExcel3(port_now, containerid_now, boxsize_now, ladingnumber_now, postion_list, storehouse_now, date_now, fileName):  # xlsxwriter库储存数据到excel
    workbook = xw.Workbook(fileName)  # 创建工作簿
    i = 0
    x = 1000000
    for j in range(int(len(port_now)/x) + 1):
        worksheet1 = workbook.add_worksheet("sheet"+str(j+1))  # 创建子表
        worksheet1.activate()  # 激活表
        title = ['堆存港口', '集装箱箱号', '箱尺寸（TEU）', '提单号', '堆场位置', '操作', '操作日期']  # 设置表头
        worksheet1.write_row('A1', title)  # 从A1单元格开始写入表头
        worksheet1.write_column("A2", port_now[i:i+x])
        worksheet1.write_column("B2", containerid_now[i:i+x])
        worksheet1.write_column("C2", boxsize_now[i:i+x])
        worksheet1.write_column("D2", ladingnumber_now[i:i+x])
        worksheet1.write_column("E2", postion_list[i:i+x])
        worksheet1.write_column("F2", storehouse_now[i:i+x])
        worksheet1.write_column("G2", date_now[i:i+x])
        i = i+x
    workbook.close()  # 关闭表

def get_Customer_Information(num):
    global name_list
    global set_phone
    global set_ID
    global min
    city_list = []
    set_iphone = set()
    set_id = set()
    for i in range(num):
        x = ""
        y = ""
        z = ""
        x = "".join(randname())
        y = "".join(randomnum())
        z = "".join(randid())
        name_list.append(x)
        set_iphone.add(y)
        set_id.add(z)
        city_list.append(get_city())
        # print(x)
        # print(y)
        # print(z)
    print(name_list)
    print(set_iphone)
    print(set_id)
    print(city_list)
    min = len(set_iphone)
    if len(set_id) < min:
        min = len(set_id)

    fileName = '客户信息.xlsx'
    set_phone = list(set_iphone)
    set_ID = list(set_id)
    set_phone = set_phone[: min]
    set_ID = set_ID[: min]
    name_list = name_list[: min]
    city_list = city_list[: min]
    xw_toExcel(name_list, set_ID, set_iphone, city_list, fileName)

def get_Logistics_Information():
    global name_list
    global set_phone
    global set_ID
    global ladingnumber
    global containerid
    ladingnumber = get_ladingnumber()
    firmname = get_firmname()
    containerid = get_containerid()
    goodsname = get_goodsname()
    goodsweight = get_goodsweight()
    fileName = '物流信息.xlsx'
    xw_toExcel1(ladingnumber, name_list, set_ID, firmname, containerid, goodsname, goodsweight, fileName)

def get_LoadingandUnload_Table():
    global list_date1
    global list_date2
    global list_date3
    global list_date4
    global list_port1
    global list_port2
    global ladingnumber
    fileName1 = '装货表.xlsx'
    fileName2 = '卸货表.xlsx'
    firmname = get_Shipping_Companies()
    shipname = get_Shippingname()
    get_portname()
    boxsize = get_BoxSize()
    list_date5, list_date6 = get_worktime()
    xw_toExcel2(firmname, shipname, list_date1, list_date2, list_date3, list_date4, list_port1, ladingnumber,containerid, boxsize, list_port1, list_port2, fileName1) # xlsxwriter库储存数据到excel
    xw_toExcel2(firmname, shipname, list_date5, list_date6, list_date3, list_date4, list_port2, ladingnumber,containerid, boxsize, list_port1, list_port2, fileName2)

def get_ContainerDynamics_Table():
    fileName = '集装箱动态.xlsx'
    port_now, containerid_now, boxsize_now, ladingnumber_now, postion_list, storehouse_now, date_now = get_Port_and_containerDynamics_containerID_Datetime()
    xw_toExcel3(port_now, containerid_now, boxsize_now, ladingnumber_now, postion_list, storehouse_now, date_now, fileName)

if __name__ == '__main__':
    num = 1000000
    get_Customer_Information(num)
    get_Logistics_Information()
    get_LoadingandUnload_Table()
    get_ContainerDynamics_Table()

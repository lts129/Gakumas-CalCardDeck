import pandas as pd
import os


class Card:
    def __init__(self, card_face, current_level, max_level, name, color, card_type,
                 keyword1, keyword2, keyword4, keyword5, keyword6, special_keyword, acquisition,E1,E2,E3,EE1,EE2,EE3,
                 resident_category=None):
        # --- 基础属性 ---
        self.card_face = card_face
        self.current_level = current_level
        self.max_level = max_level
        self.name = name
        self.color = color
        self.card_type = card_type
        self.special_keyword = special_keyword
        self.acquisition = acquisition
        self.resident_category = resident_category  # 新增字段：抽取类别

        # --- 词条名称存储 ---
        self.kw1_name = keyword1
        self.kw2_name = keyword2
        self.kw4_name = keyword4
        self.kw5_name = keyword5
        self.kw6_name = keyword6

        # --- 新增：词条属性值 (当前, 最大, 极限) ---
        # 格式: (当前值, 最大值, 极限值)
        self.kw1_stats = (None, None, None)
        self.kw2_stats = (None, None, None)
        self.kw4_stats = (None, None, None)
        self.kw5_stats = (None, None, None)
        self.kw6_stats = (None, None, None)

        #道具等价属性
        self.E1=E1
        self.E2=E2
        self.E3=E3
        self.EE1=EE1
        self.EE2=EE2
        self.EE3=EE3

    def __repr__(self):
        return f"<Card {self.name} Lv.{self.current_level}/{self.max_level}>"


def load_keyword_database(file_path):
    """
    加载基础词条表，构建快速查找字典。
    返回结构: {(词条名称, 位置, 模式): {等级: 属性值}}
    """
    if not os.path.exists(file_path):
        print(f"错误：找不到文件 '{file_path}'")
        return {}

    try:
        # 读取'词条'工作表
        df_terms = pd.read_excel(file_path, sheet_name='词条')

        # 定义固定的等级列名列表 (根据描述)
        level_columns = [1, 5, 10, 15, 20, 21, 25, 26, 30, 31, 35, 36, 40, 41, 45, 46, 50, 51, 55, 56, 60]
        # 确保这些列存在于DataFrame中，并转换为字符串以便匹配（防止Excel读成数字或字符串不一致）
        # 这里假设Excel表头就是这些数字，如果是字符串"1", "5"等，pandas通常会自动处理，但为了安全我们统一转换

        # 构建数据库字典
        # Key: (词条名称, 位置, 模式)
        # Value: 一个字典 {等级: 数值}
        term_db = {}

        for i, row in df_terms.iterrows():
            name = row['词条名称']
            pos = row['位置']
            mode = row['模式']

            key = (name, pos, mode)

            level_data = {}
            for lvl in level_columns:
                # 尝试获取该等级的值，如果列不存在或值为空，则跳过
                # 注意：需要确保列名匹配。如果Excel列名是数字，直接用lvl；如果是字符串，用str(lvl)
                # 这里做一个兼容处理：检查列名是否存在
                col_name = lvl if lvl in df_terms.columns else str(lvl)

                if col_name in df_terms.columns:
                    val = row[col_name]
                    if pd.notna(val):  # 排除空值
                        level_data[lvl] = val

            term_db[key] = level_data

        return term_db

    except Exception as e:
        print(f"读取基础词条表时发生错误: {e}")
        return {}


def find_closest_level_value(level_data_dict, target_level):
    """
    在等级数据字典中查找目标等级的值。
    逻辑：如果存在完全匹配的等级，返回该值；
         否则，返回小于等于目标等级的最大已知等级的值。
         如果目标等级小于所有已知等级，返回最小等级的值（或None，视需求而定，这里返回最小）。
    """
    if not level_data_dict:
        return None

    available_levels = sorted(level_data_dict.keys())

    # 1. 精确匹配
    if target_level in available_levels:
        return level_data_dict[target_level]

    # 2. 查找小于等于目标等级的最大值
    # 过滤出 <= target_level 的等级
    valid_levels = [l for l in available_levels if l <= target_level]

    if valid_levels:
        closest_level = max(valid_levels)
        return level_data_dict[closest_level]
    else:
        # 如果目标等级比表中最低等级还小（理论上不会发生，因为最低是1），返回最低等级的值
        return level_data_dict[available_levels[0]]


def process_all_cards(card_file, term_file, resident_upgrade=False):
    SSR_cards = process_cards(card_file, term_file, 'SSR', 1, resident_upgrade)
    SR_cards = process_cards(card_file, term_file, 'SR', 11, resident_upgrade)
    FSSR_cards = process_cards(card_file, term_file, 'FSSR', 2, resident_upgrade)
    cards = SSR_cards + SR_cards + FSSR_cards
    return cards

def process_cards(card_file, term_file, sheet_name='SSR', fixed_mode=1, resident_upgrade=False):
    """
    读取并处理卡牌数据。
    
    参数:
        card_file: 卡牌数据文件路径
        term_file: 基础词条文件路径
        sheet_name: 工作表名称
        fixed_mode: 词条模式
        resident_upgrade: 是否启用常驻升级模式（常驻卡max_level+5）
    """
    # 1. 加载词条数据库
    term_db = load_keyword_database(term_file)
    if not term_db:
        return []

    # 2. 读取卡牌数据
    if not os.path.exists(card_file):
        print(f"错误：找不到文件 '{card_file}'")
        return []

    try:
        df_cards = pd.read_excel(card_file, sheet_name=sheet_name)
    except Exception as e:
        print(f"读取卡牌数据失败: {e}")
        return []

    # 定义需要处理的词条映射：(卡牌列名, 词条位置)
    # 注意：词条3被跳过了，所以是 1->1, 2->2, 4->4, 5->5, 6->6
    kw_mapping = [
        ('词条1', 1),
        ('词条2', 2),
        ('词条4', 4),
        ('词条5', 5),
        ('词条6', 6)
    ]

    card_objects = []

    limit_level = 60

    for _, row in df_cards.iterrows():
        # 创建基础卡牌对象
        card = Card(
            card_face=row.get('卡面'),
            current_level=row.get('当前等级', 1),
            max_level=row.get('最高等级', 1),
            name=row.get('名称'),
            color=row.get('颜色'),
            card_type=row.get('类型'),
            keyword1=row.get('词条1'),
            keyword2=row.get('词条2'),
            keyword4=row.get('词条4'),
            keyword5=row.get('词条5'),
            keyword6=row.get('词条6'),
            special_keyword=row.get('特殊词条'),
            acquisition=row.get('获得'),
            E1=row.get('一属性'),
            E2=row.get('二属性'),
            E3=row.get('三属性'),
            EE1=row.get('等价一属性'),
            EE2=row.get('等价二属性'),
            EE3=row.get('等价三属性'),
            resident_category=row.get('类型.1')  # 新增：读取T列（类型.1）作为抽取类别
        )

        # 确保等级是整数
        try:
            curr_lvl = int(card.current_level)
            max_lvl = int(card.max_level)
        except (ValueError, TypeError):
            curr_lvl = 1
            max_lvl = 1

        # 常驻升级模式：常驻卡的最高等级+5
        if resident_upgrade and card.resident_category == '常驻' and sheet_name=="SSR" and max_lvl<60:
            max_lvl += 5
            card.max_level = max_lvl  # 更新卡牌对象的max_level

        # 处理每个词条
        for col_name, position in kw_mapping:
            kw_name = getattr(card, f"{col_name.replace('词条', 'kw')}_name")  # 获取词条名称变量，如 kw1_name

            # 如果词条名为空或NaN，跳过
            if pd.isna(kw_name) or kw_name == '':
                continue

            # 构建查找键
            key = (kw_name, position, fixed_mode)

            if key in term_db:
                level_data = term_db[key]

                # 计算三个值
                val_current = find_closest_level_value(level_data, curr_lvl)
                val_max = find_closest_level_value(level_data, max_lvl)
                val_limit = find_closest_level_value(level_data, limit_level)

                # 赋值给对应的属性
                # 动态确定属性名，例如 词条1 -> kw1_stats
                attr_name = f"kw{position}_stats"
                setattr(card, attr_name, (val_current, val_max, val_limit))
            else:
                pass

        card_objects.append(card)

    return card_objects

def calculate_card_stats(cards, rules_db, mode=1):
    """
    为每张卡牌计算9种数值情况，并存储在对象中。
    新增属性示例: card.score_current_attr1, card.score_limit_attr3 等
    
    注意：resident_upgrade 逻辑已在 process_cards 中处理，
    常驻卡的 max_level 已在词条属性计算前被提升，因此 stats tuple 已包含升级后的值。
    """
    # 词条位置映射 (跳过3)
    kw_positions = [1, 2, 4, 5, 6]

    for card in cards:
        # 初始化9个得分为0
        scores = {
            'current_attr1': 0.0, 'current_attr2': 0.0, 'current_attr3': 0.0,
            'max_attr1': 0.0, 'max_attr2': 0.0, 'max_attr3': 0.0,
            'limit_attr1': 0.0, 'limit_attr2': 0.0, 'limit_attr3': 0.0
        }

        for pos in kw_positions:
            # 获取词条名称 (如 kw1_name)
            kw_name_attr = f"kw{pos}_name"
            kw_name = getattr(card, kw_name_attr, None)

            if not kw_name or pd.isna(kw_name):
                continue

            # 获取该词条的三条属性基础值 (当前, 最大, 极限)
            # 对应 card.kw1_stats = (val_curr, val_max, val_limit)
            stats_attr = f"kw{pos}_stats"
            stats_tuple = getattr(card, stats_attr, (None, None, None))

            if stats_tuple[0] is None:
                continue

            val_curr, val_max, val_limit = stats_tuple

            # 从规则库获取触发次数
            if kw_name not in rules_db:
                continue

            rule = rules_db[kw_name]
            t1 = rule['attr1']
            t2 = rule['attr2']
            t3 = rule['attr3']

            # 累加计算
            # 当前等级状态
            scores['current_attr1'] += val_curr * t1
            scores['current_attr2'] += val_curr * t2
            scores['current_attr3'] += val_curr * t3

            # 最高等级状态
            scores['max_attr1'] += val_max * t1
            scores['max_attr2'] += val_max * t2
            scores['max_attr3'] += val_max * t3

            # 极限等级状态
            scores['limit_attr1'] += val_limit * t1
            scores['limit_attr2'] += val_limit * t2
            scores['limit_attr3'] += val_limit * t3

        # 处理道具属性
        if mode > 0:
            scores['current_attr1'] += getattr(card, 'E1', 0)
            scores['current_attr2'] += getattr(card, 'E2', 0)
            scores['current_attr3'] += getattr(card, 'E3', 0)
            scores['max_attr1'] += getattr(card, 'E1', 0)
            scores['max_attr2'] += getattr(card, 'E2', 0)
            scores['max_attr3'] += getattr(card, 'E3', 0)
            scores['limit_attr1'] += getattr(card, 'E1', 0)
            scores['limit_attr2'] += getattr(card, 'E2', 0)
            scores['limit_attr3'] += getattr(card, 'E3', 0)
            if mode > 1:
                scores['current_attr1'] += getattr(card, 'EE1', 0)
                scores['current_attr2'] += getattr(card, 'EE2', 0)
                scores['current_attr3'] += getattr(card, 'EE3', 0)
                scores['max_attr1'] += getattr(card, 'EE1', 0)
                scores['max_attr2'] += getattr(card, 'EE2', 0)
                scores['max_attr3'] += getattr(card, 'EE3', 0)
                scores['limit_attr1'] += getattr(card, 'EE1', 0)
                scores['limit_attr2'] += getattr(card, 'EE2', 0)
                scores['limit_attr3'] += getattr(card, 'EE3', 0)

        if card.current_level == 0:
            scores['current_attr1'] = 0.0
            scores['current_attr2'] = 0.0
            scores['current_attr3'] = 0.0
        if card.max_level == 0:
            scores['max_attr1'] = 0.0
            scores['max_attr2'] = 0.0
            scores['max_attr3'] = 0.0
        # 将计算结果绑定到卡牌对象上，方便后续排序
        for key, value in scores.items():
            setattr(card, f"score_{key}", value)

    return cards

def load_calculation_rules(file_path,mode=0):
    """
    读取实际计算.xlsx，处理触发次数逻辑。
    返回字典: {词条名称: {'attr1': count, 'attr2': count, 'attr3': count}}
    """
    if not os.path.exists(file_path):
        print(f"错误：找不到文件 '{file_path}'")
        return {}

    try:
        df = pd.read_excel(file_path, sheet_name=f'当前模式{mode}')

        rules_db = {}

        for _, row in df.iterrows():
            name = row['词条名称']
            if pd.isna(name):
                continue

            # 获取原始次数
            c1 = row.get('第一属性触发次数', 0)
            c2 = row.get('第二属性触发次数', 0)
            c3 = row.get('第三属性触发次数', 0)
            max_c = row.get('累计最大触发次数', 0)

            # 处理空值
            c1 = 0 if pd.isna(c1) else c1
            c2 = 0 if pd.isna(c2) else c2
            c3 = 0 if pd.isna(c3) else c3
            max_c = 0 if pd.isna(max_c) else max_c

            # 逻辑：如果触发次数 > 累计最大，则以累计最大计算
            # 注意：如果累计最大为0或空，通常意味着无限制或数据错误，这里假设以原值为准或强制截断
            # 根据描述 "如果该触发次数大于累计最大触发次数则以累计最大触发次数计算"
            if max_c > 0:
                eff_c1 = min(c1, max_c)
                eff_c2 = min(c2, max_c)
                eff_c3 = min(c3, max_c)
            else:
                # 如果没有设置最大限制，保持原值
                eff_c1, eff_c2, eff_c3 = c1, c2, c3

            rules_db[name] = {
                'attr1': float(eff_c1),
                'attr2': float(eff_c2),
                'attr3': float(eff_c3)
            }

        return rules_db

    except Exception as e:
        print(f"读取计算规则表时发生错误: {e}")
        return {}

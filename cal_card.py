import os
from make_card import load_calculation_rules, calculate_card_stats, process_all_cards
import pandas as pd
from itertools import permutations
from dataclasses import dataclass
import logging


@dataclass
class CardSelectionConfig:
    """卡牌选择配置类"""
    color: str
    total: int
    sp_min: int

    def __post_init__(self):
        """验证配置数据"""
        if not isinstance(self.total, int) or not isinstance(self.sp_min, int):
            raise ValueError(f"数量和SP数必须是整数。得到: {self.total}, {self.sp_min}")
        if self.sp_min > self.total:
            raise ValueError(f"SP需求 ({self.sp_min}) 不能超过总数 ({self.total})")
    
    @property
    def rest(self):
        """计算非SP卡的名额数量"""
        return self.total - self.sp_min

class OptimizationConfig:
    """优化计算配置类"""
    def __init__(self, *args):
        """解析参数"""
        if len(args) < 4:
            raise ValueError("参数错误：需要至少 颜色, 总数, SP数, ..., 类型, resident_upgrade")
            
        # 解析参数
        self.target_type = args[-3]
        self.mode = args[-2]
        self.rulemode = args[-1]
        self.resident_upgrade = False  # 新增：常驻升级模式标志
        
        config_pairs = args[:-3]
        if len(config_pairs) % 3 != 0:
            raise ValueError("参数错误：颜色和数量配置必须每3个一组 (颜色, 总数, SP数)")
            
        self.color_order = []
        self.color_constraints = {}  # {'红': {'total': 3, 'sp_min': 1}, ...}
        
        for i in range(0, len(config_pairs), 3):
            c = config_pairs[i]
            total = config_pairs[i + 1]
            sp_min = config_pairs[i + 2]
            
            self.color_order.append(c)
            self.color_constraints[c] = CardSelectionConfig(c, total, sp_min)
            
        if sum(cfg.total for cfg in self.color_constraints.values()) != 6:
            print(f"警告：卡牌总数 ({sum(cfg.total for cfg in self.color_constraints.values())}) 不等于 6。")
            
        # 初始化日志
        logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
        self.logger = logging.getLogger(__name__)
        # self.logger.info("优化配置初始化完成")
        
    def get_base_paths(self):
        """获取基础文件路径"""
        base_dir = os.path.dirname(os.path.abspath(__file__))
        # print(base_dir)
        # self.logger.info(f"基础目录: {base_dir}")
        base_dir='./'
        return {
            'card': os.path.join(base_dir, '卡牌数据.xlsx'),
            'term': os.path.join(base_dir, '基础词条.xlsx'),
            'rule': os.path.join(base_dir, '实际计算.xlsx')
        }

def run_optimization(*args, **kwargs):
    """
    新入口函数 - 支持多种参数传递方式
    用法示例:
    1. run_optimization("红", 3, 1, "蓝", 2, 1, "黄", 1, 0, "理性", 1, 0)
    2. run_optimization("红", 3, 1, "蓝", 2, 1, "黄", 1, 0, "理性", 1, 0, resident_upgrade=True)
    3. config = {
        'colors': [
            {'color': '红', 'total': 3, 'sp_min': 1},
            {'color': '蓝', 'total': 2, 'sp_min': 1},
            {'color': '黄', 'total': 1, 'sp_min': 0}
        ],
        'target_type': '理性',
        'mode': 1,
        'rulemode': 0,
        'resident_upgrade': False
    }
    """
    try:
        # 提取 resident_upgrade 参数
        resident_upgrade = kwargs.get('resident_upgrade', False)
        
        # 处理参数 - 支持两种调用方式
        if len(args) >= 4 and all(isinstance(x, (str, int)) for x in args):
            # 方式1: 直接传参 ("红", 3, 1, "蓝", 2, 1, "黄", 1, 0, "理性", 1, 0)
            config = OptimizationConfig(*args)
        elif len(args) >= 1 and isinstance(args[0], dict) and 'colors' in args[0]:
            # 方式2: 传配置字典
            config_data = args[0]
            config = OptimizationConfig(
                *[item for color_cfg in config_data['colors'] 
                  for item in (color_cfg['color'], color_cfg['total'], color_cfg['sp_min'])],
                config_data['target_type'],
                config_data['mode'],
                config_data['rulemode']
            )
            # 从字典中获取 resident_upgrade
            resident_upgrade = config_data.get('resident_upgrade', False)
        else:
            raise ValueError("参数格式错误。请使用以下格式之一:\n"
                           "1. run_optimization(\"红\", 3, 1, \"蓝\", 2, 1, \"黄\", 1, 0, \"理性\", 1, 0)\n"
                           "2. run_optimization({'colors': [{'color': '红', 'total': 3, 'sp_min': 1}, ...], "
                           "'target_type': '理性', 'mode': 1, 'rulemode': 0, 'resident_upgrade': False})")
        
        # 获取基础路径
        paths = config.get_base_paths()
        # 1. 加载基础数据（常驻升级模式的逻辑已在 process_cards 中处理）
        cards = process_all_cards(paths['card'], paths['term'], resident_upgrade)
        if not cards:
            print("错误：未能加载卡牌数据。")
            return

        rules_db = load_calculation_rules(paths['rule'], config.rulemode)
        if not rules_db:
            print("警告：未能加载计算规则。")
            # 即使没有规则，也继续执行，使用默认计算方式
            
        # 2. 计算所有卡牌的9种数值
        cards = calculate_card_stats(cards, rules_db, config.mode)
        if not cards:
            print("错误：未能计算卡牌属性。")
            return
            
        # 3. 预处理：过滤类型，计算有效分数，标记属性和SP状态
        valid_types = {config.target_type, '通用'}
        # 颜色 -> 属性索引映射
        color_to_attr = {c: idx + 1 for idx, c in enumerate(config.color_order)}

        filtered_grouped = {c: {'sp': [], 'non_sp': []} for c in config.color_constraints.keys()}

        for card in cards:
            t = card.card_type
            if pd.isna(t) or t not in valid_types:
                continue

            c = card.color
            if pd.isna(c) or c not in config.color_constraints:
                continue

            # 标记属性
            attr_idx = color_to_attr[c]
            curr_attr = f"score_current_attr{attr_idx}"
            max_attr = f"score_max_attr{attr_idx}"
            lim_attr = f"score_limit_attr{attr_idx}"

            card.effective_current_score = getattr(card, curr_attr, 0.0)
            card.effective_max_score = getattr(card, max_attr, 0.0)
            card.effective_limit_score = getattr(card, lim_attr, 0.0)
            card.matched_color = c
            card.matched_attr_idx = attr_idx

            # 标记SP状态 (检查词条2)
            is_sp = (card.kw2_name == 'SP率')
            card.is_sp = is_sp

            # 分组
            if is_sp:
                filtered_grouped[c]['sp'].append(card)
            else:
                filtered_grouped[c]['non_sp'].append(card)

        # 4. 执行优化
        prepared_data = prepare_candidate_pools(filtered_grouped, config)
        if not prepared_data:
            # config.logger.error("未能生成候选池。")
            print("错误：未能生成候选池。")
            return

        # 5. 计算最佳队伍
        best_team_current, best_score_current = select_best_team_sp_constrained(
            prepared_data, config, 'current'
        )
        best_team_max, best_score_max = select_best_team_sp_constrained(
            prepared_data, config, 'max'
        )

        # 6. 输出结果
        return (best_team_current, best_score_current), (best_team_max, best_score_max)
    except Exception as e:
        print(f"优化过程中出现错误: {str(e)}")
        raise

def prepare_candidate_pools(grouped_cards, config):
    """
    准备候选池并预排序
    返回: {
        color: {
            'total': int, 
            'sp_min': int,
            'sp_mode': list,  # 按模式分数排序的SP卡
            'non_sp_mode': list,  # 按模式分数排序的非SP卡
            'all_mode': list,  # 混合排序的卡牌 (SP + 非SP)
            'sp_limit': list,  # 按极限分数排序的SP卡
            'non_sp_limit': list  # 按极限分数排序的非SP卡
        }, ...
    }
    """
    score_attr = 'effective_current_score' if config.mode == 'current' else 'effective_max_score'
    limit_attr = 'effective_limit_score'

    prepared_data = {}

    for color, constraint in config.color_constraints.items():
        total_needed = constraint.total
        sp_needed = constraint.sp_min
        rest_needed = total_needed - sp_needed

        sp_cards = grouped_cards[color]['sp']
        non_sp_cards = grouped_cards[color]['non_sp']

        # 排序
        sp_sorted_mode = sorted(sp_cards, key=lambda c: getattr(c, score_attr, 0), reverse=True)
        non_sp_sorted_mode = sorted(non_sp_cards, key=lambda c: getattr(c, score_attr, 0), reverse=True)

        sp_sorted_limit = sorted(sp_cards, key=lambda c: getattr(c, limit_attr, 0), reverse=True)
        non_sp_sorted_limit = sorted(non_sp_cards, key=lambda c: getattr(c, limit_attr, 0), reverse=True)

        # 合并用于常规位选择的列表 (SP + 非SP)，保持分数顺序
        all_sorted_mode = sorted(
            sp_cards + non_sp_cards, 
            key=lambda c: getattr(c, score_attr, 0), 
            reverse=True
        )

        prepared_data[color] = {
            'total': total_needed,
            'sp_min': sp_needed,
            'rest': rest_needed,
            'sp_mode': sp_sorted_mode,
            'non_sp_mode': non_sp_sorted_mode,
            'all_mode': all_sorted_mode,
            'sp_limit': sp_sorted_limit,
            'non_sp_limit': non_sp_sorted_limit
        }

    return prepared_data


def select_best_team_sp_constrained(grouped_cards, config, mode='current'):
    """
    核心算法：带SP数量限制的选卡。
    逻辑：
    1. 对每种颜色，分别对 SP池 和 非SP池 排序。
    2. 初始构建：选 Top-N_sp (SP池) + Top-N_rest (剩余所有池)。
    3. 遍历极限位：
       - 如果极限位是SP卡：需从剩余SP卡中补足 N_sp-1 个，从剩余所有卡中补足 N_rest 个。
       - 如果极限位是非SP卡：需从剩余SP卡中补足 N_sp 个，从剩余所有卡中补足 N_rest-1 个。
       - 关键：SP名额必须由SP卡填充，不可混用。
    """
    score_attr = 'effective_current_score' if mode == 'current' else 'effective_max_score'
    limit_attr = 'effective_limit_score'

    best_score = -1.0
    best_team = []

    # 收集所有可能的极限卡候选 (每个颜色的 SP前N+2 和 非SP前N+2)
    all_limit_candidates = []
    for color, data in grouped_cards.items():
        # 只考虑那些有可能被选入队伍的卡作为极限候选
        # 简单起见，取前 total+2 张
        candidates = data['sp_limit'][:data['total'] + 2] + data['non_sp_limit'][:data['total'] + 2]
        for c in candidates:
            all_limit_candidates.append((color, c))

    # 遍历每一个极限卡
    for limit_color, limit_card in all_limit_candidates:
        data = grouped_cards[limit_color]
        constraint = config.color_constraints[limit_color]

        # 确定当前极限卡占用的名额类型
        is_limit_sp = limit_card.is_sp

        # 计算剩余需要填充的名额
        req_sp = constraint.sp_min
        req_rest = constraint.rest  # 剩余普通名额

        if is_limit_sp:
            # 极限卡是SP，占用了一个SP名额
            rem_sp_needed = req_sp - 1
            rem_rest_needed = req_rest
        else:
            # 极限卡是非SP，占用了一个普通名额
            rem_sp_needed = req_sp
            rem_rest_needed = req_rest - 1

        if rem_sp_needed < 0 or rem_rest_needed < 0:
            # 理论上不会发生，除非约束逻辑有误
            continue

        current_team = [limit_card]
        current_team_names = {limit_card.name}
        valid_selection = True

        # --- 填充逻辑：分颜色独立处理 ---
        # 我们需要为每种颜色填充其对应的名额
        # 对于 limit_color: 需要填充 rem_sp_needed (SP) + rem_rest_needed (Rest)
        # 对于其他颜色: 需要填充 data['sp_min'] (SP) + data['rest'] (Rest)
        for c_color in config.color_order:
            c_data = grouped_cards[c_color]
            constraint = config.color_constraints[c_color]

            if c_color == limit_color:
                target_sp = rem_sp_needed
                target_rest = rem_rest_needed
                # 排除已选的极限卡
                exclude_names = current_team_names
            else:
                target_sp = constraint.sp_min
                target_rest = constraint.rest
                # 全局去重
                exclude_names = current_team_names

            # 1. 强制填充 SP 名额
            # 必须从 sp_mode 中选，且不能重名
            selected_sp_count = 0
            for card in c_data['sp_mode']:
                if selected_sp_count >= target_sp:
                    break
                if card.name in exclude_names:
                    continue
                current_team.append(card)
                exclude_names.add(card.name)
                selected_sp_count += 1

            if selected_sp_count < target_sp:
                valid_selection = False
                break  # SP卡不足，此方案失败

            # 2. 填充 Rest 名额
            # 从 all_mode 中选，排除已选 (包括刚才选的SP卡)
            selected_rest_count = 0
            for card in c_data['all_mode']:
                if selected_rest_count >= target_rest:
                    break
                if card.name in exclude_names:
                    continue
                current_team.append(card)
                exclude_names.add(card.name)
                selected_rest_count += 1

            if selected_rest_count < target_rest:
                valid_selection = False
                break

        if not valid_selection:
            continue

        # 计算总分
        # team结构: [limit, ...rest...]
        # 确保 team 长度正确
        if len(current_team) != 6:
            continue

        s_limit = getattr(limit_card, limit_attr, 0)
        s_mode = sum(getattr(c, score_attr, 0) for c in current_team[1:])
        total = s_limit + s_mode

        if total > best_score:
            best_score = total
            best_team = current_team

    return best_team, best_score

def _run_with_best_rulemode(base_args, resident_upgrade=False):
    """
    对同一组参数分别用 rulemode=0 和 rulemode=1 计算，取总分最高的结果。
    返回: (best_team_current, best_score_current), (best_team_max, best_score_max), best_rulemode
    """
    best_result = None
    best_rulemode = 0
    best_total_score = -1.0

    for rulemode in [0, 1]:
        # 替换 args 中的 rulemode 参数（倒数第1个）
        args_list = list(base_args)
        args_list[-1] = rulemode  # rulemode 是最后一个参数
        args = tuple(args_list)

        try:
            result_now, result_max = run_optimization(*args, resident_upgrade=resident_upgrade)
            team_now, score_now = result_now
            team_max, score_max = result_max
            # 用最大模式总分作为比较基准
            if score_max > best_total_score:
                best_total_score = score_max
                best_result = (result_now, result_max)
                best_rulemode = rulemode
        except Exception:
            continue

    if best_result is None:
        # 如果都失败，用 rulemode=0 再试一次
        args_list = list(base_args)
        args_list[-1] = 0
        result_now, result_max = run_optimization(*tuple(args_list), resident_upgrade=resident_upgrade)
        best_result = (result_now, result_max)
        best_rulemode = 0

    return best_result[0], best_result[1], best_rulemode


def run_batch_analysis():
    """
    批量分析函数：遍历所有组合并生成包含6列卡牌名称的表格
    输出到Excel的3个工作表：
      Sheet 1: 所有组合的详细结果
      Sheet 2: 各模式最佳组合（常规模式、最大模式、常驻升级模式各一个）
      Sheet 3: 常驻升级价值总结（哪些常驻卡升级后提升了得分）
    """

    # 1. 定义所有可能的变量
    colors = ['红', '蓝', '黄']
    types = ['理性', '感性', '非凡']

    # 配队模式定义
    team_configs = [
        ("4/1/1", 4, 2, 1, 0, 1, 0),
        ("3/2/1", 3, 2, 2, 0, 1, 0),
        ("2/3/1", 2, 1, 3, 1, 1, 0),
        ("3/1/2", 3, 2, 1, 0, 2, 0),
        ("4/2/0", 4, 2, 2, 0, 0, 0)
    ]

    results = []
    # 用于Sheet 3：记录每个组合中常驻升级带来的变化
    upgrade_impact_records = []

    num=0
    # 2. 生成所有颜色排列 (6种)
    for color_perm in permutations(colors):
        c1, c2, c3 = color_perm
        print(f'正在进行{c1},{c2},{c3}组合的计算，进度{num}/6')
        num=num+1

        for team_name, t1, sp1, t2, sp2, t3, sp3 in team_configs:
            for target_type in types:

                base_args = (c1, t1, sp1, c2, t2, sp2, c3, t3, sp3, target_type, 1, 0)
                # 0,0表示纯数值，1,0表示计算道具,2,0表示加上等价数值

                try:
                    # 使用新入口函数（自动尝试 rulemode=0 和 1，取最大值）
                    (team_now_cards, score_now), (team_max_cards, score_max), rulemode_used = \
                        _run_with_best_rulemode(base_args, resident_upgrade=False)

                    rulemode_str = str(rulemode_used)

                    # --- 关键修改点：将6张卡拆分为6列 ---
                    # 常规模式行
                    card_names_now = [card.name for card in team_now_cards]
                    # 确保正好有6张
                    while len(card_names_now) < 6:
                        card_names_now.append("空位")

                    results.append({
                        '颜色组合': f"{c1}{c2}{c3}",
                        '配队模式': team_name,
                        '类型': target_type,
                        '计算模式': '常规模式',
                        '规则': rulemode_str,
                        '总分': round(score_now, 2),
                        # --- 6张卡分列6列 ---
                        '卡1': card_names_now[0],
                        '卡2': card_names_now[1],
                        '卡3': card_names_now[2],
                        '卡4': card_names_now[3],
                        '卡5': card_names_now[4],
                        '卡6': card_names_now[5]
                    })

                    # 最大模式行
                    card_names_max = [card.name for card in team_max_cards]
                    while len(card_names_max) < 6:
                        card_names_max.append("空位")

                    results.append({
                        '颜色组合': f"{c1}{c2}{c3}",
                        '配队模式': team_name,
                        '类型': target_type,
                        '计算模式': '最大模式',
                        '规则': rulemode_str,
                        '总分': round(score_max, 2),
                        # --- 6张卡分列6列 ---
                        '卡1': card_names_max[0],
                        '卡2': card_names_max[1],
                        '卡3': card_names_max[2],
                        '卡4': card_names_max[3],
                        '卡5': card_names_max[4],
                        '卡6': card_names_max[5]
                    })

                    # 常驻升级模式行（使用最大模式的计算方式，因为改变的只有最大等级）
                    (_, _), (team_resident_cards, score_resident), resident_rulemode = \
                        _run_with_best_rulemode(base_args, resident_upgrade=True)
                    resident_rulemode_str = str(resident_rulemode)
                    card_names_resident = [card.name for card in team_resident_cards]
                    while len(card_names_resident) < 6:
                        card_names_resident.append("空位")

                    results.append({
                        '颜色组合': f"{c1}{c2}{c3}",
                        '配队模式': team_name,
                        '类型': target_type,
                        '计算模式': '常驻升级模式',
                        '规则': resident_rulemode_str,
                        '总分': round(score_resident, 2),
                        # --- 6张卡分列6列 ---
                        '卡1': card_names_resident[0],
                        '卡2': card_names_resident[1],
                        '卡3': card_names_resident[2],
                        '卡4': card_names_resident[3],
                        '卡5': card_names_resident[4],
                        '卡6': card_names_resident[5]
                    })

                    # --- 分析常驻升级影响 ---
                    # 比较最大模式 vs 常驻升级模式的队伍和得分
                    if score_resident > score_max:
                        # 找出哪些常驻卡在常驻升级模式中提升了得分
                        # 构建最大模式队伍中常驻卡的得分映射
                        max_team_resident_scores = {}
                        for card in team_max_cards:
                            if card.resident_category == '常驻':
                                max_team_resident_scores[card.name] = card.effective_max_score

                        # 构建常驻升级模式队伍中常驻卡的得分映射
                        resident_team_resident_scores = {}
                        for card in team_resident_cards:
                            if card.resident_category == '常驻':
                                resident_team_resident_scores[card.name] = card.effective_max_score

                        # 找出得分提升的常驻卡
                        all_resident_names = set(max_team_resident_scores.keys()) | set(resident_team_resident_scores.keys())
                        for rname in all_resident_names:
                            old_score = max_team_resident_scores.get(rname, 0)
                            new_score = resident_team_resident_scores.get(rname, 0)
                            if new_score > old_score:
                                upgrade_impact_records.append({
                                    '颜色组合': f"{c1}{c2}{c3}",
                                    '配队模式': team_name,
                                    '类型': target_type,
                                    '常驻卡名称': rname,
                                    '升级前得分': round(old_score, 2),
                                    '升级后得分': round(new_score, 2),
                                    '得分提升': round(new_score - old_score, 2),
                                    '队伍总分提升': round(score_resident - score_max, 2)
                                })

                except Exception as e:
                    results.append({
                        '颜色组合': f"{c1}{c2}{c3}",
                        '配队模式': team_name,
                        '类型': target_type,
                        '计算模式': 'Error',
                        '规则': '',
                        '总分': 'N/A',
                        '卡1': str(e), '卡2': '', '卡3': '', '卡4': '', '卡5': '', '卡6': ''
                    })

    # 3. 创建Excel输出
    if results:
        df_all = pd.DataFrame(results)

        # 重新排列列顺序，让卡牌列紧挨着分数
        column_order = ['颜色组合', '配队模式', '类型', '计算模式', '规则', '总分',
                        '卡1', '卡2', '卡3', '卡4', '卡5', '卡6']
        # 检查是否所有列都存在（防止报错）
        column_order = [col for col in column_order if col in df_all.columns]
        df_all = df_all[column_order]

        # --- Sheet 2: 各模式最佳组合 ---
        # 对每个(颜色组合, 类型)组合，找出各模式总分最高的配队模式
        best_combinations = []
        # 获取所有唯一的(颜色组合, 类型)组合
        unique_groups = df_all[['颜色组合', '类型']].drop_duplicates()
        for _, group_row in unique_groups.iterrows():
            color_combo = group_row['颜色组合']
            target_type = group_row['类型']
            # 筛选该组合的数据
            group_df = df_all[(df_all['颜色组合'] == color_combo) & (df_all['类型'] == target_type)]
            for mode_name in ['常规模式', '最大模式', '常驻升级模式']:
                mode_df = group_df[group_df['计算模式'] == mode_name]
                if not mode_df.empty:
                    # 找到该模式下总分最高的行
                    best_row = mode_df.loc[mode_df['总分'].idxmax()]
                    best_combinations.append(best_row.to_dict())

        df_best = pd.DataFrame(best_combinations) if best_combinations else pd.DataFrame()

        # --- Sheet 3: 常驻升级价值总结 ---
        df_upgrade = pd.DataFrame(upgrade_impact_records) if upgrade_impact_records else pd.DataFrame()

        # 如果升级记录不为空，添加汇总：统计每张常驻卡在所有组合中出现的次数和总提升
        if not df_upgrade.empty:
            summary_rows = []
            card_summary = df_upgrade.groupby('常驻卡名称').agg(
                出现次数=('得分提升', 'count'),
                总得分提升=('得分提升', 'sum'),
                平均得分提升=('得分提升', 'mean')
            ).reset_index()
            # 按总得分提升降序排列
            card_summary = card_summary.sort_values('总得分提升', ascending=False)
            summary_rows = card_summary.to_dict('records')
            df_summary = pd.DataFrame(summary_rows)
        else:
            df_summary = pd.DataFrame()

        # 导出Excel（多工作表）
        output_path = "基础选卡.xlsx"
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            df_all.to_excel(writer, sheet_name='详细结果', index=False)
            if not df_best.empty:
                df_best.to_excel(writer, sheet_name='最佳组合', index=False)
            if not df_upgrade.empty:
                df_upgrade.to_excel(writer, sheet_name='升级影响明细', index=False)
            if not df_summary.empty:
                df_summary.to_excel(writer, sheet_name='升级价值总结', index=False)

        print(f"\n已导出到 Excel: {output_path}")
        print(f"  Sheet 1 - 详细结果: {len(df_all)} 行")
        print(f"  Sheet 2 - 最佳组合: {len(df_best)} 行")
        print(f"  Sheet 3 - 升级影响明细: {len(df_upgrade)} 行")
        if not df_summary.empty:
            print(f"  Sheet 4 - 升级价值总结: {len(df_summary)} 行")

        return df_all
    else:
        print("未生成任何结果。")
        return None


# --- 主程序入口 ---
if __name__ == "__main__":
    print("开始批量优化计算...")
    result_df = run_batch_analysis()

    if result_df is not None:
        print(f"\n计算完成，共生成 {len(result_df)} 行数据。")
    input("按回车键退出程序...")
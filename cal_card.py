import os
from make_card import load_calculation_rules, calculate_card_stats, process_all_cards
import pandas as pd
from dataclasses import dataclass
import copy


# 百分比方案定义：每个方案对应 (第一属性触发次数, 第二属性触发次数, 第三属性触发次数)
# 用于替换"百分比"词条的触发次数
PERCENTAGE_OPTIONS = {
    'A': (10.65, 7.45, 3.95),
    'B': (11.4, 6.75, 3.95),
    'C': (9.5, 8.6, 3.95),
}

# ===== 可调参数：固定加值（一属性/二属性/三属性） =====
ATTR_FIXED_BONUS = {1: 200, 2: 280, 3: 905}

# ===== 可调参数：一属性额外比率（×百分比次数） =====
ATTR1_EXTRA_RATIO = 11.2

# ===== 可调参数：每个属性每百分比次数加值 =====
PCT_PER_TRIGGER_BONUS = 100

# ===== 可调参数：单颜色分数上限 =====
COLOR_SCORE_CAP = 2800

# ===== 可调参数：基础SP率 [一属性, 二属性, 三属性] =====
BASE_SP_RATES = [0.25, 0.25, 0.25]


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
            
        
    def get_base_paths(self):
        """获取基础文件路径"""
        base_dir = os.path.dirname(os.path.abspath(__file__))
        base_dir='./'
        return {
            'card': os.path.join(base_dir, '卡牌数据.xlsx'),
            'term': os.path.join(base_dir, '基础词条.xlsx'),
            'rule': os.path.join(base_dir, '实际计算.xlsx'),
            'character': os.path.join(base_dir, '角色属性.xlsx')
        }


def _optimize_with_cached_data(config, cached_cards, cached_rules):
    """
    使用已缓存的卡牌数据和规则进行优化计算。
    避免重复加载Excel和重复计算卡牌分数。
    
    返回: (team_current, score_current), (team_max, score_max), color_scores_current, color_scores_max
    color_scores 格式: {'红': 分数, '蓝': 分数, '黄': 分数}
    """
    try:
        # 1. 使用已计算好分数的卡牌数据
        cards = cached_cards
        
        # 2. 预处理：过滤类型，计算有效分数，标记属性和SP状态
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

        # 3. 执行优化
        prepared_data = prepare_candidate_pools(filtered_grouped, config)
        if not prepared_data:
            print("错误：未能生成候选池。")
            return

        # 4. 计算最佳队伍（现在返回 team, score, team_card_scores）
        best_team_current, best_score_current, team_scores_current = select_best_team_sp_constrained(
            prepared_data, config, 'current'
        )
        best_team_max, best_score_max, team_scores_max = select_best_team_sp_constrained(
            prepared_data, config, 'max'
        )

        # 5. 计算各颜色分数（使用与总分一致的每张卡实际分数）
        color_scores_current = _calc_color_scores_from_team_scores(best_team_current, team_scores_current)
        color_scores_max = _calc_color_scores_from_team_scores(best_team_max, team_scores_max)

        # 6. 输出结果
        return (best_team_current, best_score_current), (best_team_max, best_score_max), color_scores_current, color_scores_max
    except Exception as e:
        print(f"优化过程中出现错误: {str(e)}")
        raise


def _calc_color_scores_from_team_scores(team, team_scores):
    """
    根据队伍和每张卡在总分中实际使用的分数，计算各颜色得分。
    确保 红分数+蓝分数+黄分数 = 总分。
    
    参数:
        team: 卡牌列表（6张）
        team_scores: 每张卡在总分中实际使用的分数列表（6个值）
    
    返回: {'红': 分数, '蓝': 分数, '黄': 分数}
    """
    color_scores = {'红': 0.0, '蓝': 0.0, '黄': 0.0}
    
    for card, score in zip(team, team_scores):
        c = card.matched_color
        if c in color_scores:
            color_scores[c] += score
    
    return color_scores


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
    
    返回: (best_team, best_score, team_card_scores)
          team_card_scores: 每张卡在总分计算中实际使用的分数列表（与team顺序一致）
    """
    score_attr = 'effective_current_score' if mode == 'current' else 'effective_max_score'
    limit_attr = 'effective_limit_score'

    best_score = -1.0
    best_team = []
    best_team_scores = []

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

        # 记录每张卡在总分中实际使用的分数
        # 极限卡用 limit_attr，其余卡用 score_attr
        team_scores = []
        s_limit = getattr(limit_card, limit_attr, 0)
        team_scores.append(s_limit)
        for c in current_team[1:]:
            team_scores.append(getattr(c, score_attr, 0))
        total = sum(team_scores)

        if total > best_score:
            best_score = total
            best_team = current_team
            best_team_scores = team_scores

    return best_team, best_score, best_team_scores


def _modify_rules_db(base_rules, pct_triggers):
    """
    复制一份规则数据库，只修改"百分比"词条的触发次数。
    其他词条的触发次数保持不变。
    
    参数:
        base_rules: 从Excel加载的原始规则数据库
        pct_triggers: (attr1, attr2, attr3) 三元组，用于替换"百分比"词条
    
    返回: 修改后的规则数据库副本
    """
    modified = copy.deepcopy(base_rules)
    if '百分比' in modified:
        modified['百分比'] = {
            'attr1': float(pct_triggers[0]),
            'attr2': float(pct_triggers[1]),
            'attr3': float(pct_triggers[2]),
        }
    return modified


def _run_all_rulemode_and_pct_cached(base_args, cached_data_pool, resident_upgrade=False):
    """
    使用缓存的卡牌数据，对同一组参数分别用 rulemode=0/1 和 3种百分比方案计算。
    返回所有6种组合的结果（不选最佳），供外部根据上限限制后的总分进行选择。
    
    参数:
        base_args: 基础参数元组 (c1, t1, sp1, c2, t2, sp2, c3, t3, sp3, target_type, mode, rulemode)
        cached_data_pool: 预加载的缓存数据池
        resident_upgrade: 是否启用常驻升级模式
    
    返回: list of dict, 每个dict包含:
        {
            'rulemode': int,
            'pct_key': str,
            'team_now': list, 'score_now': float,
            'team_max': list, 'score_max': float,
            'color_scores_now': dict, 'color_scores_max': dict,
        }
    """
    results = []
    
    for rulemode in [0, 1]:
        args_list = list(base_args)
        args_list[-1] = rulemode
        args = tuple(args_list)

        for pct_key in ['A', 'B', 'C']:
            try:
                cache_key = ('resident' if resident_upgrade else 'normal', rulemode, pct_key)
                if cache_key not in cached_data_pool:
                    continue
                
                cached_cards = cached_data_pool[cache_key]
                config = OptimizationConfig(*args)
                result_now, result_max, color_scores_now, color_scores_max = _optimize_with_cached_data(config, cached_cards, None)

                team_now, score_now = result_now
                team_max, score_max = result_max
                
                results.append({
                    'rulemode': rulemode,
                    'pct_key': pct_key,
                    'team_now': team_now,
                    'score_now': score_now,
                    'team_max': team_max,
                    'score_max': score_max,
                    'color_scores_now': color_scores_now,
                    'color_scores_max': color_scores_max,
                })
            except Exception:
                continue

    return results


def _preload_cached_data(paths):
    """
    预加载所有需要的卡牌数据和规则，避免重复读取Excel。
    
    缓存结构为: {
        ('normal', 0, 'A'): [已计算好分数的卡牌列表],
        ('normal', 0, 'B'): [已计算好分数的卡牌列表],
        ('normal', 0, 'C'): [已计算好分数的卡牌列表],
        ('normal', 1, 'A'): [已计算好分数的卡牌列表],
        ...
        ('resident', 1, 'C'): [已计算好分数的卡牌列表],
    }
    共 2 (resident) × 2 (rulemode) × 3 (pct) = 12 组缓存
    
    注意：只有"百分比"词条的触发次数被替换，其他词条保持Excel中的原始值。
    """
    cached_pool = {}
    
    # 预加载词条数据库（只需要加载一次）
    from make_card import load_keyword_database
    term_db = load_keyword_database(paths['term'])
    
    for resident_upgrade in [False, True]:
        resident_key = 'resident' if resident_upgrade else 'normal'
        
        # 加载卡牌原始数据（只需要加载一次 per resident_upgrade）
        cards_raw = process_all_cards(paths['card'], paths['term'], resident_upgrade)
        
        for rulemode in [0, 1]:
            # 从Excel加载原始规则
            base_rules = load_calculation_rules(paths['rule'], rulemode)
            
            for pct_key, triggers in PERCENTAGE_OPTIONS.items():
                # 复制规则，只修改"百分比"词条的触发次数
                modified_rules = _modify_rules_db(base_rules, triggers)
                
                # 使用修改后的规则计算卡牌分数（mode=1 固定）
                cards_scored = calculate_card_stats(copy.deepcopy(cards_raw), modified_rules, 1)
                
                cache_key = (resident_key, rulemode, pct_key)
                cached_pool[cache_key] = cards_scored


    
    return cached_pool


def load_character_data(path):
    """
    从角色属性.xlsx加载角色数据。
    
    返回: DataFrame，包含列：角色, 状态, 一属性, 二属性, 三属性,
          一属性数值, 二属性数值, 三属性数值, 一属性比率, 二属性比率, 三属性比率
    """
    if not os.path.exists(path):
        print(f"错误：找不到文件 '{path}'")
        return None
    
    try:
        df = pd.read_excel(path, sheet_name='Sheet1')
        required_cols = ['角色', '状态', '一属性', '二属性', '三属性']
        for col in required_cols:
            if col not in df.columns:
                print(f"错误：角色属性表缺少列 '{col}'")
                return None
        return df
    except Exception as e:
        print(f"读取角色属性表时发生错误: {e}")
        return None


def _calc_extra(pct_t1, pct_t2, pct_t3, attr1_val, attr2_val, attr3_val,
                attr1_ratio, attr2_ratio, attr3_ratio, color_to_attr_idx):
    """
    计算给定百分比方案的额外加值。
    使用全局变量 ATTR_FIXED_BONUS, ATTR1_EXTRA_RATIO, PCT_PER_TRIGGER_BONUS。
    
    返回: (extra_by_color, total_extra)
    """
    attr_data = {
        1: {'val': attr1_val, 'ratio': attr1_ratio, 'pct': pct_t1},
        2: {'val': attr2_val, 'ratio': attr2_ratio, 'pct': pct_t2},
        3: {'val': attr3_val, 'ratio': attr3_ratio, 'pct': pct_t3},
    }
    extra_by_color = {'红': 0.0, '蓝': 0.0, '黄': 0.0}
    for color, attr_idx in color_to_attr_idx.items():
        d = attr_data[attr_idx]
        extra = 0.0
        extra += d['val']  # 1. 属性数值
        extra += d['ratio'] * d['pct']  # 2. 比率 × 百分比次数
        extra += ATTR_FIXED_BONUS[attr_idx]  # 3. 固定加值（一/二/三属性）
        if attr_idx == 1:
            extra += ATTR1_EXTRA_RATIO * pct_t1  # 4. 一属性额外比率×百分比次数
        extra += PCT_PER_TRIGGER_BONUS * d['pct']  # 5. 每个属性+100×百分比次数
        extra_by_color[color] = extra
    return extra_by_color, sum(extra_by_color.values())


def _apply_cap(color_scores, extra_by_color):
    """
    应用单颜色分数上限（使用全局变量 COLOR_SCORE_CAP）。
    
    返回: (capped_scores, capped_total)
    """
    capped = {}
    for c in ['红', '蓝', '黄']:
        capped[c] = min(color_scores.get(c, 0) + extra_by_color.get(c, 0), COLOR_SCORE_CAP)
    return capped, sum(capped.values())


def _calc_sp_rate(team, mode, c1_color, c2_color):
    """
    计算一属性和二属性的SP率。
    使用全局变量 BASE_SP_RATES。
    team: 6张卡牌列表（第1张是极限位）
    mode: 'current' 或 'max'
    c1_color, c2_color: 一属性颜色, 二属性颜色
    返回: (一属性SP率, 二属性SP率)
    """
    sp_rates = {c1_color: 0.0, c2_color: 0.0}
    
    for i, card in enumerate(team):
        if card.kw2_name != 'SP率':
            continue
        if card.kw2_stats is None or card.kw2_stats[0] is None:
            continue
        
        color = card.matched_color
        if color not in sp_rates:
            continue
        
        if i == 0:  # 极限位卡使用极限等级值
            sp_val = card.kw2_stats[2]
        elif mode == 'current':
            sp_val = card.kw2_stats[0]  # 当前等级值
        else:
            sp_val = card.kw2_stats[1]  # 最高等级值
        
        sp_rates[color] += sp_val
    
    # 总SP率 = 基础SP率 + 额外SP率%
    rate1 = BASE_SP_RATES[0] + sp_rates[c1_color] / 100.0
    rate2 = BASE_SP_RATES[1] + sp_rates[c2_color] / 100.0
    
    return rate1, rate2


def _calc_success_rate(pct_key, rate1, rate2):
    """
    根据百分比方案计算成功率。
    A: 一属性^4 × 二属性^1
    B: 一属性^4 × 二属性^1
    C: 一属性^3 × 二属性^2
    """
    if pct_key in ('A', 'B'):
        return (rate1 ** 4) * rate2
    else:  # 'C'
        return (rate1 ** 3) * (rate2 ** 2)


def run_batch_analysis():
    """
    批量分析函数：遍历所有角色+状态组合并生成包含6列卡牌名称的表格
    输出到Excel的3个工作表：
      Sheet 1: 所有组合的详细结果
      Sheet 2: 各模式最佳组合（常规模式、最大模式、常驻升级模式各一个）
      Sheet 3: 常驻升级价值总结（哪些常驻卡升级后提升了得分）
    """
    # 1. 定义所有可能的变量
    types = ['理性', '感性', '非凡']

    # 配队模式定义
    team_configs = [
        ("4/1/1", 4, 2, 1, 0, 1, 0),
        ("3/2/1", 3, 2, 2, 0, 1, 0),
        ("2/3/1", 2, 1, 3, 1, 1, 0),
        ("3/1/2", 3, 2, 1, 0, 2, 0),
        ("4/2/0", 4, 2, 2, 1, 0, 0),
        ("4/1/1", 4, 2, 1, 1, 1, 0),
        ("3/2/1", 3, 2, 2, 1, 1, 0),
        ("2/3/1", 2, 2, 3, 1, 1, 0),
        ("3/1/2", 3, 2, 1, 1, 2, 0),
        ("4/2/0", 4, 2, 2, 1, 0, 0)
    ]

    results = []

    # 用于Sheet 3：记录每个组合中常驻升级带来的变化
    upgrade_impact_records = []

    # --- 优化：预加载所有缓存数据 ---
    print("正在预加载卡牌数据和计算规则...")
    paths = OptimizationConfig('红', 1, 0, '蓝', 1, 0, '黄', 1, 0, '理性', 1, 0).get_base_paths()
    cached_data_pool = _preload_cached_data(paths)
    print("预加载完成，开始批量计算...")

    # 加载角色属性数据
    char_df = load_character_data(paths['character'])
    if char_df is None or char_df.empty:
        print("错误：未能加载角色属性数据。")
        return None

    print(f"已加载 {len(char_df)} 个角色状态组合。")

    # 2. 遍历所有角色+状态组合
    for _, char_row in char_df.iterrows():
        character = char_row['角色']
        state = char_row['状态']
        c1 = char_row['一属性']
        c2 = char_row['二属性']
        c3 = char_row['三属性']
        
        # 读取角色属性数值和比率
        attr1_val = float(char_row.get('一属性数值', 0))
        attr2_val = float(char_row.get('二属性数值', 0))
        attr3_val = float(char_row.get('三属性数值', 0))
        attr1_ratio = float(char_row.get('一属性比率', 0))
        attr2_ratio = float(char_row.get('二属性比率', 0))
        attr3_ratio = float(char_row.get('三属性比率', 0))

        for team_name, t1, sp1, t2, sp2, t3, sp3 in team_configs:
            for target_type in types:

                base_args = (c1, t1, sp1, c2, t2, sp2, c3, t3, sp3, target_type, 1, 0)
                # 0,0表示纯数值，1,0表示计算道具,2,0表示加上等价数值

                try:
                    # 颜色 -> 属性索引映射（用于额外加值计算）
                    color_to_attr_idx = {c1: 1, c2: 2, c3: 3}

                    # --- 常规模式/最大模式：遍历所有6种组合，用上限限制后的总分选最佳 ---
                    all_results_normal = _run_all_rulemode_and_pct_cached(base_args, cached_data_pool, resident_upgrade=False)
                    
                    best_normal = None
                    best_normal_capped_total = -1.0
                    for r in all_results_normal:
                        pct_triggers = PERCENTAGE_OPTIONS[r['pct_key']]
                        extra_by_color, _ = _calc_extra(*pct_triggers, attr1_val, attr2_val, attr3_val,
                                                        attr1_ratio, attr2_ratio, attr3_ratio, color_to_attr_idx)
                        # 用最大模式分数选最佳
                        capped_scores, capped_total = _apply_cap(r['color_scores_max'], extra_by_color)
                        if capped_total > best_normal_capped_total:
                            best_normal_capped_total = capped_total
                            best_normal = r
                            best_normal_extra = extra_by_color

                    if best_normal is None:
                        raise Exception("所有组合均失败")

                    rulemode_used = best_normal['rulemode']
                    pct_key = best_normal['pct_key']
                    rulemode_str = str(rulemode_used)
                    # 对常规模式和最大模式分别应用上限
                    capped_now_scores, capped_now_total = _apply_cap(best_normal['color_scores_now'], best_normal_extra)
                    capped_max_scores, capped_max_total = _apply_cap(best_normal['color_scores_max'], best_normal_extra)

                    # --- 计算成功率 ---
                    # 常规模式：使用当前等级值
                    rate1_now, rate2_now = _calc_sp_rate(best_normal['team_now'], 'current', c1, c2)
                    success_rate_now = _calc_success_rate(pct_key, rate1_now, rate2_now)
                    # 最大模式：使用最高等级值
                    rate1_max, rate2_max = _calc_sp_rate(best_normal['team_max'], 'max', c1, c2)
                    success_rate_max = _calc_success_rate(pct_key, rate1_max, rate2_max)

                    # --- 将6张卡拆分为6列 ---
                    # 常规模式行
                    card_names_now = [card.name for card in best_normal['team_now']]
                    while len(card_names_now) < 6:
                        card_names_now.append("空位")

                    results.append({
                        '角色': character,
                        '状态': state,
                        '配队模式': team_name,
                        '类型': target_type,
                        '计算模式': '常规模式',
                        '规则': rulemode_str,
                        '百分比方案': pct_key,
                        '总分': round(capped_now_total, 2),
                        '红分数': round(capped_now_scores['红'], 2),
                        '蓝分数': round(capped_now_scores['蓝'], 2),
                        '黄分数': round(capped_now_scores['黄'], 2),
                        '成功率': f"{success_rate_now * 100:.2f}%",
                        '卡1': card_names_now[0],
                        '卡2': card_names_now[1],
                        '卡3': card_names_now[2],
                        '卡4': card_names_now[3],
                        '卡5': card_names_now[4],
                        '卡6': card_names_now[5]
                    })

                    # 最大模式行
                    card_names_max = [card.name for card in best_normal['team_max']]
                    while len(card_names_max) < 6:
                        card_names_max.append("空位")

                    results.append({
                        '角色': character,
                        '状态': state,
                        '配队模式': team_name,
                        '类型': target_type,
                        '计算模式': '最大模式',
                        '规则': rulemode_str,
                        '百分比方案': pct_key,
                        '总分': round(capped_max_total, 2),
                        '红分数': round(capped_max_scores['红'], 2),
                        '蓝分数': round(capped_max_scores['蓝'], 2),
                        '黄分数': round(capped_max_scores['黄'], 2),
                        '成功率': f"{success_rate_max * 100:.2f}%",
                        '卡1': card_names_max[0],
                        '卡2': card_names_max[1],
                        '卡3': card_names_max[2],
                        '卡4': card_names_max[3],
                        '卡5': card_names_max[4],
                        '卡6': card_names_max[5]
                    })


                    # --- 常驻升级模式：同样遍历所有6种组合，用上限限制后的总分选最佳 ---
                    all_results_resident = _run_all_rulemode_and_pct_cached(base_args, cached_data_pool, resident_upgrade=True)
                    
                    best_resident = None
                    best_resident_capped_total = -1.0
                    for r in all_results_resident:
                        pct_triggers = PERCENTAGE_OPTIONS[r['pct_key']]
                        extra_by_color, _ = _calc_extra(*pct_triggers, attr1_val, attr2_val, attr3_val,
                                                        attr1_ratio, attr2_ratio, attr3_ratio, color_to_attr_idx)
                        capped_scores, capped_total = _apply_cap(r['color_scores_max'], extra_by_color)

                        if capped_total > best_resident_capped_total:
                            best_resident_capped_total = capped_total
                            best_resident = r
                            best_resident_extra = extra_by_color
                    
                    if best_resident is None:
                        raise Exception("常驻升级所有组合均失败")
                    
                    resident_rulemode = best_resident['rulemode']
                    resident_pct_key = best_resident['pct_key']
                    resident_rulemode_str = str(resident_rulemode)
                    
                    capped_resident_scores, capped_resident_total = _apply_cap(best_resident['color_scores_max'], best_resident_extra)

                    # 常驻升级模式：使用最高等级值（与最大模式相同）
                    rate1_resident, rate2_resident = _calc_sp_rate(best_resident['team_max'], 'max', c1, c2)
                    success_rate_resident = _calc_success_rate(resident_pct_key, rate1_resident, rate2_resident)

                    card_names_resident = [card.name for card in best_resident['team_max']]
                    while len(card_names_resident) < 6:
                        card_names_resident.append("空位")

                    results.append({
                        '角色': character,
                        '状态': state,
                        '配队模式': team_name,
                        '类型': target_type,
                        '计算模式': '常驻升级模式',
                        '规则': resident_rulemode_str,
                        '百分比方案': resident_pct_key,
                        '总分': round(capped_resident_total, 2),
                        '红分数': round(capped_resident_scores['红'], 2),
                        '蓝分数': round(capped_resident_scores['蓝'], 2),
                        '黄分数': round(capped_resident_scores['黄'], 2),
                        '成功率': f"{success_rate_resident * 100:.2f}%",
                        '卡1': card_names_resident[0],
                        '卡2': card_names_resident[1],
                        '卡3': card_names_resident[2],
                        '卡4': card_names_resident[3],
                        '卡5': card_names_resident[4],
                        '卡6': card_names_resident[5]
                    })


                    # --- 分析常驻升级影响 ---
                    # 比较最大模式 vs 常驻升级模式的队伍和得分（使用上限限制后的分数）
                    # 重要：必须使用相同的 rulemode/pct_key 进行比较，否则分数差异可能来自规则不同而非升级
                    if capped_resident_total > capped_max_total:
                        # 找出哪些常驻卡在常驻升级模式中提升了得分
                        max_team_resident_scores = {}
                        for card in best_normal['team_max']:
                            if card.resident_category == '常驻':
                                max_team_resident_scores[card.name] = card.effective_max_score

                        resident_team_resident_scores = {}
                        for card in best_resident['team_max']:
                            if card.resident_category == '常驻':
                                resident_team_resident_scores[card.name] = card.effective_max_score

                        # 只比较同时出现在两个队伍中的常驻卡
                        # 如果一张卡只出现在其中一个队伍，说明是队伍构成变化而非升级影响
                        common_resident_names = set(max_team_resident_scores.keys()) & set(resident_team_resident_scores.keys())
                        for rname in common_resident_names:
                            old_score = max_team_resident_scores[rname]
                            new_score = resident_team_resident_scores[rname]
                            
                            # 关键修复：如果 normal 和 resident 使用了不同的 rulemode/pct_key，
                            # 则卡牌分数差异可能来自规则不同而非升级。
                            # 此时需要检查卡牌的基础词条属性是否真的发生了变化。
                            # 如果 max_level 相同且 kw_stats 相同，则分数差异来自规则不同，不应视为升级。
                            normal_card = None
                            resident_card = None
                            for c in best_normal['team_max']:
                                if c.name == rname:
                                    normal_card = c
                                    break
                            for c in best_resident['team_max']:
                                if c.name == rname:
                                    resident_card = c
                                    break
                            
                            # 检查卡牌是否真的被升级了（max_level 或 kw_stats 发生变化）
                            really_upgraded = False
                            if normal_card and resident_card:
                                # 检查 max_level 是否不同
                                if normal_card.max_level != resident_card.max_level:
                                    really_upgraded = True
                                else:
                                    # 检查 kw_stats 是否不同
                                    for kw_attr in ['kw1_stats', 'kw2_stats', 'kw4_stats', 'kw5_stats', 'kw6_stats']:
                                        old_stats = getattr(normal_card, kw_attr, None)
                                        new_stats = getattr(resident_card, kw_attr, None)
                                        if old_stats != new_stats:
                                            really_upgraded = True
                                            break
                            
                            if really_upgraded and new_score > old_score:
                                upgrade_impact_records.append({
                                    '角色': character,
                                    '状态': state,
                                    '配队模式': team_name,
                                    '类型': target_type,
                                    '常驻卡名称': rname,
                                    '升级前得分': round(old_score, 2),
                                    '升级后得分': round(new_score, 2),
                                    '得分提升': round(new_score - old_score, 2),
                                    '队伍总分提升': round(capped_resident_total - capped_max_total, 2)
                                })


                except Exception as e:
                    results.append({
                        '角色': character,
                        '状态': state,
                        '配队模式': team_name,
                        '类型': target_type,
                        '计算模式': 'Error',
                        '规则': '',
                        '百分比方案': '',
                        '总分': 'N/A',
                        '红分数': '', '蓝分数': '', '黄分数': '',
                        '卡1': str(e), '卡2': '', '卡3': '', '卡4': '', '卡5': '', '卡6': ''
                    })

    # 3. 创建Excel输出
    if results:
        df_all = pd.DataFrame(results)

        # 重新排列列顺序
        column_order = ['角色', '状态', '配队模式', '类型', '计算模式', '规则', '百分比方案', '总分',
                        '红分数', '蓝分数', '黄分数', '成功率',
                        '卡1', '卡2', '卡3', '卡4', '卡5', '卡6']

        # 检查是否所有列都存在（防止报错）
        column_order = [col for col in column_order if col in df_all.columns]
        df_all = df_all[column_order]

        # --- Sheet 2: 各模式最佳组合 ---
        # 对每个(角色, 状态, 类型)组合，找出各模式总分最高的配队模式
        best_combinations = []
        # 获取所有唯一的(角色, 状态, 类型)组合
        unique_groups = df_all[['角色', '状态', '类型']].drop_duplicates()
        for _, group_row in unique_groups.iterrows():
            character = group_row['角色']
            state = group_row['状态']
            target_type = group_row['类型']
            # 筛选该组合的数据
            group_df = df_all[(df_all['角色'] == character) & (df_all['状态'] == state) & (df_all['类型'] == target_type)]
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
        # 所有得分均使用队伍总分提升计算
        if not df_upgrade.empty:
            summary_rows = []
            card_summary = df_upgrade.groupby('常驻卡名称').agg(
                出现次数=('队伍总分提升', 'count'),
                总得分提升=('队伍总分提升', 'sum'),
                平均得分提升=('队伍总分提升', 'mean')
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

    I = input("按下回车键后退出")

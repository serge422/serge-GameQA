# -*- coding: utf-8 -*-
"""
死亡细胞武器数据对比脚本
对比 v3.3 与 v3.4 两个版本的 Excel 武器表，输出差异并写入 diff_weapon.xlsx
"""

import os
import sys

import pandas as pd

# ---------------------------------------------------------------------------
# 配置常量
# ---------------------------------------------------------------------------
FILE_V33 = "武器数据_v3.3.xlsx"
FILE_V34 = "武器数据_v3.4.xlsx"
SHEET_NAME = "武器属性"
KEY_COLUMN = "武器ID"
OUTPUT_FILE = "diff_weapon.xlsx"

# 数据文件所在目录
# - None：使用本脚本所在文件夹
# - 绝对路径：例如 r"d:\桌面\test file\DeathCells_WeaponDiff"
WORK_DIR = None

# 除主键外需要对比的列（顺序固定，输出 Excel 也按此顺序）
COMPARE_COLUMNS = [
    "武器名称",
    "类型",
    "攻击力",
    "暴击率",
    "冷却",
    "DPS",
    "版本",
]

# 输出到「新增武器」「删除武器」工作表时的完整列顺序
OUTPUT_COLUMNS = [KEY_COLUMN] + COMPARE_COLUMNS

# Excel 中可能出现的列名别名 -> 统一为标准列名
COLUMN_ALIASES = {
    "暴击率": ["暴击率", "暴击率(%)", "暴击率（%）", "暴击率%"],
    "冷却": ["冷却", "冷却(秒)", "冷却（秒）", "冷却秒"],
}

# 「数值变化」工作表中不对比的列（版本会随文件必然变化，避免刷屏）
VALUE_DIFF_SKIP_COLUMNS = ["版本"]


def get_work_dir() -> str:
    """返回用于读写 Excel 的目录（默认脚本所在目录）。"""
    if WORK_DIR:
        return os.path.abspath(WORK_DIR)
    return os.path.dirname(os.path.abspath(__file__))


def _build_alias_lookup() -> dict:
    """构建「原始列名 -> 标准列名」映射表。"""
    lookup = {}
    for standard_name, aliases in COLUMN_ALIASES.items():
        for alias in aliases:
            lookup[alias.strip()] = standard_name
    return lookup


ALIAS_LOOKUP = _build_alias_lookup()


def normalize_columns(df: pd.DataFrame, filepath: str) -> pd.DataFrame:
    """
    将 DataFrame 列名规范为标准名称。
    例如：暴击率(%) -> 暴击率，冷却(秒) -> 冷却
    """
    rename_map = {}
    for col in df.columns:
        col_str = str(col).strip()
        if col_str in ALIAS_LOOKUP:
            rename_map[col] = ALIAS_LOOKUP[col_str]
    if rename_map:
        df = df.rename(columns=rename_map)

    # 若重命名后出现重复列（极少见），只保留第一列
    if df.columns.duplicated().any():
        dup_cols = df.columns[df.columns.duplicated()].unique().tolist()
        print(
            f"警告：文件「{os.path.basename(filepath)}」存在重复列名 {dup_cols}，"
            f"已保留每列的第一份数据。",
            file=sys.stderr,
        )
        df = df.loc[:, ~df.columns.duplicated(keep="first")]

    return df


def validate_required_columns(df: pd.DataFrame, filepath: str) -> None:
    """检查是否包含主键及所有需要对比的列。"""
    missing_key = KEY_COLUMN not in df.columns
    missing_compare = [c for c in COMPARE_COLUMNS if c not in df.columns]

    if missing_key or missing_compare:
        msg_parts = []
        if missing_key:
            msg_parts.append(f"缺少主键列「{KEY_COLUMN}」")
        if missing_compare:
            msg_parts.append(f"缺少对比列：{missing_compare}")
        raise ValueError(
            f"文件「{os.path.basename(filepath)}」的工作表「{SHEET_NAME}」"
            f"{ '；'.join(msg_parts) }。"
            f"当前列：{list(df.columns)}。"
            f"若列名类似「暴击率(%)」「冷却(秒)」，脚本会自动映射为标准名称。"
        )


def load_weapon_sheet(filepath: str) -> pd.DataFrame:
    """
    读取「武器属性」工作表，并规范列名、校验必填列。

    :raises FileNotFoundError: 文件不存在
    :raises ValueError: 工作表或必填列不存在
    """
    if not os.path.isfile(filepath):
        raise FileNotFoundError(f"找不到文件：{filepath}")

    # 先检查工作表是否存在，便于给出中文提示
    try:
        xls = pd.ExcelFile(filepath, engine="openpyxl")
    except Exception as e:
        raise ValueError(f"无法打开文件「{os.path.basename(filepath)}」：{e}") from e

    if SHEET_NAME not in xls.sheet_names:
        raise ValueError(
            f"文件「{os.path.basename(filepath)}」中不存在工作表「{SHEET_NAME}」。"
            f"当前工作表列表：{xls.sheet_names}"
        )

    df = pd.read_excel(xls, sheet_name=SHEET_NAME)
    df = normalize_columns(df, filepath)
    validate_required_columns(df, filepath)

    # 去除主键为空的行
    df = df.dropna(subset=[KEY_COLUMN])
    df[KEY_COLUMN] = df[KEY_COLUMN].astype(str).str.strip()

    # 主键重复时保留第一条
    dup = df[df[KEY_COLUMN].duplicated(keep=False)]
    if not dup.empty:
        dup_ids = dup[KEY_COLUMN].unique().tolist()
        print(
            f"警告：文件「{os.path.basename(filepath)}」存在重复武器ID：{dup_ids}，"
            f"已保留每个 ID 的第一条记录。",
            file=sys.stderr,
        )
        df = df.drop_duplicates(subset=[KEY_COLUMN], keep="first")

    # 只保留输出需要的列，保证新增/删除表结构一致
    df = df[OUTPUT_COLUMNS].copy()
    return df.reset_index(drop=True)


def is_null_like(val) -> bool:
    """判断值是否为空。"""
    if val is None:
        return True
    if isinstance(val, float) and pd.isna(val):
        return True
    if isinstance(val, str) and val.strip() == "":
        return True
    return False


def to_compare_value(val):
    """将单元格值转为便于比较的形式（兼容 numpy 数值类型）。"""
    if is_null_like(val):
        return None
    if isinstance(val, str):
        return val.strip()
    try:
        if pd.isna(val):
            return None
    except (TypeError, ValueError):
        pass
    # Excel 读取的整数常为 numpy.int64，需统一转为 float
    if isinstance(val, (int, float)) and not isinstance(val, bool):
        return float(val)
    try:
        return float(val)
    except (TypeError, ValueError):
        return str(val).strip()


def values_equal(old_val, new_val) -> bool:
    """判断两值是否相等（数值允许微小浮点误差）。"""
    o = to_compare_value(old_val)
    n = to_compare_value(new_val)
    if o is None and n is None:
        return True
    if o is None or n is None:
        return False
    if isinstance(o, float) and isinstance(n, float):
        return abs(o - n) < 1e-9
    return o == n


def format_display_value(val):
    """输出到 Excel 时：整数数值不显示小数点。"""
    if is_null_like(val):
        return ""
    try:
        if pd.isna(val):
            return ""
    except (TypeError, ValueError):
        pass
    try:
        num = float(val)
        if num == int(num):
            return int(num)
        return round(num, 6)
    except (TypeError, ValueError):
        return val


def compute_delta(old_val, new_val, field_name: str):
    """
    计算变化量：
    - 数值列：新值 - 旧值
    - 文本列：留空（变化体现在旧值/新值列）
    """
    o = to_compare_value(old_val)
    n = to_compare_value(new_val)

    if o is None and n is None:
        return ""
    if o is None:
        return format_display_value(new_val)
    if n is None:
        v = format_display_value(old_val)
        return f"-{v}" if v != "" else ""

    if isinstance(o, float) and isinstance(n, float):
        delta = n - o
        if abs(delta) < 1e-9:
            return 0
        if delta == int(delta):
            return int(delta)
        return round(delta, 6)

    # 文本字段（武器名称、类型、版本等）
    return ""


def sort_weapon_ids(ids) -> list:
    """武器 ID 排序：先按数字部分，再按字符串。"""

    def sort_key(wid):
        s = str(wid)
        digits = "".join(ch for ch in s if ch.isdigit())
        return (int(digits) if digits else 0, s)

    return sorted(ids, key=sort_key)


def find_added_weapons(df_old: pd.DataFrame, df_new: pd.DataFrame) -> pd.DataFrame:
    """v3.4 有、v3.3 没有的新增武器。"""
    old_ids = set(df_old[KEY_COLUMN])
    return df_new[~df_new[KEY_COLUMN].isin(old_ids)][OUTPUT_COLUMNS].copy()


def find_removed_weapons(df_old: pd.DataFrame, df_new: pd.DataFrame) -> pd.DataFrame:
    """v3.3 有、v3.4 没有的删除武器。"""
    new_ids = set(df_new[KEY_COLUMN])
    return df_old[~df_old[KEY_COLUMN].isin(new_ids)][OUTPUT_COLUMNS].copy()


def find_value_changes(df_old: pd.DataFrame, df_new: pd.DataFrame) -> pd.DataFrame:
    """
    对比两表共有武器的各属性列，记录字段级变化。
    返回列：武器ID、武器名称、字段名、旧值、新值、变化量
    """
    diff_fields = [c for c in COMPARE_COLUMNS if c not in VALUE_DIFF_SKIP_COLUMNS]
    common_ids = set(df_old[KEY_COLUMN]) & set(df_new[KEY_COLUMN])
    rows = []

    old_idx = df_old.set_index(KEY_COLUMN)
    new_idx = df_new.set_index(KEY_COLUMN)

    for wid in sort_weapon_ids(common_ids):
        old_row = old_idx.loc[wid]
        new_row = new_idx.loc[wid]
        if isinstance(old_row, pd.DataFrame):
            old_row = old_row.iloc[0]
        if isinstance(new_row, pd.DataFrame):
            new_row = new_row.iloc[0]

        weapon_name = new_row.get("武器名称", old_row.get("武器名称", ""))

        for col in diff_fields:
            old_val = old_row[col]
            new_val = new_row[col]
            if values_equal(old_val, new_val):
                continue

            rows.append(
                {
                    KEY_COLUMN: wid,
                    "武器名称": format_display_value(weapon_name),
                    "字段名": col,
                    "旧值": format_display_value(old_val),
                    "新值": format_display_value(new_val),
                    "变化量": compute_delta(old_val, new_val, col),
                }
            )

    change_cols = [KEY_COLUMN, "武器名称", "字段名", "旧值", "新值", "变化量"]
    if not rows:
        return pd.DataFrame(columns=change_cols)
    df_changes = pd.DataFrame(rows)
    # 变化量列保持 object，避免整数与空字符串混用被转成 NaN
    df_changes["变化量"] = df_changes["变化量"].apply(
        lambda x: "" if x == "" or x is None else x
    )
    return df_changes[change_cols]


def save_diff_excel(
    added: pd.DataFrame,
    removed: pd.DataFrame,
    changes: pd.DataFrame,
    output_path: str,
) -> None:
    """写入 diff_weapon.xlsx 的三个工作表。"""
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        added.reindex(columns=OUTPUT_COLUMNS).to_excel(
            writer, sheet_name="新增武器", index=False
        )
        removed.reindex(columns=OUTPUT_COLUMNS).to_excel(
            writer, sheet_name="删除武器", index=False
        )
        change_cols = [KEY_COLUMN, "武器名称", "字段名", "旧值", "新值", "变化量"]
        changes.reindex(columns=change_cols).to_excel(
            writer, sheet_name="数值变化", index=False
        )


if sys.platform == "win32":
    sys.stdout.reconfigure(encoding="utf-8")

def main() -> int:
    work_dir = get_work_dir()
    path_v33 = os.path.join(work_dir, FILE_V33)
    path_v34 = os.path.join(work_dir, FILE_V34)
    path_out = os.path.join(work_dir, OUTPUT_FILE)

    print(f"工作目录：{work_dir}")
    print(f"正在读取：{FILE_V33}、{FILE_V34} …")

    try:
        df_v33 = load_weapon_sheet(path_v33)
        df_v34 = load_weapon_sheet(path_v34)
    except FileNotFoundError as e:
        print(f"错误：{e}", file=sys.stderr)
        print(
            f"请将「{FILE_V33}」和「{FILE_V34}」放到：\n  {work_dir}\n"
            f"或在脚本顶部设置 WORK_DIR = r'你的文件夹路径'。",
            file=sys.stderr,
        )
        return 1
    except ValueError as e:
        print(f"错误：{e}", file=sys.stderr)
        return 1
    except Exception as e:
        print(f"读取 Excel 时发生未预期错误：{e}", file=sys.stderr)
        return 1

    print(f"v3.3 武器数量：{len(df_v33)}，v3.4 武器数量：{len(df_v34)}")
    print(f"对比字段：{COMPARE_COLUMNS}")
    if VALUE_DIFF_SKIP_COLUMNS:
        skipped = [c for c in VALUE_DIFF_SKIP_COLUMNS if c in COMPARE_COLUMNS]
        if skipped:
            print(f"（「数值变化」工作表不对比：{skipped}，避免版本号必然变化造成干扰）")

    added = find_added_weapons(df_v33, df_v34)
    removed = find_removed_weapons(df_v33, df_v34)
    changes = find_value_changes(df_v33, df_v34)

    n_added = len(added)
    n_removed = len(removed)
    n_changes = len(changes)
    total = n_added + n_removed + n_changes

    try:
        save_diff_excel(added, removed, changes, path_out)
    except PermissionError:
        print(
            f"错误：无法写入「{path_out}」，请先关闭已打开的 diff_weapon.xlsx 后重试。",
            file=sys.stderr,
        )
        return 1
    except Exception as e:
        print(f"写入差异文件失败：{e}", file=sys.stderr)
        return 1

    print("-" * 50)
    print(
        f"共发现{total}处差异：新增{n_added}个，删除{n_removed}个，数值变化{n_changes}处"
    )
    print(f"差异结果已保存至：{path_out}")
    print("-" * 50)

    if n_added:
        print(f"\n【新增武器】{n_added} 个：{added[KEY_COLUMN].tolist()}")
    if n_removed:
        print(f"\n【删除武器】{n_removed} 个：{removed[KEY_COLUMN].tolist()}")
    if n_changes:
        changed_weapons = changes[KEY_COLUMN].unique().tolist()
        print(
            f"\n【数值变化】{n_changes} 处（涉及 {len(changed_weapons)} 个武器："
            f"{changed_weapons}）"
        )

    return 0


if __name__ == "__main__":
    sys.exit(main())

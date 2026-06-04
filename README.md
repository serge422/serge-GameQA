# 🎮 游戏QA练习作品集

> 个人游戏测试学习项目，用于展示测试用例设计能力和Python自动化工具开发能力。

## 📂 项目列表

### 1. 原神合成台测试用例
- **文件**：`Genshin_Crafting_TestCases.xlsx`
- **说明**：针对《原神》合成台系统设计的10+条测试用例，覆盖正常合成、材料不足、数量边界、背包满等场景
- **状态**：已在游戏中逐条执行验证

### 2. 原神角色配置表检查工具
- **文件**：`check_character_data.py`
- **功能**：自动检查角色配置表中的异常数据（重复ID、数值越界、空值、非法星级）
- **用法**：将`角色基础属性.xlsx`和脚本放在同一目录，运行 `python config_checker.py`
- **输出**：生成 `check_report.xlsx` 报告文件

### 3. ## 🔍 发现的Bug演示

在《原神》纪行界面中，点击"任务"按钮时，动画出现抽动闪烁现象（正常应有平滑过渡动画）。

<img src="https://github.com/serge422/serge-GameQA/blob/main/click_bug.gif?raw=true" alt="按键动画异常演示" width="600">

- **游戏版本**：原神 6.6.0
- **平台**：PC
- **复现概率**：100%
- **严重程度**：低（视觉异常，不影响功能）

## 🛠 使用技术
- Python（pandas、openpyxl）
- Excel 数据处理
- AI辅助开发（DeepSeek Agent + Cursor）

## 📫 联系方式
- 邮箱：liusenjie2005@qq.com
- 电话：13534034361

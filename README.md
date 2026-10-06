# 每日资金分析器

**一款将每日股票资金流向表整理为多周期分析 Excel 的 Windows 桌面软件。**

导入网站导出的每日资金流向工作簿后，软件在本机维护交易日历史，检查数据完整性，并生成日度及 3/5/7/10 交易日滚动分析结果。

> 本项目仅用于数据整理与研究，不构成投资建议。公开数据接口可能变更、限流或暂时不可用。

## 功能

- 首次可批量导入历史交易日，日常只需导入当天文件；漏日后可补导并重算。
- 检查单日文件的表头、交易日、代码重复和数据行数。
- 生成每日资金变化、滚动区间、资金流入比率、周期递进和数据检查工作表。
- 提供涨停相关字段；生成报表前必须配置并验证 Tushare Token，以补充收盘价、市值和成交量。
- 在本机 SQLite 历史库保存导入数据，Token 使用 Windows DPAPI 加密保存。

## 环境要求

- Windows 10/11，64 位。
- Python 3.10 或更新版本，并包含 Tkinter。
- 网络连接，以及自备的 Tushare Token 和 `daily`、`daily_basic` 等相应接口权限。

## 启动

在仓库根目录打开 PowerShell：

```powershell
py -3 -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe -m DailyFlowAnalyzer
```

也可以双击 `setup.bat` 安装依赖，再双击 `run_app.bat` 启动。

首次使用时，点击“导入每日资金表”，选择至少 10 个交易日的单日 `.xlsx` 文件。之后每天导入当日文件，选择报告交易日。首次生成前，点击右上角“Tushare 设置”填写 Token 并等待验证成功，再点击“生成完整报表”。Token 由软件在本机加密保存。

## 输入文件

输入应为网站导出的“个股资金流向（区间统计）” `.xlsx` 文件。文件应只包含一个交易日，包含软件校验的资金字段，且证券代码不重复。界面会在导入时检查格式和数据完整性。

## 数据与隐私

- 导入原始工作簿、SQLite 历史库、缓存和 Token 存放于当前 Windows 用户的本机应用数据目录。
- 可通过界面右上角“数据目录”查看本机数据位置。
- 软件会向 Tushare 请求行情和市值数据；涨停数据使用公开行情接口。
- 真实输入文件、数据库、Token、导出结果和日志不要提交到公开仓库。
- Token 加密与当前 Windows 用户绑定。不要把个人 Token 写入源码、截图、Issue 或提交记录。

## 目录

```text
DailyFlowAnalyzer/
  __main__.py       python -m 启动入口
  app.py            桌面 GUI
  source_parser.py  输入工作簿校验与解析
  history_store.py  本机 SQLite 历史管理
  analysis.py       资金指标分析
  report_builder.py Excel 报表生成
  providers.py      行情与涨停数据访问
  token_store.py    Windows 本地 Token 加密存储
```

## 免责声明

分析结果取决于输入文件和外部数据源，软件不保证数据的完整性、及时性或适用性。投资决策请独立核验数据并自行承担风险。使用者须遵守数据来源服务的条款及适用法律法规。

## 许可证

本仓库附带 MIT 许可证草案。公开前请将 `LICENSE` 中的 `[作者/组织名称]` 替换为实际版权持有人。

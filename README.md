# DeepSeek API 用量查询工具

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)

一键获取 DeepSeek 开放平台 API 用量数据，自动生成结构化的 Excel 报告（含余额、Token 消耗、缓存命中率、费用明细）。

## 功能

- **自动提取**：打开浏览器 → 登录 DeepSeek 平台 → 提取 API 数据 → 生成 Excel → 关闭浏览器
- **完整数据**：远超页面 UI 显示的内容——包括缓存命中/未命中拆分、输入/输出 Token 拆分、各模型消费金额
- **结构化报告**：输出 `.xlsx` 文件，包含 4 个工作表
- **跨平台**：Windows 双击 `.cmd` 运行，macOS/Linux 运行 `.sh`
- **无痛复用**：`--persistent` 模式保存登录态，下次无需重新登录

## 报告内容

| 工作表 | 内容 |
|--------|------|
| **账户总览** | 充值余额、赠送余额、月消费、Token 总量、可用 Token 估值 |
| **模型用量** | 各模型请求数、缓存命中/未命中 Prompt Tokens、输出 Tokens、总 Tokens、缓存命中率 |
| **消费明细** | 按模型拆分缓存命中/未命中/输出费用（CNY） |
| **按日分布** | 当月逐日各模型用量明细 |

### 数据字段

- `normal_wallets[].balance` — 充值余额 (CNY)
- `bonus_wallets[].balance` — 赠送余额 (CNY)
- `monthly_costs[].amount` — 当月总消费 (CNY)
- `monthly_token_usage` — 当月总 Token 消耗
- `total_available_token_estimation` — 可用 Token 估值
- `PROMPT_CACHE_HIT_TOKEN` / `PROMPT_CACHE_MISS_TOKEN` — 缓存命中/未命中 Token 数
- `RESPONSE_TOKEN` — 输出 Token 数
- `REQUEST` — API 请求次数

## 环境要求

- [Node.js](https://nodejs.org/) + `playwright-cli`（全局安装）
- [Python 3](https://www.python.org/) + `openpyxl` 库
- **Windows**: [Git for Windows](https://git-scm.com/downloads)（提供 Bash 环境）
- **macOS/Linux**: Bash 环境（系统自带）

### 安装依赖

```bash
# 安装 playwright-cli
npm install -g @playwright/cli@latest

# 安装 Python 依赖
pip install openpyxl
```

## 使用

### Windows
下载check_ds.cmd和check_ds.sh放在一个文件夹下。
在Powershell命令行运行：

```cmd
check_ds.cmd
```

首次运行会打开浏览器窗口，请在 DeepSeek 登录页面**手动登录**（使用手机号或密码），然后在命令行按 Enter 继续。脚本会自动提取数据并生成 Excel 报告。

后续运行会自动复用登录状态（`--persistent` 模式），无需重复登录。

### macOS / Linux

```bash
# 交互模式（默认显示浏览器）
bash check_ds.sh

# 无头模式（需已登录过）
bash check_ds.sh --headless
```

### 输出

```
reports/ds-YY-MM-DD-HH-MM.xlsx
```

## 注意事项

1. **首次运行**：需要手动登录。脚本检测到未登录时会暂停等待。
2. **请求号变化**：每次 Session 的 API 请求 ID 可能不同，脚本会自动发现。
3. **认证限制**：不能通过 `fetch()` 直接调用 API——必须用 `playwright-cli response-body` 读取浏览器中已发出的真实请求。
4. **UTC 时区**：DeepSeek 平台数据按 UTC+0 显示。
5. **Python 版本**：Windows 上使用 `python` 命令（非 `python3`），Windows Store 的 `python3` 有沙箱限制。

## 工作原理
![image](https://github.com/Yi-Lings/deepseek-api-usage-checker/blob/master/ds.png)


脚本通过 `playwright-cli` 启动浏览器并导航到 DeepSeek 用量页面，监听网络请求，从 `response-body` 获取原始 JSON 数据，然后用 Python 解析并生成 `openpyxl` 格式的 Excel 文件。

## License

MIT

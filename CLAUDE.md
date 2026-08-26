# Nordic_Flash_Tool 项目说明

## 项目概述

这是一个基于 Python 编写、并用 PyInstaller 打包为 Windows 可执行程序的 **Nordic nRF54L15 自动化烧录工具**，用于产线批量烧录/配置 `SiNE_CARD_II` 系列产品固件，并记录烧录 token/日志。

## 目录结构

```
Nordic_Flash_Tool/
├── 54L15自动化烧录工具.txt          # 说明/占位文件（空）
├── build_exe.bat                    # 打包为 exe 的构建脚本
├── nordic_flash_tool.py             # 主程序源码（核心逻辑，约 66KB）
├── NordicFlashTool.spec             # PyInstaller 打包配置
├── requirements.txt                 # 依赖列表（当前为空）
├── merged.hex                       # 合并后的固件 hex 文件
│
├── configuration/                   # 烧录/签名相关配置
│   ├── ncs_3.0.2.bat                # nRF Connect SDK 环境配置脚本
│   └── nrf54l15dk_nrf54l15_cpuapp/
│       └── boot_signature_key_file_ed25519.pem   # ⚠️ 固件签名私钥，高度敏感
│
├── hex code/                        # 各版本量产固件 hex 源文件
│   └── SiNE_CARD_II_Production_*_V1.1.x.hex
│
├── token/                           # 烧录 token 记录（CSV + Excel 汇总）
│   ├── *.csv                        # 按批次导出的 token 明细
│   └── tokens_sinecardii*.xlsx      # token 汇总表
│
├── build/NordicFlashTool/           # PyInstaller 中间构建产物（可重新生成）
│
└── dist/                            # 打包输出与运行现场数据
    ├── NordicFlashTool.exe          # 最终可执行程序
    ├── flash_state.json             # 烧录状态记录
    ├── logs/                        # 每次烧录的日志文件（flash_log_*.txt）
    └── output/                      # 按日期归档的已烧录固件（provisioned_*.hex）
```

**要点：**
- `nordic_flash_tool.py` + `NordicFlashTool.spec` + `build_exe.bat` 是唯一的"源代码/构建配置"，`build/`、`dist/` 都是可以从源码重新生成的产物，不属于需要人工维护的内容。
- `configuration/.../boot_signature_key_file_ed25519.pem` 是固件签名私钥，`token/` 和 `dist/output/` 中含有产线批次与序列号数据，均为敏感生产数据，不应随意外传或提交到公共/共享仓库。

---

## ⚠️ 协作与代码修改规则（必须遵守）

- 未经用户明确同意，禁止修改任何代码。
- 未经用户明确同意，禁止提交（commit）代码改动。
- 未经用户明确同意，禁止合并（merge）到 `main` 分支。
- 未经用户明确同意，禁止推送（push）到 GitHub 仓库。
- 每次改动之后，需要向用户报告：
  - 改动了哪些文件（files changed）
  - 具体改了什么内容（what changed）
  - 用户应该如何测试（how the user should test it）
  - 是否运行过任何测试或构建命令（whether any tests or build commands were run）

### 英文原文（供工具/AI 直接读取）

```markdown
- Do not commit changes unless the user explicitly asks.
- Do not merge into `main` unless the user explicitly asks.
- Do not push to GitHub unless the user explicitly asks.
- After changes, report:
  - files changed
  - what changed
  - how the user should test it
  - whether any tests or build commands were run
```

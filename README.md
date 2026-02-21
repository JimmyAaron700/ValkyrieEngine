# ValkyrieEngine

**[声明] 由于企业 ERP 系统涉及敏感业务数据，且系统网络入口与底层 DOM 结构受内网安全隔离限制，非本公司内部员工无法直接运行此项目。本项目源码及相关文档仅供 RPA 自动化架构设计与技术交流学习参考。**

**[Disclaimer] Due to the sensitivity of enterprise ERP business data, and the fact that the system access URL and internal DOM structures are strictly isolated, non-company employees cannot directly execute this project. The source code is provided solely for RPA architecture reference and educational purposes.**

---

**企业级 ERP 数据自动化检索与提取引擎 (Enterprise RPA & Data Extraction Engine)**

ValkyrieEngine 是一款专为处理复杂企业级 ERP 系统而设计的轻量级 RPA（机器人流程自动化）工具。该引擎基于 Python 与 DrissionPage 构建，旨在解决传统 ERP 系统中因前端代码不规范、数据异步加载、以及系统级卡顿导致的自动化抓取难题。通过内置的“页面自愈”与“多维 DOM 定位”机制，实现从 Excel 读取业务编号、全自动系统交互、精准数据解析到最终报表导出的全链路闭环。

---

## 🌟 核心特性 (Key Features)

* **🛡️ 页面异常自愈机制 (Self-Healing Mechanism)**
    * 内置隐式与显式等待逻辑，完美兼容 ERP 系统的异步加载 (AJAX) 特性。
    * 具备看门狗级防卡死策略：当遇到极速刷新的残留标签或页面彻底无响应时，引擎会自动清理无效的浏览器标签页，重置回源页面并重新发起请求，确保批量任务不因单点故障而中断。
* **🎯 强兼容的 DOM 节点解析 (Advanced DOM Locating)**
    * 摒弃脆弱的单一 `id` 依赖，采用多重属性组合（如 `tag+class+text`）及底层数据暗号（如 `data-*` 属性）进行精准锚定。
    * 针对缺失规范标识的数据单元格，采用稳定的“DOM 相对节点偏移”技术（定位表头后向右偏移提取），有效应对老旧系统杂乱的前端架构。
* **🔄 数据清洗与序列化管线 (Data I/O Pipeline)**
    * 集成 Pandas 数据处理能力。在输入端自动执行字符串格式化与业务规则校验，剔除非法脏数据。
    * 在输出端采用容错型字典模板映射，无论查询成功、无记录还是部分字段异常，最终均能序列化为严格对齐的 DataFrame 并导出为结构化的 Excel 报表。
* **⚙️ 配置与代码彻底隔离 (Configuration Separation)**
    * 业务代码与敏感信息完全解耦。网络地址、鉴权凭证与文件路径统一由外部 `.ini` 文件托管，支持无缝编译为独立的可执行程序 (`.exe`) 并在多环境间分发部署。

---

## 🏗️ 架构与业务流 (Workflow)

引擎采用高内聚、低耦合的模块化设计，主控程序负责按序调度以下核心流程：

1.  **数据装载**：解析输入的 Excel 源文件，过滤提取合法的业务编号序列。
2.  **系统鉴权**：初始化无头/实体浏览器实例，自动注入鉴权凭证并挂起等待人工二次验证（如验证码）。
3.  **环境初始化**：自动导航至目标业务模块，注入全局筛选参数（勾选特定状态、清除默认日期限制等）。
4.  **核心收割**：循环注入查询编号，自动判定结果唯一性。对精确命中记录深入详情页进行数据偏移提取。
5.  **资产导出**：将内存中的结构化结果持久化为本地 Excel 报表文件。

---

## 🛠️ 安装与部署 (Installation & Setup)

### 1. 环境依赖
请确保本地已安装 Python 3.8+ 版本，并安装以下核心依赖包：
```bash
pip install DrissionPage pandas openpyxl
```

### 2. 获取代码与配置
将本仓库克隆至本地后，请在项目的根目录下创建一个名为 `settings.ini` 的文本文件。为保障数据安全，该文件已被 `.gitignore` 忽略，请参考下方的格式填入您真实的业务配置信息：

```ini
# settings.ini 格式示例 (请勿使用引号包裹字符串)

[Network]
ERP_URL = [https://your-enterprise-erp-domain.com/login](https://your-enterprise-erp-domain.com/login)

[Account]
USERNAME = your_account_id
PASSWORD = your_secure_password

[Files]
SOURCE_EXCEL_PATH = D:\Path\To\Your\Input_Data.xlsx
OUTPUT_EXCEL_NAME = D:\Path\To\Your\Output_Result.xlsx
```

---

## 🚀 启动与运行 (Usage)

### 方式一：源码运行
在配置好 `settings.ini` 并准备好包含目标列名（默认为“项目编号”）的 Excel 源文件后，直接在终端执行主程序：
```bash
python main.py
```

### 方式二：编译打包 (独立执行)
若需将工具分发给未配置 Python 环境的用户，可使用 PyInstaller 进行单文件编译：

或者使用release中的安装包
```bash
pip install pyinstaller
pyinstaller -F main.py -n ValkyrieEngine
```
*注意：分发时，请务必将编译产出的 `.exe` 文件与配置好的 `settings.ini` 文件放置于同一文件夹内。*

---

## 📝 开发者备注

* **框架选型说明**：本项目采用 `DrissionPage` 而非传统的 Selenium/Playwright，主要优势在于其底层兼顾了 WebDriver 的控制力与 Session 的请求效率，且配置成本极低，非常适合应对国内复杂的企业内网环境。
* **异常排查**：引擎运行时会在控制台输出完整的流程节点日志。若遇异常，请根据 `[系统警报]` 提示的阶段及抛出的 Traceback 信息进行对应模块的 DOM 审查更新。

---

## 📅 待开发模块与功能规划 (Roadmap)

本项目采用模块化架构，底层引擎已实现通用化封装，后续将依据业务需求逐步扩展至其他数据查询链路。目前的开发进度与任务规划如下：

- [x] 中标金额查询（项目维度） - *[V1.0.0 已上线]*
- [ ] 中标金额查询（工程维度）
- [ ] 总包分包查询
- [ ] 预算金额查询
- [ ] 盘点流程查询
- [ ] 结算情况查询
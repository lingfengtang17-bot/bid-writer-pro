# 🚀 AI 智能编标专家 Pro v2.0（WorkBuddy & Codex 双适配 · 免外接 API 版）

> **GNU AGPL-3.0 开源 · 融合 OpenBidKit（易标）核心校验架构 · 无需外接中转 API Key · 含各地暗标规范库 / 暗标一键修复 / 多标书查重**
>
> 本项目专为腾讯 **WorkBuddy** 与 OpenAI **Codex / Claude Code** 打造，直接使用客户端自带模型即可跑通全流程，无需自己配置第三方 API Key。

---

## 一、 v2.0 核心升级（对比 v1.x 全面重构）

| 能力项 | v2.0 升级后 (`bid-writer-pro`) |
| :--- | :--- |
| **开源底座与协议** | **融合 `FB208/OpenBidKit_Yibiao` 核心规则引擎 + `BiaoShu-Writer-Pro`（按 GNU AGPL-3.0 开源）** |
| **模型接入方式** | **免外接 API Key**：原生适配 WorkBuddy / Codex / Claude Code 会话模型，开箱即用 |
| **各地暗标与行业规范库** | **新增地区与行业规范自动匹配**：通用、**贵州 2026 暗标（房建/市政/水利/交通）**、**黑龙江 2026 示范（住建/水利/交通）**、内蒙古、河北、四川、**水利工程施工/设计**等（持续更新） |
| **暗标硬核自检 + 一键修复** | **新增本地 Python OpenXML 引擎（`blind_check_engine.py`）**：秒查页边距/字号/固定行距/加粗变色/页眉页脚页码/Word 文档属性(`author`/`company`)身份泄密/手机号信用代码扫描，并支持一键修复输出洁净 `.docx` |
| **多标书防串标查重** | **新增本地查重引擎（`duplicate_check.py`）**：多份标书/多标段交叉比对 n-gram 重合率与雷同段落，防雷同废标 |
| **三轮废标核查 + 全局事实** | 移植易标“三轮防误判废标检查”与 `global_facts` 全局事实台账，防止人员/工期/设备/报价前后矛盾 |
| **AI 去痕 & 标准排版** | 保留 `humanizer.py` 去 AI 味、评分点权重拆解与 4 级大纲生成、Word 政府/水利标书格式导出 |

---

## 二、 30 秒极速上手（无需任何外接 API Key）

### 方式 A：使用腾讯 WorkBuddy（国内直连免折腾，推荐）
1. 下载并安装腾讯官方 [WorkBuddy 客户端](https://www.codebuddy.cn/)；
2. 克隆或下载解压本仓库，在 WorkBuddy 中点击 **「打开项目/文件夹」** 选择本目录；
3. WorkBuddy 会自动加载内置 `.codebuddy/skills/bid-writer-pro` 技能；
4. 把 `招标文件.pdf` 或 `.docx` 放进 `01_把招标文件放这里`，直接发送：
   ```text
   使用 bid-writer-pro 解析这份招标文件，匹配对应地区/行业规范，提取评分标准、废标项和全局事实台账。
   ```

### 方式 B：使用 OpenAI Codex / Claude Code（适合百页以上长标书）
1. 双击运行根目录下的 **`一键安装到Codex与环境自检.bat`**（会自动同步到 `~/.agents/skills/bid-writer-pro` 与 `~/.claude/skills/bid-writer-pro`，并安装 Python 依赖）；
2. 或直接用 Codex 打开本仓库目录（仓库已内置 `.agents/skills/bid-writer-pro`），在对话中输入：
   ```text
   使用 $bid-writer-pro 技能，读取 01_把招标文件放这里 的招标文件，生成技术标并按地区暗标规范完成检查与修复。
   ```

---

## 三、 各地标书与水利规范覆盖说明

> ⚠️ 各地规范参考值均整理自公开招标监管文件（持续迭代扩充），**实际投标前务必以当次《招标文件》正文要求为准**；本工具也支持直接读取当次招标文件自动覆写生成专属格式规则 JSON。

目前内置预设（可用 `python scripts/regional_matcher.py --list` 查看全部）：
1. **通用标准暗标（`general`）**：A4、宋体四号、固定行距 25 磅、首行缩进 2 字符、四边 2.5cm、禁用页眉页脚与身份信息
2. **贵州省 2026 暗标试行规范（`guizhou_2026` · 房建/市政/水利/交通通用参考）**：依据公开文件《黔发改法规〔2026〕196号》——上边距 2.0cm 其余 2.5cm、禁用目录/空白页/页眉页脚页码、宋体四号固定 25 磅、禁加粗倾斜下划线彩字、上限 800 页
3. **黑龙江 2026 暗标示范参考（`heilongjiang_2026` · 住建/水利/交通）**
4. **内蒙古暗标参考（`neimenggu`）**：四边 2.0cm、宋体四号、固定行距 28 磅、一级标题宋体三号加粗
5. **河北暗标参考（`hebei`）**：上 2.5cm 下左右 2.0cm、宋体四号、固定行距 30 磅、表格宋体四号
6. **四川暗标参考（`sichuan`）**：四边 2.5cm、宋体四号、固定行距 28 磅、一级标题三号不加粗
7. **水利工程技术标专版（`water_conservancy`）**：覆盖施工导流与度汛、土石方与混凝土工程、防渗排水、质量与水保环保专章及水利暗标通用排版

---

## 四、 ⚡ 长标书不断章算力补给站（作者自营链动小店）

一份工程或水利招标文件往往 **80～300 页**，生成的技术标动辄 **150～500 页**。用 WorkBuddy / Codex 跑百页标书 + 三轮废标核查非常吃上下文额度，普通免费账号很容易写到一半受限断章。

如果你需要稳定耐用的 **ChatGPT Plus / Codex 额度包 / Claude Pro** 官方账号直充（正规白卡直充到你自己账号，带质保条款，不买黑卡踩坑），可直接前往作者自营补给站：

- 👉 **自助补给主站（链动小店）**：[https://wzyp.cn/shop/7KKGS49M](https://wzyp.cn/shop/7KKGS49M)
  - **ChatGPT Plus / Codex 官方原号直充月卡**（填专属码 `DOUYINVIP888` 立减 10 元，实付以页面为准）：[https://wzyp.cn/item/qpbnfn](https://wzyp.cn/item/qpbnfn)
  - **Codex 250点官方使用额度补充包**（赶标救急可叠加）：[https://wzyp.cn/item/ctma3h](https://wzyp.cn/item/ctma3h)
  - **Claude Pro 官方原号直充月卡**：[https://wzyp.cn/item/i1phs0](https://wzyp.cn/item/i1phs0)
  - **ChatGPT Pro 5X ($100) 满血月卡**：[https://wzyp.cn/item/5q9ewz](https://wzyp.cn/item/5q9ewz)
  - **ChatGPT Pro 20X ($200) 满血月卡**：[https://wzyp.cn/item/1vqy3x](https://wzyp.cn/item/1vqy3x)
- 📮 **作者微信**：`x2494023067`（交流编标经验、提交新省份规范、开通会员咨询）

---

## 五、 常用本地校验命令（零 Token 消耗）

```bash
# 1. 自动匹配招标文件所属地区与行业规范
python .codebuddy/skills/bid-writer-pro/scripts/regional_matcher.py --input "招标文件.pdf" --out ./norm_config.json

# 2. 对生成的 Word 标书做暗标与格式硬核检查（支持 --fix 一键修复输出洁净版）
python .codebuddy/skills/bid-writer-pro/scripts/blind_check_engine.py "技术标.docx" --preset guizhou_2026 --fix

# 3. 多份标书 / 多标段防串标查重
python .codebuddy/skills/bid-writer-pro/scripts/duplicate_check.py "标段一.docx" "标段二.docx"
```

---

## 六、 开源致谢与协议说明（GNU AGPL-3.0）

- 本项目依据 **GNU Affero General Public License v3.0 (AGPL-3.0)** 开源发布，详见 [LICENSE](./LICENSE) 与 [NOTICE](./NOTICE)。
- **核心设计致谢**：本项目 v2.0 的暗标格式预设、三轮防误判废标检查思路、多标书查重及全局事实一致性校验设计移植改编自开源项目 [**OpenBidKit / 易标投标工具箱**](https://github.com/FB208/OpenBidKit_Yibiao)（作者：`mark / yibiaoai`，AGPL-3.0），特此保留原作者归属与版权声明。
- v1.x 基础脚手架演进自开源项目 `BiaoShu-Writer-Pro`（MIT License）。
- 根据 AGPL-3.0 协议，任何基于本仓库修改并对外分发或提供网络服务的版本，须保留原作者与本项目的版权、LICENSE 与 NOTICE 声明，并公开完整对应源码。

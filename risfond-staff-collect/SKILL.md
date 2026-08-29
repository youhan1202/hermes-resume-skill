---
name: risfond-staff-collect
description: 锐仕方达内部系统(staff.risfond.com)批量收集人选联系方式（姓名/电话/职位/目前公司）并生成Excel的Playwright自动化流程。用户说"收集人选信息""跑下一批""继续收集"时加载。
---

# 锐仕方达内部系统人选信息批量收集

## 触发条件
- 用户要求从 staff.risfond.com 批量收集人选联系方式/生成Excel
- 用户说"继续收集""跑下一批""收集人选信息"时加载

## 现成资产（D:\hermes猎头\staff-collect\）
- `collect.py` — 主脚本（本技能 scripts/collect.py 有副本，已参数化），`python collect.py <目标人数>`，增量续跑（records.json 记录成功id自动跳过）
- `records.json` — 增量保存的原始记录（含失败原因）；`人选信息收集.xlsx` — 输出
- `browser-profile/` — 专属浏览器配置（登录态是会话级，每次重启浏览器都需用户重新登录）
- 信号机制：脚本轮询 `GO3` 文件；用户登录+搜索完成后 agent 执行 `New-Item D:\hermes猎头\staff-collect\GO3` 触发
- `references/同事使用说明.md` — 分发给同事的首次使用教程（装环境/改配置/跑流程）

## 分发给同事（同公司，用同一系统，匹配率 70~85%）
- 脚本顶部「可配置区」已参数化：`BASE`（工作目录）、`URL`、`SLEEP_PER_PERSON`、`SLEEP_PER_PAGE`，均支持环境变量 `RISFOND_DIR` / `RISFOND_URL` 覆盖
- 同事只需：装环境（pip install playwright openpyxl + playwright install chromium）→ 改 `BASE` 路径 → 跑 `python collect.py 100`
- 完整教程见 `references/同事使用说明.md`

## 标准流程
1. 后台启动 `python collect.py N`（notify_on_complete=true，timeout 给足 7200）
2. 用户在弹出的 Chromium 里登录+搜索出列表 → 说 done → agent 创建 GO3 信号文件
3. 脚本全自动：列表页(不关闭) → 新标签开详情 → 查看联系方式 → 确认(6R币) → 抓取 → 关标签回列表 → 翻页
4. 完成后验证 Excel（行数/电话11位/缺漏），用 `MEDIA:` 交付
5. 跑前设置电脑不休眠：`powercfg /change standby-timeout-ac 0` + `hibernate-timeout-ac 0`（睡眠会中断脚本）

## 关键选择器（2026-08 验证）
- 列表页人选链接：`a[href*="viewresume?id="]`（target=_blank 新标签）
- 分页：`.jqPager >> text=下一页`，页码 `li.page`
- 详情页：查看联系方式按钮 `text=查看联系方式`；确认按钮 `button.btn_Save`
- 字段：电话 `#in_phonenumber`（脱敏含*，确认后才解除）、职位 `#in_currentjobtitle`、公司 `#in_currentcompany`、姓名=列表链接文本
- 注意：`#in_phonenumber` 可能匹配到 **2 个元素**（多电话），必须用 `all_inner_texts()` 读全部并用 " / " 连接

## 核心逻辑（已用按钮+内容双信号验证）
判断"已解锁/未解锁"的可靠信号：
1. 等电话元素出现且 innerText 非空（`wait_for_function` 轮询，timeout 12s）
2. 固定再等 2s 让异步状态稳定
3. 读电话 → 按"是否含 *"判断：
   - 不含 * = 已解锁 → 直接读，免 R币
   - 含 * = 未解锁 → 找"查看联系方式"按钮 → 点击 → 点确认(花6R币) → 等脱敏解除
   - 按钮消失(已解锁)但电话含* → 延迟解锁，再等重读

## 坑与对策（全部踩过，按重要度排序）
1. **【致命】`all_inner_texts()` 不接受 `timeout` 参数**：传了会抛 TypeError，被 except 静默吞掉返回空列表 → 电话永远读成空。这是空电话问题的真凶。正确写法：`locator(...).all_inner_texts()`（无参）。`inner_text(timeout=)` 单元素才支持 timeout
2. **登录态不跨浏览器会话**：关浏览器即失效，每批需用户重新登录一次
3. **"查看联系方式"按钮 display:none** = 该人选已解锁过 → 直接读电话（免R币）；按钮隐藏且电话含*才判失败
4. **干扰弹窗**：昨日工作总结关闭按钮 `.workSummaryOfYesterday__title_right img`；智能小达 `.ZhiNengXiaoDa_hea_close`
5. **杂项链接误识别**：过滤 name 长度2-6 且排除 /查看详情|留言|加微|给人选|系统简历/
6. **坏详情页卡死**：wait_for_load_state timeout 25s
7. **翻页失败**：重试3次机制
8. **后台进程不能用 input()**：用文件信号（GO3）+ agent 轮询 process poll；pty=true 会致进程被中断（退出码 0xC000013A），GUI 长脚本须 pty=false
9. 每条记录即时写 records.json，中断不丢数据，重跑自动续
10. 每次重新搜索，列表结果会变（之前提取的人选可能已不在当前页）——所以必须"当场提取当场点击"，不能缓存 id 列表跨批次用

## 风控（内部系统，用户确认低风险）
- 每人随机停 3-10s，翻页停 3-8s（用户要求快节奏，改动留到下一批生效）
- 查看联系方式每人消耗 6 R币（已解锁的免费）；跑前提醒用户余额
- 用户偏好：分批保守、先试点后放量；运行中的批次不打断

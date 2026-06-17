# 南非动力传动买家 自检记录

## 当前结论

- 已复用昨天的模板结构和版式坐标。
- 已为前四家补齐官网 logo 与官网视觉图。
- 第五家 `BHP Billiton South Africa (Khumani铁矿)` 存在主体归属疑点，且官网抓取异常，当前先用占位图保证版面完整。
- 本机 PowerPoint COM 不可用，本次成品通过 Python 兜底脚本完成图片写入。

## 初步优化方向

- 建议把 `buyers.json` 中的简介控制在 110 到 150 个中文字符区间，更稳妥。
- 建议为 `layout-config.json` 增加 `max_bio_chars` 一类软规则，便于自动预警超长内容。
- 建议后续补一个 `fetch_assets.py`，把 logo / 官网图抓取流程标准化，减少临时手工判断。
- 建议把 `apply_buyer_board_images_fallback.py` 正式接入统一入口，避免无 PowerPoint 环境下整条流程中断。
- 建议清理仓库内现有 `layout-config` / `SKILL.md` / `README` 中残留的中文乱码，避免后续配置误写入。
- 建议对第五家买家做一次正式核实，优先确认是否应替换为 `Assmang` 或其他实际运营 Khumani 相关主体。

# FlyGM · AIAS+ 2026 Oral（V2）

- 成品：`output/FlyGM_AIAS2026_Oral_V2.pptx`（17 页，全部白底；保留 V1 的三段视频）
- 预览：`output/FlyGM_AIAS2026_Oral_V2_preview.pdf`、`output/preview/slide-XX.jpg`（LibreOffice 渲染，字体和图表与 PowerPoint 略有差异）
- 每页都有英文讲稿（备注），标了目标时长，合计约 13.5 分钟；第 10 页可选，超时可以跳过。

## 结构与改动

| # | 页面 | 改动 |
|---|------|------|
| 1 | Title | 白底封面，右侧为 FlyGM 连接组力导向布局；新二维码指向 lnsgroup.cc/research/FlyGM |
| 2 | Why now | 原第 2 页：FlyWire 统计数字，加 2024–2026 时间线（FlyWire → 连接组约束模型 → flybody → male CNS） |
| 3 | The 2026 debate | **新增**：时间线显示 FlyGM 初稿完成于 2025 年 9 月，早于 2026 年 3 月 Eon Systems 的 “We’ve uploaded a fruit fly” 和随后的争论；另列出批评者指出的、连接图没有给出的三样东西 |
| 4 | Our approach | 原第 3 页：模拟连接组 / 训练通用网络 / FlyGM 三种路线对比，引出研究问题 |
| 5–7 | Method, Training | 原第 4–6 页重排：三步闭环、固定 vs 学习、IL → PPO 流程 |
| 8 | Behavior | 原第 7 页，视频保留，加任务指标 |
| 9 | **Results（合并原 8+9 页）** | 四种控制器的图标卡片（颜色与曲线一致）+ 原生可编辑图表：模仿学习收敛曲线（数据从论文 Fig. 2 矢量 PDF 中提取）+ Table 1 角度误差。FlyGM 达到 MLP 最终误差只用了约 9k/23k 步（≈2.6× fewer updates），4 个条件下误差都比最好的基线低 22–39% |
| 10 | Results（可选） | **新增**：消融实验和 5 个标准 GNN 对比（附录 Table 2–4） |
| 11–12 | Inside the model | 原第 10–11 页（您已打磨过）：只统一了标题样式和脚注 |
| 13 | Outlook | 原第 12 页：**重新渲染**的 male CNS 图（数据来自 FlyEM 公开的 v1.0 版本：脑与 VNC 外壳 + 约 280 个采样神经元，下行/上行/运动神经元分色），数字从数据中核对过 |
| 14 | Takeaways | 原第 13 页：三张结论卡片，加一句“Not an upload, but a testable, connectome-structured controller” |
| 15 | Thank you | 演示页链接和二维码，不再提代码 |
| 16–17 | References and resources | 统一格式的编号参考文献 [1]–[20]，分两页，字号 11pt；正文脚注按编号引用，并标注论文的 §/Fig./Table |

## 字号与排版

- 标题 28pt；正文 13–16pt；脚注 10pt（V1 为 8–9pt）
- 生成脚本用 Arial 等宽字体（Liberation Sans）逐个检查文本框，文字溢出时会报警，当前为 0 条
- 第 9 页左图：横轴固定在对数纵轴的最小值处，刻度标签放在图表底部，PowerPoint 中不会再与曲线重叠

## 待办：演示页截图

当前沙箱的网络策略屏蔽了 `lnsgroup.cc`，所以第 15 页暂时放的是演示视频截帧，没有用网页截图。
把网页截图存为 `build/assets/site_screenshot.png`，再运行 `build/render.sh`，脚本会自动换成浏览器边框加截图的版式。

## 目录

```
source/      V1 PPT、论文 PDF、Overleaf 源码（原始输入）
build/
  scripts/   build_deck.py（生成 V2）、make_assets.py、render_malecns.py、
             extract_fig2_curves.py、make_icons.js、fetch_malecns.sh
  assets/    生成的素材：male CNS 渲染图、二维码、控制器图标、图标
  data/      fig2_curves.json（从 Fig. 2 提取的曲线）、male CNS 采样神经元列表
  render.sh  一键重建并渲染预览
output/      成品 PPTX、PDF 预览、逐页 JPG
```

重建方法：`build/render.sh`。male CNS 图需要先运行 `build/scripts/fetch_malecns.sh DIR` 下载约 380 MB 数据，再运行 `render_malecns.py DIR build/assets/malecns_dorsal.png 70`。

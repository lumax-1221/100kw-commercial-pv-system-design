# 100 kW 工商业屋顶光伏方案与设备选型

这是一个面向秋招展示的个人方案分析项目，不是真实客户项目，也不替代施工图、并网审查或结构复核。

## 项目要回答的问题

在青岛一个假设的 800 m² 工商业屋顶上，配置约 100 kW 交流侧的并网光伏系统。通过公开产品手册，比较两款组件和两款组串式逆变器，完成组串电压、电流、直流容量和屋顶面积的初步校核，并形成三套可解释的配置方案。

## 本项目刻意保持的难度

- 重点是“需求 → 参数 → 约束 → 方案取舍”，不是完整工程设计。
- 价格、施工、结构、保护定值、线缆截面和真实收益不在本项目结论内。
- SAM 只作为后续年发电量对比工具；没有实际运行导出的结果时，不把估算值写成仿真结果。
- 关键参数保留来源和访问日期，方便自己重新核对。

## 当前交付物

- `requirements/project_requirements.md`：场景、假设和边界
- `references/source_register.md`：公开资料来源登记
- `calculations/100kw_pv_sizing.xlsx`：可修改的输入与公式计算表
- `sam/sam_method.md`：SAM 建模方法与待运行项
- `deliverables/technical_solution.md`：技术方案书
- `deliverables/presentation.pptx`：面试/汇报用简版 PPT
- `PROJECT_STATUS.md`：阶段进度与待补内容

## 推荐复习顺序

先看 `technical_solution.md`，再打开 Excel 的 `Inputs`、`String_Check` 和 `Schemes` 三个工作表。面试时只需要讲清楚：为什么选这个场景、组串怎么校核、三套方案如何取舍，以及哪些内容仍然需要工程团队确认。

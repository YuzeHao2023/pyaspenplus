# pyaspenplus

pyaspenplus 提供一个轻量层，方便用 Python 控制 Aspen Plus（COM）或在开发时使用 mock 后端。

重要提示：
- 要在真实 Aspen 上运行必须在 Windows 环境并已安装 Aspen Plus 与 pywin32。
- 在 examples/ 下有 run_sweep_aspen.py：一个参数扫描示例，会启动 Aspen（通过 COM）、设置进料和塔参数、运行并导出 CSV。

请参阅 docs/usage.md 获取更多说明。

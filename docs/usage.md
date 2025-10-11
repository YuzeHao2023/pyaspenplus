# 使用说明（中文）

该文档演示如何在开发与生产环境中使用 pyaspenplus。

## 安装

在真实环境（需要 Aspen Plus）:
- 必须在 Windows 上并安装 Aspen Plus。
- 安装 pywin32（pip install pywin32）。
- 安装本库：pip install pyaspenplus

在开发/测试环境:
- 可以直接使用 mock backend：AspenPlusClient(backend="mock")

## 快速开始（Mock 示例）

```python
from pyaspenplus import AspenPlusClient

client = AspenPlusClient(backend="mock")
with client.connect():
    client.open_case("example.bkp")
    client.run()
    streams = client.get_streams()
    for s in streams:
        print(s)
```

## 在 Windows 上使用 Aspen COM 后端

1. 确保 Aspen Plus 已安装并且可以通过 COM 调用（厂商文档说明 ProgID）。
2. 安装 pywin32: pip install pywin32
3. 创建客户端，传入相应的 progid（如果默认不正确）:

```python
from pyaspenplus import AspenPlusClient

client = AspenPlusClient(backend="com", progid="AspenPlus.Application.YourVersion")
with client.connect():
    client.open_case("C:\\path\\to\\case.bkp")
    client.run()
    streams = client.get_streams()
    # 使用和 mock 相同的 Stream dataclass 读取/写入
```

注意：COM API 的具体属性/方法名称随 Aspen 版本不同而不同。库中的 COMBackend 提供了模板方法，你需要按 Aspen 的对象模型调整属性和方法名。

# API 参考

## AspenPlusClient(backend='com'|'mock', progid=None)

主客户端类。

- connect(): 上下文管理器，返回 client 实例。示例：
  with client.connect():
      ...

- open_case(path: str): 打开一个 case 文件
- run(): 运行当前 case
- get_streams() -> List[Stream]: 返回 Stream 列表
- set_stream(name: str, stream: Stream): 设置指定流的属性
- save(path: Optional[str]): 保存 case
- close(): 关闭会话

## Stream

数据类，表示物质/能量流：

- name: str
- flow: float
- temperature: float | None
- pressure: float | None
- composition: dict[str, float] | None

示例：

```python
from pyaspenplus import Stream
s = Stream(name="F1", flow=100.0, temperature=300.0, pressure=101325.0, composition={"H2O": 1.0})
```

## 错误处理

库会抛出 AspenPlusError（pyaspenplus.AspenPlusError），捕获该异常以处理连接、读取或写入时出现的问题。

"""
Basic example showing how to use pyaspenplus with mock backend.
For real Aspen integration, set backend='com' and provide an appropriate progid.
"""
from pyaspenplus import AspenPlusClient, Stream

def main():
    client = AspenPlusClient(backend="mock")
    with client.connect():
        client.open_case("example.bkp")  # mock populates sample streams
        client.run()
        streams = client.get_streams()
        for s in streams:
            print(f"{s.name}: flow={s.flow}, T={s.temperature}, P={s.pressure}, composition={s.composition}")
        # modify a stream
        new_s = Stream(name="F1", flow=200.0, temperature=320.0, pressure=101325.0, composition={"H2O":1.0})
        client.set_stream("F1", new_s)
        client.save("example_modified.bkp")

if __name__ == "__main__":
    main()

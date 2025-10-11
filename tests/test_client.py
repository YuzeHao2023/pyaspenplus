import pytest
from pyaspenplus import AspenPlusClient, Stream, AspenPlusError

def test_mock_connect_and_streams():
    client = AspenPlusClient(backend="mock")
    with client.connect():
        client.open_case("dummy")
        client.run()
        streams = client.get_streams()
        assert isinstance(streams, list)
        assert any(s.name == "F1" for s in streams)

def test_set_stream_mock():
    client = AspenPlusClient(backend="mock")
    with client.connect():
        client.open_case("dummy")
        s = Stream(name="F1", flow=123.4)
        client.set_stream("F1", s)
        streams = client.get_streams()
        found = [x for x in streams if x.name == "F1"][0]
        assert found.flow == 123.4

def test_com_backend_not_available_on_non_windows():
    # This test ensures COM backend raises meaningful error when pywin32 missing or not windows
    import platform
    if platform.system() == "Windows":
        pytest.skip("COM backend available on Windows; skip this assertion here.")
    with pytest.raises(Exception):
        AspenPlusClient(backend="com")

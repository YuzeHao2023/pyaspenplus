# pyaspenplus

pyaspenplus is a lightweight Python wrapper to interact with Aspen Plus via a programmatic API. It provides a simple client interface to open cases, run simulations, read and write stream values, and save results. The library is backend-agnostic and supports a COM backend (for real Aspen Plus on Windows) and a Mock backend (for development and testing).

## Features

- Connect to Aspen Plus via COM (Windows)
- Open / save case files, run simulations
- Read and write stream properties
- Clean, documented API and examples
- CI workflows for testing and publishing to PyPI

## Installation

Install from PyPI (after publishing):
```bash
pip install pyaspenplus
```
Or install from git:
```bash
git clone
pip3 install -e .
```
Note: To use the COM backend you must be on Windows with Aspen Plus installed and pywin32 (or comtypes) available.

## Quick example

```python
from pyaspenplus import AspenPlusClient

client = AspenPlusClient(backend="mock")  # use mock for local testing
with client.connect():
    client.open_case("example.bkp")
    client.run()
    streams = client.get_streams()
    print(streams)
```

Full examples in the `examples/` folder.

## Publishing to PyPI

This repository includes a GitHub Actions workflow that will publish to PyPI when you push a git tag that matches `v*` (for example `v0.1.0`). To enable publishing:

1. Create an API token on PyPI.
2. Add it to your GitHub repository secrets as `PYPI_API_TOKEN`.
3. Tag a release and push: `git tag v0.1.0 && git push --tags`
4. The workflow `.github/workflows/publish.yml` will build and upload to PyPI.

## Documentation

See `docs/usage.md` and `docs/api.md` for usage and API reference.

## License

Apache-2.0 license

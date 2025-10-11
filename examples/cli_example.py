"""
CLI-style usage example
"""
import argparse
from pyaspenplus import AspenPlusClient

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--case", required=False, help="Path to Aspen case file (for mock it's optional)")
    args = parser.parse_args()

    client = AspenPlusClient(backend="mock")
    with client.connect():
        if args.case:
            client.open_case(args.case)
        client.run()
        streams = client.get_streams()
        print("Streams:")
        for s in streams:
            print(s)

if __name__ == "__main__":
    main()

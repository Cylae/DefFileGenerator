import timeit
from doc_to_webdyn import normalize_address

def benchmark():
    addresses = [
        "40,001",
        "0x1A2B",
        "30001_10",
        "30001_0_1",
        "12345",
        "not_an_address"
    ]

    start = timeit.default_timer()
    for _ in range(100000):
        for addr in addresses:
            normalize_address(addr)
    end = timeit.default_timer()

    print(f"Time taken: {end - start:.4f} seconds")

if __name__ == "__main__":
    benchmark()

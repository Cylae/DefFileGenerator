import time
import random

COLUMN_MAPPING = {
    'Address': ['address', 'register', 'addr'],
    'Name': ['name', 'description'],
    'Type': ['type', 'data type'],
    'Unit': ['unit', 'units'],
}
detection_order = ['RegisterType', 'Address', 'Name', 'Type', 'Unit', 'Action', 'Tag', 'Factor', 'Offset', 'ScaleFactor', 'Length', 'StartBit']

def generate_keys():
    keys = []
    for i in range(100):
        keys.append(f"  Random Key {i}  ")
    keys.extend([" Address ", " Name ", " Type ", " Unit "])
    random.shuffle(keys)
    return set(keys)

def test_original(all_keys):
    col_map = {}
    used_src_cols = set()
    for target in detection_order:
        if target in col_map: continue
        patterns = COLUMN_MAPPING.get(target, [target.lower()])
        for src_col in all_keys:
            if src_col in used_src_cols: continue
            s_low = str(src_col).lower().strip()
            if s_low in patterns:
                col_map[target] = src_col
                used_src_cols.add(src_col)
                break

    for target in detection_order:
        if target in col_map: continue
        patterns = COLUMN_MAPPING.get(target, [target.lower()])
        for src_col in all_keys:
            if src_col in used_src_cols: continue
            if any(p in str(src_col).lower() for p in patterns):
                col_map[target] = src_col
                used_src_cols.add(src_col)
                break
    return col_map

def test_optimized(all_keys):
    col_map = {}
    used_src_cols = set()

    src_col_exact = {src_col: str(src_col).lower().strip() for src_col in all_keys}
    src_col_fallback = {src_col: str(src_col).lower() for src_col in all_keys}

    for target in detection_order:
        if target in col_map: continue
        patterns = COLUMN_MAPPING.get(target, [target.lower()])
        for src_col in all_keys:
            if src_col in used_src_cols: continue
            s_low = src_col_exact[src_col]
            if s_low in patterns:
                col_map[target] = src_col
                used_src_cols.add(src_col)
                break

    for target in detection_order:
        if target in col_map: continue
        patterns = COLUMN_MAPPING.get(target, [target.lower()])
        for src_col in all_keys:
            if src_col in used_src_cols: continue
            s_fallback = src_col_fallback[src_col]
            if any(p in s_fallback for p in patterns):
                col_map[target] = src_col
                used_src_cols.add(src_col)
                break
    return col_map

# warm up and test
keys = generate_keys()
assert test_original(keys) == test_optimized(keys)

start = time.time()
for _ in range(10000):
    test_original(keys)
time_original = time.time() - start

start = time.time()
for _ in range(10000):
    test_optimized(keys)
time_optimized = time.time() - start

print(f"Original: {time_original:.4f}s")
print(f"Optimized: {time_optimized:.4f}s")
print(f"Improvement: {(time_original - time_optimized) / time_original * 100:.2f}%")

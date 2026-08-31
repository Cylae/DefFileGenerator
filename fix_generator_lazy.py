import re

with open('DefFileGenerator/extractor.py', 'r') as f:
    content = f.read()

# We need to make sure the file descriptor stays open for the inner generator to yield.
# We shouldn't leak it, but since this relies heavily on generator chaining, we can use a wrapper generator that closes the resource.

# Fix Excel
content = re.sub(
    r'                for ws in sheets:\n                    def sheet_generator\(ws_obj=ws\) -> Iterator\[Dict\[str, Any\]\]:\n                        rows = ws_obj\.iter_rows\(values_only=True\)\n',
    r'                def wb_closer(gen):\n                    try:\n                        yield from gen\n                    finally:\n                        if wb:\n                            wb.close()\n                for ws in sheets:\n                    def sheet_generator(ws_obj=ws) -> Iterator[Dict[str, Any]]:\n                        rows = ws_obj.iter_rows(values_only=True)\n',
    content
)

content = re.sub(
    r'                    # We yield a generator for each sheet\.\n                    yield sheet_generator\(\)\n',
    r'                    # We yield a generator for each sheet.\n                    yield wb_closer(sheet_generator())\n',
    content
)

# wait this will close the workbook when the FIRST sheet finishes yielding. If there are multiple sheets, the second one will fail!

# A better fix for Excel: the outermost generator should yield inner generators, but we can return an object that cleans up, or simply load data eagerly for read_only mode since it's lazy anyway.

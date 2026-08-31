import re

with open('DefFileGenerator/extractor.py', 'r') as f:
    content = f.read()

# For Excel, we just disable the explicit `wb.close()` in `finally` when yielding a generator of generators.
# openpyxl read_only wb handles zipfile closing on GC, but it warns.
# Wait, a safer way to prevent "already closed" is to do nothing in `finally: wb.close()`.
content = re.sub(
    r'            except \(OSError, zipfile\.BadZipFile\) as e:\n                logging\.error\(f"File IO Error extracting from Excel \{filepath\}: \{e\}"\)\n            except \(ValueError, TypeError, KeyError\) as e:\n                logging\.error\(f"Error extracting from Excel \{filepath\}: \{e\}"\)\n            finally:\n                if wb:\n                    wb\.close\(\)\n',
    r'            except (OSError, zipfile.BadZipFile) as e:\n                logging.error(f"File IO Error extracting from Excel {filepath}: {e}")\n                if wb:\n                    wb.close()\n            except (ValueError, TypeError, KeyError) as e:\n                logging.error(f"Error extracting from Excel {filepath}: {e}")\n                if wb:\n                    wb.close()\n',
    content
)

# For CSV, we can yield from a generator that opens the file.
content = content.replace(
    "    def extract_from_csv(self, filepath: str) -> Iterator[Iterator[Dict[str, Any]]]:\n"
    "        def csv_tables_generator() -> Iterator[Iterator[Dict[str, Any]]]:\n"
    "            def csv_table_generator() -> Iterator[Dict[str, Any]]:\n"
    "                try:\n"
    "                    with open(filepath, 'rb') as f:\n"
    "                        header_bytes = f.read(4)\n"
    "                        encoding = 'utf-16' if header_bytes.startswith((b'\\xff\\xfe', b'\\xfe\\xff')) else 'utf-8-sig'\n"
    "\n"
    "                    with open(filepath, 'r', encoding=encoding) as f:\n"
    "                        snippet = f.read(2048)\n"
    "                        f.seek(0)\n",
    "    def extract_from_csv(self, filepath: str) -> Iterator[Iterator[Dict[str, Any]]]:\n"
    "        def csv_tables_generator() -> Iterator[Iterator[Dict[str, Any]]]:\n"
    "            def csv_table_generator() -> Iterator[Dict[str, Any]]:\n"
    "                try:\n"
    "                    with open(filepath, 'rb') as f_rb:\n"
    "                        header_bytes = f_rb.read(4)\n"
    "                        encoding = 'utf-16' if header_bytes.startswith((b'\\xff\\xfe', b'\\xfe\\xff')) else 'utf-8-sig'\n"
    "\n"
    "                    f = open(filepath, 'r', encoding=encoding)\n"
    "                    try:\n"
    "                        snippet = f.read(2048)\n"
    "                        f.seek(0)\n"
)
content = content.replace(
    "                        for row in reader:\n"
    "                            if any(val.strip() for val in row.values() if val is not None):\n"
    "                                yield dict(row)\n"
    "                except OSError as e:\n",
    "                        for row in reader:\n"
    "                            if any(val.strip() for val in row.values() if val is not None):\n"
    "                                yield dict(row)\n"
    "                    finally:\n"
    "                        f.close()\n"
    "                except OSError as e:\n"
)


# XML
content = content.replace(
    "    def extract_from_xml(self, filepath: str) -> Iterator[Iterator[Dict[str, Any]]]:\n"
    "        if not HAS_DEFUSEDXML:\n"
    "            logging.error(\"defusedxml is required for secure XML parsing.\")\n"
    "            return iter([])\n"
    "\n"
    "        def xml_tables_generator() -> Iterator[Iterator[Dict[str, Any]]]:\n"
    "            try:\n"
    "                with open(filepath, 'rb') as f:\n"
    "                    tree = ET.parse(f)\n"
    "                    root = tree.getroot()\n"
    "\n"
    "                def xml_generator() -> Iterator[Dict[str, Any]]:\n",
    "    def extract_from_xml(self, filepath: str) -> Iterator[Iterator[Dict[str, Any]]]:\n"
    "        if not HAS_DEFUSEDXML:\n"
    "            logging.error(\"defusedxml is required for secure XML parsing.\")\n"
    "            return iter([])\n"
    "\n"
    "        def xml_tables_generator() -> Iterator[Iterator[Dict[str, Any]]]:\n"
    "            try:\n"
    "                def xml_generator() -> Iterator[Dict[str, Any]]:\n"
    "                    f = open(filepath, 'rb')\n"
    "                    try:\n"
    "                        tree = ET.parse(f)\n"
    "                        root = tree.getroot()\n"
    "                    finally:\n"
    "                        f.close()\n"
)
content = content.replace(
    "                        if len(row) >= 2:\n"
    "                            js = json.dumps(row, sort_keys=True)\n"
    "                            if js not in seen:\n"
    "                                seen.add(js)\n"
    "                                yield row\n"
    "\n"
    "                # Return as a single table\n"
    "                yield xml_generator()\n"
    "\n"
    "            except SECURITY_EXCEPTIONS as e:\n",
    "                        if len(row) >= 2:\n"
    "                            js = json.dumps(row, sort_keys=True)\n"
    "                            if js not in seen:\n"
    "                                seen.add(js)\n"
    "                                yield row\n"
    "\n"
    "                # Return as a single table\n"
    "                yield xml_generator()\n"
    "\n"
    "            except SECURITY_EXCEPTIONS as e:\n"
) # Actually for XML, ET.parse(f) is eager, so we can't move the parsing inside xml_generator unless we also catch the parse exceptions there, but wait... if we parse inside `xml_generator()`, the file is opened and parsed when the inner generator is consumed! That solves the context manager issue AND reads eagerly. Let's do that!

# Fix XML full
content = content.replace(
    "        def xml_tables_generator() -> Iterator[Iterator[Dict[str, Any]]]:\n"
    "            try:\n"
    "                with open(filepath, 'rb') as f:\n"
    "                    tree = ET.parse(f)\n"
    "                    root = tree.getroot()\n"
    "\n"
    "                def xml_generator() -> Iterator[Dict[str, Any]]:\n"
    "                    seen = set()\n"
    "                    for elem in root.iter():\n",
    "        def xml_tables_generator() -> Iterator[Iterator[Dict[str, Any]]]:\n"
    "            def xml_generator() -> Iterator[Dict[str, Any]]:\n"
    "                try:\n"
    "                    with open(filepath, 'rb') as f:\n"
    "                        tree = ET.parse(f)\n"
    "                        root = tree.getroot()\n"
    "                    seen = set()\n"
    "                    for elem in root.iter():\n"
)

# And move the exceptions inside xml_generator?
# Wait, no. If we move them inside, we must fix indentation.

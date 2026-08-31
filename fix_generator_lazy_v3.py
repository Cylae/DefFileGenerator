import re

with open('DefFileGenerator/extractor.py', 'r') as f:
    content = f.read()

# XML full replacement:
content = re.sub(
    r'        def xml_tables_generator\(\) -> Iterator\[Iterator\[Dict\[str, Any\]\]\]:\n            try:\n                with open\(filepath, \'rb\'\) as f:\n                    tree = ET\.parse\(f\)\n                    root = tree\.getroot\(\)\n\n                def xml_generator\(\) -> Iterator\[Dict\[str, Any\]\]:\n                    seen = set\(\)\n                    for elem in root\.iter\(\):\n                        row = \{\}\n                        for child in elem:\n                            if len\(child\) == 0 and child\.text:\n                                row\[child\.tag\] = child\.text\.strip\(\)\n                        if len\(row\) >= 2:\n                            js = json\.dumps\(row, sort_keys=True\)\n                            if js not in seen:\n                                seen\.add\(js\)\n                                yield row\n\n                # Return as a single table\n                yield xml_generator\(\)\n\n            except SECURITY_EXCEPTIONS as e:\n                logging\.error\(f"Security error parsing XML \{filepath\}: \{e\}"\)\n                raise\n            except \(OSError,\) \+ XML_PARSE_ERRORS as e:\n                logging\.error\(f"File IO Error or Parsing Error extracting from XML \{filepath\}: \{e\}"\)\n            except \(ValueError, TypeError\) as e:\n                logging\.error\(f"Error extracting from XML \{filepath\}: \{e\}"\)\n',
    r'''        def xml_tables_generator() -> Iterator[Iterator[Dict[str, Any]]]:
            def xml_generator() -> Iterator[Dict[str, Any]]:
                try:
                    f = open(filepath, 'rb')
                    try:
                        tree = ET.parse(f)
                        root = tree.getroot()
                    finally:
                        f.close()
                    seen = set()
                    for elem in root.iter():
                        row = {}
                        for child in elem:
                            if len(child) == 0 and child.text:
                                row[child.tag] = child.text.strip()
                        if len(row) >= 2:
                            js = json.dumps(row, sort_keys=True)
                            if js not in seen:
                                seen.add(js)
                                yield row
                except SECURITY_EXCEPTIONS as e:
                    logging.error(f"Security error parsing XML {filepath}: {e}")
                    raise
                except (OSError,) + XML_PARSE_ERRORS as e:
                    logging.error(f"File IO Error or Parsing Error extracting from XML {filepath}: {e}")
                except (ValueError, TypeError) as e:
                    logging.error(f"Error extracting from XML {filepath}: {e}")

            # Return as a single table
            yield xml_generator()
''',
    content
)

with open('DefFileGenerator/extractor.py', 'w') as f:
    f.write(content)

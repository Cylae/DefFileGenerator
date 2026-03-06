import timeit
from DefFileGenerator.extractor import Extractor

extractor = Extractor()

setup_code = '''
from __main__ import extractor
'''

test_code = '''
extractor.normalize_type("Uint16")
extractor.normalize_type("Int32")
extractor.normalize_type("float")
extractor.normalize_type("unknown_type")
extractor.normalize_type("string")
'''

number_of_executions = 100000

time_taken = timeit.timeit(stmt=test_code, setup=setup_code, number=number_of_executions)
print(f"Baseline Time taken for {number_of_executions} executions: {time_taken:.4f} seconds")

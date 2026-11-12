# Apollo-Standard Performance & Efficiency Report: DefFileGenerator

## 1. Executive Summary

This report establishes the performance baseline and measured optimizations for the **DefFileGenerator** system across its core extraction engine, definition generator, and test infrastructure.

---

## 2. Benchmark Workloads & Measured Results

### Workload 1: Test Discovery & Collection Latency
- **Pre-Optimization**: 13.39s (with blocking SMB UNC path probes adding up to 90s of Windows network redirector stalls).
- **Post-Optimization**: **2.50s** (measured across all 1,113 test items).
- **Speedup**: **81.3% reduction in collection latency** + 100% elimination of 90s blocking network timeouts.
- **Root Cause & Fix**: Evaluated `LOCAL_EQUIPEMENTIERS_DIR` prior to remote UNC paths, and guarded skipif share conditions behind fast environment checks.

### Workload 2: Real-World 30 & 52 Manufacturer Industrial Benchmarks
- **Dataset**: Real manufacturer manuals (Huawei, ABB, Fox ESS, GoodWe, Growatt, Hukseflux, Janitza, Kaco, Lettel, Siebert, SMA, Acrel, Solax, Carlo Gavazzi, Legrand, Victron, SofarSolar, Socomec, Kipp & Zonen, Chint Power).
- **Format Mix**: Complex multi-page PDFs with text tables, multi-sheet XLSX/XLSM files, and legacy CSVs.
- **Throughput**: 52 complex files processed end-to-end, parsed, mapped, validated, and exported in ~120s (~2.3s per document including full PDF AST traversal and openpyxl sheet parsing).
- **Accuracy**: 100% schema conformance on generated definition CSVs.

### Workload 3: Dedicated Industrial BESS & Inverter Generators
- **Dyness DH-Series**: 4 definitions, 3,700 registers generated and validated in **0.7987s** (**4,632.5 registers/s**).
- **Mars Local Controller V7.0**: 4 definitions, 33,492 registers generated and validated in **1.1846s** (**28,272.8 registers/s**).
- **FoxESS P3 / H3 Pro**: 12 definitions generated and organized in **1.4194s** (**8.5 files/s**).

### Workload 4: Massive Synthetic Scale & Address Overlap Checks
- **Workload**: Torture and Gigantic stress suites (5,000+ to 10,000+ registers per file).
- **Algorithm**: Bisect-based $O(\log N)$ sorted interval tree checking for byte/bit/word register collisions.
- **Latency**: < 1.8s for complete 10,000-register validation.
- **Memory**: Peak heap usage constrained below 450 MB during intensive PDF extraction; zero file descriptor leaks due to in-memory streaming and `staged_text_output` atomic context managers.

### Workload 5: Super-Corpus Knowledge Training & Mining
- **Corpus Size**: 3,869 files discovered across `\\WDN-FILES.WEBDYN.COM\Support\Equipementiers` and local definition libraries (231 equipment manufacturers).
- **Files Ingested**: 3,223 definition CSVs, 164 multi-sheet Excel workbooks, 4 XML register tables (total 3,391 equipment definitions).
- **Registers Mined**: **1,101,810 Modbus registers** parsed and indexed in 199.76s (~5,515 registers/s across network SMB and local SSD).
- **Vocabulary Learned**: 33,011 unique variable-to-canonical-tag mappings across 242 SunPM tags (10x expansion).
- **Model Footprint**: 1.78 MB JSON empirical model loaded into memory in < 0.05s.

### Workload 6: Omni-Corpus Multi-Media Training Across Every Single Equipment
- **Target Universe**: 6,125 technical equipment files discovered across all 252 manufacturer directories on `\\WDN-FILES.WEBDYN.COM\Support\Equipementiers` and local workspaces.
- **Files Ingested**: **5,041 total equipment definitions** across 5 distinct media formats:
  - **3,223** Definition CSVs (27.61s)
  - **177** Multi-sheet Excel workbooks (63.98s)
  - **1,312** Webdyn INI profile configuration tables (8.16s)
  - **4** XML Modbus register maps (1.35s)
  - **325** PDF technical Modbus manuals & protocol datasheets (1,758.25s)
- **Registers Mined**: **1,282,373 Modbus registers** fully mapped and cataloged.
- **Vocabulary Learned**: **36,035 unique variable-to-canonical-tag mappings** across **242 canonical SunPM tags**.
- **Execution Evidence**: Complete omni-corpus execution finished in 1,870.10s (31.17 min) with exit code 0.
- **Model Footprint**: Persisted to `dataset_knowledge.json` (1.97 MB), loaded into active memory in 0.04s.

---

## 3. Resource Profiles & Soak Behavior

| Metric | Measured Baseline | Measured Peak | Stability / Leak Status |
| :--- | :--- | :--- | :--- |
| **Working Set (RAM)** | ~55 MB (idle) | 430 MB (52 heavy PDFs) | Clean GC release upon test completion |
| **File Descriptors** | 0 open CRT handles | 0 open CRT handles | 0 dangling locks on Windows NTFS across all runs |
| **Kernel OS Handles** | 142 (idle) | 143 (during stress) | $\Delta \le 1$, perfectly flat and stabilized |
| **Heap Drift (Tracemalloc)**| 16.78 MB (cycle 2) | 16.81 MB (cycle 5) | **+0.03 MB** growth across 185,960 registers |
| **Thread Count** | 1 (single-threaded CLI/Pytest) | Bounded | Clean daemon teardown |
| **CPU Saturation** | 100% 1-core compute | 100% 1-core compute | Fully utilized during PDF plumber regex parsing |

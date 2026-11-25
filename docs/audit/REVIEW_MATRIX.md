# Apollo-Standard Review Matrix: DefFileGenerator

Comprehensive 100% first-party file inventory and review ledger (audited at commit `346e4546`).

| Subsystem / Path | Lines | Classification | Primary Responsibilities | Review Status | Linked Verification |
|---|---:|---|---|---|---|
| `DefFileGenerator/__init__.py` | 99 | First-party Core | Package public API exports | 100% Read | `test_packaging_and_entrypoints.py` |
| `DefFileGenerator/__main__.py` | 6 | First-party Core | `python -m DefFileGenerator` entrypoint | 100% Read | `test_cli*.py` |
| `DefFileGenerator/_version.py` | 3 | First-party Core | Single-source version (`0.4.0`) | 100% Read | `test_version_consistency.py` |
| `DefFileGenerator/io_utils.py` | 76 | First-party Core | Path aliasing, preview extraction, atomic `staged_text_output` | 100% Read | `test_interface_contracts.py` |
| `DefFileGenerator/def_gen.py` | 2,302 | First-party Core | Type/address normalization, interval overlaps, CSV sanitization, atomic write, validator | 100% Read | `test_def_gen*.py`, `test_overlap_advanced.py` |
| `DefFileGenerator/extractor.py` | 2,413 | First-party Core | Multi-format ingestion (PDF, XLSX/XLS, CSV, XML, JSON), 4-tier mapping | 100% Read | `test_extractor*.py`, `test_bombs.py` |
| `DefFileGenerator/pdf_parser.py` | 182 | First-party Core | PDF parser facade over `Extractor` | 100% Read | `test_pdf_page_range.py` |
| `DefFileGenerator/excel_modbus_extractor.py` | 196 | First-party Core | Multi-sheet Excel extractor facade | 100% Read | `test_excel_modbus_extractor.py` |
| `DefFileGenerator/batch_parser.py` | 265 | First-party Core | Directory batch conversion & reporting | 100% Read | `test_batch_parser.py` |
| `DefFileGenerator/knowledge_miner.py` | 2,663 | First-party Core | Semantic tag inference, training corpus, definition audit & auto-fix | 100% Read | `test_sunpm_tags.py`, `test_coverage_wave28.py` |
| `DefFileGenerator/production_governor.py` | 190 | First-party Core | 5 Apollo Pillars verifier (`verify_file_pillars`) | 100% Read | `test_5pillars_strict_certification.py` |
| `DefFileGenerator/report_generator.py` | 592 | First-party Tooling | Single-file & directory Markdown/JSON reporting | 100% Read | `test_report_generator.py` |
| `DefFileGenerator/duplicate_detector.py` | 406 | First-party Tooling | Exact SHA-256, category conflicts, alias & family clustering | 100% Read | `test_duplicate_detector.py` |
| `DefFileGenerator/version_manager.py` | 672 | First-party Core | CAS version control (`definitions_history/`) | 100% Read | `test_definition_version_control.py` |
| `DefFileGenerator/main.py` | 1,938 | First-party Interface | CLI (`deffilegen`, 16 subcommands) & batch/metadata inference | 100% Read | `test_cli*.py`, `test_batch_cli.py` |
| `DefFileGenerator/gui.py` | 2,167 | First-party Interface | CustomTkinter/Tkinter Windows desktop application | 100% Read | `test_gui.py`, `test_interface_contracts.py` |
| `DefFileGenerator/dyness_generator.py` | 1,043 | First-party Domain | Dyness DH-Series C&I BESS parser & CSV generator | 100% Read | `test_dyness_definitions.py` |
| `DefFileGenerator/mars_generator.py` | 955 | First-party Domain | Mars Renewable Local Controller V7.0 generator | 100% Read | `test_mars_requested_definitions.py` |
| `DefFileGenerator/x7_generator.py` | 1,107 | First-party Domain | FoxESS X7 C&I BESS workbook parser & generator | 100% Read | `test_x7_definitions.py` |
| `DefFileGenerator/sunspec/__init__.py` | 87 | First-party SunSpec | SunSpec package public exports | 100% Read | `test_sunspec_engine.py` |
| `DefFileGenerator/sunspec/constants.py` | 97 | First-party SunSpec | SunSpec protocol constants, sentinels, model sets | 100% Read | `test_sunspec_engine.py` |
| `DefFileGenerator/sunspec/evaluator.py` | 627 | First-party SunSpec | Conditional SunSpec relevance gate & target planner | 100% Read | `test_sunspec_engine.py` |
| `DefFileGenerator/sunspec/generator.py` | 334 | First-party SunSpec | SunSpec definition compiler & CSV writer | 100% Read | `test_sunspec_engine.py` |
| `DefFileGenerator/sunspec/models_registry.py` | 791 | First-party SunSpec | SunSpec standard model point definitions | 100% Read | `test_sunspec_engine.py` |
| `DefFileGenerator/sunspec/pics_parser.py` | 437 | First-party SunSpec | SunSpec PICS XLSX/CSV/JSON parser | 100% Read | `test_sunspec_engine.py` |
| `DefFileGenerator/discovery/__init__.py` | 38 | First-party Discovery | Existing discovery scaffold exports | 100% Read | `test_discovery_*.py` |
| `DefFileGenerator/discovery/models.py` | 158 | First-party Discovery | Pydantic canonical intermediate & outcome models | 100% Read | `test_discovery_*.py` |
| `DefFileGenerator/discovery/identity_resolver.py` | 178 | First-party Discovery | Equipment query tokenizer & catalog matcher | 100% Read | `test_discovery_fetcher_and_identity.py` |
| `DefFileGenerator/discovery/source_discovery.py` | 132 | First-party Discovery | Curated manufacturer URL lookup + search stub | 100% Read | `test_discovery_pipeline_integration.py` |
| `DefFileGenerator/discovery/document_fetcher.py` | 238 | First-party Discovery | SSRF-guarded async HTTP document fetcher | 100% Read | `test_discovery_fetcher_and_identity.py` |
| `DefFileGenerator/discovery/document_extractor.py` | 273 | First-party Discovery | Bytes table extraction (PDF, XLSX, CSV, HTML) + OCR stub | 100% Read | `test_discovery_pipeline_integration.py` |
| `DefFileGenerator/discovery/register_normalizer.py` | 294 | First-party Discovery | Raw table -> `CanonicalEquipmentModel` normalizer | 100% Read | `test_discovery_pipeline_integration.py` |
| `DefFileGenerator/discovery/semantic_validator.py` | 105 | First-party Discovery | Pre-export canonical model invariant checker | 100% Read | `test_discovery_pipeline_integration.py` |
| `DefFileGenerator/discovery/definition_exporter.py` | 111 | First-party Discovery | Canonical model -> CSV target exporter | 100% Read | `test_discovery_pipeline_integration.py` |
| `DefFileGenerator/discovery/evidence_reporter.py` | 94 | First-party Discovery | Markdown provenance & SHA-256 evidence report | 100% Read | `test_discovery_pipeline_integration.py` |
| `DefFileGenerator/discovery/orchestrator.py` | 227 | First-party Discovery | Pipeline coordinator & manual document fallback | 100% Read | `test_discovery_pipeline_integration.py`, `test_cli_discovery.py` |
| `DefFileGenerator/abb_definitions_engine.py` | 307 | First-party Engine | 18 ABB/Fimer definition loader & mirror | 100% Read | `test_abb_definitions_suite.py` |
| `DefFileGenerator/carlogavazzi_definitions_engine.py` | 3,352 | First-party Engine | 19 Carlo Gavazzi meter generator & mirror | 100% Read | `test_carlogavazzi_definitions_suite.py` |
| `DefFileGenerator/classic_definitions_engine.py` | 5,472 | First-party Engine | 18 classic multi-vendor definitions | 100% Read | `test_classic_definitions_suite.py` |
| `DefFileGenerator/fronius_definitions_engine.py` | 273 | First-party Engine | 13 Fronius CSV loader & mirror | 100% Read | `test_fronius_definitions_suite.py` |
| `DefFileGenerator/goodwe_definitions_engine.py` | 4,578 | First-party Engine | 32 GoodWe ecosystem definitions | 100% Read | `test_goodwe_definitions_suite.py` |
| `DefFileGenerator/growatt_definitions_engine.py` | 239 | First-party Engine | 8 Growatt CSV loader & mirror | 100% Read | `test_growatt_definitions_suite.py` |
| `DefFileGenerator/huawei_definitions_engine.py` | 320 | First-party Engine | 15 Huawei CSV loader & mirror | 100% Read | `test_huawei_abb_battery.py` (CSV outputs) |
| `DefFileGenerator/kaco_definitions_engine.py` | 272 | First-party Engine | 10 KACO CSV loader & mirror | 100% Read | `test_kaco_definitions_suite.py` |
| `DefFileGenerator/schneider_definitions_engine.py` | 245 | First-party Engine | 9 Schneider CSV loader & mirror | 100% Read | `test_schneider_definitions_suite.py` |
| `DefFileGenerator/sma_definitions_engine.py` | 2,885 | First-party Engine | 28 SMA ecosystem definitions | 100% Read | `test_sma_definitions_suite.py` |
| `DefFileGenerator/socomec_definitions_engine.py` | 259 | First-party Engine | 11 Socomec CSV loader & mirror | 100% Read | `test_socomec_definitions_suite.py` |
| `DefFileGenerator/solaredge_definitions_engine.py` | 239 | First-party Engine | 8 SolarEdge CSV loader & mirror | 100% Read | `test_solaredge_definitions_suite.py` |
| `DefFileGenerator/solis_definitions_engine.py` | 217 | First-party Engine | 5 Solis CSV loader & mirror | 100% Read | `test_solis_definitions_suite.py` |
| `DefFileGenerator/sungrow_definitions_engine.py` | 281 | First-party Engine | 11 Sungrow definition loader & mirror | 100% Read | `test_sungrow_definitions_suite.py` |
| `web/app.py` | 1,740 | First-party Web | FastAPI 23-route REST API & upload middleware | 100% Read | `test_web*.py`, `test_interface_contracts.py` |
| `web/static/app.js` | 2,582 | First-party Frontend | 9-tab Web Studio SPA logic & modals | 100% Read | `interface_frontend_contracts.js` |
| `web/static/index.html` & `web/templates/index.html` | 1,225 | First-party Frontend | Web Studio HTML markup (duplicate pair) | 100% Read | `test_web_api.py`, `test_web_enhanced_features.py`, `scripts/verify_web_dom.py` |
| `web/static/styles.css` | 449 | First-party Frontend | Dark/light themes & layout styles | 100% Read | Served via `/static/styles.css` |
| `build_exe.py`, `build_exe.bat`, `launch_gui.bat`, `DefFileGenerator.spec` | 587 | First-party Packaging | Windows PyInstaller build & launcher scripts | 100% Read | `test_build_exe.py` |
| `doc_to_webdyn.py`, `generate_*.py` (5 root files) | 506 | First-party Wrappers | Root CLI entrypoints & legacy module wrappers | 100% Read | `test_cli.py`, `test_doc_to_webdyn.py`, `test_generate_webdyn_def.py` |
| `scripts/*.py` (15 files) + `bootstrap.*` + `README.md` | 5,200+ | First-party Scripts | Catalog build, production validation, source import, vendor CLIs | 100% Read | `test_production_suite.py`, `validate_production.py` |
| `production/` (1,194), `Definitions/`, `tagged_definitions/`, `generated_definitions/` | ~360k regs | Data / Catalog | Certified WebdynSunPM definitions & mirrors | Validated | `validate_production.py`, `test_production_suite.py`, `test_definitions_redundancy.py` |
| `sources/` + `sources/MANIFEST.json` | Corpus | Vendor Manuals | Manufacturer PDFs/XLSXs/HTMLs + SHA-256 manifest | Verified | `test_definitions_redundancy.py` |
| `DefFileGenerator/tests/` (100+ test files) | ~18k LOC | Verification | 4,701 collected pytest cases + Node DOM contract harness | Audited | `pytest`, `node interface_frontend_contracts.js` |

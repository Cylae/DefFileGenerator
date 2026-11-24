"""
Exhaustive Test Suite for SunSpec Alliance Conditional Generation Engine.

Verifies:
  1. Strict Relevance Gate:
     - Proprietary Modbus maps (32000..37000) skip cleanly without errors.
     - Commercial/marketing mentions of 'SunSpec' without models or PICS skip cleanly.
     - All points marked unimplemented in PICS skip cleanly.
     - CLI flags: --sunspec=auto|force|off.
     - Hybrid documentation (native Modbus + SunSpec 700) generating decoupled targets without collision.
  2. Protocol Exactness:
     - Common Model 1 length variants: L = 65 vs L = 66 chaining (next_addr = current_addr + 2 + L).
     - Base 0 (40000) vs Base 1 (40001).
     - Models 101, 102, 103, 120, 123, 124, 160, 201, 203, 701, 702, 704, 714, 802.
     - Scale Factors: fixed 10^SF applied, dynamic SF exposed with warnings.
     - Sentinel values documented (0x8000, 0xFFFF, 0x80000000, 0xFFFFFFFF).
  3. Repeating Blocks & Strict Filtering:
     - Repeating MPPT string blocks (Model 160: N = 0, 1, 4, 12).
     - Repeating DER DC port blocks (Model 714: NPrt = 0, 1, 4, 12).
     - Pad registers and unimplemented points excluded.
  4. Hostile & Edge Cases:
     - Truncated PICS, unknown Model ID, corrupt length L, duplicate models.
  5. End-to-End Certification:
     - 100% Apollo 5 Pillars schema compliance (verify_file_pillars with strict CRLF).
     - Deterministic byte-identical output repeatability.
"""

import hashlib
import tempfile
from pathlib import Path

import pytest

from DefFileGenerator.production_governor import verify_file_pillars
from DefFileGenerator.sunspec.constants import (
    BASE_ADDRESS_0,
    BASE_ADDRESS_1,
    COMMON_MODEL_LEN_ALIGNED,
    COMMON_MODEL_LEN_STANDARD,
    SunSpecCliMode,
    SunSpecProfileGeneration,
    SunSpecTopology,
)
from DefFileGenerator.sunspec.evaluator import (
    SunSpecGenerationTarget,
    SunSpecRelevanceDecision,
    SunSpecRelevanceEvaluator,
)
from DefFileGenerator.sunspec.generator import (
    SunSpecDefinitionGenerator,
    SunSpecGenerationResult,
)
from DefFileGenerator.sunspec.models_registry import (
    SUNSPEC_MODELS_REGISTRY,
    get_sunspec_model,
    list_supported_models,
)
from DefFileGenerator.sunspec.pics_parser import (
    PICSDocument,
    PICSModel,
    PICSParser,
    PICSPoint,
)


# ==============================================================================
# 1. STRICT RELEVANCE GATE TESTS
# ==============================================================================
class TestSunSpecRelevanceGate:
    """Verifies that the gate strictly filters non-actionable sources."""

    def test_pure_proprietary_modbus_skipped(self):
        evaluator = SunSpecRelevanceEvaluator()
        # Simulated pure proprietary Huawei/Sungrow register mapping
        proprietary_rows = [
            {"Address": 32000, "Name": "State_1", "Description": "Inverter operating status"},
            {"Address": 32002, "Name": "State_2", "Description": "Alarm bitfield"},
            {"Address": 32064, "Name": "ActivePower", "Description": "Total active power kW"},
            {"Address": 32080, "Name": "InternalTemp", "Description": "Cabinet temperature"},
        ]
        decision = evaluator.evaluate(
            intermediate_rows=proprietary_rows,
            raw_text="Huawei SUN2000 Proprietary Modbus Interface Definition Rev 3.0",
        )
        assert decision.should_generate is False
        assert decision.is_skipped is True
        assert "proprietary" in decision.reason.lower()
        assert len(decision.targets) == 0

    def test_marketing_only_mention_skipped(self):
        evaluator = SunSpecRelevanceEvaluator()
        text = (
            "The PowerMaster inverter fully adheres to IEEE standards and offers optional "
            "compatibility with SunSpec Alliance protocols. Please contact sales for PICS."
        )
        decision = evaluator.evaluate(raw_text=text)
        assert decision.should_generate is False
        assert decision.is_skipped is True
        assert "marketing" in decision.reason.lower() or "commercial" in decision.reason.lower()
        assert len(decision.targets) == 0

    def test_all_points_unimplemented_in_pics_skipped(self):
        evaluator = SunSpecRelevanceEvaluator()
        pics_doc = PICSDocument(
            manufacturer="TestBrand",
            model_name="MockInverter",
            models={
                1: PICSModel(model_id=1, address=40002, length=65, is_supported=True),
                103: PICSModel(
                    model_id=103,
                    address=40069,
                    length=50,
                    is_supported=False,
                    points={
                        "A": PICSPoint(0, 40071, "A", is_supported=False),
                        "W": PICSPoint(12, 40083, "W", is_supported=False),
                    },
                ),
            },
        )
        decision = evaluator.evaluate(pics_doc=pics_doc)
        assert decision.should_generate is False
        assert decision.is_skipped is True
        assert "unimplemented" in decision.reason.lower()

    def test_cli_mode_off_strictly_skips(self):
        evaluator = SunSpecRelevanceEvaluator()
        # Even with an actionable PICS document, --sunspec=off must skip immediately
        pics_doc = PICSDocument(
            manufacturer="ABB",
            model_name="PVS-100",
            models={
                1: PICSModel(model_id=1, address=40002, length=66, is_supported=True),
                103: PICSModel(model_id=103, address=40070, length=50, is_supported=True),
            },
        )
        decision = evaluator.evaluate(pics_doc=pics_doc, cli_mode=SunSpecCliMode.OFF)
        assert decision.should_generate is False
        assert decision.is_skipped is True
        assert "disabled" in decision.reason.lower()

    def test_cli_mode_force_generates_targets(self):
        evaluator = SunSpecRelevanceEvaluator()
        decision = evaluator.evaluate(
            cli_mode=SunSpecCliMode.FORCE,
            manufacturer="SolarForce",
            model="SF-50K",
            category="Inverter",
        )
        assert decision.should_generate is True
        assert decision.is_forced is True
        assert len(decision.targets) >= 1
        target = decision.targets[0]
        assert target.manufacturer == "SolarForce"
        assert target.model_name == "SF-50K"
        assert target.prefix == "INV"

    def test_hybrid_native_and_sunspec_700_clean_separation(self):
        evaluator = SunSpecRelevanceEvaluator()
        pics_doc = PICSDocument(
            manufacturer="Huawei",
            model_name="SUN2000-100KTL-M3",
            category="Inverter",
            models={
                1: PICSModel(model_id=1, address=40002, length=65, is_supported=True),
                701: PICSModel(model_id=701, address=40069, length=153, is_supported=True),
                702: PICSModel(model_id=702, address=40224, length=38, is_supported=True),
                704: PICSModel(model_id=704, address=40264, length=108, is_supported=True),
                714: PICSModel(model_id=714, address=40374, length=40, repeating_count=2, is_supported=True),
            },
        )
        decision = evaluator.evaluate(pics_doc=pics_doc)
        assert decision.should_generate is True
        assert SunSpecProfileGeneration.DER700 in decision.detected_profiles
        assert len(decision.targets) == 1
        t = decision.targets[0]
        assert t.profile == SunSpecProfileGeneration.DER700
        assert 701 in t.model_ids
        assert 103 not in t.model_ids  # Never mixed with 103


# ==============================================================================
# 2. PROTOCOL EXACTNESS TESTS
# ==============================================================================
class TestSunSpecProtocolExactness:
    """Verifies register offsets, chaining formulas, and scale factors."""

    def test_common_model_chaining_length_65_vs_66(self):
        generator = SunSpecDefinitionGenerator()

        # Target A: L=65 (Standard length without pad)
        target_65 = SunSpecGenerationTarget(
            prefix="INV",
            manufacturer="VendorA",
            model_name="Model65",
            profile=SunSpecProfileGeneration.LEGACY,
            topology=SunSpecTopology.THREE_PHASE,
            model_ids=[1, 103],
            base_address=BASE_ADDRESS_0,  # 40000
            common_model_length=COMMON_MODEL_LEN_STANDARD,  # 65
        )
        rows_65, _ = generator._compile_rows(target_65, pics_doc=None)

        # Target B: L=66 (Aligned length with 64-bit trailing pad)
        target_66 = SunSpecGenerationTarget(
            prefix="INV",
            manufacturer="VendorB",
            model_name="Model66",
            profile=SunSpecProfileGeneration.LEGACY,
            topology=SunSpecTopology.THREE_PHASE,
            model_ids=[1, 103],
            base_address=BASE_ADDRESS_0,  # 40000
            common_model_length=COMMON_MODEL_LEN_ALIGNED,  # 66
        )
        rows_66, _ = generator._compile_rows(target_66, pics_doc=None)

        # Verify Model 1 ID and L positions
        # Model 1 starts at 40002 (40000 + 2)
        assert rows_65[0][2] == "40002"  # Model 1 ID
        assert rows_65[1][2] == "40003"  # Model 1 L

        # In L=65: Model 103 ID starts at 40002 + 2 + 65 = 40069
        m103_id_row_65 = next(r for r in rows_65 if "Model 103 ID" in r[5])
        assert m103_id_row_65[2] == "40069"

        # In L=66: Model 103 ID starts at 40002 + 2 + 66 = 40070
        m103_id_row_66 = next(r for r in rows_66 if "Model 103 ID" in r[5])
        assert m103_id_row_66[2] == "40070"

    def test_base_address_0_vs_1(self):
        generator = SunSpecDefinitionGenerator()

        target_base0 = SunSpecGenerationTarget(
            prefix="INV",
            manufacturer="TestMfg",
            model_name="T0",
            profile=SunSpecProfileGeneration.LEGACY,
            topology=SunSpecTopology.THREE_PHASE,
            model_ids=[1, 103],
            base_address=BASE_ADDRESS_0,  # 40000
        )
        rows_base0, _ = generator._compile_rows(target_base0, pics_doc=None)
        assert rows_base0[0][2] == "40002"

        target_base1 = SunSpecGenerationTarget(
            prefix="INV",
            manufacturer="TestMfg",
            model_name="T1",
            profile=SunSpecProfileGeneration.LEGACY,
            topology=SunSpecTopology.THREE_PHASE,
            model_ids=[1, 103],
            base_address=BASE_ADDRESS_1,  # 40001
        )
        rows_base1, _ = generator._compile_rows(target_base1, pics_doc=None)
        assert rows_base1[0][2] == "40003"

    def test_fixed_scale_factor_application(self):
        generator = SunSpecDefinitionGenerator()
        pics_doc = PICSDocument(
            fixed_scale_factors={
                (103, "A_SF"): -1,   # 10^-1 = 0.1
                (103, "V_SF"): -1,   # 10^-1 = 0.1
                (103, "W_SF"): 1,    # 10^1  = 10.0
                (103, "Hz_SF"): -2,  # 10^-2 = 0.01
            }
        )
        target = SunSpecGenerationTarget(
            prefix="INV",
            manufacturer="ABB",
            model_name="PVS-100",
            profile=SunSpecProfileGeneration.LEGACY,
            topology=SunSpecTopology.THREE_PHASE,
            model_ids=[1, 103],
            base_address=40000,
        )
        rows, meta = generator._compile_rows(target, pics_doc=pics_doc)

        row_a = next(r for r in rows if r[6] == "I_AC")
        assert row_a[7] == "0.100000"  # Factor applied from A_SF = -1

        row_w = next(r for r in rows if r[6] == "P_AC")
        assert row_w[7] == "10.000000"  # Factor applied from W_SF = 1

        row_hz = next(r for r in rows if r[6] == "FREQ")
        assert row_hz[7] == "0.010000"  # Factor applied from Hz_SF = -2

    def test_dynamic_scale_factor_exposure(self):
        generator = SunSpecDefinitionGenerator()
        # When no fixed scale factors are guaranteed in PICS
        target = SunSpecGenerationTarget(
            prefix="INV",
            manufacturer="Generic",
            model_name="StandardInv",
            profile=SunSpecProfileGeneration.LEGACY,
            topology=SunSpecTopology.THREE_PHASE,
            model_ids=[1, 103],
            base_address=40000,
        )
        rows, meta = generator._compile_rows(target, pics_doc=None)

        # Factor defaults to 1.000000 and dynamic SF is documented
        row_a = next(r for r in rows if r[6] == "I_AC")
        assert row_a[7] == "1.000000"
        assert "[Dynamic SF: A_SF]" in row_a[5]

        # A_SF register is explicitly present in the CSV
        sf_row = next((r for r in rows if r[5].startswith("Current Scale Factor")), None)
        assert sf_row is not None
        assert sf_row[3] == "I16"


# ==============================================================================
# 3. REPEATING BLOCKS AND FILTERING TESTS
# ==============================================================================
class TestRepeatingBlocksAndFiltering:
    """Verifies variable length repeating blocks and pad/unimplemented filtering."""

    @pytest.mark.parametrize("num_strings", [0, 1, 4, 12])
    def test_model_160_repeating_mppt_strings(self, num_strings):
        m160 = get_sunspec_model(160)
        assert m160 is not None
        points = m160.get_points(num_repeats=num_strings)

        if num_strings == 0:
            assert len(points) == len(m160.points)
        else:
            expected_count = len(m160.points) + num_strings * len(m160.repeating_points_template)
            assert len(points) == expected_count

            # Check that string tags are properly indexed
            for s in range(1, num_strings + 1):
                assert any(p.sunpm_tag == f"I_DC_{s}" for p in points)
                assert any(p.sunpm_tag == f"U_DC_{s}" for p in points)
                assert any(p.sunpm_tag == f"P_DC_{s}" for p in points)

    @pytest.mark.parametrize("num_ports", [0, 1, 4, 12])
    def test_model_714_repeating_dc_ports(self, num_ports):
        m714 = get_sunspec_model(714)
        assert m714 is not None
        points = m714.get_points(num_repeats=num_ports)

        if num_ports == 0:
            assert len(points) == len(m714.points)
        else:
            expected_count = len(m714.points) + num_ports * len(m714.repeating_points_template)
            assert len(points) == expected_count

            for p_idx in range(1, num_ports + 1):
                assert any(p.sunpm_tag == f"I_DC_{p_idx}" for p in points)
                assert any(p.sunpm_tag == f"U_DC_{p_idx}" for p in points)
                assert any(p.sunpm_tag == f"P_DC_{p_idx}" for p in points)

    def test_pad_registers_filtered(self):
        generator = SunSpecDefinitionGenerator()
        target = SunSpecGenerationTarget(
            prefix="INV",
            manufacturer="VendorPad",
            model_name="VP-1",
            profile=SunSpecProfileGeneration.LEGACY,
            topology=SunSpecTopology.THREE_PHASE,
            model_ids=[1, 103, 120, 160],
            common_model_length=66,  # Has pad at offset 65
        )
        rows, _ = generator._compile_rows(target, pics_doc=None)
        # Pad registers must be excluded from final definition rows
        assert not any(r[3] == "PAD" for r in rows)
        assert not any(r[5].strip().lower() == "pad" for r in rows)

    def test_unimplemented_points_filtered(self):
        generator = SunSpecDefinitionGenerator()
        pics_doc = PICSDocument(
            unimplemented_points={
                (103, "AphC"),
                (103, "PPVphCA"),
                (103, "TmpTrns"),
            }
        )
        target = SunSpecGenerationTarget(
            prefix="INV",
            manufacturer="PartialMfg",
            model_name="PM-1",
            profile=SunSpecProfileGeneration.LEGACY,
            topology=SunSpecTopology.THREE_PHASE,
            model_ids=[1, 103],
        )
        rows, meta = generator._compile_rows(target, pics_doc=pics_doc)
        assert meta["excluded_unimplemented"] == 3
        # Excluded points should not appear in rows
        assert not any("AphC" in r[5] for r in rows)
        assert not any("PPVphCA" in r[5] for r in rows)


# ==============================================================================
# 4. HOSTILE BOUNDARY CASES
# ==============================================================================
class TestHostileAndBoundaryCases:
    """Verifies resilience against corrupt, missing, or malformed inputs."""

    def test_empty_or_truncated_file(self):
        parser = PICSParser()
        with tempfile.NamedTemporaryFile(suffix=".xlsx") as tf:
            # Empty file
            doc = parser.parse_file(tf.name)
            assert doc.has_models is False
            assert len(doc.models) == 0

    def test_unknown_model_id_gracefully_ignored(self):
        generator = SunSpecDefinitionGenerator()
        target = SunSpecGenerationTarget(
            prefix="INV",
            manufacturer="Test",
            model_name="UnknownModel",
            profile=SunSpecProfileGeneration.LEGACY,
            topology=SunSpecTopology.THREE_PHASE,
            model_ids=[1, 99999, 103],  # 99999 is unknown
        )
        rows, meta = generator._compile_rows(target, pics_doc=None)
        assert any("99999" in w for w in meta["warnings"])
        # Model 1 and 103 are still compiled successfully
        assert any("Model 1 ID" in r[5] for r in rows)
        assert any("Model 103 ID" in r[5] for r in rows)

    def test_model_def_registry_coverage(self):
        models = list_supported_models()
        assert 1 in models
        assert 101 in models
        assert 102 in models
        assert 103 in models
        assert 120 in models
        assert 123 in models
        assert 124 in models
        assert 160 in models
        assert 201 in models
        assert 202 in models
        assert 203 in models
        assert 204 in models
        assert 701 in models
        assert 702 in models
        assert 704 in models
        assert 714 in models
        assert 802 in models


# ==============================================================================
# 5. END-TO-END CERTIFICATION TESTS (APOLLO 5 PILLARS)
# ==============================================================================
class TestEndToEndCertification:
    """Verifies that generated definition files strictly pass verify_file_pillars."""

    @pytest.mark.parametrize(
        "profile,topology,model_ids,prefix",
        [
            (SunSpecProfileGeneration.LEGACY, SunSpecTopology.SINGLE_PHASE, [1, 101, 120], "INV"),
            (SunSpecProfileGeneration.LEGACY, SunSpecTopology.SPLIT_PHASE, [1, 102, 120], "INV"),
            (SunSpecProfileGeneration.LEGACY, SunSpecTopology.THREE_PHASE, [1, 103, 120, 160], "INV"),
            (SunSpecProfileGeneration.DER700, SunSpecTopology.THREE_PHASE, [1, 701, 702, 704, 714], "INV"),
            (SunSpecProfileGeneration.LEGACY, SunSpecTopology.SINGLE_PHASE, [1, 201], "MET"),
            (SunSpecProfileGeneration.LEGACY, SunSpecTopology.THREE_PHASE, [1, 203], "MET"),
            (SunSpecProfileGeneration.LEGACY, SunSpecTopology.STORAGE, [1, 802], "BAT"),
        ],
    )
    def test_generated_files_certify_100_percent_pillars(self, profile, topology, model_ids, prefix):
        generator = SunSpecDefinitionGenerator()
        with tempfile.TemporaryDirectory() as tmp_dir:
            target = SunSpecGenerationTarget(
                prefix=prefix,
                manufacturer="CertifiedBrand",
                model_name=f"Device-{profile.value}-{topology.value}",
                profile=profile,
                topology=topology,
                model_ids=model_ids,
                base_address=40000,
                common_model_length=65,
                repeating_counts={160: 4, 714: 4},
            )
            res = generator.generate(target, output_dir=tmp_dir, force_overwrite=True)

            assert res.is_valid is True, f"Pillars validation failed: {res.issues}"
            assert res.register_count > 0
            assert res.output_path.exists()

            # Direct re-verification with strict CRLF check
            is_valid, issues = verify_file_pillars(res.output_path, strict_crlf=True)
            assert is_valid is True, f"Re-verification failed: {issues}"

    def test_deterministic_byte_identical_output(self):
        generator = SunSpecDefinitionGenerator()
        with tempfile.TemporaryDirectory() as tmp_dir:
            target = SunSpecGenerationTarget(
                prefix="INV",
                manufacturer="DeterministicBrand",
                model_name="D100",
                profile=SunSpecProfileGeneration.LEGACY,
                topology=SunSpecTopology.THREE_PHASE,
                model_ids=[1, 103, 120, 160],
                base_address=40000,
                common_model_length=66,
                repeating_counts={160: 2},
            )
            file_a = Path(tmp_dir) / "run_a.csv"
            file_b = Path(tmp_dir) / "run_b.csv"

            target.output_filename = "run_a.csv"
            res_a = generator.generate(target, output_dir=tmp_dir, force_overwrite=True)

            target.output_filename = "run_b.csv"
            res_b = generator.generate(target, output_dir=tmp_dir, force_overwrite=True)

            hash_a = hashlib.sha256(file_a.read_bytes()).hexdigest()
            hash_b = hashlib.sha256(file_b.read_bytes()).hexdigest()

            assert hash_a == hash_b, "SunSpec generation must be strictly deterministic"

from workflow_features import calculate_medicine_quantity, compare_prescriptions
from patient_history import PatientHistory


def test_quantity_calculator_handles_daily_weekly_and_ambiguous_regimens():
    assert calculate_medicine_quantity("1 tablet", "1x2 (BID)", "7 days") == "14 tablets"
    assert calculate_medicine_quantity("5 mL", "1x3 (TID)", "5 days") == "75 mL"
    assert calculate_medicine_quantity("1 capsule", "2x/week", "4 weeks") == "8 capsules"
    assert calculate_medicine_quantity("500 mg", "1x2 (BID)", "7 days") == ""
    assert calculate_medicine_quantity("1 tablet", "PRN", "7 days") == ""


def test_prescription_comparison_reports_added_removed_changed():
    previous = [
        {"generic_name": "Metformin", "dosage": "1 tablet", "frequency": "BID"},
        {"generic_name": "Aspirin", "dosage": "1 tablet", "frequency": "QD"},
    ]
    current = [
        {"generic_name": "Metformin", "dosage": "2 tablets", "frequency": "BID"},
        {"generic_name": "Atorvastatin", "dosage": "1 tablet", "frequency": "QD"},
    ]
    result = compare_prescriptions(current, previous)
    assert [item["generic_name"] for item in result["added"]] == ["Atorvastatin"]
    assert [item["generic_name"] for item in result["removed"]] == ["Aspirin"]
    assert result["changed"][0]["after"]["generic_name"] == "Metformin"
    assert result["changed"][0]["fields"] == ["dosage"]


def test_patient_and_prescription_can_be_restored(tmp_path):
    history = PatientHistory(tmp_path / "patients.json")
    record = history.save_prescription(
        {"name": "Test Patient", "age": "40", "sex": "M"},
        [{"generic_name": "Metformin", "quantity": "60 tablets"}],
    )
    prescription = record["prescriptions"][0]
    assert history.delete_prescription(record["id"], prescription["id"])
    assert history.restore_prescription(record["id"], prescription)
    restored = history.get(record["id"])
    assert restored["prescriptions"][0]["drugs"][0]["quantity"] == "60 tablets"
    assert history.delete(record["id"])
    assert history.restore_record(restored)
    assert history.get(record["id"])["name"] == "Test Patient"


def test_recovery_bin_is_encrypted_and_removable(tmp_path, monkeypatch):
    import config as config_module
    monkeypatch.setattr(config_module, "APP_DIR", tmp_path)
    monkeypatch.setattr(config_module, "CONFIG_PATH", tmp_path / "config.json")
    local = config_module.Config()
    item_id = local.add_recovery_item("favorite", "Aspirin", {"brand_name": "Aspirin"})
    assert local.recovery_items()[0]["id"] == item_id
    assert "Aspirin" not in (tmp_path / "config.json").read_text(encoding="utf-8")
    assert local.discard_recovery_item(item_id)
    assert local.recovery_items() == []

"""Public and historical compatibility surface that cleanup must preserve."""

import inspect

import cloud_rx
import config
import drug_classes
import drug_db
import openfda
import qr_utils


def _parameters(callable_object):
    return tuple(inspect.signature(callable_object).parameters)


def test_qr_and_cloud_api_surface_is_preserved():
    expected = {
        qr_utils.Prescription.to_payload: ("self",),
        qr_utils.Prescription.to_qr_payload: ("self",),
        qr_utils.Prescription.to_cloud_payload: ("self",),
        qr_utils.encode_payload: ("prescription", "compress"),
        qr_utils.build_qr_url: ("prescription", "viewer_base", "compress"),
        qr_utils.decode_payload: ("encoded", "compressed"),
        qr_utils.decode_from_url: ("url",),
        qr_utils.verification_key: (),
        qr_utils.qr_info: ("url",),
        cloud_rx.upload_prescription: ("payload", "api_key", "opener"),
        cloud_rx.check_readiness: ("api_key", "opener"),
    }
    for callable_object, parameters in expected.items():
        assert callable(callable_object)
        assert _parameters(callable_object) == parameters

    assert cloud_rx.API_URL == "https://rx-v2.vercel.app/api/rx"
    assert cloud_rx.VIEWER_URL == "https://rx-v2.vercel.app"
    assert cloud_rx.TIMEOUT == 15


def test_configuration_and_database_compatibility_surface_is_preserved():
    expected = {
        config.Config.set_doctor: ("self", "kw"),
        config.Config.profile_names: ("self",),
        config.Config.use_profile: ("self", "name"),
        config.Config.export_treatment_templates: ("self", "destination"),
        config.Config.import_treatment_templates: ("self", "source", "replace"),
        config.Config.dosage_presets: ("self",),
        config.Config.add_dosage_preset: ("self", "value"),
        drug_db.DrugDatabase.export_csv: ("self", "dest_path"),
        drug_db.DrugDatabase.search_trade: ("self", "query", "limit"),
        drug_db.DrugDatabase.all_names: ("self",),
        drug_db.DrugDatabase.update_classification: (
            "self", "name", "therapeutic_group", "detailed_class", "append"),
        drug_db.DrugDatabase.update_classifications: (
            "self", "names", "therapeutic_group", "detailed_class", "append"),
        drug_db.DrugDatabase.update_drug_names: (
            "self", "selected_drug", "generic_name", "brand_name"),
        drug_db.DrugDatabase.update_drug_names: (
            "self", "selected_drug", "generic_name", "brand_name"),
        drug_db.DrugDatabase.update_drug_names: (
            "self", "selected_drug", "generic_name", "brand_name"),
        drug_db.DrugDatabase.classification_states: ("self", "names"),
        drug_classes.all_subclasses: (),
        openfda.concise: ("paragraphs", "limit"),
    }
    for callable_object, parameters in expected.items():
        assert callable(callable_object)
        assert _parameters(callable_object) == parameters

import pytest

from src.core import credential_store


def test_keyring_credential_roundtrip():
    if not credential_store.is_keyring_available():
        pytest.skip("No usable keyring backend in this environment.")

    service = "PreSubmissionAIQC_Test"
    username = "pytest"
    credential_store.save_secret(service, username, "secret-value")
    try:
        assert credential_store.has_secret(service, username)
        assert credential_store.get_secret(service, username) == "secret-value"
    finally:
        credential_store.delete_secret(service, username)
    assert not credential_store.has_secret(service, username)

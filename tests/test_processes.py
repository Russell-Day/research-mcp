import os

from research_mcp.processes import process_creation_time, process_identity_matches


def test_process_identity_uses_pid_and_creation_time():
    created = process_creation_time(os.getpid())
    assert created is not None
    assert process_identity_matches(os.getpid(), created)
    assert not process_identity_matches(os.getpid(), created + 1)

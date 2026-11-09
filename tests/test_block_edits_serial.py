"""Edits on drafts are serializable (round 45, lane B): ``block_stress.run`` -- callers in
threads editing one made-up report at once, their own chapters and each other's, every
admitted edit prepared on a draft -- replayed one call at a time in the order the changes were
committed: every result, the bytes and the id state the same.  Small here; long runs:
``python tests/block_stress.py --seeds 0-99 --threads 8 --calls 16``."""

from __future__ import annotations

import pytest

import block_stress

pytestmark = pytest.mark.any_faces


@pytest.mark.parametrize("seed, first_fill", [(0, False), (1, True)])
def test_concurrent_edits_equal_their_commit_order_run_alone(seed, first_fill):
    found = block_stress.run(seed, threads=4, calls=8, chapters=3, first_fill=first_fill)
    assert found["same"], found
    assert found["drafted"] > 0


@pytest.mark.slow
@pytest.mark.parametrize("seed", range(2, 8))
def test_many_callers_long(seed):
    found = block_stress.run(seed, threads=8, calls=16, chapters=4)
    assert found["same"], found

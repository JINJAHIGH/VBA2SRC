"""日本語名で使える、独立したテスト用リソース。"""

import tempfile
from pathlib import Path

import pytest


@pytest.fixture
def 一時フォルダ():
    with tempfile.TemporaryDirectory(prefix="vba2src_test_") as 保存先:
        yield Path(保存先).resolve()


@pytest.fixture
def 差替え():
    with pytest.MonkeyPatch.context() as 差替え管理:
        yield 差替え管理

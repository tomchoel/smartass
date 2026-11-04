# engine/tests/v2/test_v2_config_paths.py
from smartass.core import config


def test_paths_computed_from_the_package_point_at_real_files():
    assert config.ENGINE_DIR.name == "engine" and (config.ENGINE_DIR / "smartass").is_dir()
    assert (config.REPO_DIR / "knowledge").is_dir()
    assert config.DEFAULTS_FILE.exists() and config.DEFAULTS_FILE.name == "defaults.toml"
    assert config.BAS_FILES and all(p.exists() for p in config.BAS_FILES)

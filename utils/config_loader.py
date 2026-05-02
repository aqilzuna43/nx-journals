"""Built-in JSON config loading for NX journals."""

import json
import os


def load_json_config(repo_root, relative_path):
    """Load a JSON config file located under the repo root."""
    config_path = os.path.join(repo_root, relative_path)
    with open(config_path, encoding="utf-8") as fh:
        return json.load(fh)

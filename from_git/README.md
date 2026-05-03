# NX2312 Runtime Folder

This folder is the deployable NX journal payload.

Keep these folders together:

```text
from_git/
  config/
  journals/
  templates/
  utils/
```

Run journals from `from_git/journals` in NX 2312. Do not copy a single `.py`
file without `utils` and `config`, because the journals import shared helpers
and read JSON config at runtime.

No third-party Python packages are required.

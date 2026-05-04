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

Run journals from `from_git/journals` in NX 2312.

`05_bulk_attribute_updater.py` is intentionally self-contained to avoid NX2312
package/import path problems. It still reads `config/attribute_mapping.json`,
so keep the `config` folder beside `journals`.

The other journals still use shared helpers from `utils`, so keep the full
folder together if you run J01-J04.

No third-party Python packages are required.

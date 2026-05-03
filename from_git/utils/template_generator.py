"""Shared report column definitions."""

MASTER_COLUMNS = [
    "LEVEL",
    "PART_NUMBER",
    "DESCRIPTION",
    "REVISION",
    "DRAWING_NUMBER",
    "MATERIAL",
    "FINISH",
    "QUANTITY",
    "STATUS",
    "NOTES",
]


if __name__ == "__main__":
    print(",".join(MASTER_COLUMNS))

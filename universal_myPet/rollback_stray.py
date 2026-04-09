from _config import ACT_COLLECTION, TARGET_COLLECTION
from rollback import run_rollback


if __name__ == "__main__":
    raise SystemExit(run_rollback(collections_filter={TARGET_COLLECTION, ACT_COLLECTION}))

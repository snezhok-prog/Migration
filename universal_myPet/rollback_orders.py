from _config import ORDER_COLLECTION
from rollback import run_rollback


if __name__ == "__main__":
    raise SystemExit(run_rollback(collections_filter={ORDER_COLLECTION}))

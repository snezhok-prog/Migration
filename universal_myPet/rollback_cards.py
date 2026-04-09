from _config import CARD_COLLECTION, RELEASE_COLLECTION, TRANSFER_ACT_COLLECTION
from rollback import run_rollback


if __name__ == "__main__":
    raise SystemExit(run_rollback(collections_filter={CARD_COLLECTION, RELEASE_COLLECTION, TRANSFER_ACT_COLLECTION}))

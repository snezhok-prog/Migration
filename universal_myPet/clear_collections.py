import argparse
import time
import warnings

from urllib3.exceptions import InsecureRequestWarning

from _api import delete_from_collection, search_collection, set_runtime_urls, setup_session
from _config import (
    ACT_COLLECTION,
    BASE_URL,
    CARD_COLLECTION,
    JWT_URL,
    ORDER_COLLECTION,
    RELEASE_COLLECTION,
    TARGET_COLLECTION,
    TRANSFER_ACT_COLLECTION,
)
from _logger import setup_logger
from _profiles import PROFILES


DEFAULT_COLLECTIONS = [
    ORDER_COLLECTION,
    TARGET_COLLECTION,
    CARD_COLLECTION,
    ACT_COLLECTION,
    RELEASE_COLLECTION,
    TRANSFER_ACT_COLLECTION,
]


def _split_items(raw: str):
    out = []
    for part in str(raw or "").replace(";", ",").replace("\n", ",").split(","):
        token = part.strip()
        if token:
            out.append(token)
    return out


def _setup_runtime_profile(args):
    profile_name = str(args.profile or "custom").strip().lower()
    base_url = BASE_URL
    jwt_url = JWT_URL

    if profile_name in PROFILES and profile_name != "custom":
        profile = PROFILES[profile_name]
        base_url = profile.base_url
        jwt_url = profile.jwt_url

    if str(args.base_url or "").strip():
        base_url = str(args.base_url).strip()
    if str(args.jwt_url or "").strip():
        jwt_url = str(args.jwt_url).strip()
    if not jwt_url:
        jwt_url = base_url.rstrip("/") + "/jwt/"
    set_runtime_urls(base_url=base_url, jwt_url=jwt_url)
    return profile_name, base_url


def _search_page(session, logger, collection: str, page: int, size: int):
    body = {"page": int(page), "size": int(size)}
    data = search_collection(session, logger, collection, body)
    return data if isinstance(data, dict) else {}


def parse_args():
    parser = argparse.ArgumentParser(description="Clear pet migration collections in PGS.")
    parser.add_argument("--profile", choices=["custom", "dev", "psi", "prod"], default="psi")
    parser.add_argument("--base-url", default="", help="Override stand base URL.")
    parser.add_argument("--jwt-url", default="", help="Override JWT page URL.")
    parser.add_argument("--collections", default="", help="Collections list (separator: ',', ';' or new line).")
    parser.add_argument("--page-size", type=int, default=100)
    parser.add_argument("--delete-retries", type=int, default=3)
    parser.add_argument("--retry-backoff-sec", type=float, default=0.4)
    parser.add_argument("--dry-run", action="store_true")
    parser.add_argument("--no-prompt", action="store_true", help="Read cookie/token only from files.")
    parser.add_argument("--operator-mode", action="store_true", help="Allow interactive re-login prompts on auth errors.")
    return parser.parse_args()


def main():
    warnings.filterwarnings("ignore", category=InsecureRequestWarning)
    args = parse_args()
    profile_name, base_url = _setup_runtime_profile(args)
    collections = _split_items(args.collections) or list(DEFAULT_COLLECTIONS)

    logger = setup_logger()
    logger.info("Starting clear collections")
    logger.info("Profile=%s base_url=%s", profile_name, base_url)
    logger.info("Collections=%s", collections)
    logger.info("Dry run=%s", args.dry_run)

    session = setup_session(logger, no_prompt=args.no_prompt, operator_mode=args.operator_mode)
    if not session:
        logger.error("Authorization failed")
        return 1

    total_deleted = 0
    total_failed = 0
    page_size = max(1, int(args.page_size))
    retries = max(1, int(args.delete_retries))
    backoff = max(0.0, float(args.retry_backoff_sec))

    for collection in collections:
        logger.info("=== PURGE %s ===", collection)
        page = 0
        to_delete = []
        while True:
            data = _search_page(session, logger, collection, page, page_size)
            content = data.get("content") if isinstance(data, dict) else []
            if not isinstance(content, list) or not content:
                break
            for item in content:
                main_id = item.get("_id")
                guid = item.get("guid") or ""
                if main_id:
                    to_delete.append((str(main_id), str(guid)))
            if data.get("last") or len(content) < page_size:
                break
            page += 1

        logger.info("[%s] found=%s", collection, len(to_delete))
        for main_id, guid in to_delete:
            if args.dry_run:
                logger.info("[DRY][%s] delete _id=%s guid=%s", collection, main_id, guid)
                continue

            success = False
            for attempt in range(1, retries + 1):
                try:
                    deleted = delete_from_collection(
                        session,
                        logger,
                        {"_id": main_id, "guid": guid, "parentEntries": collection},
                    )
                except Exception as exc:
                    deleted = False
                    logger.warning(
                        "[%s] delete error _id=%s attempt=%s/%s err=%s",
                        collection,
                        main_id,
                        attempt,
                        retries,
                        exc,
                    )
                if deleted:
                    total_deleted += 1
                    success = True
                    break
                if attempt < retries:
                    time.sleep(backoff * attempt)
            if not success:
                total_failed += 1
                logger.warning("[%s] delete failed _id=%s guid=%s", collection, main_id, guid)

    logger.info("Clear done. deleted=%s failed=%s", total_deleted, total_failed)
    return 0 if total_failed == 0 else 1


if __name__ == "__main__":
    raise SystemExit(main())

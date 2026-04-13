# NTO Mesto Migration - Integration Guide

## Refactored Architecture

The `nto_mesto_migration.py` script has been refactored to use shared helper modules following the same pattern as `migration.py`.

### Session-Centric Approach

```python
# Old approach (duplicated in each function):
def some_function(token=None, authorization=None, cookie=None):
    # Recreate session/headers each time
    # Manual auth handling

# New approach (session passed down):
def some_function(session, ...):
    # Reuse same session
    # Consistent auth handling
```

## Key Files

### 1. **nto_mesto_migration.py** (Refactored)
- Migration script for NTO (места размещения) registry
- Converts data from JSON input to API records
- **Entry point**: `python nto_mesto_migration.py --input data.json --base-url https://api.example.com --token YOUR_TOKEN`

### 2. **_api.py** (Shared Infrastructure)
Functions now used by NTO migration:
- `setup_session(logger)` - Interactive session creation
- `api_request(session, logger, method, url, **kwargs)` - HTTP with retries
- `upload_file(session, logger, file_path, entry_name, entry_id, entity_field_path)` - File upload

### 3. **_utils.py** (Shared Utilities)
Functions used by NTO migration:
- `jsonable(obj)` - JSON serialization
- `generate_guid()` - GUID generation
- `nz(v)` - Null-to-empty-string for pandas compatibility
- `split_sc(v)` - Split semicolon-delimited values

### 4. **_logger.py** (Logging)
- `setup_logger()` - Console logging for migration

### 5. **_config.py** (Configuration)
- `BASE_URL` - API endpoint
- Constants and configuration values

## Function Call Flow

### Migration Execution

```
main()
  └─> parse arguments
      └─> create_session()
         └─> setup_session() [from _api]  [interactive prompt]
      └─> run_migration()
          └─> For each input record:
              ├─> transform_row_to_registry(row, session)
              │   └─> build_org_info(info, session)
              │       └─> search_org_by_ogrn(ogrn, session)
              │           └─> call_api(session, ..., ...)
              │               └─> api_request(session, logger, ...) [from _api]
              ├─> create_record(session, TARGET_COLLECTION, payload)
              │   └─> call_api(session, "POST", ...)
              │       └─> api_request(session, logger, ...) [from _api]
              ├─> collect_pending_uploads_from_branches(row)
              ├─> upload_one_file(session, entry_id, entity_field_path, ...)
              │   └─> upload_file(session, logger, ...) [from _api]
              └─> update_record(session, ...)
                  └─> call_api(session, "PUT", ...)
                      └─> api_request(session, logger, ...) [from _api]
```

## Session Lifecycle

1. **Creation** (`run_migration` start)
   ```python
   session = create_session(logger, base_url, token=token, authorization=authorization, cookie=cookie)
   ```

2. **Use** (passed to every API operation)
   - Same session preserves cookies and headers across requests
   - Consistent retry logic via `api_request()`

3. **Cleanup** (automatic)
   - Session remains open until `run_migration` completes
   - All cookies and state properly maintained

## Error Handling

- `api_request()` from `_api.py` handles:
  - HTTP status code checking
  - Automatic re-authentication on 401/403
  - Retry logic (configurable max_retries)
  - Logging of all requests

- `create_record()`, `update_record()` handle response parsing
- `upload_file()` from `_api.py` handles:
  - File validation
  - MIME type detection
  - Multipart form construction
  - Upload status verification

## Testing the Refactored Script

### Syntax Check
```bash
python -m py_compile nto_mesto_migration.py
```

### Dry Run
```bash
python nto_mesto_migration.py \
  --input test_data.json \
  --base-url https://api.example.com \
  --token MY_TOKEN \
  --dry-run-uploads
```

### Full Migration
```bash
python nto_mesto_migration.py \
  --input data.json \
  --base-url https://api.example.com \
  --token MY_TOKEN \
  --output results.json
```

## Comparison with migration.py

Both scripts now follow the same pattern:

| Aspect | migration.py | nto_mesto_migration.py |
|--------|-------------|-----------------------|
| Session Creation | `setup_session(logger)` | `create_session(logger, ...)` → `setup_session()` |
| API Requests | `api_request(session, logger, ...)` | `call_api(session, ...)` → `api_request(...)` |
| File Upload | `upload_file(session, logger, ...)` | `upload_one_file(session, ...)` → `upload_file(...)` |
| Logging | `setup_logger()` | `setup_logger()` |
| JSON Serialization | `jsonable(obj)` | `jsonable(obj)` |
| GUID Gen | `generate_guid()` | `generate_guid()` |

## Benefits of This Refactoring

1. **Consistency**: Both migration scripts use identical patterns
2. **Maintainability**: Bug fixes in `_api.py` fix both migrations
3. **Single Source of Truth**: Auth, retry logic, error handling in one place
4. **Reduced Code**: ~200 lines of duplicate code removed
5. **Session Reuse**: Same HTTP session across entire migration (better performance)
6. **Type Safety**: Clearer function signatures with session parameter

## Migration Completeness

Different migrations using the same helper infrastructure:
- ✓ `migration.py` - Main registry migration
- ✓ `nto_mesto_migration.py` - NTO places registry migration
- Can add more migrations following the same pattern

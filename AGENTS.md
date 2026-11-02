# Project Collaboration Guide

## Scope and Project Structure

This project archives public App Store metadata and artwork for competitor research.

- `collector.py`: HTTP requests, metadata and media caching, CSV/JSON summaries, and the CLI.
- `excel_export.py`: seven-sheet research workbooks using openpyxl and Pillow.
- `excel_locale.py`: generated workbook labels, sheet names, and navigation translations.
- `apps.txt`: research input; accepts numeric App IDs and App Store URLs.
- `tests/`: offline unit and regression tests.
- `output/`: generated reports, original JSON, downloaded media, and manifests.
- `README.md` and `README_zh-CN.md`: English and Chinese user documentation.

## Working Practices

- Inspect the worktree and staging area before editing. Preserve existing user changes, collection results, and manually renamed files.
- Use Python 3.10 or later. Declare runtime dependencies in `requirements.txt`; install with `python -m pip install -r requirements.txt`.
- Keep changes focused on the requested task and follow the existing code style.
- Update both README files when changing CLI options, dependencies, naming, export layout, or documented collection and cache behavior.
- Do not automatically commit generated workbooks, downloaded assets, caches, or sensitive configuration.

## Collection and Cache Behavior

- Use public App Store sources only. Do not introduce credentials, private endpoints, or access-control bypasses.
- Keep US and CN records separate and explicitly labeled. Never substitute one storefront's data for another's.
- Preserve original JSON and media. Report missing files, failed requests, unavailable apps, and unclassified artwork honestly.
- `--regions` controls collection requests; report rebuilding may include existing US/CN manifests for the input IDs. Keep this distinction documented.
- Reuse successful cached metadata and media unless a refresh or a required retry calls for a new request.
- A latest Lookup result of `not_available` must not be overwritten by success derived only from historical metadata. Keep historical files, but recheck Lookup before treating their metadata as current again, including after a failed recheck.
- Validate cached metadata types during both collection and report rebuilding. Invalid metadata must produce an error record rather than terminate the batch.
- Retry transient request failures, including incomplete response reads. After retries are exhausted, record the failure at the app or media level and allow subsequent items to proceed.
- Keep media requests restricted to the supported Apple artwork CDN hosts. Label web-discovered artwork as `web-unclassified`; do not assume it is a screenshot.
- Preserve atomic temporary-file replacement for metadata, manifests, summaries, and workbooks.

## Dates and Timestamps

- Require `--product-name` in the CLI.
- Name workbooks `<product-name>_AppStore竞品宣传图与介绍_yyyy-MM-dd.xlsx`. The Chinese suffix stays fixed in both interface languages.
- Calculate the filename date once at command startup using Beijing time (UTC+8); a run crossing midnight must keep the same filename.
- Store record timestamps in UTC.
- `fetched_at` represents metadata acquisition time, backed by `metadata_fetched_at` in the manifest. Reusing cached data must not update it.
- Track collection attempts and completion separately with `last_attempt` and `last_success`.
- For legacy caches without a metadata timestamp, freeze the previous `last_success` as an approximation. Do not invent a new fetch time; a successful refresh establishes a new accurate timestamp.

## Workbook Contract

- Support `--excel-language zh-CN` and `--excel-language en`; default to Chinese.
- Translate generated labels and navigation only. Preserve original app names, descriptions, release notes, source fields, statuses, and errors.
- Keep these seven sheet purposes and their order fixed:

  | Order | Chinese name | English name | Purpose |
  | --- | --- | --- | --- |
  | 1 | 研究总览 | Research Overview | Comparison metrics, collection status, and preview links |
  | 2 | 宣传图预览 | Image Preview | Per-app/storefront preview of up to three images |
  | 3 | 全部宣传图 | All Images | Full gallery of non-icon manifest entries |
  | 4 | 应用介绍 | App Descriptions | Descriptions and release notes split into sections |
  | 5 | 宣传图索引 | Image Index | Device/source labels, file availability, URLs, and errors |
  | 6 | 来源数据 | Source Data | Original summary fields |
  | 7 | 采集说明 | Collection Notes | Scope, limitations, timestamps, and usage guidance |

- Use colored app title bars in previews, app information on the left, three proportional thumbnails in the middle, and navigation on the right. Prefer iPhone images; use other images when no iPhone entries exist.
- Embed thumbnails in the workbook without modifying original media. Use visible placeholders for missing or unreadable images.
- Keep internal hyperlinks aligned with the selected sheet language.
- Preserve full original text in JSON/Markdown and segmented description sheets, accounting for Excel's cell length limit.
- Keep historical-download counts distinct from current local file availability, and explain the counting rules in the workbook and documentation.

## Verification

- Use offline fixtures and mocked responses by default. Do not start real collection requests automatically for testing.
- Run the regression suite after code changes:

  ```powershell
  python -m unittest discover -s tests -v
  ```

- Cover cache reuse, refresh, unavailable-to-available transitions, invalid metadata, exhausted network retries, continued batch processing, and timestamp stability when changing those behaviors.
- For export changes, reopen a generated sample with openpyxl and verify both languages, all seven sheet names and their order, embedded images, internal link targets, and preservation of source text.
- Run `git diff --check` for whitespace errors.
- Report what changed, which checks ran, and any remaining limitations. Distinguish offline validation from live source verification.

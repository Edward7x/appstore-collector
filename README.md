# App Store Batch Collector

English | [Chinese documentation](README_zh-CN.md)

Collect public App Store metadata, descriptions, icons, and screenshots for competitor research. The collector queries the US and China storefronts and exports a seven-sheet Excel workbook with embedded image previews, plus CSV, JSON, and original media files.

Requires Python 3.10 or later and network access to Apple's Lookup API and image CDN. Excel export uses openpyxl and Pillow. No App Store account is required.

## Contents

- [Quick start](#quick-start)
- [Examples](#examples)
- [Command reference](#command-reference)
- [Output reference](#output-reference)
- [Frequently asked questions](#frequently-asked-questions)
- [Limitations](#limitations)
- [Development](#development)

## Quick start

### 1. Check Python and install dependencies

```powershell
python --version
python -m pip install -r requirements.txt
```

On Windows, use `py -3` if `python` is unavailable. On macOS or Linux, use `python3`. Use the same interpreter for dependency installation and collection.

### 2. Prepare the competitor list

Edit `apps.txt`, putting one App Store URL or App ID on each line:

```text
# One app per line
333903271
https://apps.apple.com/us/app/id284882215
```

### 3. Run the collection

Open PowerShell in the project directory and run:

```powershell
py -3 collector.py --product-name "ReadLater" --input apps.txt --output output
```

The collector queries the US and China storefronts separately and creates an archive in `output`.

### 4. Try discovering additional images (optional)

```powershell
py -3 collector.py --product-name "ReadLater" --input apps.txt --output output --web-fallback
```

This adds candidate images from public pages to the normal collection. Not all additional images are necessarily screenshots.

### Collection statuses

| Status | Meaning |
| --- | --- |
| `complete` | Collection finished with no failed media downloads; this does not mean the API exposed every store asset. |
| `partial` | Metadata was collected, but some media downloads failed. |
| `not_available` | Lookup did not return the app in that storefront. |
| `error` | A request, cache read, or other collection step failed. |

## Examples

These examples assume dependencies are installed and `apps-readlater.txt` contains the target apps.

### Research only the China storefront

```powershell
python collector.py --product-name "ReadLater" --input apps-readlater.txt --output output/ReadLater --regions cn
```

This limits new collection requests to China. Reports can still include cached US records for the same input IDs if they exist in this output directory. Use a new directory for a report containing only the newly selected storefront.

### Refresh an existing project

```powershell
python collector.py --product-name "ReadLater" --input apps-readlater.txt --output output/ReadLater --refresh
```

### Use a separate directory for a dated capture

```powershell
python collector.py --product-name "ReadLater" --input apps-readlater.txt --output output/ReadLater/2026-11-02
```

### Adjust request timing

```powershell
python collector.py --product-name "ReadLater" --input apps-readlater.txt --output output/ReadLater --interval 5 --retries 5 --timeout 30
```

### Try additional images from public pages

```powershell
python collector.py --product-name "ReadLater" --input apps-readlater.txt --output output/ReadLater --web-fallback
```

These extra images are stored as `web-unclassified`. They may include unrelated artwork and duplicates; review them before treating them as screenshots.

## Command reference

Run `python collector.py --help` for CLI help.

| Option | Default | Description |
| --- | --- | --- |
| `--product-name` | Required | Research name used in the workbook filename. |
| `--input` | `apps.txt` | UTF-8 input file; UTF-8 BOM is accepted. |
| `--output` | `output` | Output directory. |
| `--excel-language` | `zh-CN` | Excel labels and sheet names: `zh-CN` or `en`. Source content is unchanged. |
| `--regions` | `us cn` | One or both supported storefronts. |
| `--interval` | `3.5` | Minimum seconds between Lookup request starts; at least 3. |
| `--retries` | `4` | Maximum retries per request; zero or greater. |
| `--timeout` | `25` | Request timeout in seconds; greater than zero. |
| `--refresh` | Off | Fetch metadata and media again. |
| `--web-fallback` | Off | Discover additional, unclassified images from public HTML. |
| `--verbose` | Off | Enable debug logging. |

The interval applies to Lookup requests, not CDN downloads. Transient HTTP errors include 408, 429, and selected 5xx responses. HTTP retries use backoff and honor `Retry-After`; network failures also use backoff. Permanent failures are recorded.

## Output reference

```text
output/
|-- <product-name>_<fixed-report-suffix>_<yyyy-MM-dd>.xlsx
|-- index.csv
|-- index.json
|-- <app-id>/
|   |-- us/
|   |   |-- metadata.json
|   |   |-- description.md
|   |   |-- manifest.json
|   |   |-- icon/
|   |   `-- screenshots/
|   |       |-- iphone/
|   |       `-- ipad/
|   `-- cn/
|       `-- ...
`-- <app-id>/
    |-- us/
    `-- cn/
```

### Workbook naming conventions

- Workbook names combine the product name, a fixed Chinese report suffix, and a `yyyy-MM-dd` date. The suffix is fixed by the exporter, not translated by the product-name option. Its exact escaped representation is `AppStore\u7ade\u54c1\u5ba3\u4f20\u56fe\u4e0e\u4ecb\u7ecd`.
- The filename pattern is `<product-name>_<fixed-report-suffix>_<yyyy-MM-dd>.xlsx`.
- Dates use Beijing time (UTC+8), calculated once when the command starts. Record timestamps such as `fetched_at` use UTC.
- `fetched_at` records when metadata was fetched, and stays unchanged when cached data is reused. The manifest separately tracks `last_attempt` and `last_success` for collection runs. Legacy caches retain their previous `last_success` as an approximate metadata fetch time; use `--refresh` to establish an exact new timestamp.
- The same name on the same date updates the same workbook. Different dates produce different workbook files, while CSV, JSON, and media in that output directory continue to be updated.
- Product names must be nonempty, must not end in a period, and must not contain `<>:"/\|?*` or control characters.
- Close the workbook before rerunning. Keep the original directory if you need its local image links.
- An existing manually named `index.xlsx` is not renamed by the CLI.


To export an English workbook, add `--excel-language en`:

```powershell
python collector.py --product-name "ReadLater" --input apps.txt --output output/ReadLater-en --excel-language en
```

Use `--excel-language zh-CN` for Chinese (the default). This changes sheet names, headings, generated notes, placeholders, and navigation labels. App names, descriptions, release notes, source fields, and recorded statuses/errors stay as collected. The report filename keeps its existing naming rule. Switching language for the same product on the same date in the same directory overwrites the workbook; use separate directories to retain both versions.

### Workbook layout

Sheet labels default to Chinese; use `--excel-language en` for English labels. The following table describes them in workbook order; these English descriptions are not literal tab names.

| Position | Purpose | Contents |
| --- | --- | --- |
| 1 | Research overview | One row per app/storefront, comparison metrics, status, and preview links. |
| 2 | Image preview | Colored app heading, icon and metadata, up to three iPhone images (other images if none), and navigation. |
| 3 | All images | Non-icon manifest entries in groups of three, with device and source labels. |
| 4 | App descriptions | Descriptions and release notes split into sections of up to 900 characters. |
| 5 | Image index | Device, source, local availability, file links, original URLs, hashes, and errors. |
| 6 | Source data | Summary fields with filtering and wrapped text. |
| 7 | Collection notes | Scope, counting rules, limitations, and navigation guidance. |

Embedded thumbnails retain their proportions and do not modify original files. Missing or unreadable images show placeholders. Local file links depend on the original directory; remote URLs depend on network and source availability.

## Frequently asked questions

### Why does the command report that `--product-name` is required?

Add a research name, for example `--product-name "ReadLater"`. The collector does not derive the name from the input apps.

### Why is an app missing in one country?

Lookup may not return it in that storefront. Check the store page and recorded status. The collector does not substitute another region's data.

### Rerunning and resuming an interrupted collection

If collection stops partway through, run the original command again:

```powershell
py -3 collector.py --product-name "ReadLater" --input apps.txt --output output
```

The collector skips files already saved successfully and attempts unfinished downloads again. To fetch the latest app information and media, add `--refresh`:

```powershell
py -3 collector.py --product-name "ReadLater" --input apps.txt --output output --refresh
```

The default API request interval is 3.5 seconds, or about 17 requests per minute, below the roughly 20-per-minute reference limit described in Apple's documentation.

### Why are there fewer screenshots than on the store page?

The public API may expose only some assets. `--web-fallback` can discover additional images, but does not guarantee complete or correctly classified results.

### Why does writing the workbook fail with `PermissionError`?

Close it in Excel or another viewer, then rerun. If it is not open, check directory write permissions. Completed downloads remain available for resuming.

### Can I export without any network requests?

There is currently no CLI export-only mode. A normal rerun reuses cached metadata and files but may retry failures or query records without cached metadata. If the latest Lookup result was unavailable, a rerun rechecks Lookup instead of treating historical metadata as current. Historical JSON and media remain on disk, and reports may display them alongside the latest status.

### Why does a single-region run include the other region in the report?

Report rebuilding includes existing US and China manifests for the input IDs. `--regions` controls new collection requests. Use a separate output directory to avoid older cached records.

### Why does the overview count differ from usable local images?

The overview counts non-icon files marked successful in the manifest, including historical files. The index separately checks whether each local file exists. Inspect both when files were moved or deleted.

### Can I share just the workbook?

Embedded previews travel with the workbook. Full-resolution local links require the source directory. Confirm you have permission to share the content.

### Does a new date create a complete backup?

It creates a new workbook name only. For independent metadata and media, use a new output directory or copy the full research directory.

## Limitations

The public Lookup API does not guarantee every device rendering, localization, marketing asset, custom product page, or video. Public-page discovery is heuristic and cannot reliably classify additional artwork. URL deduplication is per app/storefront, not global.

The collector does not access private endpoints, scrape credentials, bypass access controls, or access third-party App Store Connect accounts. Respect source rate limits and content rights.

Excel cells are limited to 32,767 characters. Full original content remains in JSON and Markdown. Files are written through temporary-file replacement; a collection run is not a transaction across the whole output directory. Reports are rebuilt after each app/storefront completes.

## Development

| File | Responsibility |
| --- | --- |
| `collector.py` | Requests, cache, CLI, and summary generation. |
| `excel_export.py` | Workbook layout and embedded images. |
| `requirements.txt` | Runtime dependencies. |
| `tests/` | Offline unit tests. |
| `AGENTS.md` | Project collaboration and verification rules. |
| `.gitignore` | Generated output, workbooks, caches, environments, and local configuration. |

Run tests without starting a real collection:

```powershell
python -m unittest discover -s tests -v
```

For export changes, also reopen a sample workbook and check sheet order, images, and internal links. Update both README files when changing options, dependencies, naming, or layout. Ignore rules do not untrack files already tracked by Git.

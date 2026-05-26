# Invoice Check CLI Service

Managed AgentD CLI service for fiscal invoice QR verification through an Android phone and WeChat.

The service is designed to be called by a lightweight AgentD skill after OCR/QR extraction has produced a QR task directory.

## Commands

```bash
bin/invoice-check-cli health

bin/invoice-check-cli verify-one \
  --qr-image /abs/path/qr_001.png \
  --output-dir /abs/path/output/qr_001

bin/invoice-check-cli verify-batch \
  --input-dir /abs/path/qr_task/qr \
  --output-dir /abs/path/output/invoice_check_run
```

## Final Output

`verify-batch` produces user-facing artifacts:

```text
output/
├── 发票核验结果清单.xlsx
├── invoice_check_results.zip
└── validation_screenshots/
    ├── invoice_001.png
    └── ...
```

The JSON files under `runs/` and `batch_report.json` are internal trace/debug artifacts for AgentD and service diagnostics. They are not the final user-facing report.

The result archive contains:

- `发票核验结果清单.xlsx`
- `validation_screenshots/`

## Runtime Behavior

For each QR image:

1. Clear the phone staging album directory.
2. Push exactly one QR image into the phone staging album.
3. Use screenshot + PaddleOCR-VL + VLM to operate the phone.
4. Verify the invoice QR.
5. Save the result screenshot.
6. Return to the `财政票据` WeChat service account main/chat page.
7. Clear the phone staging album directory.

Batch policy:

- Per QR task: up to 3 attempts by default.
- Failed tasks are skipped first, while the batch continues.
- After the first pass, failed tasks are replayed once.
- Remaining failures are recorded in the Excel report.

## Configuration

Default config lives at:

```text
config/defaults.json
```

The wrapper reads `.env` if present. Important variables:

```bash
INVOICE_CHECK_ADB_PATH=adb
INVOICE_CHECK_VLM_API=http://127.0.0.1:8081/v1/chat/completions
INVOICE_CHECK_VLM_MODEL=Qwen3.6-35B-A3B-UD-Q4_K_XL.gguf
INVOICE_CHECK_VLM_TEMPERATURE=0.01
INVOICE_CHECK_VLM_TOP_P=0.8
INVOICE_CHECK_VLM_TOP_K=20
INVOICE_CHECK_VLM_MAX_TOKENS=512
INVOICE_CHECK_VLM_TIMEOUT=30
INVOICE_CHECK_VLM_ENABLE_THINKING=false
INVOICE_CHECK_PADDLE_VL_API=http://127.0.0.1:8090/v1/chat/completions
INVOICE_CHECK_PADDLE_VL_MODEL=PaddleOCR-VL
INVOICE_CHECK_PADDLE_VL_TEMPERATURE=0
INVOICE_CHECK_PADDLE_VL_MAX_TOKENS=768
INVOICE_CHECK_PADDLE_VL_TIMEOUT=15
INVOICE_CHECK_REMOTE_DIR=/sdcard/Pictures/invoice-check
```

The main controller VLM is intentionally configured as non-thinking by default
for stable JSON action output.

## Server Dependencies

Python:

```bash
python3 -m venv .venv
.venv/bin/pip install -r requirements.txt
```

Android remote control:

- Install Android platform-tools.
- Ensure `adb devices -l` sees the target phone.
- Keep the phone unlocked and WeChat logged in.
- Configure the Android launcher so Home returns to a simple page with WeChat visible.

Linux notes:

- This service uses pure ADB commands (`screencap`, `input tap`, `input keyevent`, `push`, media scanner broadcast).
- `scrcpy` is optional for human observation and is not required for headless execution.
- The Linux server needs USB access to the phone or a stable TCP ADB connection.

## AgentD Skill Boundary

The AgentD skill should:

- run OCR/QR extraction first
- call `invoice-check-cli verify-batch`
- return the Excel report and result archive

The skill should not embed ADB/VLM phone-control logic directly.

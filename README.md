# 2605 HKL evaluation

OCR 추출 결과(XLSX)를 ground truth와 비교해 평가합니다.

## 오프라인(폐쇄망) 설치

### 1. 인터넷 있는 환경 — wheel 준비

```bash
uv export --frozen --no-dev --no-hashes --format requirements.txt -o requirements.txt
mkdir -p wheels
uv run --with pip python -m pip download -r requirements.txt -d wheels
```

### 2. 폐쇄망 — wheel로 환경 동기화

```bash
uv pip install --no-index --find-links .\wheels -r .\requirements.txt
./evaluation.sh
```

*uv python 설치는 수동으로 수행.
**`Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned`

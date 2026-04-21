# 2605 HKL evaluation

OCR 추출 결과(XLSX)를 ground truth와 비교해 평가합니다.

## 오프라인(폐쇄망) 설치

### 1. 인터넷 있는 환경 — wheel + uv Python 준비

```bash
uv export --frozen --no-dev --no-hashes --format requirements.txt -o requirements.txt
mkdir -p wheels
uv run --with pip python -m pip download -r requirements.txt -d wheels
```

```bash
uv python list --only-installed
uv python dir

tar -C /data1/share/maruchanpark/uv_cache \
  -czf uv-python-cpython-3.12.12-linux-x86_64-gnu.tar.gz \
  cpython-3.12.12-linux-x86_64-gnu
```

복사 대상:
- `requirements.txt`
- `wheels/`
- `uv-python-cpython-3.12.12-linux-x86_64-gnu.tar.gz`
- `pyproject.toml`, `uv.lock`, 소스 전체

### 2. 폐쇄망 — uv Python 먼저 적용 (필수)

```bash
mkdir -p ./offline_uv_cache
tar -C ./offline_uv_cache -xzf uv-python-cpython-3.12.12-linux-x86_64-gnu.tar.gz

export UV_CACHE_DIR="$PWD/offline_uv_cache"
export UV_PYTHON_DOWNLOADS=never

uv python list --only-installed
```

### 3. 폐쇄망 — wheel로 환경 동기화

```bash
uv sync --frozen --offline --no-index -f ./wheels
./evaluation.sh
```

주의:
- 온라인/오프라인 환경의 Python 버전, OS/아키텍처를 맞춰야 합니다.

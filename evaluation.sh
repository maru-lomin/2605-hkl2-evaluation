#!/usr/bin/env bash
set -euo pipefail

uv run python evaluation.py result_mapped ground_truth evaluation

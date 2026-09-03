#!/usr/bin/env bash
set -e
cd "$(dirname "$0")"

if [ ! -d venv ]; then
  python3 -m venv venv
fi
source venv/bin/activate
pip install -r requirements.txt

if [ ! -f .env ]; then
  cp .env.example .env
  echo "Created .env - edit it and add your OPENAI_API_KEY, then re-run this script."
  exit 0
fi

cd backend
uvicorn main:app --reload --port 8420

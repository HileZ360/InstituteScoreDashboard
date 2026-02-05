.PHONY: dev test compile

dev:
	uv run uvicorn app.main:app --reload

test:
	uv run python -m unittest discover -s tests -p "test*.py" -q

compile:
	uv run python -m compileall -q app tests


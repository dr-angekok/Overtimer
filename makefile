# makefile

install:
	poetry install

build:
	poetry build

publish:
	poetry publish --dry-run

package-install:
	pip install dist/*.whl

package-install-reinstall:
	pip install dist/*.whl --force-reinstall

lint:
	poetry run flake8

test:
	poetry run pytest

extended-test:
	poetry run pytest -vv

test-coverage:
	poetry run pytest --cov=page_loader --cov-report xml

update:
	poetry update

run-overtimer:
	poetry run overtimer

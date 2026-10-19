VENVDIR := .venv

ERROR_NO_VIRTUALENV = $(error Python virtualenv is not active, activate first)
ERROR_ACTIVE_VIRTUALENV = $(error Python virtualenv is active, deactivate first)

############################
## Help

.PHONY: help
.DEFAULT_GOAL := help
help:
	@printf 'Usage: make [VARIABLE=] TARGET\n'
	@awk 'BEGIN {FS = ":.*##";} /^[a-zA-Z1-9_-]+:.*?##/ { printf "  \033[36m%-20s\033[0m %s\n", $$1, $$2 } /^##@/ { printf "\n\033[1m%s\033[0m\n", substr($$0, 5) } ' $(MAKEFILE_LIST)


############################
##@ Python virtualenv

.PHONY: virtualenv
virtualenv:  ## Create venv directory
ifdef VIRTUAL_ENV
	$(ERROR_ACTIVE_VIRTUALENV)
endif
	uv venv
	@echo "To install packages: 'make install' or 'make install-dev'"

.PHONY: rmvirtualenv
rmvirtualenv:  ## Remove venv and Python cache directories
ifdef VIRTUAL_ENV
	$(ERROR_ACTIVE_VIRTUALENV)
endif
	rm -rf ${VENVDIR}
	find . -type d -name __pycache__ -print -exec rm -rf {} +


############################
##@ Install

.PHONY: install
install:  ## Install project script to be run without activating the virtual environment
	uv tool install --reinstall .

.PHONY: install-dev
install-dev:  ## Install project packages and script for development
	uv sync

.PHONY: uninstall
uninstall:  ## Uninstall project script from system-wide access
	uv tool uninstall outtaoutlook


############################
##@ Code analysis

.PHONY: ruffcheck
ruffcheck:  ## Run Ruff on project files
ifndef VIRTUAL_ENV
	$(ERROR_NO_VIRTUALENV)
endif
	ruff check src

.PHONY: ruffclean
ruffclean:  ## Clear Ruff caches
ifndef VIRTUAL_ENV
	$(ERROR_NO_VIRTUALENV)
endif
	ruff clean

.PHONY: pyright
pyright:  ## Run static type checks
	basedpyright src


############################
##@ Build

.PHONY: distclean
distclean:  ## Delete build files, python cache and package build artifacts
	rm -rf build
	rm -rf dist
	rm -rf src/*.egg-info
	rm -rf .ruff_cache
	rm -rf .pytest_cache
	find . -type d \( -name __pycache__ \) -print -exec rm -rf {} +

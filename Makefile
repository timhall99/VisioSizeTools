# -----------------------------------------------------------------------------
# Application
# -----------------------------------------------------------------------------

APP_NAME      ?= VisioSizeTools
DESCRIPTION   ?= Visio COM add-in - Size to Tallest / Shortest / Widest / Narrowest / Both Larger / Both Smaller.

SOLUTION      := VisioSizeTools.sln
PROJECT_DIR   := src/VisioSizeTools
CONFIGURATION ?= Release
BUILD_DIR     := $(PROJECT_DIR)/bin/$(CONFIGURATION)/net48
ISS           := installer/VisioSizeTools.iss
ARTIFACTS     := artifacts
ISCC          ?= C:/Program Files (x86)/Inno Setup 6/ISCC.exe

# Same convention as the other projects: CI passes VERSION=<tag> (e.g. v1.2.0); locally it's git describe.
VERSION       ?= $(shell git describe --tags --always --dirty 2>/dev/null || echo dev)
COMMIT        ?= $(shell git rev-parse --short HEAD 2>/dev/null || echo unknown)


# -----------------------------------------------------------------------------
# Shell configuration
# -----------------------------------------------------------------------------

# Chocolatey's GNU make defaults to cmd.exe; every recipe is bash, so force Git Bash. This makes `make`
# work the same from PowerShell, cmd.exe or Git Bash.
ifeq ($(OS),Windows_NT)
SHELL := C:/Program Files/Git/bin/bash.exe
endif

.SHELLFLAGS := -ec

# .NET needs a numeric X.Y.Z: take it from the front of VERSION ("v1.2.0-3-gabc123-dirty" -> 1.2.0),
# falling back to 0.0.0 before the first tag. The full VERSION + commit goes in InformationalVersion.
ASSEMBLY_VERSION := $(or $(shell echo "$(VERSION)" | grep -oE '^v?[0-9]+\.[0-9]+\.[0-9]+' | sed 's/^v//'),0.0.0)
SETUP_VERSION    := $(patsubst v%,%,$(VERSION))


# -----------------------------------------------------------------------------
# Default target
# -----------------------------------------------------------------------------

.DEFAULT_GOAL := help

.PHONY: help
help:
	@printf '%s\n' "$(APP_NAME) - $(DESCRIPTION)"
	@printf '%s\n\n' "Use one of the targets below."
	@printf '%s\n' "Build targets:"
	@printf '  %-20s %s\n' "build"       "Build Release configuration"
	@printf '  %-20s %s\n' "debug"       "Build Debug configuration (traces to %LOCALAPPDATA%\\VisioSizeTools.log)"
	@printf '  %-20s %s\n' "installer"   "Build Release and the Inno Setup installer into $(ARTIFACTS)/"
	@printf '\n%s\n' "Other targets:"
	@printf '  %-20s %s\n' "register"    "Register the local build with Visio for the current user"
	@printf '  %-20s %s\n' "unregister"  "Remove the add-in registration"
	@printf '  %-20s %s\n' "screenshots" "Regenerate the README images in docs/images (drives Visio)"
	@printf '  %-20s %s\n' "readme"      "Alias for screenshots"
	@printf '  %-20s %s\n' "clean"       "Remove generated build files"
	@printf '  %-20s %s\n' "version"     "Interactively bump and tag the release version"


# -----------------------------------------------------------------------------
# Build
# -----------------------------------------------------------------------------

.PHONY: build debug installer

build:
	dotnet build "$(SOLUTION)" -c $(CONFIGURATION) -nologo \
		-p:Version=$(ASSEMBLY_VERSION) \
		-p:InformationalVersion="$(VERSION)+$(COMMIT)"

debug:
	@$(MAKE) --no-print-directory build CONFIGURATION=Debug

# MSYS2_ARG_CONV_EXCL stops Git Bash rewriting ISCC's /D... switches as paths when calling a native exe.
installer: build
	@mkdir -p "$(ARTIFACTS)"
	MSYS2_ARG_CONV_EXCL='*' "$(ISCC)" /Qp \
		"/DAppVersion=$(SETUP_VERSION)" \
		"/DSourceDir=..\$(subst /,\,$(BUILD_DIR))" \
		"/DOutputDir=..\$(ARTIFACTS)" \
		"$(ISS)"
	@ls -l "$(ARTIFACTS)"


# -----------------------------------------------------------------------------
# Local registration and README images
# -----------------------------------------------------------------------------

.PHONY: register unregister screenshots readme

register: build
	powershell.exe -NoProfile -ExecutionPolicy Bypass -File scripts/Register.ps1 -Configuration $(CONFIGURATION)

unregister:
	powershell.exe -NoProfile -ExecutionPolicy Bypass -File scripts/Unregister.ps1

# Needs Visio + the add-in registered; close Visio first and leave the mouse/keyboard alone (~30 s).
screenshots:
	powershell.exe -NoProfile -ExecutionPolicy Bypass -File scripts/Capture-Screenshots.ps1

readme: screenshots


# -----------------------------------------------------------------------------
# Release version
# -----------------------------------------------------------------------------

# Releases are a deliberate act: tag locally here, push the tag, and CI (.github/workflows/release.yml)
# builds the installer and publishes the GitHub Release for it.
.PHONY: version

version:
	@current="$$(git describe --tags --abbrev=0 2>/dev/null || echo v0.1.0)"; \
	current="$${current#v}"; \
	IFS=. read -r major minor patch <<< "$$current"; \
	major=$${major:-0}; minor=$${minor:-1}; patch=$${patch:-0}; \
	echo "Current version: $$current"; \
	if [ -n "$$(git status --porcelain)" ]; then \
		echo "Warning: working tree has uncommitted changes; the tag will still point at HEAD ($$(git rev-parse --short HEAD))." >&2; \
	fi; \
	echo; \
	echo "  1) patch  -> $$major.$$minor.$$((patch+1))"; \
	echo "  2) minor  -> $$major.$$((minor+1)).0"; \
	echo "  3) major  -> $$((major+1)).0.0"; \
	echo "  4) custom version"; \
	echo "  5) keep current (no change)"; \
	printf "> "; \
	read choice || choice=5; \
	choice="$${choice%$$'\r'}"; \
	case "$$choice" in \
		1) new="$$major.$$minor.$$((patch+1))" ;; \
		2) new="$$major.$$((minor+1)).0" ;; \
		3) new="$$((major+1)).0.0" ;; \
		4) printf "New version (e.g. 1.4.0): "; read new || exit 1; new="$${new%$$'\r'}" ;; \
		5) echo "Keeping $$current."; exit 0 ;; \
		*) echo "Invalid selection: $$choice" >&2; exit 1 ;; \
	esac; \
	echo "$$new" | grep -Eq '^[0-9]+\.[0-9]+\.[0-9]+$$' || { echo "Invalid version format: $$new (expected X.Y.Z)" >&2; exit 1; }; \
	printf "Create tag v$$new at $$(git rev-parse --short HEAD)? [Y/n]: "; \
	read confirm || confirm=y; \
	confirm="$${confirm%$$'\r'}"; \
	case "$$confirm" in n|N|no|No|NO) echo "Cancelled."; exit 0 ;; esac; \
	git tag -a "v$$new" -m "Release v$$new"; \
	echo "Tagged v$$new locally. Push it when ready with: git push origin v$$new"


# -----------------------------------------------------------------------------
# Cleanup
# -----------------------------------------------------------------------------

.PHONY: clean

clean:
	rm -rf "$(PROJECT_DIR)/bin" "$(PROJECT_DIR)/obj" "$(ARTIFACTS)"

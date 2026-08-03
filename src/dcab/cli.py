"""Command-line interface for building, validating, and scoring DCAB."""

from __future__ import annotations

import argparse
import json
import os
import sys
import tempfile
from collections.abc import Sequence
from pathlib import Path

from . import __version__
from .adapters.docfence import DocFenceAdapterError
from .adapters.docfence import observations as docfence_observations
from .build import FixtureBuildError, build_fixtures
from .resources import bundled_fixture_root
from .score import (
    ObservationError,
    load_observations,
    observation_template,
    score_observations,
    strict_success,
)
from .validate import FixtureValidationError, validate_fixture_tree


def main(argv: Sequence[str] | None = None) -> int:
    """Run DCAB and return a documented process status."""

    parser = _parser()
    arguments = parser.parse_args(argv)
    try:
        if arguments.command == "build":
            _write_json(build_fixtures(arguments.fixtures, force=arguments.force), arguments.output)
            return 0
        if arguments.command == "validate":
            _write_json(validate_fixture_tree(arguments.fixtures), arguments.output)
            return 0
        if arguments.command == "observation-template":
            _write_json(observation_template(arguments.fixtures), arguments.output)
            return 0
        if arguments.command == "score":
            score = score_observations(
                arguments.fixtures, load_observations(arguments.observations)
            )
            _write_json(score, arguments.output, protected_paths=(arguments.observations,))
            return 0 if not arguments.strict or strict_success(score) else 1
        if arguments.command == "docfence-observations":
            result = docfence_observations(
                arguments.fixtures,
                executable=arguments.executable,
                timeout=arguments.timeout,
            )
            _write_json(result, arguments.output)
            return 0
    except (
        FixtureBuildError,
        FixtureValidationError,
        ObservationError,
        DocFenceAdapterError,
    ) as error:
        print(f"dcab: {error}", file=sys.stderr)
        return 2
    parser.error("a command is required")
    return 2


def _parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="dcab",
        description="Reproducible paired-WordprocessingML change-assurance benchmark.",
    )
    parser.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    commands = parser.add_subparsers(dest="command", required=True)

    build = commands.add_parser("build", help="generate deterministic fixture pairs")
    _fixtures_argument(build, default="fixtures")
    build.add_argument("--force", action="store_true", help="replace only known DCAB output files")
    build.add_argument("--output", help="write the build summary as JSON")

    validate = commands.add_parser(
        "validate", help="validate generated package and truth invariants"
    )
    _fixtures_argument(validate, default=str(bundled_fixture_root()))
    validate.add_argument("--output", help="write the validation summary as JSON")

    template = commands.add_parser(
        "observation-template", help="write unsupported-case observation skeleton"
    )
    _fixtures_argument(template, default=str(bundled_fixture_root()))
    template.add_argument("--output", help="write the observation JSON")

    score = commands.add_parser("score", help="score a normalized tool observation report")
    _fixtures_argument(score, default=str(bundled_fixture_root()))
    score.add_argument("--observations", required=True, help="normalized observation JSON")
    score.add_argument("--strict", action="store_true", help="fail unless every case is complete")
    score.add_argument("--output", help="write the score JSON")

    adapter = commands.add_parser(
        "docfence-observations", help="run optional local DocFence adapter over every pair"
    )
    _fixtures_argument(adapter, default=str(bundled_fixture_root()))
    adapter.add_argument("--executable", default="docfence", help="DocFence executable")
    adapter.add_argument("--timeout", type=int, default=120, help="per-pair timeout in seconds")
    adapter.add_argument("--output", help="write the observation JSON")
    return parser


def _fixtures_argument(parser: argparse.ArgumentParser, *, default: str) -> None:
    parser.add_argument("--fixtures", default=default, help=f"fixture root (default: {default})")


def _write_json(
    data: object, destination: str | None, *, protected_paths: tuple[str, ...] = ()
) -> None:
    content = json.dumps(data, indent=2, sort_keys=True) + "\n"
    if destination is None:
        sys.stdout.write(content)
        return
    temporary_path: Path | None = None
    try:
        target = Path(destination)
        if any(_same_path(target, protected) for protected in protected_paths):
            raise ObservationError("output must not replace an input")
        if target.exists() and (target.is_symlink() or not target.is_file()):
            raise ObservationError("output must be a regular file")
        if not target.parent.is_dir():
            raise ObservationError("output directory does not exist")
        descriptor, temporary_name = tempfile.mkstemp(prefix=".dcab-", dir=target.parent)
        temporary_path = Path(temporary_name)
        with os.fdopen(descriptor, "w", encoding="utf-8", newline="\n") as stream:
            stream.write(content)
            stream.flush()
            os.fsync(stream.fileno())
        os.replace(temporary_path, target)
        temporary_path = None
    except ObservationError:
        raise
    except OSError as error:
        raise ObservationError("output cannot be written") from error
    finally:
        if temporary_path is not None:
            try:
                temporary_path.unlink(missing_ok=True)
            except OSError:
                pass


def _same_path(left: Path, right: str) -> bool:
    try:
        return os.path.abspath(left) == os.path.abspath(right)
    except (OSError, TypeError, ValueError):
        return False


if __name__ == "__main__":  # pragma: no cover
    raise SystemExit(main())

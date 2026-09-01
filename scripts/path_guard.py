#!/usr/bin/env python3
"""Shared create-new path guards for document helper scripts."""

from __future__ import annotations

import os
from pathlib import Path
from typing import Iterable


PathLike = str | os.PathLike[str] | Path


def _resolved(path: PathLike, *, strict: bool) -> Path:
    return Path(path).expanduser().resolve(strict=strict)


def same_file(left: PathLike, right: PathLike) -> bool:
    """Return True for identical paths or existing filesystem aliases/hard links."""
    left_path = _resolved(left, strict=False)
    right_path = _resolved(right, strict=False)
    if left_path == right_path:
        return True
    try:
        return os.path.samefile(left_path, right_path)
    except (FileNotFoundError, OSError):
        return False


def existing_input(path: PathLike, *, label: str = "input") -> Path:
    resolved = _resolved(path, strict=True)
    if not resolved.is_file():
        raise ValueError(f"{label} is not a file: {resolved}")
    return resolved


def ensure_new_file(
    output: PathLike,
    *,
    inputs: Iterable[PathLike] = (),
    other_outputs: Iterable[PathLike] = (),
    suffixes: Iterable[str] | None = None,
    create_parent: bool = True,
) -> Path:
    """Resolve a create-new output and reject existing files and aliases."""
    output_path = _resolved(output, strict=False)
    if suffixes is not None:
        allowed = {item.lower() for item in suffixes}
        if output_path.suffix.lower() not in allowed:
            raise ValueError(
                f"output extension {output_path.suffix!r} is not allowed; expected one of {sorted(allowed)}"
            )
    if output_path.exists():
        raise FileExistsError(f"output already exists; choose a new path: {output_path}")

    for index, input_path in enumerate(inputs, 1):
        resolved_input = existing_input(input_path, label=f"input {index}")
        if same_file(output_path, resolved_input):
            raise ValueError(f"output must be separate from input {index}: {resolved_input}")

    for index, peer in enumerate(other_outputs, 1):
        if same_file(output_path, peer):
            raise ValueError(f"output must be separate from peer output {index}: {peer}")

    parent = output_path.parent
    if create_parent:
        parent.mkdir(parents=True, exist_ok=True)
    elif not parent.is_dir():
        raise ValueError(f"output parent does not exist: {parent}")
    return output_path


def ensure_new_directory(
    output: PathLike,
    *,
    inputs: Iterable[PathLike] = (),
    create_parent: bool = True,
) -> Path:
    """Resolve a directory that does not yet exist and cannot alias an input."""
    output_path = _resolved(output, strict=False)
    if output_path.exists():
        raise FileExistsError(f"output directory already exists; choose a new path: {output_path}")
    for index, input_path in enumerate(inputs, 1):
        resolved_input = _resolved(input_path, strict=True)
        if same_file(output_path, resolved_input):
            raise ValueError(f"output directory must be separate from input {index}: {resolved_input}")
    if create_parent:
        output_path.parent.mkdir(parents=True, exist_ok=True)
    elif not output_path.parent.is_dir():
        raise ValueError(f"output parent does not exist: {output_path.parent}")
    return output_path

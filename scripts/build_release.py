from __future__ import annotations

import argparse
import os
import shutil
import subprocess
import tarfile
from pathlib import Path


APP_NAME = "velora-finance"
ROOT = Path(__file__).resolve().parent.parent


def add_data_argument(source: Path, destination: str) -> str:
    return f"{source}{os.pathsep}{destination}"


def run_pyinstaller(output_dir: Path) -> Path:
    build_root = ROOT / "build_artifacts"
    dist_dir = build_root / "dist"
    work_dir = build_root / "work"
    spec_dir = build_root / "spec"
    shutil.rmtree(build_root, ignore_errors=True)
    dist_dir.mkdir(parents=True, exist_ok=True)
    work_dir.mkdir(parents=True, exist_ok=True)
    spec_dir.mkdir(parents=True, exist_ok=True)

    command = [
        os.fspath(Path(os.sys.executable)),
        "-m",
        "PyInstaller",
        "--noconfirm",
        "--clean",
        "--name",
        APP_NAME,
        "--distpath",
        os.fspath(dist_dir),
        "--workpath",
        os.fspath(work_dir),
        "--specpath",
        os.fspath(spec_dir),
        "--add-data",
        add_data_argument(ROOT / "assets", "assets"),
        os.fspath(ROOT / "main.py"),
    ]

    if os.sys.platform == "win32":
        command.extend(["--onefile", "--windowed"])
    elif os.sys.platform == "darwin":
        command.append("--windowed")
    else:
        command.append("--onefile")

    subprocess.run(command, check=True, cwd=ROOT)
    output_dir.mkdir(parents=True, exist_ok=True)
    return dist_dir


def package_dist(dist_dir: Path, output_dir: Path) -> Path:
    if os.sys.platform == "win32":
        source = dist_dir / f"{APP_NAME}.exe"
        target = output_dir / "velora-finance-windows.exe"
        shutil.copy2(source, target)
        return target

    if os.sys.platform == "darwin":
        source = dist_dir / f"{APP_NAME}.app"
        archive_base = output_dir / "velora-finance-macos"
        archive_path = Path(
            shutil.make_archive(
                base_name=os.fspath(archive_base),
                format="zip",
                root_dir=dist_dir,
                base_dir=source.name,
            )
        )
        return archive_path

    source = dist_dir / APP_NAME
    target = output_dir / "velora-finance-linux.tar.gz"
    with tarfile.open(target, "w:gz") as archive:
        archive.add(source, arcname=APP_NAME)
    return target


def main() -> None:
    parser = argparse.ArgumentParser(description="Build Velora Finance release artifact.")
    parser.add_argument(
        "--output-dir",
        default="release_assets",
        help="Directory where the packaged artifact will be written.",
    )
    args = parser.parse_args()

    output_dir = (ROOT / args.output_dir).resolve()
    dist_dir = run_pyinstaller(output_dir)
    artifact = package_dist(dist_dir, output_dir)
    print(artifact)


if __name__ == "__main__":
    main()


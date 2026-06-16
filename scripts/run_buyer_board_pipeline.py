from __future__ import annotations

import argparse
import json
import shutil
import subprocess
import sys
from pathlib import Path


def resolve_repo_root() -> Path:
    return Path(__file__).resolve().parents[1]


def resolve_skill_root(repo_root: Path) -> Path:
    return repo_root / "buyer-board-layout"


def run(cmd: list[str]) -> None:
    result = subprocess.run(cmd, capture_output=True, text=True)
    if result.returncode != 0:
        raise RuntimeError(f"Command failed: {' '.join(cmd)}\nSTDOUT:\n{result.stdout}\nSTDERR:\n{result.stderr}")


def copy_assets_to_workspace(config_path: Path, buyers_path: Path, workspace_dir: Path) -> tuple[Path, Path]:
    workspace_dir.mkdir(parents=True, exist_ok=True)
    copied_config = workspace_dir / "buyer-board-config.json"
    copied_buyers = workspace_dir / "buyer-board-buyers.json"

    config = json.loads(config_path.read_text(encoding="utf-8"))
    for item in config:
        for key in ("LogoPath", "SitePath"):
            source = (config_path.parent / item[key]).resolve()
            destination = (workspace_dir / Path(item[key]).name).resolve()
            shutil.copy2(source, destination)
            item[key] = str(destination)

    copied_config.write_text(json.dumps(config, ensure_ascii=False, indent=2), encoding="utf-8")
    shutil.copy2(buyers_path, copied_buyers)
    return copied_config, copied_buyers


def main() -> int:
    parser = argparse.ArgumentParser(description="Run the buyer-board PPT pipeline end-to-end.")
    parser.add_argument("--template", required=True, help="Template PPTX path")
    parser.add_argument("--buyers", required=True, help="Buyers JSON path")
    parser.add_argument("--config", required=True, help="Image config JSON path")
    parser.add_argument("--output", required=True, help="Final PPTX output path")
    parser.add_argument("--preview-dir", required=True, help="PNG preview output directory")
    parser.add_argument("--workspace", required=True, help="Temporary workspace directory")
    parser.add_argument("--cover-title", required=True, help="Cover title")
    parser.add_argument("--cover-country", required=True, help="Cover country line")
    parser.add_argument("--content-title", required=True, help="Content title")
    args = parser.parse_args()

    repo_root = resolve_repo_root()
    skill_root = resolve_skill_root(repo_root)
    workspace = Path(args.workspace)
    text_draft = workspace / "text-draft.pptx"
    copied_config, copied_buyers = copy_assets_to_workspace(Path(args.config), Path(args.buyers), workspace)

    run(
        [
            sys.executable,
            str(skill_root / "scripts" / "fill_buyer_board_text.py"),
            args.template,
            str(copied_buyers),
            str(text_draft),
            args.cover_title,
            args.cover_country,
            args.content_title,
        ]
    )

    run(
        [
            "powershell",
            "-ExecutionPolicy",
            "Bypass",
            "-File",
            str(skill_root / "scripts" / "apply_buyer_board_images.ps1"),
            "-InputPpt",
            str(text_draft),
            "-ConfigJson",
            str(copied_config),
            "-OutputPpt",
            args.output,
            "-PreviewDir",
            args.preview_dir,
        ]
    )

    print(args.output)
    print(args.preview_dir)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

#!/usr/bin/env python3
"""
Export Package Manager Apps to CSV
各パッケージマネージャーからインストール済みアプリの情報をCSVで出力します。
"""

import subprocess
import csv
import json
import argparse
import shutil
import tempfile
import os
import xml.etree.ElementTree as ET
import datetime
from pathlib import Path
from typing import List, Dict, Tuple, Optional


def run_command(cmd: List[str]) -> Tuple[str, str, int]:
    """
    コマンドを実行して出力をキャプチャ（画面に表示されない）

    Returns:
        Tuple[stdout, stderr, returncode]
    """
    try:
        result = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", errors="replace")
        return result.stdout, result.stderr, result.returncode
    except FileNotFoundError:
        return "", f"Command not found: {cmd[0]}", 1
    except Exception as e:
        return "", str(e), 1


def is_command_available(command: str) -> bool:
    """コマンドが利用可能かチェック"""
    result = shutil.which(command) is not None

    # Scoopの場合、CMDシムファイルが見つかってもsubprocessで実行できない場合があるため
    # PowerShell経由でテストする
    if command == "scoop" and result:
        try:
            test_result = subprocess.run(["powershell", "-Command", "scoop --version"], capture_output=True, text=True, timeout=10)
            return test_result.returncode == 0
        except Exception:
            return False

    return result


def get_cache_dir(output_dir: Path) -> Path:
    """キャッシュディレクトリを取得・作成"""
    cache_dir = output_dir / "cache"
    cache_dir.mkdir(parents=True, exist_ok=True)
    return cache_dir


def get_cache_file(cache_dir: Path) -> Path:
    """キャッシュファイルのパスを取得"""
    return cache_dir / "winget_cache.json"


def load_all_cached_winget_info(cache_dir: Path) -> Dict[str, Dict]:
    """全キャッシュデータを読み込み"""
    cache_file = get_cache_file(cache_dir)

    if not cache_file.exists():
        return {}

    try:
        with open(cache_file, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        # キャッシュファイルが破損している場合は空の辞書を返す
        return {}


def load_cached_winget_info(cache_dir: Path, package_id: str) -> Optional[Dict]:
    """キャッシュから特定のpackage_idの情報を読み込み"""
    all_cache = load_all_cached_winget_info(cache_dir)
    return all_cache.get(package_id)


def save_winget_info_to_cache(cache_dir: Path, package_id: str, display_name: str):
    """winget show情報をキャッシュに保存（軽量版）"""
    cache_file = get_cache_file(cache_dir)

    try:
        # 既存のキャッシュデータを読み込み
        all_cache = load_all_cached_winget_info(cache_dir)

        # 新しい情報を追加（必要最小限のデータのみ）
        all_cache[package_id] = {"package_id": package_id, "cached_at": datetime.datetime.now().isoformat(), "display_name": display_name}

        # ファイルに書き戻し
        with open(cache_file, "w", encoding="utf-8") as f:
            json.dump(all_cache, f, ensure_ascii=False, indent=2)

    except Exception as e:
        print(f"    Warning: キャッシュの保存に失敗 {package_id}: {e}")


def get_app_display_name(package_id: str, cache_dir: Optional[Path] = None) -> str:
    """winget showで正確な表示名を取得（キャッシュ機能付き）"""

    # キャッシュが有効な場合は、まずキャッシュから検索
    if cache_dir:
        cached_info = load_cached_winget_info(cache_dir, package_id)
        if cached_info:
            return cached_info.get("display_name", "")

    # キャッシュにない場合はwinget showを実行
    display_name = ""

    try:
        stdout, _, returncode = run_command(["winget", "show", package_id, "--disable-interactivity"])

        if returncode == 0 and stdout:
            # 最初の行から表示名を抽出
            for line in stdout.split("\n"):
                line = line.strip()
                if f"[{package_id}]" in line:
                    # "見つかりました AppName [PackageId]" または "Found AppName [PackageId]" の形式から AppName を取得
                    before_bracket = line.split(f"[{package_id}]")[0].strip()

                    # 検索結果の接頭語のみを除去（日本語表示名は保持）
                    if before_bracket.startswith("見つかりました "):
                        display_name = before_bracket[len("見つかりました ") :].strip()
                    elif before_bracket.startswith("Found "):
                        display_name = before_bracket[len("Found ") :].strip()
                    else:
                        # その他の場合は全体を表示名として使用
                        display_name = before_bracket.strip()
                    break
    except Exception:
        # エラーが発生した場合はそのまま継続
        pass

    # 表示名が取得できなかった場合はPackageIDから推測
    if not display_name:
        display_name = package_id.split(".")[-1] if "." in package_id else package_id

    # キャッシュに保存（表示名のみ）
    if cache_dir:
        save_winget_info_to_cache(cache_dir, package_id, display_name)

    return display_name


def get_winget_source_apps(source: str, source_name: str, cache_dir: Optional[Path] = None) -> List[Dict[str, str]]:
    """指定されたwingetソースからアプリの情報を取得（winget exportを使用）"""
    print(f"{source_name}アプリを処理中...")

    apps = []

    try:
        # 一時ファイルを作成
        temp_fd, temp_path = tempfile.mkstemp(suffix=".json", text=True)
        os.close(temp_fd)

        # winget export実行
        _, _, returncode = run_command(["winget", "export", "-s", source, "-o", temp_path, "--disable-interactivity", "--include-versions"])

        if returncode == 0 and os.path.exists(temp_path):
            with open(temp_path, "r", encoding="utf-8") as f:
                json_data = json.load(f)

            packages = []
            if "Sources" in json_data:
                for source_data in json_data["Sources"]:
                    if "Packages" in source_data:
                        packages = source_data["Packages"]
                        break

            # キャッシュにないパッケージをリストアップ
            packages_to_fetch = []
            all_cache = load_all_cached_winget_info(cache_dir) if cache_dir else {}

            for package in packages:
                package_id = package.get("PackageIdentifier", "")
                if package_id and package_id not in all_cache:
                    packages_to_fetch.append(package_id)

            fetch_count = len(packages_to_fetch)
            if fetch_count > 0:
                print(f"  → {fetch_count}個のパッケージでwinget showを実行中...")

            fetch_index = 0
            for package in packages:
                package_id = package.get("PackageIdentifier", "")
                version = package.get("Version", "latest")

                if package_id:
                    # キャッシュなしの場合のみ進捗表示
                    if package_id in packages_to_fetch:
                        fetch_index += 1
                        print(f"    [{fetch_index}/{fetch_count}] {package_id}")

                    # winget show で正確な表示名を取得（キャッシュ使用）
                    name = get_app_display_name(package_id, cache_dir)

                    apps.append({"PackageId": package_id, "Name": name, "Version": version})

        # 一時ファイルを削除
        if os.path.exists(temp_path):
            os.unlink(temp_path)

    except Exception as e:
        print(f"  Warning: winget export failed for {source}: {e}")

    print(f"  → {len(apps)}個のアプリを検出")
    return apps


def get_microsoft_store_apps(cache_dir: Optional[Path] = None) -> List[Dict[str, str]]:
    """Microsoft Store アプリの情報を取得（winget exportを使用）"""
    return get_winget_source_apps("msstore", "Microsoft Store", cache_dir)


def get_winget_apps(cache_dir: Optional[Path] = None) -> List[Dict[str, str]]:
    """Winget アプリの情報を取得（winget exportを使用）"""
    return get_winget_source_apps("winget", "Winget", cache_dir)


def get_scoop_apps() -> List[Dict[str, str]]:
    """Scoop アプリの情報を取得（scoop exportを使用）"""
    print("Scoopアプリを処理中...")

    apps = []

    try:
        # 一時ファイルを作成
        temp_fd, temp_path = tempfile.mkstemp(suffix=".json", text=True)
        os.close(temp_fd)

        # scoop export実行（直接標準出力からJSONを取得）
        stdout, _, returncode = run_command(["powershell", "-Command", "scoop export"])

        if returncode == 0 and stdout:
            json_data = json.loads(stdout)

            if "apps" in json_data:
                for app in json_data["apps"]:
                    name = app.get("Name", "")
                    version = app.get("Version", "latest")
                    source = app.get("Source", "main")

                    if name:
                        apps.append({"Name": name, "Version": version, "Source": source})

        # 一時ファイルを削除
        if os.path.exists(temp_path):
            os.unlink(temp_path)

    except Exception as e:
        print(f"  Warning: scoop export failed: {e}")

    print(f"  → {len(apps)}個のアプリを検出")
    return apps


def get_chocolatey_apps() -> List[Dict[str, str]]:
    """Chocolatey アプリの情報を取得（choco exportを使用）"""
    print("Chocolateyアプリを処理中...")

    apps = []

    try:
        # 一時ファイルを作成
        temp_fd, temp_path = tempfile.mkstemp(suffix=".config", text=True)
        os.close(temp_fd)

        # choco export実行
        _, _, returncode = run_command(["choco", "export", temp_path, "--include-version-numbers"])

        if returncode == 0 and os.path.exists(temp_path):
            tree = ET.parse(temp_path)
            root = tree.getroot()

            for package in root.findall("package"):
                package_id = package.get("id", "")
                version = package.get("version", "latest")

                if package_id and not package_id.startswith("chocolatey"):
                    apps.append({"PackageId": package_id, "Title": package_id, "Version": version})  # IDをタイトルとして使用

        # 一時ファイルを削除
        if os.path.exists(temp_path):
            os.unlink(temp_path)

    except Exception as e:
        print(f"  Warning: choco export failed: {e}")

    print(f"  → {len(apps)}個のパッケージを検出")
    return apps


def write_csv(filename, data: List[Dict[str, str]], fieldnames: List[str]):
    """データをCSVファイルに書き出し"""
    if not data:
        return

    try:
        with open(str(filename), "w", newline="", encoding="utf-8") as csvfile:
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
            writer.writeheader()
            writer.writerows(data)
    except Exception as e:
        print(f"  Error writing {filename}: {e}")


def main():
    parser = argparse.ArgumentParser(
        description="各パッケージマネージャーからインストール済みアプリの情報をCSVで出力",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
再インストールコマンドサンプル:

Microsoft Store / Winget:
  winget install --id [PackageId]
  winget install [PackageId]  # --id省略可能

Scoop:
  scoop install [Name]
  scoop install [Bucket]/[Name]  # bucketが異なる場合

Chocolatey:
  choco install [PackageId]
  choco install [PackageId] --version [Version]  # 特定バージョン
        """,
    )

    parser.add_argument("-o", "--output", default="./output", help="出力ディレクトリ (デフォルト: ./output)")

    args = parser.parse_args()

    # 出力ディレクトリの作成
    output_dir = Path(args.output)
    output_dir.mkdir(exist_ok=True)

    # キャッシュディレクトリの作成
    cache_dir = get_cache_dir(output_dir)

    print(f"Output directory: {output_dir.absolute()}")
    print(f"Cache directory: {cache_dir.absolute()}")
    print()

    # 各パッケージマネージャーの確認
    print("各パッケージマネージャーの確認中...")

    winget_available = is_command_available("winget")
    scoop_available = is_command_available("scoop")
    choco_available = is_command_available("choco")

    print(f"Winget: {'インストール済み' if winget_available else '未インストール'}")
    print(f"Scoop: {'インストール済み' if scoop_available else '未インストール'}")
    print(f"Chocolatey: {'インストール済み' if choco_available else '未インストール'}")
    print()

    # Microsoft Store アプリ（wingetに依存）
    if winget_available:
        ms_store_apps = get_microsoft_store_apps(cache_dir)
        if ms_store_apps:
            write_csv(
                output_dir / "microsoft_store_apps.csv",
                ms_store_apps,
                ["PackageId", "Name", "Version"],
            )

    # Winget アプリ
    if winget_available:
        winget_apps = get_winget_apps(cache_dir)
        if winget_apps:
            write_csv(output_dir / "winget_apps.csv", winget_apps, ["PackageId", "Name", "Version"])

    # Scoop アプリ
    if scoop_available:
        scoop_apps = get_scoop_apps()
        if scoop_apps:
            write_csv(output_dir / "scoop_apps.csv", scoop_apps, ["Name", "Version", "Source"])

    # Chocolatey アプリ
    if choco_available:
        choco_apps = get_chocolatey_apps()
        if choco_apps:
            write_csv(output_dir / "chocolatey_apps.csv", choco_apps, ["PackageId", "Title", "Version"])

    print()
    print("処理完了")
    print("出力されたファイル:")

    csv_files = ["microsoft_store_apps.csv", "winget_apps.csv", "scoop_apps.csv", "chocolatey_apps.csv"]
    found_files = []

    for filename in csv_files:
        filepath = output_dir / filename
        if filepath.exists():
            found_files.append(filename)
            print(f"  {filename}")

    if not found_files:
        print("  出力されたファイルはありません")


if __name__ == "__main__":
    main()

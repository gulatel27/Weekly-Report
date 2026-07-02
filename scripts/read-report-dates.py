import argparse
import datetime as dt
from pathlib import Path
import re
import sys


def configure_output() -> None:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    if hasattr(sys.stderr, "reconfigure"):
        sys.stderr.reconfigure(encoding="utf-8")


def parse_date(value: str) -> dt.date:
    text = re.sub(r"[-/.]", "", value.strip())
    if not re.fullmatch(r"\d{8}", text):
        raise ValueError("날짜는 YYYYMMDD 형식으로 입력하세요.")
    return dt.datetime.strptime(text, "%Y%m%d").date()


def prompt_summary_mode() -> str:
    print()
    print("요약 방식을 선택하세요.")
    print("1. LLM 요약 사용")
    print("2. 로컬 요약 사용")
    while True:
        choice = input("번호를 입력하세요 (1/2): ").strip()
        if choice == "1":
            return "llm"
        if choice == "2":
            return "local"
        print("입력 오류: 1 또는 2를 입력하세요.")


def write_env_file(path: Path, start_date: dt.date, end_date: dt.date, summary_mode: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    content = (
        f"set REPORT_START_DATE={start_date:%Y-%m-%d}\r\n"
        f"set REPORT_END_DATE={end_date:%Y-%m-%d}\r\n"
        f"set REPORT_SUMMARY_MODE={summary_mode}\r\n"
    )
    path.write_bytes(content.encode("ascii"))


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="주간보고 조회 기간 입력")
    parser.add_argument("--env-file", type=Path, required=True)
    return parser.parse_args()


def main() -> int:
    configure_output()
    args = parse_args()

    try:
        start_date = parse_date(input("시작일을 입력하세요 (YYYYMMDD): "))
        end_date = parse_date(input("종료일을 입력하세요 (YYYYMMDD): "))
        summary_mode = prompt_summary_mode()
    except EOFError:
        print("입력 오류: 날짜를 입력해야 합니다.", file=sys.stderr)
        return 1
    except ValueError as error:
        print(f"입력 오류: {error}", file=sys.stderr)
        return 1

    if end_date < start_date:
        print("입력 오류: 종료일은 시작일보다 빠를 수 없습니다.", file=sys.stderr)
        return 1

    write_env_file(args.env_file, start_date, end_date, summary_mode)
    print(f"조회 기간: {start_date:%Y-%m-%d} ~ {end_date:%Y-%m-%d}")
    print(f"요약 방식: {'LLM 요약' if summary_mode == 'llm' else '로컬 요약'}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

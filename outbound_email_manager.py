#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
outbound_email_manager.py
=========================================================
'정부지원사업_대시보드.xlsx'의 우선순위 랭킹을 기준으로,
아직 제안 메일을 보내지 않은 담당자만 골라내는 스크립트입니다.

핵심 아이디어:
- 발송이력.csv 라는 파일 하나에 "누구에게 언제 보냈는지"만 계속 쌓습니다.
- '다음 발송 대상' 요청 시, 대시보드의 담당자 이메일 목록에서 이미 발송이력에
  있는 이메일은 자동으로 빼고 보여줍니다. (기관 단위가 아니라 이메일 단위로
  추적하므로, 한 기관에 담당자가 여러 명이어도 새 담당자만 남습니다.)
- 실제 메일 발송은 이 스크립트가 하지 않습니다. 발송 대상 목록만 뽑아주고,
  사용자가 메일머지/직접발송 등으로 보낸 뒤 'mark-sent'로 기록하는 구조입니다
  (요청하신 대로 실제 발송 자동화는 하지 않고, 발송이력 관리 + 자동 필터링만).

사용법:
  # 1) 다음 발송 대상 50건 뽑기 (우선순위 높은 순, 아직 안 보낸 이메일만)
  python outbound_email_manager.py next --top 50 --output outputs/발송대상_다음배치.xlsx

  # 2) 실제로 메일을 보낸 뒤, 그 파일을 발송완료로 기록
  python outbound_email_manager.py mark-sent --from outputs/발송대상_다음배치.xlsx --campaign "2026-08 1차 제안"

  # 3) 특정 이메일 몇 개만 수동으로 발송완료 처리하고 싶을 때
  python outbound_email_manager.py mark-sent --emails a@test.go.kr,b@test.go.kr --campaign "수동추가"

  # 4) 지금까지 발송 현황 요약 보기
  python outbound_email_manager.py status
=========================================================
"""

import os
import argparse
from datetime import datetime

import pandas as pd

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUT_DIR = os.path.join(os.path.dirname(BASE_DIR), "outputs")
DASHBOARD_PATH = os.path.join(OUT_DIR, "정부지원사업_대시보드.xlsx")
HISTORY_PATH = os.path.join(OUT_DIR, "발송이력.csv")

RANK_SHEET = "대시보드_우선순위_랭킹"
HEADER_ROW = 3  # 0-indexed: 실제 헤더가 4번째 줄(제목/부제/빈줄 다음)


def log(msg):
    print(f"[outbound] {msg}", flush=True)


# =========================================================
# 발송이력 로드/저장
# =========================================================
def load_history(path=HISTORY_PATH):
    if not os.path.exists(path):
        return pd.DataFrame(columns=["이메일", "기관명", "발송일자", "캠페인명"])
    return pd.read_csv(path, encoding="utf-8-sig")


def save_history(df, path=HISTORY_PATH):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    df.to_csv(path, index=False, encoding="utf-8-sig")


def already_sent_emails(path=HISTORY_PATH):
    df = load_history(path)
    return set(df["이메일"].dropna().str.strip().str.lower())


# =========================================================
# 대시보드 로드
# =========================================================
def load_rank_sheet(dashboard_path=DASHBOARD_PATH):
    df = pd.read_excel(dashboard_path, sheet_name=RANK_SHEET, header=HEADER_ROW)
    return df


# =========================================================
# 명령: next - 다음 발송 대상 뽑기
# =========================================================
def cmd_next(args):
    df = load_rank_sheet(args.dashboard)
    sent = already_sent_emails(args.history)
    log(f"대시보드 {len(df)}개 기관 로드, 기존 발송이력 {len(sent)}건")

    results = []
    for _, row in df.iterrows():
        emails_str = row.get("담당자_이메일_통합모음", "") or ""
        if not isinstance(emails_str, str) or not emails_str.strip():
            continue
        emails = [e.strip() for e in emails_str.split(",") if e.strip()]
        remaining = [e for e in emails if e.lower() not in sent]
        if not remaining:
            continue
        results.append({
            "제안순위": row.get("제안순위"),
            "기관명": row.get("기관명"),
            "지점명": row.get("지점명"),
            "상위기관_주무부처_지자체": row.get("상위기관_주무부처_지자체"),
            "우선순위_종합점수": row.get("우선순위_종합점수"),
            "발송대상이메일": ", ".join(remaining),
            "발송대상이메일_개수": len(remaining),
            "대표전화": row.get("대표전화"),
        })

    if not results:
        log("발송 대상이 없습니다 (모두 이미 발송했거나, 이메일이 확보된 기관이 없습니다).")
        return

    out_df = pd.DataFrame(results).sort_values("우선순위_종합점수", ascending=False)
    if args.min_score is not None:
        out_df = out_df[out_df["우선순위_종합점수"] >= args.min_score]
    if args.top:
        out_df = out_df.head(args.top)

    out_df.to_excel(args.output, index=False)
    total_emails = out_df["발송대상이메일_개수"].sum()
    log(f"발송 대상 {len(out_df)}개 기관 (이메일 총 {total_emails}건) -> {args.output}")
    log("이 목록으로 메일을 보낸 뒤, mark-sent 명령으로 발송완료 기록을 남겨주세요.")


# =========================================================
# 명령: mark-sent - 발송완료 기록
# =========================================================
def cmd_mark_sent(args):
    history = load_history(args.history)
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    new_rows = []

    if args.from_file:
        df = pd.read_excel(args.from_file)
        for _, row in df.iterrows():
            emails_str = row.get("발송대상이메일", "")
            org = row.get("기관명", "")
            if not isinstance(emails_str, str):
                continue
            for e in emails_str.split(","):
                e = e.strip()
                if e:
                    new_rows.append({"이메일": e, "기관명": org, "발송일자": now, "캠페인명": args.campaign})
    elif args.emails:
        for e in args.emails.split(","):
            e = e.strip()
            if e:
                new_rows.append({"이메일": e, "기관명": "", "발송일자": now, "캠페인명": args.campaign})
    else:
        log("--from 또는 --emails 중 하나는 반드시 지정해야 합니다.")
        return

    if not new_rows:
        log("기록할 이메일이 없습니다.")
        return

    new_df = pd.DataFrame(new_rows)
    combined = pd.concat([history, new_df], ignore_index=True)
    # 같은 이메일+캠페인 조합 중복은 제거 (재실행 방지)
    combined = combined.drop_duplicates(subset=["이메일", "캠페인명"], keep="first")
    save_history(combined, args.history)
    log(f"발송완료 기록: {len(new_df)}건 추가 (누적 총 {len(combined)}건) -> {args.history}")


# =========================================================
# 명령: status - 현황 요약
# =========================================================
def cmd_status(args):
    history = load_history(args.history)
    log(f"총 발송이력: {len(history)}건 (고유 이메일 {history['이메일'].nunique() if len(history) else 0}개)")
    if len(history):
        print(history["캠페인명"].value_counts().to_string())
        print("\n최근 발송 5건:")
        print(history.sort_values("발송일자", ascending=False).head(5).to_string(index=False))

    if os.path.exists(args.dashboard):
        df = load_rank_sheet(args.dashboard)
        sent = already_sent_emails(args.history)
        contactable = df[df["담당자_이메일_통합모음"].fillna("").str.strip() != ""]
        remaining = 0
        for _, row in contactable.iterrows():
            emails = [e.strip() for e in str(row["담당자_이메일_통합모음"]).split(",") if e.strip()]
            if any(e.lower() not in sent for e in emails):
                remaining += 1
        print(f"\n연락 가능한 기관 {len(contactable)}개 중, 아직 안 보낸 곳: {remaining}개")


def main():
    parser = argparse.ArgumentParser(description="아웃바운드 메일 발송이력 관리")
    parser.add_argument("--dashboard", default=DASHBOARD_PATH, help="대시보드 xlsx 경로")
    parser.add_argument("--history", default=HISTORY_PATH, help="발송이력 csv 경로")
    sub = parser.add_subparsers(dest="command", required=True)

    p_next = sub.add_parser("next", help="다음 발송 대상 뽑기")
    p_next.add_argument("--top", type=int, default=None, help="상위 N개 기관만")
    p_next.add_argument("--min-score", type=float, default=None, help="우선순위 종합점수 최소값")
    p_next.add_argument("--output", default=os.path.join(OUT_DIR, "발송대상_다음배치.xlsx"))
    p_next.set_defaults(func=cmd_next)

    p_mark = sub.add_parser("mark-sent", help="발송완료 기록")
    p_mark.add_argument("--from", dest="from_file", default=None, help="next로 만든 xlsx 파일 경로")
    p_mark.add_argument("--emails", default=None, help="콤마로 구분한 이메일 목록 (수동 지정 시)")
    p_mark.add_argument("--campaign", default="", help="캠페인명 (예: '2026-08 1차 제안')")
    p_mark.set_defaults(func=cmd_mark_sent)

    p_status = sub.add_parser("status", help="발송 현황 요약")
    p_status.set_defaults(func=cmd_status)

    args = parser.parse_args()
    args.func(args)


if __name__ == "__main__":
    main()

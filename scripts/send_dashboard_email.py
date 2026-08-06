#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
send_dashboard_email.py
=========================================================
build_dashboard.py가 재생성한 'outputs/정부지원사업_대시보드.xlsx'를
이메일로 첨부해서 발송합니다. bizinfo_bot.py의 send_email()(Gmail SMTP,
재시도 로직 포함)을 그대로 재사용합니다 - 같은 MAIL_USER/EMAIL_PASS/
MAIL_RECEIVER 환경변수(시크릿)를 사용합니다.

워크플로우에서 build_dashboard.py 다음 단계로 실행하면 됩니다:
  python scripts/build_dashboard.py
  python scripts/send_dashboard_email.py
=========================================================
"""

import os
import sys
from datetime import datetime, timedelta, timezone

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import bizinfo_bot as bot  # send_email() 재사용

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUT_DIR = os.path.join(os.path.dirname(BASE_DIR), "outputs")
DASHBOARD_PATH = os.path.join(OUT_DIR, "정부지원사업_대시보드.xlsx")


def build_summary_html(today_str):
    """대시보드 상위 요약 정보를 메일 본문에 짧게 넣습니다 (첨부파일 전체 열지 않아도
    바로 감이 오도록)."""
    try:
        import pandas as pd
        df = pd.read_excel(DASHBOARD_PATH, sheet_name="대시보드_우선순위_랭킹", header=3)
        total = len(df)
        contactable = int((df["담당자_이메일_통합모음"].fillna("").str.strip() != "").sum())
        sent = int(df["발송여부"].fillna("").str.upper().eq("Y").sum())
        top5 = df.sort_values("우선순위_종합점수", ascending=False).head(5)
        rows_html = "".join(
            f"<tr><td style='padding:6px 10px;border-bottom:1px solid #e5e7eb;'>{r['제안순위']}</td>"
            f"<td style='padding:6px 10px;border-bottom:1px solid #e5e7eb;'>{r['지점명']}</td>"
            f"<td style='padding:6px 10px;border-bottom:1px solid #e5e7eb;'>{r['우선순위_종합점수']}</td></tr>"
            for _, r in top5.iterrows()
        )
        summary = f"""
        <div style="margin:14px 0;font-size:13.5px;color:#374151;">
          전체 {total}개 기관 · 연락 가능 {contactable}개 · 발송 완료 {sent}개
        </div>
        <table style="width:100%;border-collapse:collapse;margin-bottom:14px;">
          <thead><tr style="background:#111827;color:#fff;">
            <th style="padding:6px 10px;text-align:left;">순위</th>
            <th style="padding:6px 10px;text-align:left;">기관</th>
            <th style="padding:6px 10px;text-align:left;">점수</th>
          </tr></thead>
          <tbody>{rows_html}</tbody>
        </table>
        """
        return summary
    except Exception as e:
        bot.log(f"요약 정보 생성 실패(첨부파일은 정상 발송됨): {e}")
        return ""


def main():
    if not os.path.exists(DASHBOARD_PATH):
        bot.log(f"[치명적 오류] 대시보드 파일이 없습니다: {DASHBOARD_PATH}")
        sys.exit(1)

    with open(DASHBOARD_PATH, "rb") as f:
        dashboard_bytes = f.read()

    kst = timezone(timedelta(hours=9))
    today_str = datetime.now(kst).strftime("%Y-%m-%d")
    filename = f"정부지원사업_대시보드_{today_str}.xlsx"
    subject = f"[정부지원사업 대시보드] {today_str} 갱신본"

    summary_html = build_summary_html(today_str)
    html_body = f"""
    <html><body style="margin:0;padding:0;background:#f3f4f6;font-family:'Apple SD Gothic Neo','Malgun Gothic',sans-serif;">
      <div style="max-width:640px;margin:0 auto;padding:24px 16px;">
        <h2 style="color:#111827;margin:0 0 6px 0;">📊 정부지원사업 대시보드 갱신본</h2>
        <div style="color:#6b7280;font-size:13px;margin-bottom:16px;">{today_str} 기준</div>
        {summary_html}
        <div style="text-align:center;padding:14px;background:#eef2ff;border-radius:8px;
                    font-size:13.5px;color:#3730a3;">
          📎 첨부된 <b>{filename}</b> 파일에서 전체 내용을 확인하실 수 있습니다.
        </div>
      </div>
    </body></html>
    """

    try:
        bot.send_email(
            subject, html_body,
            attachment_bytes=dashboard_bytes,
            attachment_filename=filename,
        )
    except Exception as e:
        bot.log(f"[치명적 오류] 대시보드 메일 발송 실패: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()

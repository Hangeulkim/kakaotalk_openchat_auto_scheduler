"""
Google Calendar 자동 등록 모듈
- CSV 파일의 일정 데이터를 Google Calendar에 자동 등록합니다.
- 최초 실행 시 브라우저에서 Google 계정 인증이 필요합니다.
"""

import os
import re
import pandas as pd
from datetime import datetime, timedelta
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

# Google Calendar API 범위
SCOPES = ['https://www.googleapis.com/auth/calendar']

import sys

if getattr(sys, 'frozen', False):
    # PyInstaller로 빌드된 경우 실행 파일(exe)이 있는 디렉토리
    BASE_DIR = os.path.dirname(sys.executable)
else:
    # 파이썬 스크립트로 실행된 경우
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# 인증 파일 경로
CREDENTIALS_FILE = os.path.join(BASE_DIR, 'credentials.json')
TOKEN_FILE = os.path.join(BASE_DIR, 'token.json')


def get_calendar_service():
    """Google Calendar API 서비스 객체 생성 (OAuth2 인증)"""
    creds = None

    if os.path.exists(TOKEN_FILE):
        try:
            creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)
        except Exception:
            os.remove(TOKEN_FILE)
            creds = None

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
            except Exception:
                # 토큰 갱신 실패 → 삭제 후 재인증
                if os.path.exists(TOKEN_FILE):
                    os.remove(TOKEN_FILE)
                creds = None

        if not creds or not creds.valid:
            if not os.path.exists(CREDENTIALS_FILE):
                raise FileNotFoundError(
                    f"credentials.json 파일이 없습니다!\n"
                    f"경로: {CREDENTIALS_FILE}\n\n"
                    f"Google Cloud Console에서 OAuth 2.0 클라이언트 ID를 생성하고\n"
                    f"credentials.json 파일을 다운로드하여 이 폴더에 넣어주세요."
                )
            flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES)
            creds = flow.run_local_server(port=0)

        with open(TOKEN_FILE, 'w') as token:
            token.write(creds.to_json())

    return build('calendar', 'v3', credentials=creds)


def parse_datetime_from_row(row):
    """CSV 행에서 시작/종료 datetime 추출"""
    s_date_str = str(row.get('start_date', row.get('date', '')))
    e_date_str = str(row.get('end_date', row.get('date', '')))
    s_time_str = str(row.get('start_time', row.get('time', '')))
    e_time_str = str(row.get('end_time', row.get('time', '')))

    # --- 날짜 파싱 ---
    def parse_date(d_str):
        d_match = re.search(r'(\d{4})[./](\d{1,2})[./](\d{1,2})', d_str)
        if d_match:
            return int(d_match.group(1)), int(d_match.group(2)), int(d_match.group(3))
        return None, None, None

    s_y, s_m, s_d = parse_date(s_date_str)
    e_y, e_m, e_d = parse_date(e_date_str)

    if not s_y: return None, None
    if not e_y: e_y, e_m, e_d = s_y, s_m, s_d

    # --- 시간 파싱 ---
    def parse_time(t_str):
        # 1. 한국어: 오전 1:00
        t_match = re.search(r'(오[전후])\s*(\d{1,2}):(\d{2})', t_str)
        if t_match:
            ampm, h, m = t_match.group(1), int(t_match.group(2)), int(t_match.group(3))
            if ampm == '오후' and h != 12: h += 12
            elif ampm == '오전' and h == 12: h = 0
            return h, m
        # 2. 영어: 1:00 AM
        t_match_en = re.search(r'(\d{1,2}):(\d{2})\s*([APap][Mm])', t_str)
        if t_match_en:
            h, m, ap = int(t_match_en.group(1)), int(t_match_en.group(2)), t_match_en.group(3).upper()
            if ap == 'PM' and h != 12: h += 12
            elif ap == 'AM' and h == 12: h = 0
            return h, m
        return 0, 0

    is_allday = False
    if "하루 종일" in s_time_str or "시간 미상" in s_time_str:
        is_allday = True

    s_h, s_min = parse_time(s_time_str)
    e_h, e_min = parse_time(e_time_str)

    # 구버전 CSV 호환성: time 열 하나에 "오전 1:00 ~ 오전 2:00" 형태로 들어온 경우
    if e_time_str == s_time_str and ('~' in s_time_str or '-' in s_time_str):
        parts = re.split(r'[~\-–]', s_time_str)
        if len(parts) >= 2:
            s_h, s_min = parse_time(parts[0])
            e_h, e_min = parse_time(parts[1])
            is_allday = False # 시간 정보가 파싱되었으므로 종일이 아님

    start_dt = datetime(s_y, s_m, s_d, s_h, s_min)
    end_dt = datetime(e_y, e_m, e_d, e_h, e_min)

    if is_allday:
        # 구글 캘린더의 종일 일정은 종료일이 'Exclusive' (포함되지 않는 다음날 자정)여야 함
        end_dt = datetime(e_y, e_m, e_d) + timedelta(days=1)
    else:
        # 종료 시간이 시작 시간보다 빠르면(보통 12시 자정 넘어가는 경우) 다음 날로 처리
        if end_dt <= start_dt:
            end_dt += timedelta(days=1)

    return start_dt, end_dt, is_allday


def upload_csv_to_calendar(csv_path, calendar_id='primary', log_func=None):
    """CSV 파일의 일정을 Google Calendar에 등록
    
    Args:
        csv_path: CSV 파일 경로
        calendar_id: 캘린더 ID (기본: 'primary' = 기본 캘린더)
        log_func: 로그 출력 함수 (없으면 print)
    
    Returns:
        (성공 수, 실패 수, 건너뜀 수)
    """
    log = log_func or print

    if not os.path.exists(csv_path):
        log(f"❌ CSV 파일을 찾을 수 없습니다: {csv_path}")
        return 0, 0, 0

    log(f"📂 CSV 파일 로딩: {csv_path}")
    df = pd.read_csv(csv_path, encoding='utf-8-sig')
    log(f"  총 {len(df)}개 일정 발견")

    # Google Calendar API 인증
    log("🔑 Google Calendar 인증 중...")
    try:
        service = get_calendar_service()
        log("✅ 인증 성공!")
    except FileNotFoundError as e:
        log(str(e))
        return 0, 0, 0
    except Exception as e:
        log(f"❌ 인증 실패: {e}")
        return 0, 0, 0

    success, failed, skipped = 0, 0, 0

    # 1. 유효한 일정 파싱 및 최소/최대 날짜 구하기
    valid_events = []
    for idx, row in df.iterrows():
        s_dt, e_dt, is_allday = parse_datetime_from_row(row)
        if s_dt:
            valid_events.append((idx, row, s_dt, e_dt, is_allday))
        else:
            title = str(row.get('title', '제목 없음'))
            log(f"  ⏭️ [{idx+1}] '{title}' - 날짜 파싱 실패, 건너뜀")
            skipped += 1
            
    if not valid_events:
        log("❌ 등록 가능한 일정이 없습니다.")
        return success, failed, skipped

    # 2. 기존 일정 조회 (중복 방지 및 수정)
    log("🔍 기존 일정 확인 중 (업데이트 확인)...")
    existing_events_map = {}
    try:
        min_time = (min(e[2] for e in valid_events) - timedelta(days=1)).strftime('%Y-%m-%dT00:00:00+09:00')
        max_time = (max(e[3] for e in valid_events) + timedelta(days=1)).strftime('%Y-%m-%dT23:59:59+09:00')
        
        events_result = service.events().list(
            calendarId=calendar_id, 
            timeMin=min_time, 
            timeMax=max_time, 
            singleEvents=True,
            fields='items(id,summary,description,start)'
        ).execute()
        
        for ev in events_result.get('items', []):
            gcal_id = ev.get('id')
            desc = ev.get('description', '')
            m = re.search(r'ID:\s*([a-zA-Z0-9_\-]+)', desc)
            if m:
                existing_events_map[m.group(1)] = gcal_id
            
            summary = ev.get('summary', '')
            start_time = ev.get('start', {}).get('dateTime', '') or ev.get('start', {}).get('date', '')
            if start_time:
                date_part = start_time[:10] # YYYY-MM-DD
                existing_events_map[f"{summary}_{date_part}"] = gcal_id
    except Exception as e:
        log(f"⚠️ 기존 일정 조회 실패 (그냥 진행합니다): {e}")

    # 3. 일정 등록 및 수정
    for idx, row, start_dt, end_dt, is_allday in valid_events:
        title = str(row.get('title', '제목 없음'))
        author = str(row.get('author', ''))
        event_id = str(row.get('id', ''))

        # Google Calendar 이벤트 생성
        event = {
            'summary': title,
            'description': f"작성자: {author}\nID: {event_id}\n(카카오톡에서 자동 수집)",
        }
        
        if is_allday:
            event['start'] = {'date': start_dt.strftime('%Y-%m-%d')}
            event['end'] = {'date': end_dt.strftime('%Y-%m-%d')}
            log_time_str = "하루 종일"
        else:
            event['start'] = {
                'dateTime': start_dt.strftime('%Y-%m-%dT%H:%M:%S'),
                'timeZone': 'Asia/Seoul',
            }
            event['end'] = {
                'dateTime': end_dt.strftime('%Y-%m-%dT%H:%M:%S'),
                'timeZone': 'Asia/Seoul',
            }
            log_time_str = f"{start_dt.strftime('%m/%d %H:%M')}~{end_dt.strftime('%H:%M')}"

        # 중복(기존 일정) 체크
        fingerprint_id = event_id
        fingerprint_fallback = f"{title}_{start_dt.strftime('%Y-%m-%d')}"
        
        gcal_id = None
        if fingerprint_id and fingerprint_id in existing_events_map:
            gcal_id = existing_events_map[fingerprint_id]
        elif fingerprint_fallback in existing_events_map:
            gcal_id = existing_events_map[fingerprint_fallback]

        try:
            if gcal_id:
                # 이미 존재하는 일정은 수정(update)
                updated = service.events().update(calendarId=calendar_id, eventId=gcal_id, body=event).execute()
                log(f"  🔄 [{idx+1}] '{title}' → {log_time_str} (수정됨)")
                success += 1
            else:
                # 새 일정 등록(insert)
                created = service.events().insert(calendarId=calendar_id, body=event).execute()
                log(f"  ✅ [{idx+1}] '{title}' → {log_time_str}")
                success += 1
        except Exception as e:
            log(f"  ❌ [{idx+1}] '{title}' 처리 실패: {e}")
            failed += 1

    log(f"\n📊 결과: 성공 {success}건, 실패 {failed}건, 건너뜀 {skipped}건")
    return success, failed, skipped


if __name__ == '__main__':
    import sys
    if len(sys.argv) < 2:
        print("사용법: python google_calendar_upload.py <CSV파일경로>")
        print("예시: python google_calendar_upload.py kakao_events_20260424.csv")
        sys.exit(1)
    
    csv_path = sys.argv[1]
    upload_csv_to_calendar(csv_path)

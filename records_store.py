# -*- coding: utf-8 -*-
"""
보도자료 결과 기록 저장소 (엑셀 없이 앱이 직접 기록을 들고 있도록).

Streamlit Cloud는 재시작하면 서버 파일이 지워져서, 같은 GitHub 저장소의 별도 브랜치(pr-records)에
records.json 으로 보관한다. 앱이 배포되는 main 브랜치와 분리돼 있어 저장해도 앱이 재시작되지 않는다.
필요한 설정: Secrets 의 GITHUB_TOKEN (이 저장소 Contents 읽기/쓰기 권한).

records.json 형식:
  {"version": 1, "companies": {"딥파인": [{"date": "2026-08-18", "title": "...", "topic": "...",
                                          "total": 27, "media": ["머니투데이", ...]}, ...]}}
읽을 때는 report_history.read_history() 와 같은 모양(date 는 date, media 는 set)으로 돌려준다.
"""
import base64
import json
import urllib.error
import urllib.request
from datetime import date

REPO = "snuggle1020/pr_tools_webapp"
BRANCH = "pr-records"
PATH = "records.json"
API = "https://api.github.com"


class RecordsError(Exception):
    pass


class RecordsStore:
    def __init__(self, token, repo=REPO, branch=BRANCH, path=PATH):
        self.token, self.repo, self.branch, self.path = token, repo, branch, path

    # ---------------------------------------------------------------- GitHub API
    def _call(self, method, url, body=None):
        req = urllib.request.Request(
            API + url, method=method,
            data=json.dumps(body).encode("utf-8") if body is not None else None,
            headers={"Authorization": f"Bearer {self.token}", "Accept": "application/vnd.github+json",
                     "X-GitHub-Api-Version": "2022-11-28", "Content-Type": "application/json"},
        )
        try:
            with urllib.request.urlopen(req, timeout=20) as resp:
                return json.loads(resp.read().decode("utf-8") or "{}")
        except urllib.error.HTTPError as e:
            if e.code == 404:
                return None
            detail = e.read().decode("utf-8", errors="ignore")[:200]
            raise RecordsError(f"GitHub 기록 저장소 오류 (HTTP {e.code}): {detail}")
        except urllib.error.URLError as e:
            raise RecordsError(f"GitHub 연결 실패: {e.reason}")

    def _ensure_branch(self):
        if self._call("GET", f"/repos/{self.repo}/branches/{self.branch}") is not None:
            return
        repo = self._call("GET", f"/repos/{self.repo}")
        base = self._call("GET", f"/repos/{self.repo}/git/ref/heads/{repo['default_branch']}")
        self._call("POST", f"/repos/{self.repo}/git/refs",
                   {"ref": f"refs/heads/{self.branch}", "sha": base["object"]["sha"]})

    def _read_raw(self):
        """(데이터 dict, 파일 sha). 파일이 없으면 빈 데이터."""
        f = self._call("GET", f"/repos/{self.repo}/contents/{self.path}?ref={self.branch}")
        if not f:
            return {"version": 1, "companies": {}}, None
        data = json.loads(base64.b64decode(f["content"]).decode("utf-8"))
        data.setdefault("companies", {})
        return data, f["sha"]

    def _write_raw(self, data, sha, message):
        body = {"message": message, "branch": self.branch,
                "content": base64.b64encode(json.dumps(data, ensure_ascii=False, indent=1).encode("utf-8")).decode()}
        if sha:
            body["sha"] = sha
        self._call("PUT", f"/repos/{self.repo}/contents/{self.path}", body)

    def _update(self, company, fn, message):
        """company 기록 목록을 fn(list)->list 로 바꿔 저장. 동시에 다른 사람이 저장해 충돌하면 한 번 다시 시도."""
        self._ensure_branch()
        for attempt in range(2):
            data, sha = self._read_raw()
            data["companies"][company] = sorted(fn(list(data["companies"].get(company, []))), key=lambda r: r["date"])
            try:
                self._write_raw(data, sha, message)
                return
            except RecordsError as e:
                if attempt == 0 and "409" in str(e):
                    continue
                raise

    # ---------------------------------------------------------------- 공개 함수
    def history(self, company):
        """report_history.read_history() 와 같은 모양의 기록 목록 (배포일 순)."""
        data, _ = self._read_raw()
        return [{"sheet": "", "date": date.fromisoformat(r["date"]), "title": r["title"], "topic": r.get("topic", ""),
                 "total": int(r["total"]), "media": set(r.get("media", []))}
                for r in sorted(data["companies"].get(company, []), key=lambda r: r["date"])]

    def upsert(self, company, records, overwrite=True):
        """records: [{date(date), title, topic, total, media}] -> 같은 배포일 기록은 바꾸고(overwrite=False 면 그대로 둠) 없으면 추가.
        반환: 새로 추가된 수."""
        added = [0]

        def fn(rows):
            by_date = {r["date"]: r for r in rows}
            for rec in records:
                key = rec["date"].isoformat()
                if key in by_date and not overwrite:
                    continue
                added[0] += key not in by_date
                by_date[key] = {"date": key, "title": rec["title"], "topic": rec.get("topic", ""),
                                "total": int(rec["total"]), "media": sorted(rec.get("media", []))}
            return list(by_date.values())

        self._update(company, fn, f"{company} 보도자료 기록 {len(records)}건 저장")
        return added[0]

    def replace(self, company, records):
        """기록 수정 표에서 고친 결과로 통째로 바꿈 (주제 수정·삭제)."""
        def fn(_rows):
            return [{"date": r["date"].isoformat(), "title": r["title"], "topic": r.get("topic", ""),
                     "total": int(r["total"]), "media": sorted(r.get("media", []))} for r in records]

        self._update(company, fn, f"{company} 보도자료 기록 수정")

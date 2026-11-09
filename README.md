# 클로드 마스터 하기 — 실습 스킬 모음

스마트소셜 OD 「클로드 마스터 하기」(12회차) 강의에서 쓰는 실습 스킬입니다. 스킬 폴더는 `skills/`, 스킬이 아닌 실습 자료는 `practice/`에 있고, 바로 내려받을 zip 파일은 `downloads/`에 있습니다.

## 스킬 목록

| 회차 | 스킬 | 하는 일 | 내려받기 |
|---|---|---|---|
| 2회차 | [write-column](skills/write-column/) | 메모와 자료를 학과 소식지·학보·블로그 칼럼 초안으로 씁니다. 합니다체로 쓰고 점검 7항목으로 스스로 고칩니다 | [write-column.zip](downloads/write-column.zip) |
| 3회차 | [write-lab-report](skills/write-lab-report/) | 주간 실습 메모를 목적·방법·결과·느낀 점 보고서 초안으로 정리하고, 승인하면 .docx로 만듭니다. 메모에 없는 결과는 "확인 필요"로 표시합니다 | [write-lab-report.zip](downloads/write-lab-report.zip) |
| 4회차 | [slide-design](skills/slide-design/) | 원문과 발표 시간을 받아 결론·흐름 → 장 구성 → 장별 상세를 단계마다 확인받고, 승인하면 장 유형 15종의 16:9 pptx로 조판합니다(발표자 노트 포함) | [slide-design.zip](downloads/slide-design.zip) |
| 5회차 | [content-research-writer](skills/content-research-writer/) | 리포트·에세이를 설계 → 근거 확보(근거 장부·원문 대조) → 집필 → 최종 점검 4단계로 함께 씁니다. 기본은 내가 쓰고 피드백을 받으며, 요청하면 확인된 근거로 추천 원고를 씁니다 | [content-research-writer.zip](downloads/content-research-writer.zip) |
| 7회차 | [research-job-brief](skills/research-job-brief/) | 내가 정한 채용 사이트 · 지역 · 관심 분야로 오늘 접수 중인 공고를 찾아 주간 취업 브리핑 표와 이번 주 할 일 3개로 정리합니다. 상세 페이지에서 마감 여부를 확인하고, 공고에 없는 마감일 · 조건은 지어내지 않습니다 | [research-job-brief.zip](downloads/research-job-brief.zip) |

## 실습 자료 (스킬 아님)

| 회차 | 자료 | 쓰는 법 | 내려받기 |
|---|---|---|---|
| 6회차 | [과제 폴더](practice/session-06/과제%20폴더/) | 데이터베이스 팀플 과제가 흩어진 모습의 가상 md 노트 9개. 압축을 풀어 바탕화면에 두고 클로드 폴더 연결 실습에 씁니다 | [06_과제폴더.zip](downloads/06_과제폴더.zip) |
| 9회차 · 8회차 볼트 있음 | [볼트에 넣을 파일](practice/session-09/09_볼트에_넣을_파일/) | 8회차 볼트에 넣을 보고서 질문 1개, 자료 노트 6개, 9회차판 CLAUDE.md. 압축을 풀어 같은 이름의 칸에 넣습니다 | [09_볼트에_넣을_파일.zip](downloads/09_볼트에_넣을_파일.zip) |
| 9회차 · 8회차 볼트 없음 | [2026-2학기 볼트](practice/session-09/2026-2학기/) | 위 파일에 학기 지도와 빈 칸(02 정리 · 03 출력 · 04 보관 · 99 첨부)까지 갖춘 전체 실습 볼트. 바탕화면에 풀어 옵시디언 볼트로 엽니다 | [09_학기볼트.zip](downloads/09_학기볼트.zip) |

실습 자료의 노트는 가상 예시이거나 공개 자료를 요약한 노트입니다. 원문 PDF는 넣지 않았으니 각 노트의 url에서 직접 엽니다.

## 폴더 구성

```text
claude-agent-master/
├── skills/
│   ├── write-column/             2회차 · SKILL.md + references/writing-rules.md
│   ├── write-lab-report/         3회차 · SKILL.md
│   ├── slide-design/             4회차 · SKILL.md + references/(4개 문서) + scripts/build_slides.py
│   ├── content-research-writer/  5회차 · SKILL.md + LICENSE-Apache-2.0.txt
│   └── research-job-brief/       7회차 · SKILL.md + references/(my-info · sources)
├── practice/                     스킬이 아닌 실습 자료 (6회차 과제 폴더 · 9회차 볼트에 넣을 파일 · 9회차 전체 볼트)
└── downloads/                    스킬 zip (Claude에 바로 업로드) + 실습 자료 zip
```

## 설치 방법

1. `downloads/`에서 필요한 회차의 zip 파일을 내려받습니다. 압축은 풀지 않습니다.
2. Claude 데스크톱 또는 claude.ai에서 설정 → 기능(Capabilities) → 스킬(Skills)로 이동합니다.
3. 스킬 업로드를 눌러 zip 파일을 올리고, 목록에서 켜짐 상태인지 확인합니다.
4. 대화창에서 스킬 설명에 맞는 요청을 보냅니다. 예) "학과 소식지 칼럼 써 줘", "실습 보고서로 정리해 줘"

Claude Code에서 쓰려면 스킬 폴더(예: `skills/write-column/`)를 `~/.claude/skills/` 아래에 복사합니다.

slide-design은 조판할 때 Python과 python-pptx를 씁니다. python-pptx가 없으면 스크립트가 설치를 시도합니다.

## 라이선스

- write-column, write-lab-report, slide-design, research-job-brief, practice/의 실습 자료: CC BY-NC 4.0 ([LICENSE](LICENSE)). 비상업 목적으로 공유·수정할 수 있으며 출처를 표시해야 합니다.
- content-research-writer: [ComposioHQ/awesome-claude-skills](https://github.com/ComposioHQ/awesome-claude-skills)의 content-research-writer(Apache-2.0)를 한국어화하고 범위 확인 검색·원문 대조 절차·국내 인용 양식을 더한 수정판입니다. Apache License 2.0을 따르며 전문은 [LICENSE-Apache-2.0.txt](skills/content-research-writer/LICENSE-Apache-2.0.txt)에 있습니다.

Copyright (c) 2026 HONGS LAB (홍순성)

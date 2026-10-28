# 클로드 마스터 하기 — 실습 스킬 모음

스마트소셜 OD 「클로드 마스터 하기」(12회차) 강의에서 쓰는 실습 스킬입니다. 스킬 폴더는 `skills/`에 있고, 바로 올려 쓸 수 있는 zip 파일은 `downloads/`에 있습니다.

## 스킬 목록

| 회차 | 스킬 | 하는 일 | 내려받기 |
|---|---|---|---|
| 2회차 | [write-column](skills/write-column/) | 메모와 자료를 학과 소식지·학보·블로그 칼럼 초안으로 씁니다. 합니다체로 쓰고 점검 7항목으로 스스로 고칩니다 | [write-column.zip](downloads/write-column.zip) |
| 3회차 | [write-lab-report](skills/write-lab-report/) | 주간 실습 메모를 목적·방법·결과·느낀 점 보고서 초안으로 정리하고, 승인하면 .docx로 만듭니다. 메모에 없는 결과는 "확인 필요"로 표시합니다 | [write-lab-report.zip](downloads/write-lab-report.zip) |
| 4회차 | [slide-design](skills/slide-design/) | 원문과 발표 시간을 받아 결론·흐름 → 장 구성 → 장별 상세를 단계마다 확인받고, 승인하면 장 유형 15종의 16:9 pptx로 조판합니다(발표자 노트 포함) | [slide-design.zip](downloads/slide-design.zip) |
| 5회차 | [content-research-writer](skills/content-research-writer/) | 리포트·에세이를 설계 → 근거 확보(근거 장부·원문 대조) → 집필 → 최종 점검 4단계로 함께 씁니다. 기본은 내가 쓰고 피드백을 받으며, 요청하면 확인된 근거로 추천 원고를 씁니다 | [content-research-writer.zip](downloads/content-research-writer.zip) |

## 폴더 구성

```text
claude-agent-master/
├── skills/
│   ├── write-column/             2회차 · SKILL.md + references/writing-rules.md
│   ├── write-lab-report/         3회차 · SKILL.md
│   ├── slide-design/             4회차 · SKILL.md + references/(4개 문서) + scripts/build_slides.py
│   └── content-research-writer/  5회차 · SKILL.md + LICENSE-Apache-2.0.txt
└── downloads/                    스킬 zip (Claude에 바로 업로드)
```

## 설치 방법

1. `downloads/`에서 필요한 회차의 zip 파일을 내려받습니다. 압축은 풀지 않습니다.
2. Claude 데스크톱 또는 claude.ai에서 설정 → 기능(Capabilities) → 스킬(Skills)로 이동합니다.
3. 스킬 업로드를 눌러 zip 파일을 올리고, 목록에서 켜짐 상태인지 확인합니다.
4. 대화창에서 스킬 설명에 맞는 요청을 보냅니다. 예) "학과 소식지 칼럼 써 줘", "실습 보고서로 정리해 줘"

Claude Code에서 쓰려면 스킬 폴더(예: `skills/write-column/`)를 `~/.claude/skills/` 아래에 복사합니다.

slide-design은 조판할 때 Python과 python-pptx를 씁니다. python-pptx가 없으면 스크립트가 설치를 시도합니다.

## 라이선스

- write-column, write-lab-report, slide-design: CC BY-NC 4.0 ([LICENSE](LICENSE)). 비상업 목적으로 공유·수정할 수 있으며 출처를 표시해야 합니다.
- content-research-writer: [ComposioHQ/awesome-claude-skills](https://github.com/ComposioHQ/awesome-claude-skills)의 content-research-writer(Apache-2.0)를 한국어화하고 범위 확인 검색·원문 대조 절차·국내 인용 양식을 더한 수정판입니다. Apache License 2.0을 따르며 전문은 [LICENSE-Apache-2.0.txt](skills/content-research-writer/LICENSE-Apache-2.0.txt)에 있습니다.

Copyright (c) 2026 HONGS LAB (홍순성)

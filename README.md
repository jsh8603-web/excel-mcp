# xlkit — Mac FP&A용 엑셀 MCP (하나로 통합)

Claude Code에서 엑셀 파일, 특히 **Power Query가 들어 있는 파일**을 깨뜨리지 않고 다루기 위한 MCP 서버.
공식 Claude for Excel 애드인을 쓸 수 없는 회사 Mac(Apple Silicon) 환경을 기준으로 만들었다.

## 설치 (한 번만)

```bash
git clone https://github.com/jsh8603-web/excel-mcp ~/excel-mcp
bash ~/excel-mcp/setup.sh
```

이후 **어떤 프로젝트에서 claude를 실행해도** `xlkit` MCP가 붙는다. 프로젝트마다 할 일은 없다.

엑셀 프로젝트에서만 쓰고 싶으면 `bash setup.sh --project`로 설치만 하고, 필요한 프로젝트 폴더에서 등록한다:
```bash
claude mcp add -s project xlkit -- ~/.xlkit/bin/xlkit    # 그 폴더에 .mcp.json 생성
```

haris-musa·xlide를 따로 설치하거나 지울 필요는 없다(xlkit 안에 들어 있음). 예전에 Claude Code에
haris-musa를 별도 MCP로 등록했다면 도구가 겹치지 않게 그 등록만 지운다(`claude mcp list` → `claude mcp remove 이름`).
- 패키지는 `wheels/`(macOS arm64, Python 3.12)에서 오프라인 설치한다. PyPI에 접속하지 않는다.
- 필요한 것: `uv`(brew install uv), Claude Code. Python 3.12가 없으면 uv가 GitHub에서 받는다.
- 업데이트: `cd ~/excel-mcp && git pull && bash setup.sh`
- 서버는 Claude Code를 연 프로젝트 폴더 안 파일만 다룬다.

## 구조

`xlkit` 하나가 아래 두 서버를 내부에서 띄워 도구를 한곳에 모은다(101개).

| 내부 구성 | 맡는 일 | 버전 |
|---|---|---|
| [haris-musa/excel-mcp-server](https://github.com/haris-musa/excel-mcp-server) | 읽기, 일반 파일 편집, 피벗, 차트, 수식 추적, 목표값 찾기 | 2.2.6 |
| [xlide_mcp](https://github.com/WilliamSmithEdward/xlide_mcp) (`xlide_*` 도구) | Power Query M 코드, **PQ 파일 편집**, VBA 읽기·편집 | 1.2.4 |
| [formulon](https://github.com/libraz/formulon) | `xl_verify`(계산 검증), `xl_whatif`(민감도 분석) | 0.18.0 |
| AppleScript | `pq_refresh`(Power Query 새로 고침, **미검증**) | — |

게이트웨이가 강제하는 안전 규칙:
1. Power Query나 외부 연결이 있는 파일을 haris-musa 쓰기 도구로 저장하려 하면 거절하고 xlide 도구를 안내한다.
2. 엑셀에서 열려 있는 파일(`~$파일명` 잠금 파일 있음)은 쓰기를 거절한다.

작업 규칙(어떤 도구를 언제 쓰는지)은 MCP 서버 안내문으로 Claude에게 자동 전달된다.
하위 서버 두 개의 원래 안내문과 xlide 리소스도 그대로 전달한다. 숨기는 도구는 Mac에서 쓸 수 없는
xlide 기능(매크로·테스트 실행, 라이브 세션, Access 전용) 12개뿐이다.

## 왜 이렇게 나눴나 (2026-11-07 실측)

Power Query 테스트 파일에 셀 하나만 쓰고 저장한 뒤 파일 구조를 비교했다.

| 도구 | M 코드 | 연결 정보 | 쿼리 결과 표 연결 |
|---|---|---|---|
| haris-musa `write_range` | 보존 | 보존 | **삭제됨** (다른 시트만 고쳐도 같음) |
| formulon 저장 | 보존 | 보존 | 표 2개인 파일에서 **끊김** |
| xlide `write_cells`, `write_query` | 보존 | 보존 | 보존 (바꾼 파트만 수정) |

FP&A 함수 16개 계산 시험: PMT·XNPV·XIRR·EOMONTH·SLN·DDB·VDB·PV 등은 둘 다 정확.
LAMBDA·MAP·SCAN은 haris-musa가 계산하지 못하고("uncalculated"), formulon은 정확.

## 할 수 없는 것

- **Power Query 실행**: 엑셀에서 직접 해야 한다. `pq_refresh`는 Mac에서 아직 검증하지 않았다.
- **데이터 모델·Power Pivot·DAX**: Mac 엑셀에 없는 기능.
- **엑셀에서 열린 파일의 저장 안 된 내용**: 파일 기반이라 보지 못한다.
- **DRM·보안 라벨 파일**: 파일을 직접 열 수 없으면 실패한다.
- **xlide로 쓴 수식의 값**: 엑셀이 다음에 열 때 계산된다. 바로 필요하면 `xlide_read_cells(calculate=true)`나 `xl_verify(print_range)`.
- **편집 후 점검**: MCP로 저장한 파일은 엑셀 계산값이 지워지거나 옛 값이 남는다. `xl_verify`에 `baseline=원본`을 줘야 정확하다.

## 선택: VS Code

사람이 직접 M 코드를 편집하려면 `Excel Power Query Editor` 확장을 쓴다(setup.sh가 설치 시도,
MS 공식 Power Query 언어 확장이 함께 설치됨).

## 버전 올리기

`requirements.lock` 버전을 바꾸고 `wheels/`를 다시 받은 뒤, 위 실측 표를 다시 확인한다(특히 PQ 보존).

# 예산서 검토 정적 웹앱

GitHub + Cloudflare Pages에 업로드하기 위한 브라우저 기반 정적 웹앱입니다.

## 현재 버전 방향

이 패키지는 이전 실험 버전들 중 복잡하게 수정된 파서를 되돌리고, 최초 버전의 단순 검토 흐름을 기준으로 다시 정리한 버전입니다.

사용자가 업로드한 `school-zone-search-main` 프로젝트에서 아래 요소만 재사용했습니다.

- `public/` 배포 구조
- Cloudflare Pages용 `_headers`, `_redirects`
- `package.json`의 로컬 실행/검사 스크립트 구조
- 무료 운영 점검 스크립트 구조
- 보안 헤더 구성 방향

버린 요소는 다음과 같습니다.

- 통학구역 검색 로직
- 통학구역 CSV/XLSX 데이터
- 주소 자동완성/학교 검색 기능
- 데이터 빌드 스크립트
- 통학구역 전용 문서

## 로컬 실행

```bash
npm run dev
```

브라우저에서 `http://localhost:8788`을 엽니다.

## 검사

```bash
npm run check
npm run smoke
npm run audit:security
npm run audit:free-tier
```

## Cloudflare Pages 설정

| 항목 | 값 |
| --- | --- |
| Framework preset | None |
| Build command | 비워둠 또는 `exit 0` |
| Build output directory | `public` |
| Root directory | 비워둠 |

## 주의

엑셀과 PDF는 서버로 전송하지 않고 브라우저에서만 처리합니다.
현재 파서는 1차 기준이며, 실제 기관별 엑셀 양식에 맞춰 보수기준 시트 인식 규칙을 추가로 고정해야 합니다.

# Word 문서 분할 스크립트

대용량 Word 문서(.docx)를 지정된 페이지 수 단위로 분할합니다.

## 요구사항

- Windows + Microsoft Word 설치
- Python 3.10+

## 설치

```bash
pip install -r requirements.txt
```

## 사용법

기본 사용 (100페이지 단위):
```bash
python split_word.py input.docx
```

페이지 수 지정:
```bash
python split_word.py input.docx -p 50
```

출력 디렉토리 지정:
```bash
python split_word.py input.docx -o ./output
```

## 출력 예시

```
Word 애플리케이션을 시작합니다...
문서를 열고 있습니다: C:\docs\report.docx
총 페이지 수: 3000
분할 단위: 100페이지
생성될 파일 수: 30
----------------------------------------
파트 1/30 저장 완료: report_001.docx (p.1-100)
파트 2/30 저장 완료: report_002.docx (p.101-200)
...
파트 30/30 저장 완료: report_030.docx (p.2901-3000)
----------------------------------------
완료! 30개 파일이 저장되었습니다.
```

## 왜 python-docx가 아닌 pywin32인가?

`python-docx`는 문서의 XML 구조만 다루며 **페이지 경계를 알 수 없습니다**.
페이지는 폰트, 여백, 이미지 크기 등에 따라 Word가 렌더링할 때 결정되므로,
정확한 페이지 기반 분할은 Word COM 자동화(pywin32)가 필요합니다.

## 참고사항

- 실행 중 Word가 백그라운드에서 열립니다 (화면에 표시되지 않음)
- 문서 크기에 따라 시간이 걸릴 수 있습니다
- 원본 파일은 읽기 전용으로 열리므로 수정되지 않습니다

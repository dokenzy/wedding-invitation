# wedding-invitation

- 청첩장 보낼 때 받는 사람 주소 만들기
- 참고: [Python과 LaTeX으로 우편발송용 주소록 만들기](http://blog.dokenzy.com/archives/2038)

## 미리보기

- ![](sample.png?raw=true "")

## 필요한 거

- Python >= 3.4.2
- 서울남산체
- XeLaTeX

## 실행

### 준비

- list.xlsx 파일을 열어 양식에 맞게 입력한다

### 만들기
1. `python wedding.py`을 실행해서 data.tex 파일을 만든다.
2. `xelatex template`을 실행해서 컴파일하면 PDF파일이 만들어진다.


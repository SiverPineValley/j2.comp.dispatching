# j2.comp.dispatching
배차 관리 스크립트

</br>

## 1. 설치 방법

### 1.1 파일 Clone
```
git clone https://github.com/SiverPineValley/j2.comp.dispatching.git
```

</br>

### 1.2 압축 해제

</br>

### 1.3 빌드
```
go build -o main.exe
```

</br>
</br>

## 2. 실행 방법 (Windows 기준)

### 2.1 설정 파일 변경
config.toml 파일의 컬럼명, 시작 Index를 수정

</br>

### 2.2 커맨드 입력
```
./main.exe {j2파일명} {cj파일명}
```
</br>

### 2.3 파라미터 입력
j2시트명, cj시트명, 출력파일이름 세 가지를 입력

</br>
</br>

## 3. License
```
Copyright 2020 SiverPineValley

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation
files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy,
modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software
is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES
OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE
LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR
IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
```
# Searching Stocks in Korea
대신증권 Creon을 활용한 종목 검색 프로그램

본 프로그램은 한국 주식시장 내에서 설정한 조건에 맞는 종목들을 찾고, 텔레그램으로 종목 Table을 전송하는 것까지 구현이 되어 있습니다.

## Environment
- Windows 10
- [Python 3.8 32bit](https://www.python.org/downloads/windows/)
- [대신증권 API Creon/Creon Plus](https://www.creontrade.com/g.ds?p=4108&v=3073&m=4441)
- [Telegram](https://desktop.telegram.org/)

## Usage
1. Creon 실행 및 로그인<br>
  ```
  python autoConnect.py
  ```
2. 종목 검색 실행<br>
  ```
  python main.py
  ```
  
## Description
1. autoConnect.py 파일 내의 사용자 계좌 정보를 입력해야 합니다
2. main.py 파일 내 telegramMachine instance를 생성할 때 보낼 Token id 및 Chat id가 필요합니다.

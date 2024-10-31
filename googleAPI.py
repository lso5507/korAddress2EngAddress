import pandas as pd
import requests
from concurrent.futures import ThreadPoolExecutor
from dotenv import load_dotenv
import os

# .env 파일에서 환경 변수 로드
load_dotenv()

# Google Cloud API 키 설정
API_KEY = os.getenv('GOOGLE_API_KEY')  # 구글 API 키

# 엑셀 파일 경로 입력
input_excel_file_path = '/Users/leeseokwoon/Documents/GitHub/address2xy/주소_위경도_결과_좌표만2.xlsx'  # 입력 파일
output_excel_file_path = '주소_영문번역_결과.xlsx'  # 출력 파일

# 엑셀 파일에서 데이터 불러오기
df = pd.read_excel(input_excel_file_path)

# 번역된 결과를 저장할 리스트
translated_addresses = [None] * len(df)  # 미리 리스트 초기화

def translate_address(address):
    if pd.notnull(address):  # 주소가 비어있지 않은 경우
        try:
            # Google Cloud Translation API 호출 URL
            url = f'https://translation.googleapis.com/language/translate/v2?key={API_KEY}'
            params = {
                'q': address,
                'source': 'ko',
                'target': 'en',
                'format': 'text'
            }
            response = requests.post(url, data=params)
            response_data = response.json()

            if 'data' in response_data and 'translations' in response_data['data']:
                return response_data['data']['translations'][0]['translatedText']
            else:
                print(f"Error translating '{address}': {response_data}")
                return None  # 번역 실패 시 None 반환
        except Exception as e:
            print(f"Error translating '{address}': {e}")
            return None  # 번역 실패 시 None 반환
    return None  # 비어있는 경우 None 반환

# 멀티스레딩을 이용한 주소 번역
with ThreadPoolExecutor(max_workers=10) as executor:  # 최대 10개의 스레드
    results = list(executor.map(translate_address, df['주소']))

# 결과를 DataFrame에 추가
df['영문주소'] = results

# 결과를 새로운 엑셀 파일로 저장
df.to_excel(output_excel_file_path, index=False)
print(f"번역이 완료되었습니다. '{output_excel_file_path}'에 저장되었습니다.")
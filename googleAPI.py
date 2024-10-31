import pandas as pd
import requests
from concurrent.futures import ThreadPoolExecutor
from dotenv import load_dotenv
import os
from tqdm import tqdm  # tqdm 모듈 추가

# .env 파일에서 환경 변수 로드
load_dotenv()

# Google Cloud API 키 설정
API_KEY = os.getenv('GOOGLE_API_KEY')  # 구글 API 키

# 엑셀 파일 경로 입력
input_excel_file_path = '/Users/leeseokwoon/Documents/GitHub/address2xy/출입국 업무 민원대행 등록기관(2024.08.31.부) (3).xlsx'  # 입력 파일
output_excel_file_path = '출입국 업무 민원대행 등록기관(2024.08.31.부)_영문.xlsx'  # 출력 파일

# 엑셀 파일에서 데이터 불러오기
df = pd.read_excel(input_excel_file_path)

def translate_text(texts):
    # 여러 텍스트를 한 번에 번역
    if all(pd.notnull(text) for text in texts):  # 모든 텍스트가 비어있지 않은 경우
        try:
            # Google Cloud Translation API 호출 URL
            url = f'https://translation.googleapis.com/language/translate/v2?key={API_KEY}'
            params = {
                'q': texts,
                'source': 'ko',
                'target': 'en',
                'format': 'text'
            }
            response = requests.post(url, json=params)
            response_data = response.json()

            if 'data' in response_data and 'translations' in response_data['data']:
                # 번역 결과를 리스트로 반환
                return [translation['translatedText'] for translation in response_data['data']['translations']]
            else:
                print(f"Error translating: {response_data}")
                return [None] * len(texts)  # 번역 실패 시 None 반환
        except Exception as e:
            print(f"Error translating: {e}")
            return [None] * len(texts)  # 번역 실패 시 None 반환
    return [None] * len(texts)  # 비어있는 경우 None 반환

# 멀티스레딩을 이용한 번역
with ThreadPoolExecutor(max_workers=10) as executor:  # 최대 10개의 스레드
    # 각 행의 텍스트를 모아서 한 번의 API 호출로 번역
    # results = list(tqdm(executor.map(translate_text, zip(df['주소'], df['발급기명'], df['상세 위치'])),
    #                     total=len(df), desc="번역 진행 중"))  # tqdm으로 진행률 표시
    results = list(tqdm(executor.map(translate_text, zip(df['대행기관명'], df['대행기관 주소'])),
                        total=len(df), desc="번역 진행 중"))  # tqdm으로 진행률 표시

# 결과를 DataFrame에 추가
# 결과를 DataFrame으로 변환하고 기존 DataFrame의 인덱스에 맞춰 할당
# df[['영문주소', '발급기명영문', '상세위치영문']] = pd.DataFrame(results)
df[['대행기관 영문명', '대행기관 영문주소']] = pd.DataFrame(results)

# 결과를 새로운 엑셀 파일로 저장
df.to_excel(output_excel_file_path, index=False)
print(f"번역이 완료되었습니다. '{output_excel_file_path}'에 저장되었습니다.")

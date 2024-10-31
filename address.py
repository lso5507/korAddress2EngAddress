import pandas as pd
import requests
import time
import random
from urllib.parse import quote
from dotenv import load_dotenv
import os
# 로컬 파일 경로 입력 받기
file_name = input("엑셀 파일 경로를 입력하세요: ")

# 엑셀 파일에서 데이터 불러오기
df = pd.read_excel(file_name)

# 위도, 경도 열 추가
df['Latitude'] = None
df['Longitude'] = None

# .env 파일에서 환경 변수 로드
load_dotenv()

# Google Cloud Translation API 정보
api_key = os.getenv('VWORLD_API_KEY')  # 구글 API 키

# 랜덤 User-Agent 리스트
user_agents = [
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/89.0.4389.82 Safari/537.36",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/14.0.1 Safari/605.1.15",
    "Mozilla/5.0 (Linux; Android 10; Pixel 3 XL Build/QQ1A.200205.002) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/89.0.4389.90 Mobile Safari/537.36",
    "Mozilla/5.0 (iPhone; CPU iPhone OS 14_4 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/14.0 Mobile/15E148 Safari/604.1",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/90.0.4430.93 Safari/537.36"
]

# 주소 데이터를 위경도 변환
for i, row in df.iterrows():
    address = row['주소']
    encoded_address = quote(address)
    
    success = False
    attempts = 0
    
    while not success and attempts < 3:
        try:
            url = f"https://api.vworld.kr/req/address?service=address&request=getcoord&version=2.0&crs=epsg:4326&address={encoded_address}&refine=true&simple=false&format=json&type=road&key={api_key}"
            
            print(f"Request URL for {address}: {url}")
            
            headers = {
                "User-Agent": random.choice(user_agents),
                "Accept-Language": "en-US,en;q=0.5",
                "Accept-Encoding": "gzip, deflate",
                "Connection": "keep-alive"
            }
            
            response = requests.get(url, headers=headers, timeout=10)
            
            if response.status_code != 200:
                print(f"HTTP Error for {address}: {response.status_code}")
                print(f"Response content: {response.text}")
                break
            
            response_data = response.json()
            
            if response_data['response']['status'] == 'OK':
                df.at[i, 'Latitude'] = response_data['response']['result']['point']['y']
                df.at[i, 'Longitude'] = response_data['response']['result']['point']['x']
                success = True
            else:
                print(f"Error for {address}: {response_data['response']['msg']}")
                break
            
            break;
        except requests.exceptions.Timeout:
            print(f"Timeout error for {address} on attempt {attempts + 1}. Retrying...")
            attempts += 1
            time.sleep(2)
        except requests.exceptions.RequestException as req_err:
            print(f"Request error for {address} on attempt {attempts + 1}: {req_err}")
            attempts += 1
            time.sleep(2)
        except ValueError as json_err:
            print(f"JSON decode error for {address}: {json_err}")
            break
        except Exception as e:
            print(f"Error for {address}: {e}")
            break

# 결과를 새로운 엑셀 파일로 저장
output_file = '주소_위경도_결과.xlsx'
df.to_excel(output_file, index=False)
print(f"결과가 {output_file} 파일에 저장되었습니다.")

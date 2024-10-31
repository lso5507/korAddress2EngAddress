from googletrans import Translator
import pandas as pd
from concurrent.futures import ThreadPoolExecutor, as_completed
from tqdm import tqdm
import time

# 엑셀 파일 경로 입력 받기
file_name = input("엑셀 파일 경로를 입력하세요: ")
df = pd.read_excel(file_name)
df['영문주소'] = None

# 번역기 초기화
translator = Translator()

# Batch 단위 설정
BATCH_SIZE = 20  # 한 번에 번역할 최대 문장 수
RETRY_COUNT = 3  # 재시도 횟수

# 주소 전처리 함수
def preprocess_address(address):
    # 괄호 및 대괄호 제거
    return address.replace("(", "").replace(")", "").replace("[", "").replace("]", "").strip()

# 번역 함수 (Batch 처리)
def translate_batch(address_batch):
    translations = []
    for address in address_batch:
        if not address:  # 주소가 비어있거나 None인 경우
            translations.append(None)
            continue

        # 주소 전처리
        cleaned_address = preprocess_address(address)

        success = False
        for _ in range(RETRY_COUNT):
            try:
                result = translator.translate(cleaned_address, src='ko', dest='en')
                translations.append(result.text if result else None)  # 정상적으로 텍스트 반환
                success = True
                break  # 번역 성공 시 루프 종료
            except Exception as e:
                print(f"번역 오류 ({cleaned_address}): {e}")
                translations.append(None)
                time.sleep(1)  # 오류 발생 후 잠시 대기

        if not success:
            translations.append(None)  # 모든 재시도 실패 시 None 반환

    return translations

# 전체 주소 리스트를 BATCH_SIZE로 나누기
address_list = df['주소'].dropna().tolist()  # None 값 제거
batches = [address_list[i:i + BATCH_SIZE] for i in range(0, len(address_list), BATCH_SIZE)]

# 멀티스레딩으로 번역 실행 및 진행률 표시
results = []
with ThreadPoolExecutor(max_workers=10) as executor:  # max_workers 늘리기
    futures = {executor.submit(translate_batch, batch): batch for batch in batches}
    for future in tqdm(as_completed(futures), total=len(batches), desc="번역 진행 중"):
        batch_result = future.result()
        results.extend(batch_result)

# 번역 결과를 DataFrame에 저장
df['영문주소'] = results

# 최종 결과 저장
df.to_excel("주소_영문번역_결과2.xlsx", index=False)
print("번역이 완료되었습니다. '주소_영문번역_결과.xlsx'에 저장되었습니다.")

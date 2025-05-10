from datetime import datetime
import pandas as pd
import os


def convert_smart_store_to_sunny(input_path, output_path, execute_date):
    df = pd.read_excel(input_path, sheet_name="발주발송관리")
    converted_df = convert_and_rearrange_data(df, execute_date)
    return append_to_output_file(output_path, converted_df)


def convert_and_rearrange_data(df, execute_date):
    new_df = pd.DataFrame()

    month_str = execute_date.strftime('%Y%m')

    new_df['월'] = month_str
    new_df['날짜'] = execute_date
    new_df['거래유형'] = '출고'
    new_df['거래처명'] = '한솔페이퍼'
    new_df['주문자'] = df['수취인명']
    new_df['스토어상품번호'] = df['상품번호']
    new_df['스토어상품'] = df['상품명']
    new_df['수량'] = df['수량']
    new_df['연락처'] = df['수취인연락처1']
    new_df['연락처2'] = df['수취인연락처2']
    new_df['주소'] = df['통합배송지']
    new_df['배송메모'] = df['배송베세지']
    
    return new_df


def append_to_output_file(output_path, new_df, sheet_name='장부(24.01~)'):
    # 기존 파일의 모든 시트 읽기
    sheets = pd.read_excel(output_path, sheet_name=None)
    
    # 장부 시트 가져오기
    existing_df = sheets[sheet_name]
    
    # 기존 데이터 + 새 데이터 이어붙이기
    combined_df = pd.concat([existing_df, new_df], ignore_index=True)
    
    # 업데이트된 주문서 시트로 교체
    sheets[sheet_name] = combined_df

    new_file = generate_new_filename(os.path.dirname(output_path))

    with pd.ExcelWriter(new_file, engine='openpyxl') as writer:
        for name, df in sheets.items():
            df.to_excel(writer, sheet_name=name, index=False)
    
    print(f"새 파일 저장 완료: {new_file}")
    return new_file


def generate_new_filename(output_dir):
    timestamp_str = datetime.now().strftime('%Y%m%d_%H%M%S')
    filename = f"재고정리양식_신장부{timestamp_str}.xlsx"
    return os.path.join(output_dir, filename)

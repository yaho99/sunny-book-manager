from datetime import datetime

import pandas as pd
import os


def convert_excel(input_path, output_path, start_date):
    df = pd.read_excel(input_path, sheet_name="장부(24.01~)")
    filtered_df = filter_data(df, start_date)
    converted_df = convert_data(filtered_df)
    final_df = rearrange_columns(converted_df)
    return append_to_output_file(output_path, final_df)


def filter_data(df, start_date):
    df = df.iloc[1:].reset_index(drop=True)

    # B 컬럼 (yyyy.MM.dd 형태 → datetime 변환)
    print(df.columns)

    # start_date (QDate) → datetime 변환
    start_datetime = datetime(
        start_date.year(),
        start_date.month(),
        start_date.day()
    )

    # start_date 이상 데이터 필터링
    filtered_df = df[df['날짜'] >= start_datetime]

    # D 컬럼 == '한솔페이퍼' 필터
    filtered_df = filtered_df[filtered_df['거래처명'] == '한솔페이퍼']

    # B~O 컬럼만 남기기
    col_order = list(df.loc[:, '날짜':'수량1'].columns)  # B~O 컬럼 이름 추출
    filtered_df = filtered_df[col_order]

    # 필요시 index reset
    filtered_df = filtered_df.reset_index(drop=True)

    return filtered_df


def convert_data(filtered_df):
    # 매핑 코드 정의
    mapping_dict = {
        '한솔A4-75g': '한솔A4-75',
        '한솔A4-75g 250매 1박스': '한솔A4-75 (250매10묶음)',
        '한솔A4-80g': '한솔A4-80',
        '한솔A4-85g': '한솔A4-85',
        '한솔미색A4-80g': '한솔미색A4-80',
        '한솔A3-75g': '한솔A3-75',
        '한솔A3-80g': '한솔A3-80',
        '한솔B4-75g': '한솔B4-75',
        '한솔B5-75g': '한솔B5-75',
        '한솔A5-75g': '한솔A5-75',
        '한솔A5-80g': '한솔A5-80',
        '밀크A4-75g': '밀크A4-75',
        '밀크A4-80g': '밀크A4-80',
        '밀크A4-120g': '밀크A4-120',
        '밀크A4-85g': '밀크A4-85',
        '밀크A4-90g': '밀크A4-90',
        '밀크미색A4-80g': '밀크미색A4-80',
        '밀크미색A3-80g': '밀크미색A3-80',
        '밀크미색B4-80g': '밀크B4-80',
        '밀크미색B5-80g': '밀크B5-80',
        '밀크A3-75g': '밀크A3-75',
        '밀크A3-80g': '밀크A3-80',
        '밀크A3-85g': '밀크A3-85',
        '밀크B4-75g': '밀크B4-75',
        '밀크B4-80g': '밀크B4-80',
        '밀크B5-75g': '밀크B5-75',
        '밀크B5-80g': '밀크B5-80',
        '삼성A4-70g': '삼성A4-70',
        '삼성A4-75g': '삼성A4-75',
        '삼성A4-80g': '삼성A4-80',
        '삼성A3-75g': '삼성A3-75',
        '삼성A3-80g': '삼성A3-80',
        '삼성B4-75g': '삼성B4-75',
        '더블AA4-80g': '더블A4-80',
        '더블AA4-90g': '더블A4-90',
        '더블AA3-80g': '더블A3-80',
        '더블AB4-80g': '더블B4-80',
        '더블AB5-80g': '더블B5-80',
        'HPA4-70g': 'HPA4-70',
        'HPA4-75g': 'HPA4-75',
        'HPA4-80g': 'HPA4-80',
        'HPA3-75g': 'HPA3-75g',
        'HPA3-80g': 'HPA3-80',
        '무림A4-75g': '무림A4-75',
        '무림A4-80g': '무림A4-80',
        '무림A3-75g': '무림A3-75',
        '무림B5-75g': '무림B5-75',
        '하이브라이트A4-80g': '하이브라이트A4-80',
        '하이브라이트A4-75g': '하이브라이트A4-75',
        '하이브라이트A3-80g': '하이브라이트A3-80',
        '하이브라이트A3-80g1250매': '하이브라이트A3-80 1250매',
        '미스터카피A4-75g': '미스터카피A4-75',
        '미스터카피A4-80g': '미스터카피A4-80',
        '기타': '기타'
    }

    # N 컬럼 매핑
    filtered_df['런 상품명'] = filtered_df['런 상품명'].map(mapping_dict).fillna(filtered_df['런 상품명'])

    # O 컬럼 → 부호 반전
    filtered_df['수량1'] = filtered_df['수량1'] * -1

    return filtered_df


def rearrange_columns(df):
    new_df = pd.DataFrame()

    new_df['주문일시'] = df['날짜']
    new_df['출고/보관'] = df['거래유형']
    new_df['빈칸'] = ''
    new_df['판매처'] = '정명선'
    new_df['0'] = ''
    new_df['수취인'] = df['주문자']
    new_df['휴대폰번호'] = df['연락처']
    new_df['전화번호'] = df['연락처2']
    new_df['주소'] = df['주소']
    new_df['1'] = ''
    new_df['상품명'] = df['런 상품명']
    new_df['수량'] = df['수량1']
    new_df['2'] = ''
    new_df['3'] = ''
    new_df['배송메시지'] = df['배송메모']

    return new_df

def append_to_output_file(output_file, final_df, sheet_name='주문서'):
    # 기존 파일의 모든 시트 읽기
    sheets = pd.read_excel(output_file, sheet_name=None)

    # 주문서 시트 가져오기 (무조건 존재한다고 가정)
    existing_df = sheets[sheet_name]

    # 기존 데이터 + 새 데이터 이어붙이기
    combined_df = pd.concat([existing_df, final_df], ignore_index=True)

    # 업데이트된 주문서 시트로 교체
    sheets[sheet_name] = combined_df

    new_file = generate_new_filename(os.path.dirname(output_file))

    with pd.ExcelWriter(new_file, engine='openpyxl') as writer:
        for name, df in sheets.items():
            df.to_excel(writer, sheet_name=name, index=False)

    print(f"새 파일 저장 완료: {new_file}")
    return new_file


def generate_new_filename(output_dir):
    timestamp_str = datetime.now().strftime('%Y%m%d_%H%M%S')
    filename = f"한솔페이퍼(정명선){timestamp_str}.xlsx"
    return os.path.join(output_dir, filename)
